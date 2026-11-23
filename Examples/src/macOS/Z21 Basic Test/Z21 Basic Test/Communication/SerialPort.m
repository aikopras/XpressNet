//
//  SerialPort.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  2026: Reusable serial port, based on the POSIX (termios) interface of macOS. See SerialPort.h.
//
//  Notes on the implementation:
//  - The port is opened non-blocking. A dispatch source on the main queue reports that data can be read,
//    so the main thread never waits.
//  - Data that can not be written immediately (for example because CTS is off) is kept, and written as soon as
//    the port accepts it (a second dispatch source).
//  - The baudrate is set with IOSSIOSPEED (IOKit/serial/ioss.h), which also allows non-standard baudrates.
//

#import "SerialPort.h"
#import <IOKit/serial/ioss.h>
#include <termios.h>
#include <fcntl.h>
#include <unistd.h>
#include <sys/ioctl.h>
#include <errno.h>
#include <string.h>

@implementation SerialPortClass {
  int _fd;                                // File descriptor of the open port; -1 if the port is closed
  dispatch_source_t _readSource;          // Reports that data can be read
  dispatch_source_t _writeSource;         // Reports that data can be written; only exists if _pendingData exists
  NSMutableData *_pendingData;            // Data that could not be written yet; nil if none
}

@synthesize delegate = _delegate;
@synthesize lastError = _lastError;

#define READ_BUFFER_SIZE 1024


// ******************************************************************************************
// ************************************** PUBLIC METHODS ************************************
// ******************************************************************************************
+ (NSArray *)availablePorts {
  NSArray *names = [[NSFileManager defaultManager] contentsOfDirectoryAtPath:@"/dev" error:NULL];
  NSMutableArray *ports = [NSMutableArray array];
  for (NSString *name in names) {
    if ([name hasPrefix:@"cu."]) [ports addObject:[@"/dev" stringByAppendingPathComponent:name]];
  }
  return [ports sortedArrayUsingSelector:@selector(compare:)];
}


- (id)init {
  self = [super init];
  if (self) _fd = -1;
  return self;
}


- (BOOL)openPort:(NSString *)path settings:(NSString *)settings handshake:(BOOL)handshake {
  [self close];
  [self setError:nil];
  // Analyse the settings: "baudrate,parity,databits,stopbits"
  NSArray *fields = [settings componentsSeparatedByString:@","];
  speed_t baudrate = 57600;
  NSString *parity = @"N";
  int dataBits = 8;
  int stopBits = 1;
  if ([fields count] > 0) baudrate = (speed_t)[[fields objectAtIndex:0] integerValue];
  if ([fields count] > 1) parity = [[[fields objectAtIndex:1] stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]] uppercaseString];
  if ([fields count] > 2) dataBits = [[fields objectAtIndex:2] intValue];
  if ([fields count] > 3) stopBits = [[fields objectAtIndex:3] intValue];
  if ((baudrate == 0) || (dataBits < 5) || (dataBits > 8) || (stopBits < 1) || (stopBits > 2) ||
      !([parity isEqualToString:@"N"] || [parity isEqualToString:@"E"] || [parity isEqualToString:@"O"])) {
    [self setError:[NSString stringWithFormat:@"invalid settings \"%@\"", settings]];
    return NO;
  }
  // Open the port. O_NONBLOCK: open returns immediately, and read / write never wait
  int fd = open([path fileSystemRepresentation], O_RDWR | O_NOCTTY | O_NONBLOCK);
  if (fd < 0) {[self setError:[NSString stringWithUTF8String:strerror(errno)]]; return NO;}
  // No other program may open the port while we use it
  if (ioctl(fd, TIOCEXCL) < 0) {[self failOpen:fd]; return NO;}
  // Raw mode: no processing of special characters
  struct termios options;
  if (tcgetattr(fd, &options) < 0) {[self failOpen:fd]; return NO;}
  cfmakeraw(&options);
  options.c_cflag |= (CLOCAL | CREAD);                       // Ignore modem status lines, enable the receiver
  options.c_cflag &= ~(CSIZE | PARENB | PARODD | CSTOPB | CRTSCTS);
  switch (dataBits) {
    case 5:  options.c_cflag |= CS5; break;
    case 6:  options.c_cflag |= CS6; break;
    case 7:  options.c_cflag |= CS7; break;
    default: options.c_cflag |= CS8; break;
  }
  if ([parity isEqualToString:@"E"]) options.c_cflag |= PARENB;
  if ([parity isEqualToString:@"O"]) options.c_cflag |= (PARENB | PARODD);
  if (stopBits == 2) options.c_cflag |= CSTOPB;
  if (handshake) options.c_cflag |= CCTS_OFLOW;              // CTS controls our output; RTS is not used
  options.c_cc[VMIN] = 0;
  options.c_cc[VTIME] = 0;
  cfsetspeed(&options, B9600);                               // Placeholder; the real baudrate is set below
  if (tcsetattr(fd, TCSANOW, &options) < 0) {[self failOpen:fd]; return NO;}
  if (ioctl(fd, IOSSIOSPEED, &baudrate) < 0) {[self failOpen:fd]; return NO;}
  // As in the VB6 version: DTR and RTS off
  int modemLines = TIOCM_DTR | TIOCM_RTS;
  ioctl(fd, TIOCMBIC, &modemLines);
  tcflush(fd, TCIOFLUSH);
  _fd = fd;
  // Keep this object alive until the read source has been cancelled (see close)
  [self retain];
  __block SerialPortClass *blockSelf = self;
  _readSource = dispatch_source_create(DISPATCH_SOURCE_TYPE_READ, (uintptr_t)fd, 0, dispatch_get_main_queue());
  dispatch_source_set_event_handler(_readSource, ^{[blockSelf readAvailableData];});
  dispatch_source_set_cancel_handler(_readSource, ^{
    // The file descriptor is closed only after the source has been cancelled
    close(fd);
    [blockSelf release];
  });
  dispatch_resume(_readSource);
  return YES;
}


- (void)sendBytes:(const uint8_t *)bytes length:(NSUInteger)length {
  if ((_fd < 0) || (length == 0)) return;
  // If data is still waiting, the new data must wait as well (to keep the order)
  if (_pendingData != nil) {[_pendingData appendBytes:bytes length:length]; return;}
  ssize_t written = write(_fd, bytes, length);
  if (written < 0) {
    if ((errno != EAGAIN) && (errno != EINTR)) {[self closeWithReason:[NSString stringWithUTF8String:strerror(errno)]]; return;}
    written = 0;
  }
  if ((NSUInteger)written < length) [self keepData:(bytes + written) length:(length - (NSUInteger)written)];
}


- (void)close {
  if (_fd < 0) return;
  [self stopWriteSource];
  _fd = -1;
  dispatch_source_cancel(_readSource);      // The cancel handler closes the file descriptor
  dispatch_release(_readSource);
  _readSource = nil;
}


- (BOOL)isOpen {
  return (_fd >= 0);
}


// ******************************************************************************************
// ************************************** INTERNAL METHODS **********************************
// ******************************************************************************************
- (void)setError:(NSString *)error {
  [_lastError release];
  _lastError = [error copy];
}


- (void)failOpen:(int)fd {
  [self setError:[NSString stringWithUTF8String:strerror(errno)]];
  close(fd);
}


// Called by the read source. One read per event: if more data is available, the source reports this again
- (void)readAvailableData {
  if (_fd < 0) return;
  uint8_t buffer[READ_BUFFER_SIZE];
  ssize_t received = read(_fd, buffer, READ_BUFFER_SIZE);
  if (received > 0) {[_delegate serialPort:self didReceiveBytes:buffer length:(NSUInteger)received]; return;}
  if ((received < 0) && ((errno == EAGAIN) || (errno == EINTR))) return;   // No data after all
  // 0 bytes (end of file) while the source reported data, or an error: the device has gone (for example the
  // USB cable was removed)
  [self closeWithReason:(received == 0) ? @"device removed" : [NSString stringWithUTF8String:strerror(errno)]];
}


// Keep data that could not be written yet, and write it as soon as the port accepts it
- (void)keepData:(const uint8_t *)bytes length:(NSUInteger)length {
  _pendingData = [[NSMutableData alloc] initWithBytes:bytes length:length];
  __block SerialPortClass *blockSelf = self;
  _writeSource = dispatch_source_create(DISPATCH_SOURCE_TYPE_WRITE, (uintptr_t)_fd, 0, dispatch_get_main_queue());
  dispatch_source_set_event_handler(_writeSource, ^{[blockSelf writePendingData];});
  dispatch_resume(_writeSource);
}


- (void)writePendingData {
  if ((_fd < 0) || (_pendingData == nil)) return;
  ssize_t written = write(_fd, [_pendingData bytes], [_pendingData length]);
  if (written < 0) {
    if ((errno != EAGAIN) && (errno != EINTR)) [self closeWithReason:[NSString stringWithUTF8String:strerror(errno)]];
    return;
  }
  [_pendingData replaceBytesInRange:NSMakeRange(0, (NSUInteger)written) withBytes:NULL length:0];
  if ([_pendingData length] == 0) [self stopWriteSource];
}


- (void)stopWriteSource {
  if (_writeSource != nil) {
    dispatch_source_cancel(_writeSource);
    dispatch_release(_writeSource);
    _writeSource = nil;
  }
  [_pendingData release];
  _pendingData = nil;
}


// The port is closed because of an error: close it and inform the delegate
- (void)closeWithReason:(NSString *)reason {
  [self setError:reason];
  [self close];
  [_delegate serialPort:self didCloseWithReason:reason];
}


- (void)dealloc {
  [_pendingData release];
  [_lastError release];
  [super dealloc];
}

@end
