//
//  SerialPort.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  2026: Reusable serial port, based on the POSIX (termios) interface of macOS. It is used for USB interfaces
//  (FTDI, CH340, CP210x, CDC-ACM, ...) as well as for RS-232 interfaces via a USB-serial adapter: on macOS both
//  appear as a device /dev/cu.xxx, provided macOS has a driver for the USB chip. The class has no knowledge of
//  DCC or XpressNet; it only opens the port, sends, receives and reports.
//
//  Usage:
//  - availablePorts returns the devices that can be opened
//  - create the object, set its delegate and call openPort:settings:handshake:
//  - the delegate receives all data, and is informed if the port is lost (for example because the USB cable is removed)
//  - all delegate methods are called on the main thread
//  - call close to close the port
//  Settings have the same format as for the MSComm control of VB6: "baudrate,parity,databits,stopbits",
//  for example "57600,N,8,1". Parity may be N (none), E (even) or O (odd). Any baudrate is allowed.
//  If handshake is YES, the CTS line of the interface controls when data may be sent (hardware flow control);
//  the RTS line is not used. As in the VB6 version (MSComm: DTREnable and RTSEnable False), DTR and RTS are
//  switched off after the port is opened.
//

#import <Foundation/Foundation.h>

@class SerialPortClass;

// The object that uses the serial port should implement the following two methods
@protocol SerialPortDelegate <NSObject>
// Called for every chunk of data received. The bytes are only valid during this call
- (void)serialPort:(SerialPortClass *)port didReceiveBytes:(const uint8_t *)bytes length:(NSUInteger)length;
// Called if the port is closed because of an error (for example because the USB cable is removed)
- (void)serialPort:(SerialPortClass *)port didCloseWithReason:(NSString *)reason;
@end


@interface SerialPortClass : NSObject

// The following methods should be used
+ (NSArray *)availablePorts;                           // Full paths of all /dev/cu.* devices, sorted
- (BOOL)openPort:(NSString *)path settings:(NSString *)settings handshake:(BOOL)handshake;   // NO: see lastError
- (void)sendBytes:(const uint8_t *)bytes length:(NSUInteger)length;
- (void)close;
- (BOOL)isOpen;

@property (assign) id <SerialPortDelegate> delegate;   // Not retained
@property (readonly) NSString *lastError;              // Description of the last error, nil if none

@end
