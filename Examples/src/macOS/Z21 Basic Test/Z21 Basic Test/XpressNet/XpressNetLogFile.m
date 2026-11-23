//
//  XpressNetLogFile.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the LogFileClass of the VB6 XpressNet DLL. See XpressNetLogFile.h.
//  Differences with the VB6 version:
//  - the default directory is ~/Library/Logs/TMC-Logfiles (VB6: %PUBLIC%\TMC-Logfiles). 1.2: ~/Library/Logs/TMC
//  - hex values are always written with two digits (VB6 Hex(): "1" instead of "01")
//

#import "XpressNetLogFile.h"
#import "XpressNet.h"

@implementation LogFileClass {
  NSDateFormatter *_timeFormatter;      // Time with ms precision, as in the VB6 version
}

@synthesize LogDirectory      = _LogDirectory;
@synthesize logFileName       = _logFileName;
@synthesize logFileNameOld1   = _logFileNameOld1;
@synthesize logFileNameOld2   = _logFileNameOld2;
@synthesize logFileNameOld3   = _logFileNameOld3;
@synthesize myHost            = _myHost;
@synthesize Parent            = _Parent;

// The maximum Logfile size as well as the names of the Logfile and backups may be modified
#define MAX_LOGFILE_SIZE 100000                    // Size is in bytes


- (id)init {
  self = [super init];
  if (self) {
    _timeFormatter = [[NSDateFormatter alloc] init];
    [_timeFormatter setLocale:[NSLocale localeWithLocaleIdentifier:@"en_US_POSIX"]];
    [_timeFormatter setDateFormat:@"yyyy/MM/dd HH:mm:ss.SSS"];
    _myHost = [@"" copy];
    // In case the application doesn't provide a directory, use the default
    // [self setLogDirectory:[NSHomeDirectory() stringByAppendingPathComponent:@"Library/Logs/TMC"]];
    [self setLogDirectory:[NSHomeDirectory() stringByAppendingPathComponent:@"Library/Logs/TMC-Logfiles"]];   // 2.0
  }
  return self;
}


- (void)dealloc {
  [_timeFormatter release];
  [_LogDirectory release];
  [_logFileName release];
  [_logFileNameOld1 release];
  [_logFileNameOld2 release];
  [_logFileNameOld3 release];
  [_myHost release];
  [super dealloc];
}


- (NSString *)LogDirectory {
  return _LogDirectory;
}


// Setting the directory creates it, if needed. If that fails, the previous directory is kept
- (void)setLogDirectory:(NSString *)value {
  NSString *directory = [[value stringByExpandingTildeInPath] stringByStandardizingPath];
  if (![[NSFileManager defaultManager] createDirectoryAtPath:directory withIntermediateDirectories:YES attributes:nil error:NULL]) return;
  [_LogDirectory release];
  _LogDirectory = [directory copy];
  [self setLogFileName:[directory stringByAppendingPathComponent:@"TMC-LogFile.txt"]];
  [self setLogFileNameOld1:[directory stringByAppendingPathComponent:@"TMC-LogFile-1.txt"]];
  [self setLogFileNameOld2:[directory stringByAppendingPathComponent:@"TMC-LogFile-2.txt"]];
  [self setLogFileNameOld3:[directory stringByAppendingPathComponent:@"TMC-LogFile-3.txt"]];
}


// ****************************************************************************************************************
// ************************************************ PUBLIC METHODS ************************************************
// ****************************************************************************************************************
- (void)OpenLogfile {
  // The goal of this method is to check the size of the existing Logfile.
  // If that size exceeds MAX_LOGFILE_SIZE, the Logfiles will be rotated and renamed.
  NSFileManager *manager = [NSFileManager defaultManager];
  NSDictionary *attributes = [manager attributesOfItemAtPath:_logFileName error:NULL];
  if ((attributes != nil) && ([attributes fileSize] > MAX_LOGFILE_SIZE)) {
    // Delete the oldest file and rename the others
    [manager removeItemAtPath:_logFileNameOld3 error:NULL];
    [manager moveItemAtPath:_logFileNameOld2 toPath:_logFileNameOld3 error:NULL];
    [manager moveItemAtPath:_logFileNameOld1 toPath:_logFileNameOld2 error:NULL];
    [manager moveItemAtPath:_logFileName toPath:_logFileNameOld1 error:NULL];
  }
  // Write an empty line, to indicate a restart
  [self writeLine:@""];
}


- (void)Append:(NSString *)data {
  [self writeLine:[NSString stringWithFormat:@"%@: %@", [self timeText], data]];
}


- (void)AppendWithHost:(NSString *)data {
  // 2.0: preceded by the letter of the XpressNet object (1.2: followed by the host)
  // [self writeLine:[NSString stringWithFormat:@"%@: %@ (%@)", [self timeText], data, _myHost]];
  [self writeLine:[NSString stringWithFormat:@"%@: %@ %@", [self timeText], [_Parent instanceId], data]];
}


// 2.0: the connection of this XpressNet object
- (void)Connection:(NSString *)data {
  [self AppendWithHost:[@"= " stringByAppendingString:data]];
}


- (void)Details:(NSString *)data  {if ([_Parent LogDetails])  [self AppendWithHost:data];}
// 2.0: the description of a message gets the same direction marker as its hex values
- (void)DetailsSent:(NSString *)data     {[self Details:[@"=> " stringByAppendingString:data]];}
- (void)DetailsReceived:(NSString *)data {[self Details:[@"<= " stringByAppendingString:data]];}
- (void)Errors:(NSString *)data   {if ([_Parent LogErrors])   [self AppendWithHost:data];}
- (void)Warnings:(NSString *)data {if ([_Parent LogWarnings]) [self AppendWithHost:data];}


// 2.0: preceded by the letter of the XpressNet object (1.2: followed by the host)
- (void)SendHex:(const uint8_t *)byteArray length:(int)length {
  [self writeLine:[NSString stringWithFormat:@"%@: %@ => %@", [self timeText], [_Parent instanceId], [self hexStringFor:byteArray from:0 to:(length - 1)]]];
}


- (void)ReceiveHex:(const uint8_t *)byteArray length:(int)length {
  [self writeLine:[NSString stringWithFormat:@"%@: %@ <= %@", [self timeText], [_Parent instanceId], [self hexStringFor:byteArray from:0 to:(length - 1)]]];
}


- (void)ReceiveHexRange:(const uint8_t *)byteArray from:(int)msgStart to:(int)msgEnd {
  [self writeLine:[NSString stringWithFormat:@"%@: %@ <= %@", [self timeText], [_Parent instanceId], [self hexStringFor:byteArray from:msgStart to:msgEnd]]];
}


// ****************************************************************************************************************
// ************************************************ LOCAL METHODS *************************************************
// ****************************************************************************************************************
- (NSString *)timeText {
  return [_timeFormatter stringFromDate:[NSDate date]];
}


- (NSString *)hexStringFor:(const uint8_t *)byteArray from:(int)msgStart to:(int)msgEnd {
  NSMutableString *hexString = [NSMutableString string];
  for (int i = msgStart; i <= msgEnd; i++) [hexString appendFormat:@"%02X ", byteArray[i]];
  // 2.0: without the last space (1.2: followed by the host)
  return [hexString stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];
}


// Append a line to the Logfile. As in the VB6 version, the file is opened and closed for every line
- (void)writeLine:(NSString *)line {
  if (_logFileName == nil) return;
  NSFileManager *manager = [NSFileManager defaultManager];
  if (![manager fileExistsAtPath:_logFileName]) [manager createFileAtPath:_logFileName contents:nil attributes:nil];
  NSFileHandle *file = [NSFileHandle fileHandleForWritingAtPath:_logFileName];
  if (file == nil) return;
  [file seekToEndOfFile];
  [file writeData:[[line stringByAppendingString:@"\n"] dataUsingEncoding:NSUTF8StringEncoding]];
  [file closeFile];
}

@end
