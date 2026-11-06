//
//  XpressNetLogFile.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the LogFileClass of the VB6 XpressNet DLL (version 1.1: the directory can be changed).
//  The default directory is ~/Library/Logs/TMC-Logfiles (1.2: ~/Library/Logs/TMC). The Logfile is rotated if its
//  size exceeds 100000 bytes.
//
//  THE FOLLOWING METHODS ARE DEFINED:
//  OpenLogfile                              Check the size of the Logfile, rotate if needed, and write an empty line
//  Append:          some Text               Append "some text" to the Logfile
//  AppendWithHost:  some text               Append "some text", preceded by the letter of the XpressNet object
//  Connection:      some text               2.0: Append "= some text" (the connection of the XpressNet object)
//  SendHex:length:  bytes                   Append "=>" followed by the hex values of the bytes
//  ReceiveHex:length:  bytes                Append "<=" followed by the hex values of the bytes
//  ReceiveHexRange:from:to:  bytes          Append "<=" followed by the hex values between from and to
//
//  Details:         some Text               Same as AppendWithHost, but checks Parent.LogDetails first
//  DetailsSent:     some Text               2.0: as Details, preceded by "=>" (description of a sent message)
//  DetailsReceived: some Text               2.0: as Details, preceded by "<=" (description of a received message)
//  Errors:          some Text               Same as AppendWithHost, but checks Parent.LogErrors first
//  Warnings:        some Text               Same as AppendWithHost, but checks Parent.LogWarnings first
//
//  2.0: every line of an XpressNet object starts with its letter (A, B, ...; see the attribute instanceId of the
//  XpressNet object); when a connection is made, a line "A = 192.168.0.200, port 5550, local port 54321" tells which
//  connection the letter belongs to. 1.2: every line ended with the IP address or port (myHost, which is still set)
//

#import <Foundation/Foundation.h>
@class XpressNetClass;

@interface LogFileClass : NSObject

- (void)OpenLogfile;
- (void)Append:(NSString *)data;
- (void)AppendWithHost:(NSString *)data;
- (void)Details:(NSString *)data;
- (void)DetailsSent:(NSString *)data;                  // 2.0
- (void)DetailsReceived:(NSString *)data;              // 2.0
- (void)Connection:(NSString *)data;                   // 2.0
- (void)Errors:(NSString *)data;
- (void)Warnings:(NSString *)data;
- (void)SendHex:(const uint8_t *)byteArray length:(int)length;
- (void)ReceiveHex:(const uint8_t *)byteArray length:(int)length;
- (void)ReceiveHexRange:(const uint8_t *)byteArray from:(int)msgStart to:(int)msgEnd;

@property (copy) NSString *LogDirectory;           // Setting a new directory creates it, if needed
@property (copy) NSString *logFileName;
@property (copy) NSString *logFileNameOld1;
@property (copy) NSString *logFileNameOld2;
@property (copy) NSString *logFileNameOld3;
@property (copy) NSString *myHost;                 // Allows the Logfile Hex dumps to include the IP address / USB port

// For internal use: a pointer to the XpressNet object, for the Log... attributes
@property (assign) XpressNetClass *Parent;

@end
