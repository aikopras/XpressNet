//
//  XpressNetReceive.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the ReceiveClass of the VB6 XpressNet DLL. For internal library use.
//

#import <Foundation/Foundation.h>
@class XpressNetClass;

@interface ReceiveClass : NSObject

// Called with the bytes received from the TCP, USB or serial connection
- (void)addToIndata:(const uint8_t *)inBytes length:(int)bytesReceived;
// Called to analyse an individual XpressNet message. The first byte should be the XpressNet header;
// framing bytes and the XOR byte should already be removed. length is the number of bytes of the message.
- (void)DecodeXpressNetMessage:(int)length message:(const uint8_t *)message broadcast:(BOOL)broadcast;
- (void)emptyIndata;                                   // Removes all bytes that are still in the input buffer

@property (assign) XpressNetClass *Parent;

@end
