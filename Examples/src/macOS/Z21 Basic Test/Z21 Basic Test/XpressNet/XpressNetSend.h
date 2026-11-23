//
//  XpressNetSend.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the SendClass of the VB6 XpressNet DLL: the message and nextMessage methods,
//  the FIFO queue, the flow control and the timers. For internal library use.
//  2.0: retransmission and pause (see XpressNetSend.m)
//

#import <Foundation/Foundation.h>
#import "XpressNet.h"
@class XpressNetClass;

@interface SendClass : NSObject

// message is called by the XpressNet object to send (prepare the sending of) an XpressNet message.
// numberOfBytes is followed by the individual bytes of the XpressNet message, separated by comma's.
// These bytes should NOT include the two Framing bytes (FF and FE), nor the XOR byte at the end.
// Example: [send message:2, 0x21, 0x81] sends "Alles An"
- (void)message:(int)numberOfBytes, ...;
- (void)nextMessage;                                   // Called by the receive object after a response is received
- (void)stoptcpResponseTimer;
- (void)emptyFifo;
- (int)size;                                           // The number of messages in the FIFO
- (void)stop;                                          // Stops all timers; must be called before the object is released

// 2.0: a Z21 LAN message that is not an XpressNet message (LAN_SET_BROADCASTFLAGS, LAN_LOGOFF); see XpressNetSend.m
- (void)transmitZ21:(const uint8_t *)z21Message length:(int)length;
// 2.0: called by the receive object with the answers of the interface (see XpressNetSend.m)
- (void)answerError:(NSString *)answer;
- (void)answerRejected:(NSString *)answer reason:(SendErrorReasonEnum)reason;
- (void)pause:(BOOL)answer;
- (void)resume;

@property (assign) BOOL waitForAck;                    // Can we send the message now, or should it be stored in the FIFO?
@property (assign) BOOL waitForlanKeepAliveResponse;   // Flag that helps to distinguish where certain responses belong to
@property (readonly) BOOL interfaceNotAddressed;       // 2.0: is sending paused (01 05 / 01 08)? Also read by AccAutoClass
@property (assign) XpressNetClass *Parent;

@end
