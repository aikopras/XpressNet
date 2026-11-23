//
//  XpressNetReceiveZ21.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the ReceiveClassZ21 of the VB6 XpressNet DLL. For internal library use.
//

#import <Foundation/Foundation.h>
@class XpressNetClass;

@interface ReceiveClassZ21 : NSObject

// Called with every UDP datagram received from the Z21
- (void)addToIndata:(const uint8_t *)inBytes length:(int)bytesReceived;

@property (assign) XpressNetClass *Parent;

@end
