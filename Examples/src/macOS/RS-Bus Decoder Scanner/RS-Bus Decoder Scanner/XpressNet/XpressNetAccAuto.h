//
//  XpressNetAccAuto.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the AccAutoClass of the VB6 XpressNet DLL: implementation of the AccessoryAuto method.
//  For internal library use; the library user calls CmdAccessoryAuto:position: of the XpressNet object.
//

#import <Foundation/Foundation.h>
@class XpressNetClass;

@interface AccAutoClass : NSObject

- (void)handle:(int)address position:(int)position;
- (int)size;                                           // The number of accessory commands in the FIFO
- (int)emptyFifo;                                      // 2.0: switches off the activated switch; returns the number removed
- (void)resume;                                        // 2.0: called at the end of a pause of the SendClass (01 07)
- (void)stop;                                          // Stops the timer; must be called before the object is released

@property (assign) XpressNetClass *Parent;

@end
