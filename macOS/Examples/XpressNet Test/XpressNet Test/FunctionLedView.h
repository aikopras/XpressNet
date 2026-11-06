//
//  FunctionLedView.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Shows the LEDs for the loco functions F0..F68, in 9 rows (F0-F4, F5-F12, F13-F20, ..., F61-F68),
//  as the 69 LEDs (Shape LEDF0F68) of the VB6 XpressNet Test program.
//  White: unknown, green: function off, red: function on.
//  The number of rows can be reduced: the Z21 Basic Test shows only 4 rows (F0..F28).
//

#import <Cocoa/Cocoa.h>

@interface FunctionLedView : NSView

// The number of rows that is shown: 1..9 (default: 9, F0..F68). 4 rows show F0..F28.
@property (assign) int numberOfRows;

// Show the function bits of a function group (VB6: ShowFbGroup).
// Groups 1..3 contain 4 function bits (group 1 also F0 in bit 4), the groups 4..10 contain 8 function bits
- (void)setGroup:(int)group data:(int)data;
// All LEDs white (VB6: ClearLedsForFunctions)
- (void)clearLeds;

@end
