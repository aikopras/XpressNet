//
//  RSBusMatrixView.h
//  RS-Bus Monitor
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the PanelManager class of the VB6 program "RsBusMonitor".
//  The view shows the 128 RS-Bus addresses in 4 blocks of 32 rows. Each row has a small square for the type of
//  feedback message, followed by 8 LEDs (two nibbles of 4 bits) for the feedback bits.
//

#import <Cocoa/Cocoa.h>

// The state of a feedback bit (LED), as in the VB6 version
typedef enum
{ BIT_NOT_CONNECTED = 0,      // grey: no information (yet)
  BIT_OFF = 1,                // green: the bit is 0
  BIT_ON = 2,                 // red: the bit is 1
} bitState_t;

// The type of the last feedback message received for an address (colour of the square), as in the VB6 version
typedef enum
{ BLOCK_NOT_CONNECTED = 0,    // grey: no information (yet)
  BLOCK_CENTRALE = 1,         // red: TT=00, switch decoder without feedback (information from the command station)
  BLOCK_WISSEL = 2,           // green: TT=01 (or 11), switch decoder with feedback
  BLOCK_BEZETMELDING = 3,     // blue: TT=10, feedback decoder
} blockState_t;

@interface RSBusMatrixView : NSView

// Shows the 4 LEDs of one nibble, plus the type of the feedback message
// address: 1..128 / nibble: 0 (bits 0..3) or 1 (bits 4..7) / data: 0..15
- (void)setFeedbackNibble:(int)nibble withData:(int)data forAddress:(int)address type:(blockState_t)type;

@end
