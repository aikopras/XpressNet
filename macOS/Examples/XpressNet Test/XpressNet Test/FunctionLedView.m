//
//  FunctionLedView.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  Shows the LEDs for the loco functions F0..F68. See FunctionLedView.h.
//  Each row has a label, followed by two groups of 4 LEDs. As in the VB6 version:
//  - row 0 (F0-F4): F0 is the last LED of the first group, F1..F4 form the second group
//  - row 1 (F5-F12): F5..F8 and F9..F12
//  - rows 2..8: 8 functions each (F13-F20 .. F61-F68)
//  Only the first numberOfRows rows are shown (default: all 9 rows).
//  The view is flipped: the origin is at the top left corner.
//

#import "FunctionLedView.h"

@implementation FunctionLedView {
  int _ledState[69];                      // 0 = unknown (white), 1 = function off (green), 2 = function on (red)
}

@synthesize numberOfRows = _numberOfRows;

// ***************************************************************************************************************************************************
// ******************************************************* Local C type definitions and declarations *************************************************
// ***************************************************************************************************************************************************
#define NUMBER_OF_LEDS 69
#define NUMBER_OF_ROWS 9                  // Maximum number of rows
#define ROW_HEIGHT 18
#define LABEL_WIDTH 38                    // Width of the labels "F0-F4" .. "F61-F68"
#define LABEL_GAP 4
#define LED_SIZE 9
#define LED_PITCH 11                      // Distance between two LEDs within a group
#define GROUP_GAP 6                       // Extra distance between both groups of 4 LEDs
#define LED_UNKNOWN 0
#define LED_FUNCTION_OFF 1
#define LED_FUNCTION_ON 2

static NSString * const rowLabels[NUMBER_OF_ROWS] =
  {@"F0-F4", @"F5-F12", @"F13-F20", @"F21-F28", @"F29-F36", @"F37-F44", @"F45-F52", @"F53-F60", @"F61-F68"};


// ****************************************************************************************************************
// ************************************************ PUBLIC METHODS ************************************************
// ****************************************************************************************************************
// Both initialisers set the default number of rows, so the view can be created from a xib as well as in code
- (id)initWithFrame:(NSRect)frameRect {
  self = [super initWithFrame:frameRect];
  if (self) _numberOfRows = NUMBER_OF_ROWS;
  return self;
}

- (id)initWithCoder:(NSCoder *)coder {
  self = [super initWithCoder:coder];
  if (self) _numberOfRows = NUMBER_OF_ROWS;
  return self;
}


- (int)numberOfRows {
  return _numberOfRows;
}

- (void)setNumberOfRows:(int)value {
  if (value < 1) value = 1;
  if (value > NUMBER_OF_ROWS) value = NUMBER_OF_ROWS;
  _numberOfRows = value;
  [self setNeedsDisplay:YES];
}


- (void)setGroup:(int)group data:(int)data {
  int bitsInGroup;                        // Groups 1..3 contain 4 FB bits, the others contain 8 FB bits
  int offset;                             // Offset in the LED array for a certain group
  if ((group < 1) || (group > 10)) return;
  if (group < 4) {
    bitsInGroup = 4;
    offset = group * bitsInGroup - 4;
  }
  else {
    bitsInGroup = 8;
    offset = group * bitsInGroup - 20;
  }
  for (int i = 1; i <= bitsInGroup; i++) {
    _ledState[i + offset] = (data & (1 << (i - 1))) ? LED_FUNCTION_ON : LED_FUNCTION_OFF;
  }
  // F0 in Group 1 is a special case
  if (group == 1) _ledState[0] = (data & 0x10) ? LED_FUNCTION_ON : LED_FUNCTION_OFF;
  [self setNeedsDisplay:YES];
}


- (void)clearLeds {
  for (int i = 0; i < NUMBER_OF_LEDS; i++) _ledState[i] = LED_UNKNOWN;
  [self setNeedsDisplay:YES];
}


// ****************************************************************************************************************
// ************************************************* VIEW METHODS *************************************************
// ****************************************************************************************************************
- (BOOL)isFlipped {return YES;}

- (void)viewDidChangeEffectiveAppearance {
  [super viewDidChangeEffectiveAppearance];
  [self setNeedsDisplay:YES];
}


- (void)drawRect:(NSRect)dirtyRect {
  NSMutableParagraphStyle *rightAligned = [[[NSMutableParagraphStyle alloc] init] autorelease];
  [rightAligned setAlignment:NSTextAlignmentRight];
  NSDictionary *attributes = [NSDictionary dictionaryWithObjectsAndKeys:
                              [NSFont systemFontOfSize:[NSFont smallSystemFontSize]], NSFontAttributeName,
                              [NSColor labelColor], NSForegroundColorAttributeName,
                              rightAligned, NSParagraphStyleAttributeName, nil];
  CGFloat labelHeight = [@"F0" sizeWithAttributes:attributes].height;
  // Labels
  for (int row = 0; row < _numberOfRows; row++) {
    CGFloat centerY = row * ROW_HEIGHT + (ROW_HEIGHT / 2.0);
    [rowLabels[row] drawInRect:NSMakeRect(0, centerY - (labelHeight / 2.0), LABEL_WIDTH, labelHeight) withAttributes:attributes];
  }
  // LEDs
  for (int i = 0; i < NUMBER_OF_LEDS; i++) {
    int row, column;
    if (i == 0)       {row = 0; column = 3;}
    else if (i <= 4)  {row = 0; column = i + 3;}
    else if (i <= 12) {row = 1; column = i - 5;}
    else              {row = 2 + (i - 13) / 8; column = (i - 13) % 8;}
    if (row >= _numberOfRows) break;      // The LEDs are ordered by row
    CGFloat x = LABEL_WIDTH + LABEL_GAP + column * LED_PITCH + ((column >= 4) ? GROUP_GAP : 0);
    CGFloat y = row * ROW_HEIGHT + (ROW_HEIGHT - LED_SIZE) / 2.0;
    NSColor *fillColor = [NSColor whiteColor];
    if (_ledState[i] == LED_FUNCTION_OFF) fillColor = [NSColor systemGreenColor];
    if (_ledState[i] == LED_FUNCTION_ON)  fillColor = [NSColor systemRedColor];
    NSRect ledRect = NSMakeRect(x + 0.5, y + 0.5, LED_SIZE - 1, LED_SIZE - 1);
    [fillColor setFill];
    NSRectFill(ledRect);
    [[NSColor systemGrayColor] setStroke];
    [NSBezierPath strokeRect:ledRect];
  }
}

@end
