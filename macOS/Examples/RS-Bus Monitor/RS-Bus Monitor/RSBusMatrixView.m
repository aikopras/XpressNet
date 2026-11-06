//
//  RSBusMatrixView.m
//  RS-Bus Monitor
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the PanelManager class of the VB6 program "RsBusMonitor". The layout algorithm, the margins and
//  the colours of the LEDs and squares are taken from the VB6 version. Differences with the VB6 version:
//  - the layout is computed each time the view is drawn, so it follows the size of the window automatically
//  - the background and the address labels use the system colours, so the view also works in dark mode
//  The view is flipped: the origin is at the top left corner, as in VB6.
//

#import "RSBusMatrixView.h"

@implementation RSBusMatrixView

// ***************************************************************************************************************************************************
// ******************************************************* Local C type definitions and declarations *************************************************
// ***************************************************************************************************************************************************
#define NUM_BLOCKS 4
#define ROWS_PER_BLOCK 32
#define BITS_PER_ADDR 8
#define NUM_ADDRESSES (NUM_BLOCKS * ROWS_PER_BLOCK)
// Margins in points
#define MARGIN_L 12
#define MARGIN_T 12
#define MARGIN_R 12
#define MARGIN_B 12
// Extra offsets
#define LABEL_OFFSET 8
#define MIDGROUP_GAP 6
#define LABEL_FONT_SIZE 12

// State. Memory of static variables is initialised with zeros, thus with BIT_NOT_CONNECTED and BLOCK_NOT_CONNECTED
static bitState_t bitState[NUM_ADDRESSES][BITS_PER_ADDR];
static blockState_t blockState[NUM_ADDRESSES];

// Layout variables (the names are taken from the VB6 version)
static CGFloat mRowH;           // height of a row
static CGFloat mD;              // diameter of a LED
static CGFloat mR;              // radius of a LED
static CGFloat mSpacing;        // space between two LEDs
static CGFloat mBlockGap;       // space between two blocks
static CGFloat mLabelW;         // width of the label column
static CGFloat mBlockW;         // width of a block
static CGFloat mBlockPad;       // space between label column and square
static CGFloat mSq;             // size of the square


// ****************************************************************************************************************************************************
// ******************************************************************* PUBLIC METHODS *****************************************************************
// ****************************************************************************************************************************************************
- (void)setFeedbackNibble:(int)nibble withData:(int)data forAddress:(int)address type:(blockState_t)type {
  if ((address < 1) || (address > NUM_ADDRESSES)) return;
  if ((nibble < 0) || (nibble > 1)) return;
  if ((data < 0) || (data > 15)) return;
  int addr = address - 1;                         // Internally addresses are from 0..127
  // Nibble 0 holds bits 0..3, nibble 1 holds bits 4..7
  BOOL refreshNeeded = NO;
  int offset = (nibble == 0) ? 0 : 4;
  for (int i = 0; i < 4; i++) {
    bitState_t value = (data & (1 << i)) ? BIT_ON : BIT_OFF;
    if (bitState[addr][i + offset] != value) {
      bitState[addr][i + offset] = value;
      refreshNeeded = YES;
    }
  }
  // Colour of the square in front of the LEDs
  if (blockState[addr] != type) {
    blockState[addr] = type;
    refreshNeeded = YES;
  }
  if (refreshNeeded) [self setNeedsDisplay:YES];
}


// ****************************************************************************************************************************************************
// ******************************************************************* VIEW METHODS *******************************************************************
// ****************************************************************************************************************************************************
- (BOOL)isFlipped {return YES;}
- (BOOL)isOpaque {return YES;}

// Draw everything again after the window has been resized, or after a change between light and dark mode
- (void)setFrameSize:(NSSize)newSize {
  [super setFrameSize:newSize];
  [self setNeedsDisplay:YES];
}

- (void)viewDidChangeEffectiveAppearance {
  [super viewDidChangeEffectiveAppearance];
  [self setNeedsDisplay:YES];
}


- (void)drawRect:(NSRect)dirtyRect {
  // Clear the background
  [[NSColor windowBackgroundColor] setFill];
  NSRectFill([self bounds]);
  // Compute the layout of the panel
  NSDictionary *labelAttributes = [NSDictionary dictionaryWithObjectsAndKeys:
                                   [NSFont boldSystemFontOfSize:LABEL_FONT_SIZE], NSFontAttributeName,
                                   [NSColor systemBlueColor], NSForegroundColorAttributeName, nil];
  if (![self computeLayoutForSize:[self bounds].size labelAttributes:labelAttributes]) return;
  CGFloat labelH = [@"0" sizeWithAttributes:labelAttributes].height;
  // Draw all addresses, labels, squares and LEDs
  for (int addr = 0; addr < NUM_ADDRESSES; addr++) {
    int b = addr / ROWS_PER_BLOCK;
    int r = addr % ROWS_PER_BLOCK;
    CGFloat baseX = MARGIN_L + b * (mBlockW + mBlockGap);
    CGFloat centerY = MARGIN_T + r * mRowH + (mRowH / 2);
    // Labels, for every fourth address
    if ((r % 4) == 0) {
      NSString *labelText = [NSString stringWithFormat:@"%d", addr + 1];
      [labelText drawAtPoint:NSMakePoint(baseX + LABEL_OFFSET, centerY - (labelH / 2)) withAttributes:labelAttributes];
    }
    // Square for the feedback type
    [self drawBlockAtX:(baseX + mLabelW + mBlockPad + (mSq / 2)) y:centerY state:blockState[addr]];
    // LEDs
    CGFloat ledOffsetX = mBlockPad + (mSq * 3);
    for (int k = 0; k < BITS_PER_ADDR; k++) {
      CGFloat x = baseX + mLabelW + ledOffsetX + k * (mD + mSpacing) + mR;
      if (k >= 4) x = x + MIDGROUP_GAP;
      [self drawLampAtX:x y:centerY state:bitState[addr][k]];
    }
  }
}


// ****************************************************************************************************************************************************
// ****************************************************************** LAYOUT METHODS ******************************************************************
// ****************************************************************************************************************************************************
// Compute the layout of the panel, as in the VB6 version. Returns NO if the view is too small
- (BOOL)computeLayoutForSize:(NSSize)size labelAttributes:(NSDictionary *)labelAttributes {
  CGFloat contentW = size.width - (MARGIN_L + MARGIN_R);
  CGFloat contentH = size.height - (MARGIN_T + MARGIN_B);
  if ((contentW < 100) || (contentH < 100)) return NO;
  // Width of the label column
  mLabelW = [@"128" sizeWithAttributes:labelAttributes].width + LABEL_OFFSET + 6;
  // Height of a row
  mRowH = floor(contentH / ROWS_PER_BLOCK);
  if (mRowH < 8) mRowH = 8;
  // Diameter of the LEDs, and the spacing between them
  mD = mRowH - 4;
  if (mD < 6) mD = 6;
  mR = mD / 2;
  mSpacing = [self spacingForDiameter:mD];
  // Size of the squares
  mSq = mRowH / 2;
  if (mSq < 2) mSq = 2;
  if (mSq > (mR / 3)) mSq = mR / 3;
  // Width of a block: make the LEDs smaller until all blocks fit
  mBlockGap = 24;
  mBlockPad = mSpacing / 4;
  if (mBlockPad < 1) mBlockPad = 1;
  while (YES) {
    mBlockW = mLabelW + (BITS_PER_ADDR * mD) + ((BITS_PER_ADDR - 1) * mSpacing) + MIDGROUP_GAP + mBlockPad + mSq;
    CGFloat totalW = (NUM_BLOCKS * mBlockW) + ((NUM_BLOCKS - 1) * mBlockGap);
    if (totalW <= contentW) break;
    if (mD > 6) mD = mD - 1;
    mR = mD / 2;
    mSpacing = [self spacingForDiameter:mD];
    if (mBlockGap > 12) mBlockGap = mBlockGap - 1;
    if ((mD <= 6) && (mBlockGap <= 12)) break;
  }
  return YES;
}


- (CGFloat)spacingForDiameter:(CGFloat)diameter {
  CGFloat spacing = round(diameter / 3);
  if (spacing < 2) spacing = 2;
  return spacing;
}


// ****************************************************************************************************************************************************
// ****************************************************************** DRAWING METHODS *****************************************************************
// ****************************************************************************************************************************************************
+ (NSColor *)colorWithRed:(int)red green:(int)green blue:(int)blue {
  return [NSColor colorWithCalibratedRed:(red / 255.0) green:(green / 255.0) blue:(blue / 255.0) alpha:1.0];
}


// Draw the square for an address. x and y are the center of the square
- (void)drawBlockAtX:(CGFloat)x y:(CGFloat)y state:(blockState_t)state {
  NSColor *fillColor;
  switch (state) {
    case BLOCK_NOT_CONNECTED: fillColor = [RSBusMatrixView colorWithRed:220 green:220 blue:220]; break;
    case BLOCK_CENTRALE:      fillColor = [RSBusMatrixView colorWithRed:255 green:0   blue:0];   break;
    case BLOCK_WISSEL:        fillColor = [RSBusMatrixView colorWithRed:0   green:255 blue:0];   break;
    default:                  fillColor = [RSBusMatrixView colorWithRed:0   green:0   blue:255]; break;
  }
  [fillColor setFill];
  NSRectFill(NSMakeRect(x - (mSq / 2), y - (mSq / 2), mSq, mSq));
}


// Draw a LED. x and y are the center of the LED
- (void)drawLampAtX:(CGFloat)x y:(CGFloat)y state:(bitState_t)state {
  NSColor *fillColor;
  switch (state) {
    case BIT_ON:            fillColor = [RSBusMatrixView colorWithRed:220 green:30  blue:30];  break;
    case BIT_NOT_CONNECTED: fillColor = [RSBusMatrixView colorWithRed:220 green:220 blue:220]; break;
    default:                fillColor = [RSBusMatrixView colorWithRed:192 green:242 blue:154]; break;
  }
  [self drawCircleAtX:x y:y radius:mR fillColor:fillColor edgeColor:[RSBusMatrixView colorWithRed:100 green:100 blue:100]];
  // A highlight if the LED is on, and a smaller, subtle highlight if the LED is off
  if ((state == BIT_ON) && (mR >= 6))
    [self drawCircleAtX:(x - mR / 3) y:(y - mR / 3) radius:(mR / 3)
              fillColor:[RSBusMatrixView colorWithRed:255 green:200 blue:200] edgeColor:fillColor];
  else if ((state == BIT_OFF) && (mR >= 6))
    [self drawCircleAtX:(x - mR / 3) y:(y - mR / 3) radius:(mR / 4)
              fillColor:[RSBusMatrixView colorWithRed:200 green:200 blue:200] edgeColor:fillColor];
}


- (void)drawCircleAtX:(CGFloat)x y:(CGFloat)y radius:(CGFloat)radius fillColor:(NSColor *)fillColor edgeColor:(NSColor *)edgeColor {
  NSBezierPath *circle = [NSBezierPath bezierPathWithOvalInRect:NSMakeRect(x - radius, y - radius, 2 * radius, 2 * radius)];
  [fillColor setFill];
  [circle fill];
  [edgeColor setStroke];
  [circle setLineWidth:1.0];
  [circle stroke];
}


@end
