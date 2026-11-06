//
//  AppDelegate.h
//  RS-Bus Monitor
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "RS-Bus Monitor", to show the feedback information of all 128 RS-Bus addresses.
//

#import <Cocoa/Cocoa.h>
#import "XpressNet.h"
#import "RSBusMatrixView.h"
@class PreferencesController;

// The colours of the LEDs, as in the VB6 version
typedef enum
{ LED_OFF = 0,                // light grey: no information
  LED_GREEN = 1,              // Connected: connection established / NotAus and NotHalt: not active
  LED_RED = 2,                // Connected: connection closed / NotAus and NotHalt: active
} ledState_t;

// 2.0: the program uses the XpressNet library (XpressNetClass), instead of its own TCP code (TCPConnection, TCPSocket)
// @interface AppDelegate : NSObject <NSApplicationDelegate>
@interface AppDelegate : NSObject <NSApplicationDelegate, XpressNetDelegate>

// @property (retain) TCPClass *tcpConnection;
@property (retain) XpressNetClass *XpressNet;
@property (retain) PreferencesController *preferencesController;

// Declare the properties for the main user interface window
@property (retain) IBOutlet NSWindow *window;
@property (assign) IBOutlet RSBusMatrixView *matrixView;
@property (assign) IBOutlet NSTextField *labelIpAddress;
@property (assign) IBOutlet NSTextField *labelStatus;
@property (assign) IBOutlet NSBox *ledConnected;
@property (assign) IBOutlet NSBox *ledNotAus;
@property (assign) IBOutlet NSBox *ledNotHalt;
@property (assign) IBOutlet NSTextField *textRsAddress;
@property (assign) IBOutlet NSTextField *textSwitchAddress;

- (IBAction)pushedConnect:(id)sender;
- (IBAction)pushedAllesAn:(id)sender;
- (IBAction)pushedAllesAus:(id)sender;
- (IBAction)pushedGetRsBusData:(id)sender;
- (IBAction)pushedRsToSwitch:(id)sender;                    // ">" button of the address conversion
- (IBAction)pushedSwitchToRs:(id)sender;                    // "<" button of the address conversion
- (IBAction)showPreferences:(id)sender;

@end
