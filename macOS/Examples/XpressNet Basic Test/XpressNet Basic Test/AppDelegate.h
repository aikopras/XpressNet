//
//  AppDelegate.h
//  XpressNet Basic Test
//
//  Created by Aiko Pras on 31-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "XpressNet Basic Test", to test the XpressNet library and
//  command stations via TCP/IP or USB.
//

#import <Cocoa/Cocoa.h>
#import "XpressNet.h"
@class FunctionLedView;

@interface AppDelegate : NSObject <NSApplicationDelegate, XpressNetDelegate, NSTableViewDataSource, NSTableViewDelegate, NSComboBoxDelegate>

@property (retain) XpressNetClass *XpressNet;

// Declare the properties for the main user interface window
@property (retain) IBOutlet NSWindow *window;
// Connections
@property (assign) IBOutlet NSTextField *textIpAddress;
@property (assign) IBOutlet NSTextField *textTcpPort;
@property (assign) IBOutlet NSComboBox *comboSerialPort;
@property (assign) IBOutlet NSComboBox *comboBaudrate;
@property (assign) IBOutlet NSButton *checkHandshake;           // 2.0: hardware handshake for USB
@property (assign) IBOutlet NSBox *ledConnected;
@property (assign) IBOutlet NSBox *ledNotAus;
@property (assign) IBOutlet NSBox *ledNotHalt;
@property (assign) IBOutlet NSBox *ledProgrammierMode;
@property (assign) IBOutlet NSTextField *labelMasterVersion;
@property (assign) IBOutlet NSTextField *labelInterfaceVersion;
@property (assign) IBOutlet NSTextField *labelInterfaceCode;
@property (assign) IBOutlet NSTextField *labelXpressNetAddress;
@property (assign) IBOutlet NSTextField *labelXpressNetVersion;
@property (assign) IBOutlet NSTextField *labelTcpFree;
@property (assign) IBOutlet NSTextField *labelStatus;
// Switch and RS-Bus feedback
@property (assign) IBOutlet NSTextField *textSwitch;
@property (assign) IBOutlet NSTableView *tableSwitch;
@property (assign) IBOutlet NSTextField *textRsBusAddress;
@property (assign) IBOutlet NSTableView *tableRsBus;
// Loks
@property (assign) IBOutlet NSTextField *textLokAddress;
@property (assign) IBOutlet NSTextField *textLokFunction;
@property (assign) IBOutlet NSTextField *textLokSpeed;
@property (assign) IBOutlet NSComboBox *comboLokSteps;
@property (assign) IBOutlet FunctionLedView *functionLeds;

// Connections and information
- (IBAction)pushedConnectTcp:(id)sender;
- (IBAction)pushedConnectUsb:(id)sender;
- (IBAction)pushedMasterInfo:(id)sender;
- (IBAction)pushedInterfaceInfo:(id)sender;
- (IBAction)pushedLanInfo:(id)sender;
// Switches and feedback
- (IBAction)pushedSwitchLeft:(id)sender;
- (IBAction)pushedSwitchRight:(id)sender;
- (IBAction)pushedSwitchFeedback:(id)sender;
- (IBAction)pushedRsFeedback:(id)sender;
// Command station
- (IBAction)pushedAllesAn:(id)sender;
- (IBAction)pushedAllesAus:(id)sender;
- (IBAction)pushedAlleLoksAnhalten:(id)sender;
// Loks
- (IBAction)pushedEineLokAnhalten:(id)sender;
- (IBAction)pushedLokEnter:(id)sender;
- (IBAction)pushedLokAusStack:(id)sender;
- (IBAction)pushedLokForward:(id)sender;
- (IBAction)pushedLokReverse:(id)sender;
- (IBAction)pushedFunctionOn:(id)sender;
- (IBAction)pushedFunctionOff:(id)sender;
- (IBAction)pushedGetF0F12:(id)sender;
- (IBAction)pushedGetF13F28:(id)sender;
- (IBAction)pushedGetF29F68:(id)sender;

@end
