//
//  AppDelegate.h
//  Z21 Basic Test
//
//  Created by Aiko Pras on 31-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "Z21 Basic Test", to test the XpressNet library and
//  Z21 compatible command stations via UDP (Z21 LAN protocol).
//

#import <Cocoa/Cocoa.h>
#import "XpressNet.h"
@class FunctionLedView;

@interface AppDelegate : NSObject <NSApplicationDelegate, XpressNetDelegate, NSTableViewDataSource, NSTableViewDelegate>

@property (retain) XpressNetClass *XpressNet;

// Declare the properties for the main user interface window
@property (retain) IBOutlet NSWindow *window;
// Z21 Network Settings, Status and Version
@property (assign) IBOutlet NSTextField *textIpAddress;
@property (assign) IBOutlet NSTextField *textUdpPort;
@property (assign) IBOutlet NSBox *ledNotAus;
@property (assign) IBOutlet NSBox *ledNotHalt;
@property (assign) IBOutlet NSBox *ledProgrammierMode;
@property (assign) IBOutlet NSTextField *labelMasterVersion;
@property (assign) IBOutlet NSTextField *labelZ21Firmware;
@property (assign) IBOutlet NSTextField *labelMessageCounter;
@property (assign) IBOutlet NSTextField *labelStatus;
// Accessory and Extended Accessory
@property (assign) IBOutlet NSTextField *textSwitch;
@property (assign) IBOutlet NSButton *checkActivate;
@property (assign) IBOutlet NSButton *checkAutoDeactivate;
@property (assign) IBOutlet NSTableView *tableSwitch;
@property (assign) IBOutlet NSTextField *textExtAccAddress;
@property (assign) IBOutlet NSTextField *textExtAccValue;
// Loks
@property (assign) IBOutlet NSTextField *textLokAddress;
@property (assign) IBOutlet NSTextField *textLokFunction;
@property (assign) IBOutlet NSTextField *textLokSpeed;
@property (assign) IBOutlet NSComboBox *comboLokSteps;
@property (assign) IBOutlet FunctionLedView *functionLeds;

// Z21 Network Settings
- (IBAction)pushedUdpSet:(id)sender;
// Command station
- (IBAction)pushedAllesAn:(id)sender;
- (IBAction)pushedAllesAus:(id)sender;
- (IBAction)pushedAlleLoksAnhalten:(id)sender;
// Accessories
- (IBAction)pushedSwitchLeft:(id)sender;
- (IBAction)pushedSwitchRight:(id)sender;
- (IBAction)pushedSwitchFeedback:(id)sender;
- (IBAction)pushedExtAccSet:(id)sender;
// Loks
- (IBAction)pushedEineLokAnhalten:(id)sender;
- (IBAction)pushedLokEnter:(id)sender;
- (IBAction)pushedLokAusStack:(id)sender;
- (IBAction)pushedLokForward:(id)sender;
- (IBAction)pushedLokReverse:(id)sender;
- (IBAction)pushedFunctionOn:(id)sender;
- (IBAction)pushedFunctionOff:(id)sender;
- (IBAction)pushedGetLokInfo:(id)sender;

@end
