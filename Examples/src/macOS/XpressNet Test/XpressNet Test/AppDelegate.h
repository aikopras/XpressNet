//
//  AppDelegate.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "XpressNet Test" (Extended_Test), to test the XpressNet library and
//  command stations via TCP/IP, USB or a serial interface.
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
@property (assign) IBOutlet NSButton *checkCts;
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
@property (assign) IBOutlet NSTextField *labelMessageCounter;
// Switch and RS-Bus feedback
@property (assign) IBOutlet NSTextField *textSwitch;
@property (assign) IBOutlet NSButton *checkActivate;
@property (assign) IBOutlet NSButton *checkAutoDeactivate;
@property (assign) IBOutlet NSTableView *tableSwitch;
@property (assign) IBOutlet NSTextField *textRsBusAddress;
@property (assign) IBOutlet NSTableView *tableRsBus;
// Modellzeit and command station
@property (assign) IBOutlet NSTextField *labelModellzeit;
@property (assign) IBOutlet NSComboBox *comboDay;
@property (assign) IBOutlet NSTextField *textHour;
@property (assign) IBOutlet NSTextField *textMinute;
@property (assign) IBOutlet NSTextField *textFactor;
@property (assign) IBOutlet NSButton *checkRun;
@property (assign) IBOutlet NSTextField *labelBuildNr;
@property (assign) IBOutlet NSTextField *labelVersionRuckmelder;
@property (assign) IBOutlet NSTextField *labelBuildRuckmelder;
@property (assign) IBOutlet NSTextField *labelVersionBootloader;
// Loks
@property (assign) IBOutlet NSTextField *textLokAddress;
@property (assign) IBOutlet NSTextField *textLokFunction;
@property (assign) IBOutlet NSTextField *textLokSpeed;
@property (assign) IBOutlet NSComboBox *comboLokSteps;
@property (assign) IBOutlet FunctionLedView *functionLeds;
// ExtAccCmd and Programmieren
@property (assign) IBOutlet NSTextField *textExtAccAddress;
@property (assign) IBOutlet NSTextField *textExtAccValue;
@property (assign) IBOutlet NSButton *checkSM;
@property (assign) IBOutlet NSButton *radioDirect;
@property (assign) IBOutlet NSButton *radioRegister;
@property (assign) IBOutlet NSButton *radioPage;
@property (assign) IBOutlet NSTextField *textPomAddress;
@property (assign) IBOutlet NSTextField *textPomCV;
@property (assign) IBOutlet NSTextField *textPomValue;

// Connections and information
- (IBAction)pushedConnectTcp:(id)sender;
- (IBAction)pushedConnectUsb:(id)sender;
- (IBAction)pushedConnectSerial:(id)sender;
- (IBAction)pushedMasterInfo:(id)sender;
- (IBAction)pushedInterfaceInfo:(id)sender;
- (IBAction)pushedLanInfo:(id)sender;
- (IBAction)pushedErweiterteInfo:(id)sender;
- (IBAction)pushedStartMan:(id)sender;
- (IBAction)pushedStartAuto:(id)sender;
// Switches and feedback
- (IBAction)pushedSwitchLeft:(id)sender;
- (IBAction)pushedSwitchRight:(id)sender;
- (IBAction)pushedSwitchFeedback:(id)sender;
- (IBAction)pushedRsFeedback:(id)sender;
- (IBAction)pushedExtAccSet:(id)sender;
- (IBAction)pushedSendTest:(id)sender;
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
// Programming
- (IBAction)pushedPomGet:(id)sender;
- (IBAction)pushedPomSet:(id)sender;
- (IBAction)pushedPomHolen:(id)sender;
- (IBAction)selectedSmMode:(id)sender;                      // Used by the three radio buttons
// Modellzeit
- (IBAction)pushedModellzeitAnfordern:(id)sender;
- (IBAction)pushedModellzeitAnhalten:(id)sender;
- (IBAction)pushedModellzeitStarten:(id)sender;
- (IBAction)pushedModellzeitStellen:(id)sender;

@end
