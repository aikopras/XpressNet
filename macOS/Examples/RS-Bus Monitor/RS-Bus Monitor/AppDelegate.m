//
//  AppDelegate.m
//  RS-Bus Monitor
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "RS-Bus Monitor", to show the feedback information of all 128 RS-Bus addresses:
//  - CONNECT TCP/IP opens the connection to the Lenz interface, using the IP address from the preferences.
//    If a connection already exists, it is closed first.
//  - The matrix (RSBusMatrixView) shows all RS-Bus feedback messages that are received from the command station.
//  - GET RS-BUS DATA requests the feedback information of all 128 RS-Bus addresses (both nibbles).
//  - ALLES AN and ALLES AUS switch the track power on and off. The LEDs show NotAus and NotHalt.
//  - The address conversion shows which switch addresses belong to a RS-Bus address, and vice versa.
//
//  2.0: the program uses the XpressNet library (XpressNetClass), instead of its own TCP code (TCPConnection.m and
//  TCPSocket.m). The code follows the VB6 version (RsBusMonitor.frm); differences with the previous macOS version:
//  - Get RS-Bus data in two parts of 128 requests (see pushedGetRsBusData). The library sends the requests one by one.
//  - Feedback messages with TT=00 (switch decoder without feedback) are ignored while Get RS-Bus data is active
//    (as in VB6). Before, all answers with TT=00 were ignored.
//  - The events SendError, InterfaceNichtAdressiert and InterfaceWiederAdressiert are shown in the status line.
//    The interface messages and the keep-alive are handled by the library.
//  - The Logfiles are written in ~/Library/Logs/TMC-Logfiles.
//  - The default IP address is 192.168.0.200 (the default address of the Lenz interface, as in the other examples).
//

#import "AppDelegate.h"
// #import "TCPConnection.h"
#import "XpressNetLogFile.h"
#import "PreferencesController.h"

#define MIN_WINDOW_WIDTH 940                // Minimum size of the window contents, so the left panel stays visible
#define MIN_WINDOW_HEIGHT 540

@implementation AppDelegate {
  BOOL GetSchaltInfo;                       // Get RS-Bus data is active
  int AnswerCount;                          // The answers for RS-Bus address 64 and 128
}

// @synthesize tcpConnection           = _tcpConnection;
@synthesize XpressNet               = _XpressNet;
@synthesize preferencesController   = _preferencesController;
@synthesize window                  = _window;             // UI variable
@synthesize matrixView              = _matrixView;         // UI variable
@synthesize labelIpAddress          = _labelIpAddress;     // UI variable
@synthesize labelStatus             = _labelStatus;        // UI variable
@synthesize ledConnected            = _ledConnected;       // UI variable
@synthesize ledNotAus               = _ledNotAus;          // UI variable
@synthesize ledNotHalt              = _ledNotHalt;         // UI variable
@synthesize textRsAddress           = _textRsAddress;      // UI variable
@synthesize textSwitchAddress       = _textSwitchAddress;  // UI variable


- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
  // Check if the preferences file exists
  [self checkPreferences];
  [_window setContentMinSize:NSMakeSize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)];
  [self showIpAddress:[[NSUserDefaults standardUserDefaults] stringForKey:@"defaultIpAddressForReceiving"]];
  [self showStatusLine:@""];
  [self showConnectedLed:LED_OFF];
  [self showNotAusLed:LED_OFF];
  [self showNotHaltLed:LED_OFF];
  // Allocate memory and initialise the XpressNet object. As in the VB6 version, the connection is made after CONNECT is pushed
  // _tcpConnection = [[TCPClass alloc] init];
  _XpressNet = [[XpressNetClass alloc] init];
  [_XpressNet setDelegate:self];
  // Set the directory for logfiles
  [[_XpressNet Logfile] setLogDirectory:@"~/Library/Logs/TMC-Logfiles"];
  GetSchaltInfo = NO;
}

- (void)dealloc {
  // [_tcpConnection release];
  [_XpressNet setDelegate:nil];
  [_XpressNet release];
  [_preferencesController release];
  [_window release];
  [super dealloc];
}

- (BOOL)applicationShouldTerminateAfterLastWindowClosed:(NSApplication *)theApplication {return YES;}
- (BOOL)applicationSupportsSecureRestorableState:(NSApplication *)app {return YES;}

- (NSApplicationTerminateReply)applicationShouldTerminate:(NSApplication *)sender
{
  // [_tcpConnection closeTcp];
  // Release the XpressNet object, which in turn closes the connection
  [_XpressNet setDelegate:nil];
  [self setXpressNet:nil];
  return NSTerminateNow;
}


// ************************************************************************************************************
// ********************************************* PREFERENCES **************************************************
// ************************************************************************************************************
-(IBAction)showPreferences:(id)sender {
  if (self.preferencesController == nil) {
    PreferencesController *controller = [[PreferencesController alloc] initWithWindowNibName:@"PreferencesController"];
    self.preferencesController = controller;   // retain property
    [controller release];
  }
  [self.preferencesController showWindow:self];
}


- (void)checkPreferences {
  // Each preference that does not exist yet, gets its default value
  // 2.0: the default address of the Lenz interface (as in the other examples)
  // if ([defaults objectForKey:@"defaultIpAddressForReceiving"] == nil) [defaults setObject:@"192.168.24.213" forKey:@"defaultIpAddressForReceiving"];
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  if ([defaults objectForKey:@"defaultIpAddressForReceiving"] == nil) [defaults setObject:@"192.168.0.200" forKey:@"defaultIpAddressForReceiving"];
}


// ************************************************************************************************************
// ******************************************** METHODS FOR STATUS AND LEDS ***********************************
// ************************************************************************************************************
- (void)showStatusLine:(NSString *)statustext {
  [_labelStatus setStringValue:statustext];
}

- (void)showIpAddress:(NSString *)ipAddress {
  [_labelIpAddress setStringValue:(ipAddress != nil) ? ipAddress : @""];
}

- (void)showLed:(NSBox *)led state:(ledState_t)state {
  if (state == LED_OFF)   [led setFillColor:[NSColor colorWithCalibratedWhite:0.88 alpha:1.0]];
  if (state == LED_GREEN) [led setFillColor:[NSColor systemGreenColor]];
  if (state == LED_RED)   [led setFillColor:[NSColor systemRedColor]];
}

- (void)showConnectedLed:(ledState_t)state {[self showLed:_ledConnected state:state];}
- (void)showNotAusLed:(ledState_t)state     {[self showLed:_ledNotAus state:state];}
- (void)showNotHaltLed:(ledState_t)state    {[self showLed:_ledNotHalt state:state];}

// 2.0: the name of the reason of the event SendError (the names of SendErrorReasonEnum)
- (NSString *)sendErrorName:(int)Reason {
  switch (Reason) {
    case WiederholungErfolglos: return @"WiederholungErfolglos";
    case BefehlNichtVorhanden: return @"BefehlNichtVorhanden";
    case Timeout: return @"Timeout";
    case FifoVoll: return @"FifoVoll";
    case KeineVerbindung: return @"KeineVerbindung";
    case FehlerInDenBefehlsparametern: return @"FehlerInDenBefehlsparametern";
    case ZentraleAdressiertDasInterfaceNichtMehr: return @"ZentraleAdressiertDasInterfaceNichtMehr";
    default: return [NSString stringWithFormat:@"SendError %d", Reason];
  }
}


// ************************************************************************************************************
// **************************************** USER INTERFACE METHODS ********************************************
// ************************************************************************************************************
- (IBAction)pushedConnect:(id)sender {
  // [_tcpConnection openTcp];
  // The IP address is read from the preferences each time, so a new value is used after Connect TCP/IP is pushed.
  // An existing connection is closed first by the library
  NSString *ipAddress = [[NSUserDefaults standardUserDefaults] stringForKey:@"defaultIpAddressForReceiving"];
  [self showIpAddress:ipAddress];
  [self showStatusLine:@"TCP connecting ....."];
  [self showConnectedLed:LED_OFF];
  [self showNotAusLed:LED_OFF];
  [self showNotHaltLed:LED_OFF];
  // A new connection stops Get RS-Bus data: the library empties its message queues
  GetSchaltInfo = NO;
  [_XpressNet CmdConnectTcp:((ipAddress != nil) ? ipAddress : @"") Port:@"5550"];
}


- (IBAction)pushedAllesAn:(id)sender {
  // [_tcpConnection sendAllesAn];
  [_XpressNet CmdAllesAn];
}


- (IBAction)pushedAllesAus:(id)sender {
  // [_tcpConnection sendAllesAus];
  [_XpressNet CmdAllesAus];
}


// 2.0: Get RS-Bus data in two parts of 128 requests (64 RS-Bus addresses, two nibbles each). The FIFO of version 2.0 of
// the library holds 128 messages: more requests at once would give the event SendError (FifoVoll).
// The second part is requested as soon as both requests for RS-Bus address 64 have been answered; Get RS-Bus data
// ends when both requests for RS-Bus address 128 have been answered (see checkSchaltInfo). This does not depend on
// the speed of the interface. The event SendError stops Get RS-Bus data.
- (IBAction)pushedGetRsBusData:(id)sender {
  // [_tcpConnection requestAllFeedbacks];
  if (GetSchaltInfo) return;                      // Get RS-Bus data is already active
  GetSchaltInfo = YES;
  AnswerCount = 0;
  [self showStatusLine:@"Requesting RS-Bus data ....."];
  [self requestFeedbacksFrom:1 to:64];
}


// Requests both nibbles of the RS-Bus addresses firstAddress .. lastAddress
- (void)requestFeedbacksFrom:(int)firstAddress to:(int)lastAddress {
  for (int address = firstAddress; address <= lastAddress; address++) {
    [_XpressNet CmdFeedbackRequest:address Nibble:0];     // Nibble = 0
    [_XpressNet CmdFeedbackRequest:address Nibble:1];     // Nibble = 1
  }
}


// Called for every feedback message, after the matrix has been updated. The answers are counted per RS-Bus address,
// and not per nibble: for an RS-Bus address without feedback information, the command station (LZ100) answers both
// requests with nibble 0 (42 Address 00).
// Note: an unrequested feedback message (broadcast) for RS-Bus address 64 or 128 is counted as well
- (void)checkSchaltInfo:(int)RsAddress {
  if (!GetSchaltInfo) return;
  if ((RsAddress != 64) && (RsAddress != 128)) return;
  AnswerCount = AnswerCount + 1;
  if (AnswerCount == 2) [self requestFeedbacksFrom:65 to:128];  // The first part has been answered
  if (AnswerCount == 4) {                                       // The second part has been answered
    GetSchaltInfo = NO;
    [self showStatusLine:@"RS-Bus data received"];
  }
}


// ************************************************************************************************************
// *********************************************** ADDRESS CONVERSION *****************************************
// ************************************************************************************************************
// As in the VB6 version: RS-Bus address N corresponds to the switch addresses (N-1)*4+1 .. (N-1)*4+4.
// The 128 RS-Bus addresses thus correspond to the switch addresses 1..512
- (IBAction)pushedRsToSwitch:(id)sender {
  int rsAddr = [_textRsAddress intValue];
  if ((rsAddr < 1) || (rsAddr > 128)) {[self showStatusLine:@"RS address should be between 1-128"]; return;}
  int firstSw = (rsAddr - 1) * 4 + 1;
  int lastSw = firstSw + 3;
  [_textSwitchAddress setStringValue:[NSString stringWithFormat:@"%d .. %d", firstSw, lastSw]];
}


- (IBAction)pushedSwitchToRs:(id)sender {
  int swAddr = [_textSwitchAddress intValue];
  if ((swAddr < 1) || (swAddr > 512)) {[self showStatusLine:@"Switch address should be between 1-512"]; return;}
  int rsAddr = (swAddr - 1) / 4 + 1;
  [_textRsAddress setIntValue:rsAddr];
}


// ************************************************************************************************************
// ************************************************ XPRESSNET EVENTS ******************************************
// ************************************************************************************************************
// 2.0: the events of the XpressNet library (XpressNetDelegate), instead of the methods called by TCPClass
- (void)InterfaceConnected {
  [self showConnectedLed:LED_GREEN];
  [self showStatusLine:@"Connection established"];
}

- (void)TcpClosedByRemoteHost {
  [self showConnectedLed:LED_RED];
  GetSchaltInfo = NO;                             // Get RS-Bus data stops
  [self showStatusLine:@"TCP connection closed by remote host"];
}

- (void)TcpError:(NSString *)Description {
  [self showConnectedLed:LED_RED];
  GetSchaltInfo = NO;                             // Get RS-Bus data stops
  [self showStatusLine:Description];
}

- (void)AllesAn {
  [self showNotAusLed:LED_GREEN];
  [self showNotHaltLed:LED_GREEN];
}

- (void)NotAus  {[self showNotAusLed:LED_RED];}
- (void)NotHalt {[self showNotHaltLed:LED_RED];}

// A message that the library could not send. Reason: see SendErrorReasonEnum (XpressNet.h)
// Get RS-Bus data is stopped (for example: no answer, the message queues of the library have been emptied)
- (void)SendError:(int)Reason Description:(NSString *)Description {
  GetSchaltInfo = NO;
  [self showStatusLine:[NSString stringWithFormat:@"%@: %@", [self sendErrorName:Reason], Description]];
}

// The command station does not address the interface: sending is paused
- (void)InterfaceNichtAdressiert  {[self showStatusLine:@"Interface nicht adressiert"];}
// The command station addresses the interface again
- (void)InterfaceWiederAdressiert {[self showStatusLine:@"Interface wieder adressiert"];}

// We have 2 events for RS-Bus input: Feedback (feedback decoders) and SwitchFeedback (switch decoders)
// Feedback: address = 1..128, Nibble = 0 / 1, data = 0..15
- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data {
  [_matrixView setFeedbackNibble:Nibble withData:data forAddress:address type:BLOCK_BEZETMELDING];
  // Get RS-Bus data: is this the answer on the last request of a part?
  [self checkSchaltInfo:address];
}

// SwitchFeedback: switchAddress = ((address * 4) + (Nibble * 2) + 1), with address = 0..127
- (void)SwitchFeedback:(int)switchAddress ttBits:(int)ttBits iBit:(int)iBit data:(int)data {
  int RsAddress = (switchAddress - 1) / 4 + 1;
  int Nibble = ((switchAddress - 1) / 2) % 2;
  if (ttBits == 0) {     // Feedback comes from Command Station
    if (!GetSchaltInfo) [_matrixView setFeedbackNibble:Nibble withData:data forAddress:RsAddress type:BLOCK_CENTRALE];
  }
  else {                 // Feedback comes from Decoder
    [_matrixView setFeedbackNibble:Nibble withData:data forAddress:RsAddress type:BLOCK_WISSEL];
  }
  // Get RS-Bus data: is this the answer on the last request of a part?
  [self checkSchaltInfo:RsAddress];
}


@end
