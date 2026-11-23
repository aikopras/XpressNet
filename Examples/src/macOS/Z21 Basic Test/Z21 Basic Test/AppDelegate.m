//
//  AppDelegate.m
//  Z21 Basic Test
//
//  Created by Aiko Pras on 31-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "Z21 Basic Test" (form FormZ21).
//  The program uses the XpressNet library (XpressNetClass), which has the same interface as the VB6 XpressNet DLL.
//  We only use the XpressNet part of Z21 (LAN_X messages).
//  Differences with the VB6 version:
//  - The last used IP address and UDP port are remembered for the next start.
//  - The switch feedback messages are shown in a table, with the latest message at top.
//  - New: a status line, which shows the UDP settings and errors of the UDP connection.
//  - New: "Number of Z21Messages" shows the number of transmitted and received XpressNet messages
//    (as in XpressNet Test). In VB6 this field was never updated.
//  - The Steps combobox contains the values 14, 27, 28 and 128 (VB6: empty list).
//  - 2.0: the event SendError in the status line. (The events InterfaceNichtAdressiert and InterfaceWiederAdressiert
//    are not used by the Z21.)
//

#import "AppDelegate.h"
#import "FunctionLedView.h"

#define MAX_TABLE_ROWS 1000                 // Maximum number of feedback messages shown in a table
#define WINDOW_WIDTH 760                    // The window can only be resized vertically
#define MIN_WINDOW_HEIGHT 608
#define NUMBER_OF_LED_ROWS 4                // Z21 LAN_X_LOCO_INFO contains F0..F28

@implementation AppDelegate {
  NSMutableArray *_switchMessages;          // Rows of the switch feedback table, latest message first
  NSTimer *_messageCounterTimer;            // Shows the number of XpressNet messages
  // Function Groups: for each Lok Address, the values must be fetched again
  int FB[11];                               // FB: Function Befehl (on / off)
}

@synthesize XpressNet               = _XpressNet;
@synthesize window                  = _window;
@synthesize textIpAddress           = _textIpAddress;
@synthesize textUdpPort             = _textUdpPort;
@synthesize ledNotAus               = _ledNotAus;
@synthesize ledNotHalt              = _ledNotHalt;
@synthesize ledProgrammierMode      = _ledProgrammierMode;
@synthesize labelMasterVersion      = _labelMasterVersion;
@synthesize labelZ21Firmware        = _labelZ21Firmware;
@synthesize labelMessageCounter     = _labelMessageCounter;
@synthesize labelStatus             = _labelStatus;
@synthesize textSwitch              = _textSwitch;
@synthesize checkActivate           = _checkActivate;
@synthesize checkAutoDeactivate     = _checkAutoDeactivate;
@synthesize tableSwitch             = _tableSwitch;
@synthesize textExtAccAddress       = _textExtAccAddress;
@synthesize textExtAccValue         = _textExtAccValue;
@synthesize textLokAddress          = _textLokAddress;
@synthesize textLokFunction         = _textLokFunction;
@synthesize textLokSpeed            = _textLokSpeed;
@synthesize comboLokSteps           = _comboLokSteps;
@synthesize functionLeds            = _functionLeds;


- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
  _XpressNet = [[XpressNetClass alloc] init];      // We only use the XpressNet part of Z21
  [_XpressNet setDelegate:self];
  _switchMessages = [[NSMutableArray alloc] init];
  // The window can only be resized vertically: the feedback table becomes longer
  [_window setContentMinSize:NSMakeSize(WINDOW_WIDTH, MIN_WINDOW_HEIGHT)];
  [_window setContentMaxSize:NSMakeSize(WINDOW_WIDTH, 10000)];
  // The values used last time
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  if ([defaults stringForKey:@"lastIpAddress"] != nil) [_textIpAddress setStringValue:[defaults stringForKey:@"lastIpAddress"]];
  if ([defaults stringForKey:@"lastUdpPort"] != nil)   [_textUdpPort setStringValue:[defaults stringForKey:@"lastUdpPort"]];
  // The LEDs start light grey
  NSColor *grey = [NSColor colorWithCalibratedWhite:0.88 alpha:1.0];
  [self showLed:_ledNotAus color:grey];
  [self showLed:_ledNotHalt color:grey];
  [self showLed:_ledProgrammierMode color:grey];
  // Only the functions F0..F28 are shown
  [_functionLeds setNumberOfRows:NUMBER_OF_LED_ROWS];
  [_functionLeds clearLeds];
  [self showStatusLine:@""];
  [_labelMessageCounter setStringValue:@"0"];
  _messageCounterTimer = [NSTimer scheduledTimerWithTimeInterval:1.0 target:self selector:@selector(messageCounterTimerFired:) userInfo:nil repeats:YES];
}

- (void)dealloc {
  [_messageCounterTimer invalidate];
  [_XpressNet setDelegate:nil];
  [_XpressNet release];
  [_switchMessages release];
  [_window release];
  [super dealloc];
}

- (BOOL)applicationShouldTerminateAfterLastWindowClosed:(NSApplication *)theApplication {return YES;}
- (BOOL)applicationSupportsSecureRestorableState:(NSApplication *)app {return YES;}

- (NSApplicationTerminateReply)applicationShouldTerminate:(NSApplication *)sender
{
  // Release the XpressNet object, which in turn closes the connection
  [_XpressNet setDelegate:nil];
  [self setXpressNet:nil];
  return NSTerminateNow;
}


// ************************************************************************************************************
// ****************************************** SUPPORT METHODS *************************************************
// ************************************************************************************************************
- (void)showStatusLine:(NSString *)statustext {
  [_labelStatus setStringValue:statustext];
}

- (void)showLed:(NSBox *)led color:(NSColor *)color {
  [led setFillColor:color];
}

// As VB6 IsNumeric: YES if the field contains an integer value
- (BOOL)isNumeric:(NSTextField *)field {
  NSString *text = [[field stringValue] stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];
  NSScanner *scanner = [NSScanner scannerWithString:text];
  int value;
  return ([scanner scanInt:&value] && [scanner isAtEnd]);
}

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

// Transforms numbers from 0..15 (4 bits) in a bitstring
- (NSString *)printBits:(int)data {
  return [NSString stringWithFormat:@"%d%d%d%d", (data >> 3) & 1, (data >> 2) & 1, (data >> 1) & 1, data & 1];
}

- (void)messageCounterTimerFired:(NSTimer *)timer {
  long number = [_XpressNet NumberOfExpressNetMessagesReceived] + [_XpressNet NumberOfExpressNetMessagesTransmitted];
  [_labelMessageCounter setStringValue:[NSString stringWithFormat:@"%ld", number]];
}


// ************************************************************************************************************
// ********************************************* Z21 NETWORK SETTINGS *****************************************
// ************************************************************************************************************
- (IBAction)pushedUdpSet:(id)sender {
  [[NSUserDefaults standardUserDefaults] setObject:[_textIpAddress stringValue] forKey:@"lastIpAddress"];
  [[NSUserDefaults standardUserDefaults] setObject:[_textUdpPort stringValue] forKey:@"lastUdpPort"];
  [self showStatusLine:[NSString stringWithFormat:@"UDP: IP address %@, port %@", [_textIpAddress stringValue], [_textUdpPort stringValue]]];
  [_XpressNet CmdInitUdp:[_textIpAddress stringValue] Port:[_textUdpPort stringValue]];
}


// ************************************************************************************************************
// ************************************************* COMMAND STATION ******************************************
// ************************************************************************************************************
- (IBAction)pushedAllesAn:(id)sender         {[_XpressNet CmdAllesAn];}
- (IBAction)pushedAllesAus:(id)sender        {[_XpressNet CmdAllesAus];}
- (IBAction)pushedAlleLoksAnhalten:(id)sender {[_XpressNet CmdAlleLoksAnhalten];}


// ************************************************************************************************************
// ************************************************** ACCESSORIES *********************************************
// ************************************************************************************************************
- (IBAction)pushedSwitchRight:(id)sender {[self switchAccessory:1];}     // VB6: cmdRecht
- (IBAction)pushedSwitchLeft:(id)sender  {[self switchAccessory:0];}     // VB6: cmdAf

- (void)switchAccessory:(int)position {
  if (![self isNumeric:_textSwitch]) return;
  if ([_checkAutoDeactivate state] == NSControlStateValueOn) {
    [_XpressNet CmdAccessoryAuto:[_textSwitch intValue] position:position];
  }
  else {
    BOOL activate = ([_checkActivate state] == NSControlStateValueOn);
    [_XpressNet CmdAccessory:[_textSwitch intValue] position:position Activate:(activate ? 1 : 0)];
    [_checkActivate setState:(activate ? NSControlStateValueOff : NSControlStateValueOn)];
  }
}

- (IBAction)pushedSwitchFeedback:(id)sender {
  if ([self isNumeric:_textSwitch]) [_XpressNet CmdSwitchInfoRequest:[_textSwitch intValue]];
}

- (IBAction)pushedExtAccSet:(id)sender {
  if ([self isNumeric:_textExtAccAddress]) [_XpressNet CmdZ21SetExtAccessory:[_textExtAccAddress intValue] value:[_textExtAccValue intValue]];
}


// ************************************************************************************************************
// ****************************************************** LOKS ************************************************
// ************************************************************************************************************
- (IBAction)pushedEineLokAnhalten:(id)sender {
  if ([self isNumeric:_textLokAddress]) [_XpressNet CmdEineLokAnhalten:[_textLokAddress intValue]];
}

- (IBAction)pushedLokEnter:(id)sender {
  // A new lok address is entered. Get the lok info
  if (![self isNumeric:_textLokAddress]) return;
  [_functionLeds clearLeds];
  [_XpressNet CmdLokInfoAnfordern:[_textLokAddress intValue]];
}

- (IBAction)pushedLokAusStack:(id)sender {
  if ([self isNumeric:_textLokAddress]) [_XpressNet CmdLokAusStack:[_textLokAddress intValue]];
}

- (IBAction)pushedLokForward:(id)sender {
  if (![self isNumeric:_textLokAddress] || ![self isNumeric:_textLokSpeed]) return;
  [_XpressNet CmdFahrbefehl:[_textLokAddress intValue] Speed:[_textLokSpeed intValue] Steps:[_comboLokSteps intValue] Forward:YES];
}

- (IBAction)pushedLokReverse:(id)sender {
  if (![self isNumeric:_textLokAddress] || ![self isNumeric:_textLokSpeed]) return;
  [_XpressNet CmdFahrbefehl:[_textLokAddress intValue] Speed:[_textLokSpeed intValue] Steps:[_comboLokSteps intValue] Forward:NO];
}

- (IBAction)pushedFunctionOn:(id)sender  {[self handleFunctionBefehl:YES];}
- (IBAction)pushedFunctionOff:(id)sender {[self handleFunctionBefehl:NO];}

- (IBAction)pushedGetLokInfo:(id)sender {[_XpressNet CmdLokInfoAnfordern:[_textLokAddress intValue]];}

- (void)handleFunctionBefehl:(BOOL)turnOn {
  // As in VB6, all functions F0..F68 can be switched, although only F0..F28 are shown
  if (![self isNumeric:_textLokAddress] || ![self isNumeric:_textLokFunction]) return;
  int FbNumber = [_textLokFunction intValue];
  if ((FbNumber < 0) || (FbNumber > 68)) return;
  int group;
  int bitposition;
  // F0 needs to be considered as special case. The other functions are coded in a consistent way
  if (FbNumber == 0)       {group = 1; bitposition = 4;}
  else if (FbNumber <= 4)  {group = 1; bitposition = FbNumber - 1;}
  else if (FbNumber <= 8)  {group = 2; bitposition = FbNumber - 5;}
  else if (FbNumber <= 12) {group = 3; bitposition = FbNumber - 9;}
  else {group = 4 + (FbNumber - 13) / 8; bitposition = (FbNumber - 13) % 8;}   // F13-F20: group 4 ... F61-F68: group 10
  if (turnOn) FB[group] |= (1 << bitposition);
  else FB[group] &= ~(1 << bitposition);
  [_XpressNet CmdLokFunktionBefehl:[_textLokAddress intValue] group:group data:FB[group]];
}


// ************************************************************************************************************
// **************************************************** TABLES ************************************************
// ************************************************************************************************************
- (NSInteger)numberOfRowsInTableView:(NSTableView *)tableView {
  return (NSInteger)[_switchMessages count];
}

- (id)tableView:(NSTableView *)tableView objectValueForTableColumn:(NSTableColumn *)tableColumn row:(NSInteger)row {
  if ((row < 0) || (row >= (NSInteger)[_switchMessages count])) return nil;
  return [[_switchMessages objectAtIndex:(NSUInteger)row] objectForKey:[tableColumn identifier]];
}


// ************************************************************************************************************
// ************************************************ XPRESSNET EVENTS ******************************************
// ************************************************************************************************************
// The UDP connection reports its errors via TcpError (VB6: not used in the Z21 Basic test)
- (void)TcpError:(NSString *)Description {
  [self showStatusLine:Description];
}

// 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum (XpressNet.h)
- (void)SendError:(int)Reason Description:(NSString *)Description {
  [self showStatusLine:[NSString stringWithFormat:@"%@: %@", [self sendErrorName:Reason], Description]];
}

- (void)SwitchFeedback:(int)address ttBits:(int)ttBits iBit:(int)iBit data:(int)data {
  // TT: C = Feedback comes from Command Station, D = Feedback comes from Decoder
  NSDictionary *message = [NSDictionary dictionaryWithObjectsAndKeys:
                           [NSString stringWithFormat:@"%d/%d", address, address + 1], @"address",
                           [self printBits:data], @"data",
                           (ttBits == 0) ? @"C" : @"D", @"tt",
                           [NSString stringWithFormat:@"%d", iBit], @"ibit", nil];
  // The latest message at top. The number of messages is limited
  [_switchMessages insertObject:message atIndex:0];
  if ([_switchMessages count] > MAX_TABLE_ROWS) [_switchMessages removeLastObject];
  [_tableSwitch reloadData];
}

- (void)AllesAn {
  [self showLed:_ledNotAus color:[NSColor systemGreenColor]];
  [self showLed:_ledNotHalt color:[NSColor systemGreenColor]];
  [self showLed:_ledProgrammierMode color:[NSColor systemGreenColor]];
}

- (void)NotAus          {[self showLed:_ledNotAus color:[NSColor systemRedColor]];}
- (void)NotHalt         {[self showLed:_ledNotHalt color:[NSColor systemRedColor]];}
- (void)ProgrammierMode {[self showLed:_ledProgrammierMode color:[NSColor systemRedColor]];}

- (void)ZentraleVersion:(NSString *)Version {
  [_labelMasterVersion setStringValue:Version];
}

- (void)z21FirmwareVersion:(NSString *)Version {
  [_labelZ21Firmware setStringValue:Version];
}

- (void)Z21Lokinfo:(int)address Besetzt:(BOOL)Besetzt Speedsteps:(int)Speedsteps Speed:(int)Speed Forward:(BOOL)Forward F0_F4:(int)F0_F4 F5_F8:(int)F5_F8 F9_F12:(int)F9_F12 F13_F20:(int)F13_F20 F21_F28:(int)F21_F28 {
  // Set the combobox and speed, based on the received input
  [_comboLokSteps setIntValue:Speedsteps];
  [_textLokSpeed setIntValue:Speed];
  // Store the values of F0..F28, to allow changing a single Function later
  FB[1] = F0_F4;                            // Store Gruppe 1
  FB[2] = F5_F8;                            // Store Gruppe 2
  FB[3] = F9_F12;                           // Store Gruppe 3
  FB[4] = F13_F20;                          // Store Gruppe 4
  FB[5] = F21_F28;                          // Store Gruppe 5
  // Show the LEDs for F0 .. F28
  [_functionLeds setGroup:1 data:F0_F4];
  [_functionLeds setGroup:2 data:F5_F8];
  [_functionLeds setGroup:3 data:F9_F12];
  [_functionLeds setGroup:4 data:F13_F20];
  [_functionLeds setGroup:5 data:F21_F28];
}

@end
