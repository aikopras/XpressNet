//
//  AppDelegate.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 program "XpressNet Test" (Extended_Test, form FormXpressNet).
//  The program uses the XpressNet library (XpressNetClass), which has the same interface as the VB6 XpressNet DLL.
//  Differences with the VB6 version:
//  - The user interface is reorganised. USB and serial connections share the port and the baudrate.
//  - The port is the name of a serial device (/dev/cu.xxx), instead of a COM port number. The list of devices
//    is updated each time the list is opened.
//  - The last used IP address, TCP port, serial port and baudrate are remembered for the next start.
//  - The switch and RS-Bus feedback messages are shown in tables, with the latest message at top.
//  - "Master Info" requests the version of the command station (VB6: click on the label "Master Version").
//  - PomErgebnis (the answer on "Holen") is shown in the Address and Value fields.
//  - Corrected error: FunctionsZustandF29F68 stored all values in FB(4) and FB(5), instead of FB(6) .. FB(10).
//  - 2.0: the check box Handshake (USB), and the events SendError, InterfaceNichtAdressiert and
//    InterfaceWiederAdressiert in the status line
//

#import "AppDelegate.h"
#import "XpressNetLogFile.h"
#import "SerialPort.h"
#import "FunctionLedView.h"

#define MAX_TABLE_ROWS 1000                 // Maximum number of feedback messages shown in a table
#define WINDOW_WIDTH 1180                   // The window can only be resized vertically
#define MIN_WINDOW_HEIGHT 700

@implementation AppDelegate {
  NSMutableArray *_switchMessages;          // Rows of the switch feedback table, latest message first
  NSMutableArray *_rsBusMessages;           // Rows of the RS-Bus feedback table, latest message first
  NSTimer *_messageCounterTimer;            // Shows the number of XpressNet messages (VB6: TimerMessageCounter)
  // Function Groups: for each Lok Address, the values must be fetched again
  int FB[11];                               // FB: Function Befehl (on / off)
}

@synthesize XpressNet               = _XpressNet;
@synthesize window                  = _window;
@synthesize textIpAddress           = _textIpAddress;
@synthesize textTcpPort             = _textTcpPort;
@synthesize comboSerialPort         = _comboSerialPort;
@synthesize comboBaudrate           = _comboBaudrate;
@synthesize checkHandshake          = _checkHandshake;
@synthesize checkCts                = _checkCts;
@synthesize ledConnected            = _ledConnected;
@synthesize ledNotAus               = _ledNotAus;
@synthesize ledNotHalt              = _ledNotHalt;
@synthesize ledProgrammierMode      = _ledProgrammierMode;
@synthesize labelMasterVersion      = _labelMasterVersion;
@synthesize labelInterfaceVersion   = _labelInterfaceVersion;
@synthesize labelInterfaceCode      = _labelInterfaceCode;
@synthesize labelXpressNetAddress   = _labelXpressNetAddress;
@synthesize labelXpressNetVersion   = _labelXpressNetVersion;
@synthesize labelTcpFree            = _labelTcpFree;
@synthesize labelStatus             = _labelStatus;
@synthesize labelMessageCounter     = _labelMessageCounter;
@synthesize textSwitch              = _textSwitch;
@synthesize checkActivate           = _checkActivate;
@synthesize checkAutoDeactivate     = _checkAutoDeactivate;
@synthesize tableSwitch             = _tableSwitch;
@synthesize textRsBusAddress        = _textRsBusAddress;
@synthesize tableRsBus              = _tableRsBus;
@synthesize labelModellzeit         = _labelModellzeit;
@synthesize comboDay                = _comboDay;
@synthesize textHour                = _textHour;
@synthesize textMinute              = _textMinute;
@synthesize textFactor              = _textFactor;
@synthesize checkRun                = _checkRun;
@synthesize labelBuildNr            = _labelBuildNr;
@synthesize labelVersionRuckmelder  = _labelVersionRuckmelder;
@synthesize labelBuildRuckmelder    = _labelBuildRuckmelder;
@synthesize labelVersionBootloader  = _labelVersionBootloader;
@synthesize textLokAddress          = _textLokAddress;
@synthesize textLokFunction         = _textLokFunction;
@synthesize textLokSpeed            = _textLokSpeed;
@synthesize comboLokSteps           = _comboLokSteps;
@synthesize functionLeds            = _functionLeds;
@synthesize textExtAccAddress       = _textExtAccAddress;
@synthesize textExtAccValue         = _textExtAccValue;
@synthesize checkSM                 = _checkSM;
@synthesize radioDirect             = _radioDirect;
@synthesize radioRegister           = _radioRegister;
@synthesize radioPage               = _radioPage;
@synthesize textPomAddress          = _textPomAddress;
@synthesize textPomCV               = _textPomCV;
@synthesize textPomValue            = _textPomValue;


- (void)applicationDidFinishLaunching:(NSNotification *)aNotification
{
  _XpressNet = [[XpressNetClass alloc] init];
  [_XpressNet setDelegate:self];
  _switchMessages = [[NSMutableArray alloc] init];
  _rsBusMessages = [[NSMutableArray alloc] init];
  // The window can only be resized vertically: the feedback tables become longer
  [_window setContentMinSize:NSMakeSize(WINDOW_WIDTH, MIN_WINDOW_HEIGHT)];
  [_window setContentMaxSize:NSMakeSize(WINDOW_WIDTH, 10000)];
  // The values used last time
  NSUserDefaults *defaults = [NSUserDefaults standardUserDefaults];
  if ([defaults stringForKey:@"lastIpAddress"] != nil) [_textIpAddress setStringValue:[defaults stringForKey:@"lastIpAddress"]];
  if ([defaults stringForKey:@"lastTcpPort"] != nil)   [_textTcpPort setStringValue:[defaults stringForKey:@"lastTcpPort"]];
  if ([defaults stringForKey:@"lastBaudrate"] != nil)  [_comboBaudrate setStringValue:[defaults stringForKey:@"lastBaudrate"]];
  [self updateSerialPorts];
  NSString *lastPort = [defaults stringForKey:@"lastSerialPort"];
  if (lastPort != nil) [_comboSerialPort setStringValue:lastPort];
  else if ([_comboSerialPort numberOfItems] > 0) [_comboSerialPort selectItemAtIndex:0];
  [_comboDay selectItemAtIndex:0];
  // The LEDs start light grey
  NSColor *grey = [NSColor colorWithCalibratedWhite:0.88 alpha:1.0];
  [self showLed:_ledConnected color:grey];
  [self showLed:_ledNotAus color:grey];
  [self showLed:_ledNotHalt color:grey];
  [self showLed:_ledProgrammierMode color:grey];
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
  [_rsBusMessages release];
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

// The list of serial devices is updated each time the list is opened
- (void)comboBoxWillPopUp:(NSNotification *)notification {
  if ([notification object] == _comboSerialPort) [self updateSerialPorts];
}

- (void)updateSerialPorts {
  NSString *current = [[[_comboSerialPort stringValue] copy] autorelease];
  [_comboSerialPort removeAllItems];
  [_comboSerialPort addItemsWithObjectValues:[SerialPortClass availablePorts]];
  [_comboSerialPort setStringValue:current];
}


// ************************************************************************************************************
// ************************************************* CONNECTIONS **********************************************
// ************************************************************************************************************
- (IBAction)pushedConnectTcp:(id)sender {
  [self showStatusLine:@"TCP connecting ....."];
  [self showLed:_ledConnected color:[NSColor colorWithCalibratedWhite:0.88 alpha:1.0]];
  [[NSUserDefaults standardUserDefaults] setObject:[_textIpAddress stringValue] forKey:@"lastIpAddress"];
  [[NSUserDefaults standardUserDefaults] setObject:[_textTcpPort stringValue] forKey:@"lastTcpPort"];
  [_XpressNet CmdConnectTcp:[_textIpAddress stringValue] Port:[_textTcpPort stringValue]];
}

- (IBAction)pushedConnectUsb:(id)sender {
  [self showStatusLine:@"USB connecting ....."];
  [self saveSerialSettings];
  NSString *Settings = [NSString stringWithFormat:@"%@,N,8,1", [_comboBaudrate stringValue]];
  // 2.0: with the parameter Handshake
  // [_XpressNet CmdConnectUsb:[_comboSerialPort stringValue] Settings:Settings];
  [_XpressNet CmdConnectUsb:[_comboSerialPort stringValue] Settings:Settings Handshake:([_checkHandshake state] == NSControlStateValueOn)];
}

- (IBAction)pushedConnectSerial:(id)sender {
  [self showStatusLine:@"Serial connecting ....."];
  [self saveSerialSettings];
  NSString *Settings = [NSString stringWithFormat:@"%@,N,8,1", [_comboBaudrate stringValue]];
  [_XpressNet CmdConnectSerial:[_comboSerialPort stringValue] Settings:Settings Handshake:([_checkCts state] == NSControlStateValueOn)];
}

- (void)saveSerialSettings {
  [[NSUserDefaults standardUserDefaults] setObject:[_comboSerialPort stringValue] forKey:@"lastSerialPort"];
  [[NSUserDefaults standardUserDefaults] setObject:[_comboBaudrate stringValue] forKey:@"lastBaudrate"];
}

- (IBAction)pushedMasterInfo:(id)sender    {[_XpressNet CmdSoftwareversionZentraleAnfordern];}
- (IBAction)pushedInterfaceInfo:(id)sender {[_XpressNet CmdInterfaceVersion];}
- (IBAction)pushedErweiterteInfo:(id)sender {[_XpressNet CmdErweiterteZentralenVersionsinformation];}
- (IBAction)pushedStartMan:(id)sender      {[_XpressNet CmdZentralenStartmodeSetzen:NO];}
- (IBAction)pushedStartAuto:(id)sender     {[_XpressNet CmdZentralenStartmodeSetzen:YES];}

- (IBAction)pushedLanInfo:(id)sender {
  [_XpressNet CmdXpressNetAddress];
  [_XpressNet CmdXpressNetVersion];
  [_XpressNet CmdInterfaceFreieVerbindungen];
}


// ************************************************************************************************************
// ********************************************* SWITCHES AND FEEDBACK ****************************************
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

- (IBAction)pushedRsFeedback:(id)sender {
  if (![self isNumeric:_textRsBusAddress]) return;
  [_XpressNet CmdFeedbackRequest:[_textRsBusAddress intValue] Nibble:0];     // Nibble 0
  [_XpressNet CmdFeedbackRequest:[_textRsBusAddress intValue] Nibble:1];     // Nibble 1
}

- (IBAction)pushedExtAccSet:(id)sender {
  if ([self isNumeric:_textExtAccAddress]) [_XpressNet CmdZ21SetExtAccessory:[_textExtAccAddress intValue] value:[_textExtAccValue intValue]];
}

// Test of the FIFO: a fixed sequence of accessory commands, as in the VB6 version
- (IBAction)pushedSendTest:(id)sender {
  static const int test[][3] = {
    {745, 0, 1}, {745, 0, 0}, {745, 1, 1}, {745, 1, 0}, {746, 0, 1}, {746, 0, 0}, {746, 1, 1}, {746, 1, 0}, {751, 0, 1}, {751, 0, 0},
    {751, 1, 1}, {751, 1, 0}, {753, 0, 1}, {753, 0, 0}, {754, 0, 1}, {754, 0, 0}, {745, 1, 1}, {746, 0, 0}, {751, 0, 1}, {753, 0, 0},
    {745, 0, 1}, {745, 0, 0}, {745, 1, 1}, {745, 1, 0}, {746, 0, 1}, {746, 0, 0}, {746, 1, 1}, {746, 1, 0}, {751, 0, 1}, {751, 0, 0},
    {751, 1, 1}, {751, 1, 0}, {753, 0, 1}};
  for (unsigned int i = 0; i < sizeof(test) / sizeof(test[0]); i++) {
    [_XpressNet CmdAccessory:test[i][0] position:test[i][1] Activate:test[i][2]];
  }
}


// ************************************************************************************************************
// ************************************************* COMMAND STATION ******************************************
// ************************************************************************************************************
- (IBAction)pushedAllesAn:(id)sender         {[_XpressNet CmdAllesAn];}
- (IBAction)pushedAllesAus:(id)sender        {[_XpressNet CmdAllesAus];}
- (IBAction)pushedAlleLoksAnhalten:(id)sender {[_XpressNet CmdAlleLoksAnhalten];}


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

- (IBAction)pushedGetF0F12:(id)sender  {[_XpressNet CmdLokInfoAnfordern:[_textLokAddress intValue]];}
- (IBAction)pushedGetF13F28:(id)sender {[_XpressNet CmdFunktionsZustandF13F28Anfordern:[_textLokAddress intValue]];}
- (IBAction)pushedGetF29F68:(id)sender {[_XpressNet CmdFunktionsZustandF29F68Anfordern:[_textLokAddress intValue]];}

- (void)handleFunctionBefehl:(BOOL)turnOn {
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
// *************************************************** PROGRAMMING ********************************************
// ************************************************************************************************************
// Programming on Main as well as all three variants for Service Mode programming.
// For simplicity, we do not check if the input values are valid (as in the VB6 version)
- (BOOL)serviceMode {return ([_checkSM state] == NSControlStateValueOn);}

- (IBAction)selectedSmMode:(id)sender {
  // Nothing to do: the radio buttons are read when Get or Set is pushed. Having the same action makes them a group
}

- (IBAction)pushedPomGet:(id)sender {
  if ([self serviceMode]) {
    if ([_radioDirect state] == NSControlStateValueOn)   [_XpressNet CmdLeseanfrageProgrammieren:[_textPomCV intValue]];
    if ([_radioRegister state] == NSControlStateValueOn) [_XpressNet CmdLeseanfrageProgrammierenRegistermode:[_textPomCV intValue]];
    if ([_radioPage state] == NSControlStateValueOn)     [_XpressNet CmdLeseanfrageProgrammierenPagemode:[_textPomCV intValue]];
  }
  else [_XpressNet CmdPomRead:[_textPomAddress intValue] CV:[_textPomCV intValue]];
}

- (IBAction)pushedPomSet:(id)sender {
  if ([self serviceMode]) {
    if ([_radioDirect state] == NSControlStateValueOn)   [_XpressNet CmdSchreibbefehlProgrammieren:[_textPomCV intValue] Val:[_textPomValue intValue]];
    if ([_radioRegister state] == NSControlStateValueOn) [_XpressNet CmdSchreibbefehlProgrammierenRegisterMode:[_textPomCV intValue] daten:[_textPomValue intValue]];
    if ([_radioPage state] == NSControlStateValueOn)     [_XpressNet CmdSchreibbefehlProgrammierenPageMode:[_textPomCV intValue] daten:[_textPomValue intValue]];
  }
  else [_XpressNet CmdPomWrite:[_textPomAddress intValue] CV:[_textPomCV intValue] value:[_textPomValue intValue]];
}

- (IBAction)pushedPomHolen:(id)sender {
  if ([self serviceMode]) [_XpressNet CmdProgrammierergebnisAnfordern];
  else [_XpressNet CmdPoMErgebnisAnfordern];
}


// ************************************************************************************************************
// *************************************************** MODELLZEIT *********************************************
// ************************************************************************************************************
- (IBAction)pushedModellzeitAnfordern:(id)sender {[_XpressNet CmdModellzeitAnfordern];}
- (IBAction)pushedModellzeitAnhalten:(id)sender  {[_XpressNet CmdModellzeitAnhalten];}
- (IBAction)pushedModellzeitStarten:(id)sender   {[_XpressNet CmdModellzeitStarten];}

- (IBAction)pushedModellzeitStellen:(id)sender {
  int day = (int)[_comboDay indexOfSelectedItem];          // Monday = 0 .. Sunday = 6
  if (day < 0) day = 0;
  if ([_checkRun state] != NSControlStateValueOn) [_textFactor setIntValue:0];
  [_XpressNet CmdModellzeitStellen:day hour:[_textHour intValue] minute:[_textMinute intValue] factor:[_textFactor intValue]];
}


// ************************************************************************************************************
// **************************************************** TABLES ************************************************
// ************************************************************************************************************
- (NSInteger)numberOfRowsInTableView:(NSTableView *)tableView {
  if (tableView == _tableSwitch) return (NSInteger)[_switchMessages count];
  return (NSInteger)[_rsBusMessages count];
}

- (id)tableView:(NSTableView *)tableView objectValueForTableColumn:(NSTableColumn *)tableColumn row:(NSInteger)row {
  NSArray *messages = (tableView == _tableSwitch) ? _switchMessages : _rsBusMessages;
  if ((row < 0) || (row >= (NSInteger)[messages count])) return nil;
  return [[messages objectAtIndex:(NSUInteger)row] objectForKey:[tableColumn identifier]];
}

// The latest message at top. The number of messages is limited
- (void)addMessage:(NSDictionary *)message toTable:(NSTableView *)table messages:(NSMutableArray *)messages {
  [messages insertObject:message atIndex:0];
  if ([messages count] > MAX_TABLE_ROWS) [messages removeLastObject];
  [table reloadData];
}


// ************************************************************************************************************
// ************************************************ XPRESSNET EVENTS ******************************************
// ************************************************************************************************************
- (void)InterfaceConnected {
  [self showLed:_ledConnected color:[NSColor systemGreenColor]];
  [self showStatusLine:@"Connection established"];
}

- (void)TcpClosedByRemoteHost {
  [self showLed:_ledConnected color:[NSColor systemRedColor]];
  [self showStatusLine:@"TCP connection closed by remote host"];
}

- (void)TcpError:(NSString *)Description {
  [self showLed:_ledConnected color:[NSColor systemRedColor]];
  [self showStatusLine:Description];
}

// 2.0: a message that the library could not send. Reason: see SendErrorReasonEnum (XpressNet.h)
- (void)SendError:(int)Reason Description:(NSString *)Description {
  [self showStatusLine:[NSString stringWithFormat:@"%@: %@", [self sendErrorName:Reason], Description]];
}

// 2.0: the command station does not address the interface: sending is paused
- (void)InterfaceNichtAdressiert  {[self showStatusLine:@"Interface nicht adressiert"];}
// 2.0: the command station addresses the interface again
- (void)InterfaceWiederAdressiert {[self showStatusLine:@"Interface wieder adressiert"];}

- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data {
  NSDictionary *message = [NSDictionary dictionaryWithObjectsAndKeys:
                           [NSString stringWithFormat:@"%d", address], @"address",
                           [self printBits:data], @"data",
                           [NSString stringWithFormat:@"%d", Nibble], @"nibble", nil];
  [self addMessage:message toTable:_tableRsBus messages:_rsBusMessages];
}

- (void)SwitchFeedback:(int)address ttBits:(int)ttBits iBit:(int)iBit data:(int)data {
  // TT: C = Feedback comes from Command Station, D = Feedback comes from Decoder
  NSDictionary *message = [NSDictionary dictionaryWithObjectsAndKeys:
                           [NSString stringWithFormat:@"%d/%d", address, address + 1], @"address",
                           [self printBits:data], @"data",
                           (ttBits == 0) ? @"C" : @"D", @"tt",
                           [NSString stringWithFormat:@"%d", iBit], @"ibit", nil];
  [self addMessage:message toTable:_tableSwitch messages:_switchMessages];
}

- (void)AllesAn {
  [self showLed:_ledNotAus color:[NSColor systemGreenColor]];
  [self showLed:_ledNotHalt color:[NSColor systemGreenColor]];
  [self showLed:_ledProgrammierMode color:[NSColor systemGreenColor]];
}

- (void)NotAus          {[self showLed:_ledNotAus color:[NSColor systemRedColor]];}
- (void)NotHalt         {[self showLed:_ledNotHalt color:[NSColor systemRedColor]];}
- (void)ProgrammierMode {[self showLed:_ledProgrammierMode color:[NSColor systemRedColor]];}

- (void)ProgrammierinfoCV:(int)CV value:(int)value {
  [_textPomAddress setStringValue:@""];
  [_textPomCV setIntValue:CV];
  [_textPomValue setIntValue:value];
}

- (void)ProgrammierinfoPageRegister:(int)EEPromAddress value:(int)value {
  [_textPomAddress setStringValue:@""];
  [_textPomCV setIntValue:EEPromAddress];
  [_textPomValue setIntValue:value];
}

- (void)PomErgebnis:(int)LokAddress CvValue:(int)CvValue {
  [_textPomAddress setIntValue:LokAddress];
  [_textPomValue setIntValue:CvValue];
}

- (void)ZentraleVersion:(NSString *)Version {
  [_labelMasterVersion setStringValue:Version];
}

- (void)ErweiterteZentraleVersion:(NSString *)BuildZentrale VersionRuckMelder:(NSString *)VersionRuckMelder BuildRuckMelder:(NSString *)BuildRuckMelder VersionBootloader:(NSString *)VersionBootloader {
  [_labelBuildNr setStringValue:BuildZentrale];
  [_labelVersionRuckmelder setStringValue:VersionRuckMelder];
  [_labelBuildRuckmelder setStringValue:BuildRuckMelder];
  [_labelVersionBootloader setStringValue:VersionBootloader];
}

- (void)Fehlermeldung:(int)value {
  [self showStatusLine:[NSString stringWithFormat:@"Fehlermeldung: %d (V3.6: 3.1.14)", value]];
}

- (void)XpressNetAddress:(int)address           {[_labelXpressNetAddress setIntValue:address];}
- (void)XpressNetVersion:(NSString *)Version    {[_labelXpressNetVersion setStringValue:Version];}
- (void)InterfaceFreieVerbindungen:(int)number  {[_labelTcpFree setIntValue:number];}

- (void)InterfaceVersion:(NSString *)Version code:(NSString *)code {
  [_labelInterfaceVersion setStringValue:Version];
  [_labelInterfaceCode setStringValue:code];
}

- (void)Modellzeit:(int)day hour:(int)hour minute:(int)minute factor:(int)factor {
  [_labelModellzeit setStringValue:[NSString stringWithFormat:@"%02d:%02d", hour, minute]];
}

- (void)NormaleLokinfo:(BOOL)Besetzt Speedsteps:(int)Speedsteps Speed:(int)Speed Forward:(BOOL)Forward F0_F4:(int)F0_F4 F5_F8:(int)F5_F8 F9_F12:(int)F9_F12 {
  // Set the combobox and speed, based on the received input
  [_comboLokSteps setIntValue:Speedsteps];
  [_textLokSpeed setIntValue:Speed];
  // Store the values of F0..F12, to allow changing a single Function later
  FB[1] = F0_F4;                            // Gruppe 1
  FB[2] = F5_F8;                            // Gruppe 2
  FB[3] = F9_F12;                           // Gruppe 3
  // Show the LEDs for F0 .. F12
  [_functionLeds setGroup:1 data:F0_F4];
  [_functionLeds setGroup:2 data:F5_F8];
  [_functionLeds setGroup:3 data:F9_F12];
}

- (void)FunctionsZustandF13F20:(int)F13_F20 F21_F28:(int)F21_F28 {
  FB[4] = F13_F20;                          // Store Gruppe 4
  FB[5] = F21_F28;                          // Store Gruppe 5
  [_functionLeds setGroup:4 data:F13_F20];
  [_functionLeds setGroup:5 data:F21_F28];
}

- (void)FunctionsZustandF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68 {
  // 2026: VB6 stored these values in FB(4) and FB(5)
  FB[6] = F29_F36;                          // Store Gruppe 6
  FB[7] = F37_F44;                          // Store Gruppe 7
  FB[8] = F45_F52;                          // Store Gruppe 8
  FB[9] = F53_F60;                          // Store Gruppe 9
  FB[10] = F61_F68;                         // Store Gruppe 10
  [_functionLeds setGroup:6 data:F29_F36];
  [_functionLeds setGroup:7 data:F37_F44];
  [_functionLeds setGroup:8 data:F45_F52];
  [_functionLeds setGroup:9 data:F53_F60];
  [_functionLeds setGroup:10 data:F61_F68];
}

@end
