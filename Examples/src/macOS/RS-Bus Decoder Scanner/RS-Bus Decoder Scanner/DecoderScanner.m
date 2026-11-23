//
//  DecoderScanner.m
//  RS-Bus Decoder Scanner
//
//  Created by Aiko Pras on 22-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  See DecoderScanner.h for a description of the three steps of the scan.
//  Only one PoM read is outstanding at any moment, since all decoders answer via the same RS-Bus address (128).
//

#import "DecoderScanner.h"
#import "XpressNetLogFile.h"

#define CONNECT_TIMEOUT 5.0                     // Seconds to wait for the connections with the Lenz interfaces
#define RSBUS_TIMEOUT 20.0                      // Seconds to wait for the RS-Bus data of one command station
#define POM_TIMEOUT 1.0                         // Seconds to wait for the RS-Bus answer on a PoM read (as Decoder-POM)
#define MAX_ATTEMPTS 2                          // Number of PoM reads per CV (per command station for CV7)
#define ACCESSORY_POM_OFFSET 6999               // PoM address = 6999 + RS-Bus address (CV10 = decoder address + 1)
#define FEEDBACK_POM_OFFSET 6000                // PoM address = 6000 + RS-Bus address (as Programmer GBM-POM)
#define RS_BUS_FEEDBACK_ADDRESS 128             // The RS-Bus address for the PoM answers
#define RSBUS_PART 32                           // RS-Bus addresses per Get RS-Bus data request (see nextRsStation)
#define PROGRESS_RSBUS 10.0                     // Part (%) of the progress bar for step 2; step 3 takes the rest

#define QUIET_TIME 0.5                          // Seconds to ignore all answers after a timeout or a bad value
#define FEC_GUARD_TIME 0.3                      // Seconds to ignore repeated answers (CV7 and RSFEC > 0)
#define MAX_REREADS 2                           // Number of extra reads of a CV with a value that is not possible

// The names of the CVs, as in CvValues.h of the AP_DCC_Decoder_Core library
#define myAddrL        1
#define version        7
#define myAddrH        9
#define myRSAddr      10
#define CmdStation    19
#define RSFEC         20
#define SkipUnEven    21
#define DecType       27
#define SwitchCount   35
#define SwitchType    36


// ****************************************************************************************************************
// *********************************************** SCANNED DECODER ************************************************
// ****************************************************************************************************************
@implementation ScannedDecoderClass {
  int cvStore[MAX_CV_NUMBER + 1];
}

@synthesize rsStation  = _rsStation;
@synthesize rsAddress  = _rsAddress;
@synthesize rsBusType  = _rsBusType;
@synthesize pomStation = _pomStation;
@synthesize skip       = _skip;
@synthesize givenPomAddress = _givenPomAddress;
@synthesize liftIr     = _liftIr;
@synthesize liftLevel  = _liftLevel;
@synthesize pomIpAddress = _pomIpAddress;

- (id)init {
  self = [super init];
  if (self) {
    for (int i = 0; i <= MAX_CV_NUMBER; i++) cvStore[i] = -1;
    _pomStation = 0;
    _skip = NO;
  }
  return self;
}

- (void)dealloc {
  [_pomIpAddress release];
  [super dealloc];
}

- (int)pomAddress {
  if (_givenPomAddress > 0) return _givenPomAddress;
  if (_rsBusType == ACCESSORY_DECODER) return _rsAddress + ACCESSORY_POM_OFFSET;
  return _rsAddress + FEEDBACK_POM_OFFSET;
}

- (void)setCv:(int)cvNumber value:(int)value {
  if ((cvNumber > 0) && (cvNumber <= MAX_CV_NUMBER)) cvStore[cvNumber] = value;
  // A decoder with a given PoM address gets its RS-Bus address from CV10
  if ((_givenPomAddress > 0) && (cvNumber == myRSAddr)) _rsAddress = value;
}

- (int)getCv:(int)cvNumber {
  if ((cvNumber > 0) && (cvNumber <= MAX_CV_NUMBER)) return cvStore[cvNumber];
  return -1;
}

- (NSString *)cvText:(int)cvNumber {
  int value = [self getCv:cvNumber];
  if (value < 0) return @"?";
  return [NSString stringWithFormat:@"%d", value];
}

// The number of switches of an accessory decoder. Decoders with version < 20 have 4 switches. From version 20,
// as AVR-Servo-3-6.ino: SwitchType (CV36) 1 (DKW) = 2 switches, 2 (double crossover) = 4 switches, otherwise
// SwitchCount (CV35). SwitchCount 0, 255 or above the board maximum means: the board maximum
// A GBM with relays board (DecType 50) has 4 relays (NUMBER_OF_DEVICES in global.h of the GBM firmware).
// The lift decoder (DecType 24) listens to 3 decoder addresses = 12 switch addresses (Lift_Main.ino)
- (int)numberOfSwitches {
  if (_rsBusType != ACCESSORY_DECODER) return (cvStore[DecType] == 50) ? 4 : -1;
  int cv7 = cvStore[version];
  if (cv7 < 0) return -1;
  if (cvStore[DecType] == 24) return 12;
  if (cv7 < 20) return 4;
  int boardMaximum;
  switch (cvStore[DecType]) {
    case 20:  boardMaximum = 2; break;          // Servo-2
    case 21:  boardMaximum = 3; break;          // Servo-3
    case 22:  boardMaximum = 6; break;          // Servo-6
    case 33:  boardMaximum = 16; break;         // Relays decoder for 16 relais
    default:  boardMaximum = 4; break;
  }
  if (cvStore[SwitchType] == 1) return 2;
  if (cvStore[SwitchType] == 2) return 4;
  int count = cvStore[SwitchCount];
  if ((count <= 0) || (count == 255) || (count > boardMaximum)) return boardMaximum;
  return count;
}

// An RS-Bus address holds 4 switches (2 bits each). With SkipUnEven each switch has its own nibble (2 per address)
// The lift decoder uses two RS-Bus addresses: myRSAddr and myRSAddr + 1 (feedback.cpp of Lift_Main)
- (int)rsAddressCount {
  if (_liftIr) return 2;                        // RS_AddresLow and RS_AddresHigh (mySettings.h of Lift_IR)
  if (_rsBusType != ACCESSORY_DECODER) return 1;
  if (cvStore[DecType] == 24) return 2;
  int switches = [self numberOfSwitches];
  if (switches <= 0) return 1;
  if (cvStore[SkipUnEven] == 1) return (switches + 1) / 2;
  return (switches + 3) / 4;
}

- (NSString *)rsAddressText {
  if (_rsAddress == 0) return @"?";             // A decoder with a given PoM address, CV10 not received
  int count = [self rsAddressCount];
  if (count <= 1) return [NSString stringWithFormat:@"%d", _rsAddress];
  return [NSString stringWithFormat:@"%d-%d", _rsAddress, _rsAddress + count - 1];
}

// The switch addresses of the decoder (first address as getSwitchAddress of Programmer Decoder-POM).
// With SkipUnEven each switch uses two switch addresses (of which only the even one is used).
// A GBM listens to the accessory address in CV1 and CV9 for its relays (main.c and dcc_decode.c of the GBM
// firmware). It is shown for a GBM with relays board only. If bit 7 of CV9 is set, the decoder has not been
// initialised (main.c of the GBM and switch decoder firmware): "-"
- (NSString *)accessoryAddressText {
  int cv1 = cvStore[myAddrL];
  int cv9 = cvStore[myAddrH];
  int cv7 = cvStore[version];
  int cv27 = cvStore[DecType];
  int decoderAddress;
  if (_rsBusType != ACCESSORY_DECODER) {
    if (cv27 < 0) return @"?";
    if (cv27 != 50) return @"";                   // 50: GBM with relays board
  }
  if ((cv1 < 0) || (cv9 < 0) || (cv7 < 0)) return @"?";
  if (cv9 & 0x80) return @"-";
  if (_rsBusType == ACCESSORY_DECODER) {
    // Decoders with version >= 16: decoderAddress = CV9 << 6 + CV1 - 1 (as Programmer Decoder-POM).
    // The lift decoder has version 12, but uses storedAddress() of AP_DCC_Decoder_Core, which also subtracts 1
    decoderAddress = ((cv9 & 0x7F) << 6) | cv1;
    if ((cv7 >= 16) || (cv27 == 24)) decoderAddress = decoderAddress - 1;
  }
  else decoderAddress = (((cv9 & 0x07) << 6) | cv1) - 1;
  if (decoderAddress < 0) return @"-";
  int first = decoderAddress * 4 + 1;
  int switches = [self numberOfSwitches];
  if ((cvStore[SkipUnEven] == 1) && (cvStore[DecType] != 24)) switches = switches * 2;
  return [NSString stringWithFormat:@"%d-%d", first, first + switches - 1];
}

// CmdStation (CV19), as CvValues.h of the AP_DCC_Decoder_Core library
- (NSString *)cmdStationText {
  switch (cvStore[CmdStation]) {
    case -1: return @"?";
    case 0:  return @"Roco";
    case 1:  return @"Lenz";
    case 2:  return @"OpenDCC Z1";
  }
  return [NSString stringWithFormat:@"%d", cvStore[CmdStation]];
}

// The names of DecType (CV27), as in Programmer Decoder-POM (DCCDecoder.m) and CvValues.h of AP_DCC_Decoder_Core
+ (NSString *)typeName:(int)value {
  NSString *type;
  NSString *subType = @"";
  switch (value) {
    case 16:  type = @"Switch decoder"; break;
    case 17:  type = @"Switch decoder"; subType = @"with Emergency board"; break;
    case 20:  type = @"Servo-2 Decoder"; break;
    case 21:  type = @"Servo-3 Decoder"; break;
    case 22:  type = @"Servo-6 Decoder"; break;
    case 24:  type = @"Lift decoder - Level"; break;
    case 32:  type = @"Relays decoder"; subType = @"for 4 relais"; break;
    case 33:  type = @"Relays decoder"; subType = @"for 16 relais"; break;
    case 34:  type = @"Relays decoder"; subType = @"for a configurable number of relais"; break;
    case 48:  type = @"Track Occupancy decoder"; break;
    case 49:  type = @"Track Occupancy decoder"; subType = @"with reverser board"; break;
    case 50:  type = @"Track Occupancy decoder"; subType = @"with relays board"; break;
    case 52:  type = @"Track Occupancy decoder"; subType = @"with speed measurement"; break;
    case 64:  type = @"Function Decoder"; break;
    case 128: type = @"Safety decoder"; subType = @"with Watchdog"; break;
    case 193: type = @"TMC 24 Channel IO Decoder"; break;
    case 194: type = @"TMC 16 Channel Switch Decoder"; break;
    default:  return nil;
  }
  if ([subType length] == 0) return type;
  return [NSString stringWithFormat:@"%@ %@", type, subType];
}

- (NSString *)typeText {
  if (_liftIr) return @"Lift decoder - IR";
  int value = cvStore[DecType];
  if (value < 0) return @"?";
  NSString *name = [ScannedDecoderClass typeName:value];
  if (name == nil) return [NSString stringWithFormat:@"Unknown (CV27 = %d)", value];
  return name;
}

@end


// ****************************************************************************************************************
// *********************************************** COMMAND STATION ************************************************
// ****************************************************************************************************************
@implementation StationClass

@synthesize number      = _number;
@synthesize ipAddress   = _ipAddress;
@synthesize XpressNet   = _XpressNet;
@synthesize connected   = _connected;
@synthesize connectDone = _connectDone;
@synthesize scanner     = _scanner;

- (void)dealloc {
  [_XpressNet setDelegate:nil];
  [_XpressNet release];                         // Closes the connection
  [_ipAddress release];
  [super dealloc];
}

// The events of the XpressNet library, forwarded to the scanner together with this station
- (void)InterfaceConnected {
  [_scanner station:self connected:YES text:nil];
}

- (void)TcpClosedByRemoteHost {
  [_scanner station:self connected:NO text:@"TCP connection closed by remote host"];
}

- (void)TcpError:(NSString *)Description {
  [_scanner station:self connected:NO text:Description];
}

// Feedback decoder (TT=10): address = 1..128
- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data {
  [_scanner station:self rsAddress:address nibble:Nibble data:data ttBits:2];
}

// Switch decoder (TT=00, 01 or 11): switchAddress = ((address * 4) + (Nibble * 2) + 1), with address = 0..127
// The conversion back to RS-Bus address and nibble is as in the RS-Bus Monitor
- (void)SwitchFeedback:(int)switchAddress ttBits:(int)ttBits iBit:(int)iBit data:(int)data {
  int rsAddress = (switchAddress - 1) / 4 + 1;
  int nibble = ((switchAddress - 1) / 2) % 2;
  [_scanner station:self rsAddress:rsAddress nibble:nibble data:data ttBits:ttBits];
}

- (void)SendError:(int)Reason Description:(NSString *)Description {
  [_scanner station:self sendError:[NSString stringWithFormat:@"SendError %d: %@", Reason, Description]];
}

@end


// ****************************************************************************************************************
// *********************************************** DECODER SCANNER ************************************************
// ****************************************************************************************************************
typedef enum
{ IDLE = 0,
  CONNECTING = 1,                               // Step 1
  RSBUS_DATA = 2,                               // Step 2
  POM_READ = 3,                                 // Step 3
} scanPhase_t;

typedef enum                                    // The next step after the quiet time or FEC guard time
{ RESEND = 0,
  NEXT_CV = 1,
  NEXT_DECODER = 2,
} pomAction_t;

static NSInteger compareDecoders(id a, id b, void *context);

@implementation DecoderScannerClass {
  scanPhase_t phase;
  NSMutableArray *stations;                     // StationClass objects
  NSMutableArray *found;                        // All decoders found on the RS-Bus (ScannedDecoderClass)
  NSMutableArray *problems;                     // Texts for the final status line
  NSMutableArray *noAnswer;                     // RS-Bus addresses without PoM answer (NSString)
  NSTimer *timer;                               // The running timeout (one at a time). Retained by the run loop
  int numberConnected;
  // Step 2
  int rsStationIndex;                           // Index in stations
  int rsStationsDone;
  int partLastAddress;                          // The last RS-Bus address of the part that has been requested
  int answerCount;                              // The answers for partLastAddress
  // Step 3
  int decoderIndex;                             // Index in found
  NSMutableArray *userCvs;                      // The CVs chosen by the user (NSNumber)
  int liftPomAddress;                           // The lift decoder (0 = not included)
  int liftIrAddress;                            // The first RS-Bus address of the IR sensors of the lift (0 = none)
  NSMutableArray *liftRows;                     // The rows of the lift decoder, added at the end of the table
  NSMutableArray *cvList;                       // The CVs to read for the current decoder (NSNumber)
  int cvIndex;                                  // Index in cvList
  NSMutableArray *candidates;                   // The stations to try for CV7 (StationClass)
  int candidateIndex;
  int attempt;
  int rereads;                                  // Extra reads because of a value that is not possible
  pomAction_t pendingAction;
  BOOL waitingForAnswer;
  BOOL haveLowNibble;
  int lowNibble;
  int lastSuccessfulStation;                    // Number of the station that delivered the last PoM messages
}

@synthesize delegate = _delegate;
@synthesize decoders = _decoders;

- (id)init {
  self = [super init];
  if (self) {
    phase = IDLE;
    stations = [[NSMutableArray alloc] init];
    found = [[NSMutableArray alloc] init];
    problems = [[NSMutableArray alloc] init];
    noAnswer = [[NSMutableArray alloc] init];
    candidates = [[NSMutableArray alloc] init];
    userCvs = [[NSMutableArray alloc] init];
    liftRows = [[NSMutableArray alloc] init];
    cvList = [[NSMutableArray alloc] init];
    _decoders = [[NSMutableArray alloc] init];
  }
  return self;
}

- (void)dealloc {
  [self stopTimer];
  for (StationClass *station in stations) [station setScanner:nil];
  [stations release];
  [found release];
  [problems release];
  [noAnswer release];
  [candidates release];
  [userCvs release];
  [liftRows release];
  [cvList release];
  [_decoders release];
  [super dealloc];
}

- (BOOL)active {return (phase != IDLE);}


// ************************************************************************************************************
// ************************************************ START AND STOP ********************************************
// ************************************************************************************************************
- (void)startWithIpAddresses:(NSArray *)ipAddresses cvNumbers:(NSArray *)cvNumbers
                liftPomAddress:(int)pomAddress liftIrAddress:(int)irAddress {
  if (phase != IDLE) return;
  liftPomAddress = pomAddress;
  liftIrAddress = irAddress;
  [liftRows removeAllObjects];
  [userCvs removeAllObjects];
  for (NSNumber *cv in cvNumbers) {
    if (([cv intValue] >= 1) && ([cv intValue] <= MAX_CV_NUMBER) && ![userCvs containsObject:cv]) [userCvs addObject:cv];
  }
  [found removeAllObjects];
  [_decoders removeAllObjects];
  [problems removeAllObjects];
  [noAnswer removeAllObjects];
  lastSuccessfulStation = 0;
  // Step 1: create a XpressNet object for every IP address, and open the connection
  for (int i = 0; i < [ipAddresses count]; i++) {
    NSString *ipAddress = [[ipAddresses objectAtIndex:i] stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];
    if ([ipAddress length] == 0) continue;
    StationClass *station = [[StationClass alloc] init];
    [station setNumber:i + 1];
    [station setIpAddress:ipAddress];
    [station setScanner:self];
    XpressNetClass *xpressNet = [[XpressNetClass alloc] init];
    [xpressNet setDelegate:station];
    [[xpressNet Logfile] setLogDirectory:@"~/Library/Logs/TMC-Logfiles"];
    [station setXpressNet:xpressNet];
    [xpressNet release];
    [stations addObject:station];
    [station release];
  }
  if ([stations count] == 0) {
    [_delegate scannerFinished:@"No IP address in the preferences"];
    return;
  }
  phase = CONNECTING;
  [_delegate scannerStatus:[NSString stringWithFormat:@"Connecting to %d command station(s) .....", (int)[stations count]] progress:0];
  for (StationClass *station in stations) [[station XpressNet] CmdConnectTcp:[station ipAddress] Port:@"5550"];
  // A connection may already have failed in CmdConnectTcp (for example: an invalid IP address)
  if (phase == CONNECTING) [self startTimer:CONNECT_TIMEOUT selector:@selector(connectTimeout)];
}


- (void)stop {
  if (phase == IDLE) return;
  [self finish:@"Scan stopped"];
}


// The end of the scan. The XpressNet objects (and thus the connections) are released after the current event has
// been handled, since this method may be called from within an event of one of these objects
- (void)finish:(NSString *)text {
  [self stopTimer];
  phase = IDLE;
  waitingForAnswer = NO;
  for (StationClass *station in stations) [station setScanner:nil];   // Later events are ignored
  NSMutableArray *oldStations = stations;
  stations = [[NSMutableArray alloc] init];
  [self performSelector:@selector(releaseStations:) withObject:oldStations afterDelay:0];
  [oldStations release];
  [candidates removeAllObjects];
  [_decoders sortUsingFunction:compareDecoders context:NULL];
  [_decoders addObjectsFromArray:liftRows];     // The lift decoder at the end of the table
  NSString *finalText = text;
  if ([problems count] > 0) finalText = [NSString stringWithFormat:@"%@. %@", text, [problems componentsJoinedByString:@"; "]];
  [_delegate scannerFinished:finalText];
}

- (void)releaseStations:(NSArray *)oldStations {
  // Nothing to do: performSelector retained the array, and releases it (and thereby the stations) after this call
}


// ************************************************************************************************************
// **************************************************** TIMER *************************************************
// ************************************************************************************************************
- (void)startTimer:(NSTimeInterval)seconds selector:(SEL)selector {
  [self stopTimer];
  timer = [NSTimer scheduledTimerWithTimeInterval:seconds target:self selector:selector userInfo:nil repeats:NO];
}

- (void)stopTimer {
  [timer invalidate];
  timer = nil;
}


// ************************************************************************************************************
// ******************************************** STEP 1: CONNECT ***********************************************
// ************************************************************************************************************
- (void)station:(StationClass *)station connected:(BOOL)ok text:(NSString *)text {
  if (phase == CONNECTING) {
    if ([station connectDone]) return;
    [station setConnectDone:YES];
    [station setConnected:ok];
    if (!ok) [problems addObject:[NSString stringWithFormat:@"Command station %d (%@): %@", [station number], [station ipAddress], text]];
    for (StationClass *s in stations) if (![s connectDone]) return;
    [self startRsBusData];
    return;
  }
  // The connection is closed during the scan. Step 3 continues with the other stations (via the timeouts)
  if (!ok && [station connected]) {
    [station setConnected:NO];
    [problems addObject:[NSString stringWithFormat:@"Command station %d (%@): %@", [station number], [station ipAddress], text]];
    if ((phase == RSBUS_DATA) && (station == [stations objectAtIndex:rsStationIndex])) [self nextRsStation];
  }
}

- (void)connectTimeout {
  timer = nil;
  if (phase != CONNECTING) return;
  for (StationClass *station in stations) {
    if (![station connectDone]) {
      [station setConnectDone:YES];
      [station setConnected:NO];
      [problems addObject:[NSString stringWithFormat:@"Command station %d (%@): no connection", [station number], [station ipAddress]]];
    }
  }
  [self startRsBusData];
}


// ************************************************************************************************************
// ****************************************** STEP 2: RS-BUS DATA *********************************************
// ************************************************************************************************************
- (void)startRsBusData {
  [self stopTimer];
  numberConnected = 0;
  for (StationClass *station in stations) if ([station connected]) numberConnected++;
  if (numberConnected == 0) {
    [self finish:@"No connection with any command station"];
    return;
  }
  phase = RSBUS_DATA;
  rsStationIndex = -1;
  rsStationsDone = 0;
  [self nextRsStation];
}

// As Get RS-Bus data of the RS-Bus Monitor, but in four parts of 32 RS-Bus addresses (two nibbles each). The FIFO of
// the library holds 128 messages, but after the connection is established the library itself also sends messages
// (software version and status of the command station). The next part is requested after both answers for the last
// RS-Bus address of the current part; the station is ready after both answers for RS-Bus address 128
- (void)nextRsStation {
  [self stopTimer];
  if (rsStationIndex >= 0) rsStationsDone++;
  rsStationIndex++;
  while ((rsStationIndex < [stations count]) && ![[stations objectAtIndex:rsStationIndex] connected]) rsStationIndex++;
  if (rsStationIndex >= [stations count]) {
    [self startPom];
    return;
  }
  StationClass *station = [stations objectAtIndex:rsStationIndex];
  partLastAddress = RSBUS_PART;
  answerCount = 0;
  [self showRsBusProgress:0];
  [self requestFeedbacks:station from:1 to:partLastAddress];
  [self startTimer:RSBUS_TIMEOUT selector:@selector(rsBusTimeout)];
}

- (void)requestFeedbacks:(StationClass *)station from:(int)firstAddress to:(int)lastAddress {
  for (int address = firstAddress; address <= lastAddress; address++) {
    [[station XpressNet] CmdFeedbackRequest:address Nibble:0];
    [[station XpressNet] CmdFeedbackRequest:address Nibble:1];
  }
}

- (void)showRsBusProgress:(int)rsAddress {
  StationClass *station = [stations objectAtIndex:rsStationIndex];
  NSString *text = [NSString stringWithFormat:@"Command station %d (%@): requesting RS-Bus data .....", [station number], [station ipAddress]];
  double progress = PROGRESS_RSBUS * (rsStationsDone + rsAddress / 128.0) / numberConnected;
  [_delegate scannerStatus:text progress:progress];
}

- (void)rsBusTimeout {
  timer = nil;
  StationClass *station = [stations objectAtIndex:rsStationIndex];
  [problems addObject:[NSString stringWithFormat:@"Command station %d (%@): RS-Bus data incomplete", [station number], [station ipAddress]]];
  [self nextRsStation];
}

// An answer during step 2: TT=01 (accessory decoder) and TT=10 (GBM) are stored. RS-Bus address 128 is skipped,
// since it holds the PoM answers
- (void)rsBusAnswer:(StationClass *)station rsAddress:(int)rsAddress ttBits:(int)ttBits {
  if (station != [stations objectAtIndex:rsStationIndex]) return;
  if ((rsAddress < RS_BUS_FEEDBACK_ADDRESS) && ((ttBits == 1) || (ttBits == 2))) {
    BOOL known = NO;
    for (ScannedDecoderClass *d in found) if (([d rsStation] == [station number]) && ([d rsAddress] == rsAddress)) known = YES;
    if (!known) {
      ScannedDecoderClass *decoder = [[ScannedDecoderClass alloc] init];
      [decoder setRsStation:[station number]];
      [decoder setRsAddress:rsAddress];
      [decoder setRsBusType:(ttBits == 1) ? ACCESSORY_DECODER : FEEDBACK_DECODER];
      [found addObject:decoder];
      [decoder release];
    }
  }
  [self showRsBusProgress:rsAddress];
  if (rsAddress != partLastAddress) return;
  answerCount++;
  if (answerCount < 2) return;
  if (partLastAddress >= 128) {                 // The last part has been answered
    [self nextRsStation];
    return;
  }
  answerCount = 0;                              // The next part
  partLastAddress = partLastAddress + RSBUS_PART;
  [self requestFeedbacks:station from:partLastAddress - RSBUS_PART + 1 to:partLastAddress];
}


// ************************************************************************************************************
// ********************************************** STEP 3: POM *************************************************
// ************************************************************************************************************
// Sort order: first the accessory decoders, then the GBMs; per type on RS-Bus address
static NSInteger compareFound(id a, id b, void *context) {
  ScannedDecoderClass *d1 = a;
  ScannedDecoderClass *d2 = b;
  if ([d1 rsBusType] != [d2 rsBusType]) return ([d1 rsBusType] < [d2 rsBusType]) ? NSOrderedAscending : NSOrderedDescending;
  if ([d1 rsAddress] != [d2 rsAddress]) return ([d1 rsAddress] < [d2 rsAddress]) ? NSOrderedAscending : NSOrderedDescending;
  if ([d1 rsStation] != [d2 rsStation]) return ([d1 rsStation] < [d2 rsStation]) ? NSOrderedAscending : NSOrderedDescending;
  return NSOrderedSame;
}

// Sort order of the table: on the station that delivered the PoM messages, then on PoM address
static NSInteger compareDecoders(id a, id b, void *context) {
  ScannedDecoderClass *d1 = a;
  ScannedDecoderClass *d2 = b;
  if ([d1 pomStation] != [d2 pomStation]) return ([d1 pomStation] < [d2 pomStation]) ? NSOrderedAscending : NSOrderedDescending;
  if ([d1 pomAddress] != [d2 pomAddress]) return ([d1 pomAddress] < [d2 pomAddress]) ? NSOrderedAscending : NSOrderedDescending;
  return NSOrderedSame;
}

// A value that is not possible, according to the ranges in CvValues.h of the AP_DCC_Decoder_Core library.
// CV35 and CV36 differ between accessory decoders and GBMs, and are only checked for accessory decoders
static BOOL isPlausible(int cv, int value, rsBusType_t type) {
  switch (cv) {
    case version:     return (value >= 8);
    case myAddrH:     return ((value & 0x7F) <= 7);         // Bit 7: not initialised
    case myRSAddr:    return (value <= 128);
    case CmdStation:  return (value <= 2);
    case RSFEC:       return (value <= 2);
    case DecType:     return ([ScannedDecoderClass typeName:value] != nil);
    case SkipUnEven:  return (value <= 1);
    case SwitchType:  return (type != ACCESSORY_DECODER) || (value <= 2);
  }
  return YES;
}

- (void)startPom {
  [self stopTimer];
  phase = POM_READ;
  [found sortUsingFunction:compareFound context:NULL];
  // The lift decoder is read first, so its RS-Bus addresses are known (and skipped) during the rest of the scan.
  // Its rows are added at the end of the table, after the scan. Its RS-Bus address comes from CV10, and its answers
  // are accepted from any command station (rsStation 0) until the first answer has been received
  if (liftPomAddress > 0) {
    ScannedDecoderClass *decoder = [[ScannedDecoderClass alloc] init];
    [decoder setRsBusType:ACCESSORY_DECODER];
    [decoder setRsStation:0];
    [decoder setRsAddress:0];
    [decoder setGivenPomAddress:liftPomAddress];
    [found insertObject:decoder atIndex:0];
    [decoder release];
  }
  decoderIndex = -1;
  [self nextDecoder];
}

- (void)nextDecoder {
  decoderIndex++;
  while ((decoderIndex < [found count]) && [[found objectAtIndex:decoderIndex] skip]) decoderIndex++;
  if (decoderIndex >= [found count]) {
    if ([noAnswer count] > 0)
      [problems insertObject:[NSString stringWithFormat:@"No PoM answer for %@", [noAnswer componentsJoinedByString:@", "]] atIndex:0];
    [self finish:[NSString stringWithFormat:@"Scan ready: %d decoder(s) answered via PoM", (int)[_decoders count]]];
    return;
  }
  // The stations to try for CV7: the last successful one first, then the others
  [candidates removeAllObjects];
  for (StationClass *station in stations)
    if ([station connected] && ([station number] == lastSuccessfulStation)) [candidates addObject:station];
  for (StationClass *station in stations)
    if ([station connected] && ([station number] != lastSuccessfulStation)) [candidates addObject:station];
  if ([candidates count] == 0) {
    [self finish:@"No connection with any command station"];
    return;
  }
  // First CV7 only. The other CVs are added once CV7 has been received (see buildCvList)
  [cvList removeAllObjects];
  [cvList addObject:[NSNumber numberWithInt:version]];
  candidateIndex = 0;
  cvIndex = 0;
  attempt = 0;
  rereads = 0;
  [self sendPomRead];
}

// The CVs needed for the table first (RSFEC second, for the FEC guard time), then the CVs chosen by the user
- (void)buildCvList:(ScannedDecoderClass *)decoder {
  [cvList addObject:[NSNumber numberWithInt:RSFEC]];
  [cvList addObject:[NSNumber numberWithInt:DecType]];
  [cvList addObject:[NSNumber numberWithInt:SkipUnEven]];
  if ([decoder rsBusType] == ACCESSORY_DECODER) {
    if ([decoder getCv:version] >= 20) {
      [cvList addObject:[NSNumber numberWithInt:SwitchCount]];
      [cvList addObject:[NSNumber numberWithInt:SwitchType]];
    }
  }
  [cvList addObject:[NSNumber numberWithInt:myAddrL]];
  [cvList addObject:[NSNumber numberWithInt:myAddrH]];
  if ([decoder givenPomAddress] > 0) [cvList addObject:[NSNumber numberWithInt:myRSAddr]];
  for (NSNumber *cv in userCvs) if (![cvList containsObject:cv]) [cvList addObject:cv];
}

- (StationClass *)stationWithNumber:(int)number {
  for (StationClass *station in stations) if ([station number] == number) return station;
  return nil;
}

- (void)sendPomRead {
  ScannedDecoderClass *decoder = [found objectAtIndex:decoderIndex];
  StationClass *station;
  if (cvIndex == 0) station = [candidates objectAtIndex:candidateIndex];
  else station = [self stationWithNumber:[decoder pomStation]];
  int cv = [[cvList objectAtIndex:cvIndex] intValue];
  NSString *text = [NSString stringWithFormat:@"Decoder %d of %d (RS-Bus address %d, PoM address %d): CV%d via command station %d",
                    decoderIndex + 1, (int)[found count], [decoder rsAddress], [decoder pomAddress], cv, [station number]];
  double progress = PROGRESS_RSBUS + (100.0 - PROGRESS_RSBUS) * (decoderIndex + (double)cvIndex / [cvList count]) / [found count];
  [_delegate scannerStatus:text progress:progress];
  haveLowNibble = NO;
  waitingForAnswer = YES;
  [[station XpressNet] CmdPomRead:[decoder pomAddress] CV:cv];
  [self startTimer:POM_TIMEOUT selector:@selector(pomTimeout)];
}

// An answer during step 3. The decoder sends two RS-Bus messages: first the low nibble, then the high nibble
- (void)pomAnswer:(StationClass *)station rsAddress:(int)rsAddress nibble:(int)nibble data:(int)data {
  if (!waitingForAnswer) return;
  ScannedDecoderClass *decoder = [found objectAtIndex:decoderIndex];
  if (([decoder rsStation] != 0) && ([station number] != [decoder rsStation])) return;
  BOOL ownAddress = ([decoder rsBusType] == FEEDBACK_DECODER) && (rsAddress == [decoder rsAddress]);   // Old GBM firmware
  if ((rsAddress != RS_BUS_FEEDBACK_ADDRESS) && !ownAddress) return;
  if (nibble == 0) {
    lowNibble = data & 0x0F;
    haveLowNibble = YES;
    return;
  }
  if (!haveLowNibble) return;                   // The first nibble was lost: wait for the timeout
  [self stopTimer];
  waitingForAnswer = NO;
  int value = ((data & 0x0F) << 4) + lowNibble;
  int cv = [[cvList objectAtIndex:cvIndex] intValue];
  // A value that is not possible is read again, after the quiet time
  if (!isPlausible(cv, value, [decoder rsBusType]) && (rereads < MAX_REREADS)) {
    rereads++;
    [self afterTime:QUIET_TIME action:RESEND];
    return;
  }
  [decoder setCv:cv value:value];
  if ([decoder rsStation] == 0) [decoder setRsStation:[station number]];   // A decoder with a given PoM address
  if (cvIndex == 0) {
    StationClass *pomStation = [candidates objectAtIndex:candidateIndex];
    [decoder setPomStation:[pomStation number]];
    [decoder setPomIpAddress:[pomStation ipAddress]];
    lastSuccessfulStation = [pomStation number];
    [self buildCvList:decoder];
    if ([decoder givenPomAddress] > 0) [liftRows addObject:decoder];
    else {
      [_decoders addObject:decoder];
      [_delegate scannerFoundDecoder:decoder];
    }
  }
  else [_delegate scannerUpdatedDecoder:decoder];
  // The FEC guard time after CV7 (RSFEC is not known yet), and if the decoder repeats its answers (RSFEC > 0)
  if ((cvIndex == 0) || ([decoder getCv:RSFEC] > 0)) [self afterTime:FEC_GUARD_TIME action:NEXT_CV];
  else [self nextCv];
}

- (void)nextCv {
  cvIndex++;
  attempt = 0;
  rereads = 0;
  if (cvIndex < [cvList count]) {
    [self sendPomRead];
    return;
  }
  // All CVs have been read. The RS-Bus address(es) of this decoder are not scanned (again). For a decoder with a
  // given PoM address this includes its first RS-Bus address. The TT bits are not checked, since the RSbus library
  // sends "feedback decoder" messages by default (RSbus.h), for example for the lift decoder
  ScannedDecoderClass *decoder = [found objectAtIndex:decoderIndex];
  if ([decoder rsBusType] == ACCESSORY_DECODER) {
    int last = [decoder rsAddress] + [decoder rsAddressCount] - 1;
    for (ScannedDecoderClass *other in found) {
      if ((other != decoder) && ([other rsStation] == [decoder rsStation]) &&
          ([other rsAddress] >= [decoder rsAddress]) && ([other rsAddress] <= last)) [other setSkip:YES];
    }
  }
  if ([decoder givenPomAddress] > 0) [self addLiftIr:decoder];
  [self nextDecoder];
}

// The IR sensors of the lift have their own board (Lift_IR), without DCC and thus without PoM. Their RS-Bus
// addresses are given in the preferences. If they are on the RS-Bus of the lift decoder, they get their own row,
// with the PoM address and CVs of the lift decoder
- (void)addLiftIr:(ScannedDecoderClass *)lift {
  if (liftIrAddress <= 0) return;
  BOOL onRsBus = NO;
  for (ScannedDecoderClass *other in found) {
    if (([other rsStation] == [lift rsStation]) && ([other rsAddress] >= liftIrAddress) && ([other rsAddress] <= liftIrAddress + 1)) {
      [other setSkip:YES];
      onRsBus = YES;
    }
  }
  if (!onRsBus) {
    [problems addObject:[NSString stringWithFormat:@"IR sensors of the lift not found on RS-Bus address %d", liftIrAddress]];
    return;
  }
  ScannedDecoderClass *ir = [[ScannedDecoderClass alloc] init];
  [ir setLiftIr:YES];
  [ir setLiftLevel:lift];                       // Not retained: both are in liftRows
  [ir setRsStation:[lift rsStation]];
  [ir setRsAddress:liftIrAddress];
  [ir setPomStation:[lift pomStation]];
  [ir setPomIpAddress:[lift pomIpAddress]];
  [liftRows insertObject:ir atIndex:[liftRows indexOfObject:lift]];   // The lower RS-Bus addresses first
  [ir release];
}

- (void)pomTimeout {
  timer = nil;
  waitingForAnswer = NO;
  attempt++;
  if (attempt < MAX_ATTEMPTS) {
    [self afterTime:QUIET_TIME action:RESEND];
    return;
  }
  attempt = 0;
  rereads = 0;
  if (cvIndex == 0) {
    // No answer on CV7 via this station: try the next one. Without any answer the decoder is omitted
    candidateIndex++;
    if (candidateIndex < [candidates count]) [self afterTime:QUIET_TIME action:RESEND];
    else {
      ScannedDecoderClass *decoder = [found objectAtIndex:decoderIndex];
      if ([decoder givenPomAddress] > 0) [noAnswer addObject:[NSString stringWithFormat:@"PoM %d", [decoder givenPomAddress]]];
      else [noAnswer addObject:[NSString stringWithFormat:@"RS-Bus %d (%@)", [decoder rsAddress], [[self stationWithNumber:[decoder rsStation]] ipAddress]]];
      [self afterTime:QUIET_TIME action:NEXT_DECODER];
    }
    return;
  }
  // No answer for this CV: the value stays unknown ("?")
  [_delegate scannerUpdatedDecoder:[found objectAtIndex:decoderIndex]];
  [self afterTime:QUIET_TIME action:NEXT_CV];
}

// The next step after the quiet time or the FEC guard time. In the meantime all answers are ignored
- (void)afterTime:(NSTimeInterval)seconds action:(pomAction_t)action {
  pendingAction = action;
  [self startTimer:seconds selector:@selector(doPendingAction)];
}

- (void)doPendingAction {
  timer = nil;
  if (phase != POM_READ) return;
  switch (pendingAction) {
    case RESEND:       [self sendPomRead]; break;
    case NEXT_CV:      [self nextCv]; break;
    case NEXT_DECODER: [self nextDecoder]; break;
  }
}


// ************************************************************************************************************
// ******************************************** EVENTS OF THE STATIONS ****************************************
// ************************************************************************************************************
- (void)station:(StationClass *)station rsAddress:(int)rsAddress nibble:(int)nibble data:(int)data ttBits:(int)ttBits {
  if (phase == RSBUS_DATA) [self rsBusAnswer:station rsAddress:rsAddress ttBits:ttBits];
  else if (phase == POM_READ) [self pomAnswer:station rsAddress:rsAddress nibble:nibble data:data];
}

// In step 2 the library has emptied its queues: the RS-Bus data of this station is incomplete.
// In step 3 the timeout handles the missing answer
- (void)station:(StationClass *)station sendError:(NSString *)text {
  if ((phase == RSBUS_DATA) && (station == [stations objectAtIndex:rsStationIndex])) {
    [problems addObject:[NSString stringWithFormat:@"Command station %d (%@): %@", [station number], [station ipAddress], text]];
    [self nextRsStation];
  }
}

@end
