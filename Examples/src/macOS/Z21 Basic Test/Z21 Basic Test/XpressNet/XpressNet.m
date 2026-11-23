//
//  XpressNet.m
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 XpressNet DLL (XpressNetClass). Version 2.0, November 2026. See XpressNet.h.
//  The XpressNet object also implements the delegate methods of the TCP, UDP and serial connections
//  (VB6: the forms with the WinSock and MSComm controls, and the classes Tcp, Udp and Usb).
//

#import "XpressNet.h"
#import "XpressNetSend.h"
#import "XpressNetReceive.h"
#import "XpressNetReceiveZ21.h"
#import "XpressNetAccAuto.h"
#import "XpressNetLogFile.h"
#import "TCPSocket.h"
#import "UDPSocket.h"
#import "SerialPort.h"

@interface XpressNetClass () <TCPSocketDelegate, UDPSocketDelegate, SerialPortDelegate> {
  NSString *_instanceId;                  // 2.0: the letter of this object in the logfile (see init)
  NSString *_connectionPort;              // 2.0: the port of the TCP or UDP connection, for the logfile
}
@end

@implementation XpressNetClass

@synthesize delegate                              = _delegate;
@synthesize Z21Protocol                           = _Z21Protocol;
@synthesize LogMessages                           = _LogMessages;
@synthesize LogWarnings                           = _LogWarnings;
@synthesize LogErrors                             = _LogErrors;
@synthesize LogHexMessages                        = _LogHexMessages;
@synthesize LogDetails                            = _LogDetails;
@synthesize AccessoryInterval                     = _AccessoryInterval;
@synthesize ZentraleNotAus                        = _ZentraleNotAus;
@synthesize ZentraleNotHalt                       = _ZentraleNotHalt;
@synthesize ZentraleProgrammierMode               = _ZentraleProgrammierMode;
@synthesize ZentraleVersion                       = _ZentraleVersion;
@synthesize ZentraleKennung                       = _ZentraleKennung;
@synthesize ZentraleKurzSchluss                   = _ZentraleKurzSchluss;
@synthesize ZentraleStartmodeAuto                 = _ZentraleStartmodeAuto;
@synthesize ZentraleKaltstart                     = _ZentraleKaltstart;
@synthesize ZentraleRamCheckFehler                = _ZentraleRamCheckFehler;
@synthesize NumberOfExpressNetMessagesReceived    = _NumberOfExpressNetMessagesReceived;
@synthesize NumberOfExpressNetMessagesTransmitted = _NumberOfExpressNetMessagesTransmitted;
@synthesize Logfile                               = _Logfile;
@synthesize send                                  = _send;
@synthesize receive                               = _receive;
@synthesize receiveZ21                            = _receiveZ21;
@synthesize tcp                                   = _tcp;
@synthesize udp                                   = _udp;
@synthesize usb                                   = _usb;
@synthesize AccAuto                               = _AccAuto;
@synthesize connectionType                        = _connectionType;
@synthesize zentraleVersionNumber                 = _zentraleVersionNumber;
@synthesize instanceId                            = _instanceId;

// ***************************************************************************************************************************************************
// ******************************************************* Local C type definitions and declarations *************************************************
// ***************************************************************************************************************************************************
#define CONNECT_TIMEOUT 5                 // Seconds to wait for a TCP connection
#define SERIAL_START_DELAY 0.1            // In case of a serial connection, it seems we must wait 100ms before the interface is stable
#define KENNUNG_Z21 0x12                  // Zentralen-Kennung of the Z21 / Yamorc
// 2.0: known values for the Zentralen-Kennung, for the logfile (V3.6: 3.1.3)
#define KENNUNG_LZ100 0x00                // LZ 100
#define KENNUNG_LH200 0x01                // LH 200
#define KENNUNG_DPC 0x02                  // DPC - Zentrale (Compact und Commander)
#define KENNUNG_CNTRPLUS 0x03             // Control Plus - Zentrale
#define KENNUNG_LOKMAUS2 0x04             // Roco LokMaus II (source: JMRI XNetConstants)
#define KENNUNG_LDH 0x07                  // LDH (source: JMRI XNetConstants)
#define KENNUNG_MULTIMAUS 0x10            // Roco multiMAUS (source: JMRI XNetConstants)
#define KENNUNG_OPENDCC 0x20              // See: https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html
#define KENNUNG_OPENDCCXP 0x22

// 2.0: the letter of an XpressNet object in the logfile: A for the first object of the program, B for the second, and
// so on (after Z: AA, AB, ...). Two programs that write in the same logfile both start with A
static int instanceCounter = 0;

static NSString *letters(int n) {
  NSMutableString *s = [NSMutableString string];
  while (n > 0) {
    n = n - 1;
    [s insertString:[NSString stringWithFormat:@"%c", 'A' + (n % 26)] atIndex:0];
    n = n / 26;
  }
  return s;
}

// Some bit manipulation functions
static BOOL bitValue(int data, int index) {return ((data >> index) & 1) ? YES : NO;}

// The delegate implements only the events it needs
#define DELEGATE_RESPONDS(selectorName) ([_delegate respondsToSelector:@selector(selectorName)])


- (id)init {
  self = [super init];
  if (self) {
    instanceCounter = instanceCounter + 1;
    _instanceId = [letters(instanceCounter) copy];   // 2.0: first, since the logfile uses it
    _connectionType = none;               // We start without any connection
    _send = [[SendClass alloc] init];
    [_send setParent:self];               // Give the subclass a pointer to me, so it can call others
    _receive = [[ReceiveClass alloc] init];
    [_receive setParent:self];
    _receiveZ21 = [[ReceiveClassZ21 alloc] init];
    [_receiveZ21 setParent:self];
    _AccAuto = [[AccAutoClass alloc] init];
    [_AccAuto setParent:self];
    // Logfile specific initialization
    _Logfile = [[LogFileClass alloc] init];
    [_Logfile setParent:self];
    [_Logfile OpenLogfile];               // Check Logfile size, and rotate if needed
    _LogMessages = YES;                   // Set the defaults for what messages will be logged.
    _LogErrors = YES;
    _LogWarnings = YES;
    _LogHexMessages = YES;
    _LogDetails = YES;
    // Attributes the library user may change
    _Z21Protocol = NO;
    _AccessoryInterval = 125;             // Time in msec between two accessory commands
    // Attributes the library user may read
    _ZentraleVersion = [@"0.0" copy];     // Give initial version, to avoid problems if the interface is not open
    _zentraleVersionNumber = 0;
  }
  return self;
}


- (void)dealloc {
  [_Logfile Details:@"Close XpressNet interface"];
  // Before leaving, make sure we close any open connections
  [self closeConnections];
  [_tcp setDelegate:nil];
  [_udp setDelegate:nil];
  [_usb setDelegate:nil];
  [_tcp release];
  [_udp release];
  [_usb release];
  [_send setParent:nil];
  [_send release];
  [_receive release];
  [_receiveZ21 release];
  [_AccAuto release];
  [_Logfile release];
  [_ZentraleVersion release];
  [_instanceId release];
  [_connectionPort release];
  [super dealloc];
}


// ****************************************************************************************************************
// ********************************************* METHODS FOR ESTABLISHING CONNECTIONS *****************************
// ****************************************************************************************************************
- (void)CmdConnectTcp:(NSString *)Host {[self CmdConnectTcp:Host Port:@"5550"];}

- (void)CmdConnectTcp:(NSString *)Host Port:(NSString *)Port {
  [self closeConnections];
  [_Logfile setMyHost:Host];                              // Save the host (IP address) to include in the Logfile
  [_Logfile Details:[NSString stringWithFormat:@"TCP connection requested to Host: %@ at Port: %@", Host, Port]];
  _connectionType = lanTcpPort;                           // We connect via TCP over a LAN
  [self setConnectionPort:Port];                          // 2.0: for the logfile
  if (_tcp == nil) {
    _tcp = [[TCPSocketClass alloc] init];
    [_tcp setDelegate:self];
  }
  [_tcp openToHost:Host port:Port timeout:CONNECT_TIMEOUT];
}


- (void)CmdInitUdp:(NSString *)Host {[self CmdInitUdp:Host Port:@"21106"];}

- (void)CmdInitUdp:(NSString *)Host Port:(NSString *)Port {
  [self closeConnections];
  [_Logfile setMyHost:Host];
  [_Logfile Details:[NSString stringWithFormat:@"UDP initialised: Host: %@, Port: %@", Host, Port]];
  _connectionType = lanUdpPort;                           // We connect via UDP over a LAN
  if (_udp == nil) {
    _udp = [[UDPSocketClass alloc] init];
    [_udp setDelegate:self];
  }
  _Z21Protocol = YES;
  [self setConnectionPort:Port];                          // 2.0: for the logfile
  [_udp openToHost:Host port:Port];                       // Messages are kept until the connection is ready
  // 2.0: the first message is LAN_SET_BROADCASTFLAGS (Z21 spec 2.16). Without these flags the Z21 sends no
  // broadcasts, such as LAN_X_BC_TRACK_POWER_ON/OFF (events AllesAn / NotAus) and LAN_X_TURNOUT_INFO.
  // The flags must be set again after every login; the login is done with the first command (Z21 spec 2.2)
  [self CmdZ21SetBroadcastFlags:Z21BroadcastFlagsDefault];
  [self CmdSoftwareversionZentraleAnfordern];
  [self CmdStatusZentraleAnfordern];
  [self CmdZ21GetFirmwareVersion];
}


- (void)CmdConnectUsb:(NSString *)Port {[self CmdConnectUsb:Port Settings:@"115200,N,8,1"];}

- (void)CmdConnectUsb:(NSString *)Port Settings:(NSString *)Settings {[self CmdConnectUsb:Port Settings:Settings Handshake:NO];}

// 2.0: the parameter Handshake (hardware handshake RTS/CTS, default off). The Lenz documentation of the LI-USB and the
// 23151 says no handshake; JMRI uses hardware handshake (RTS/CTS) for the LI-USB 23150. See SerialPort.h
- (void)CmdConnectUsb:(NSString *)Port Settings:(NSString *)Settings Handshake:(BOOL)Handshake {
  // Lenz uses 57600, Yamorc 115200
  [self closeConnections];
  [_Logfile setMyHost:Port];                              // Save the port to include in the Logfile
  // [_Logfile Details:[NSString stringWithFormat:@"USB connect attempt using Port: %@", Port]];
  [_Logfile Details:[NSString stringWithFormat:@"USB connect attempt using Port=%@, Settings=%@ and Handshake=%@",
                     Port, Settings, Handshake ? @"True" : @"False"]];
  _connectionType = usbPort;                              // We connect via USB
  [self openSerialPort:Port settings:Settings handshake:Handshake];
}


- (void)CmdConnectSerial:(NSString *)Port Settings:(NSString *)Settings Handshake:(BOOL)Handshake {
  [self closeConnections];
  [_Logfile setMyHost:Port];
  [_Logfile Details:[NSString stringWithFormat:@"Serial connect attempt using Port=%@, Settings=%@ and Handshake=%@",
                     Port, Settings, Handshake ? @"True" : @"False"]];
  _connectionType = serialPort;                           // We connect via a serial (RS232) connection
  [self openSerialPort:Port settings:Settings handshake:Handshake];
}


// USB and serial use the same object
- (void)openSerialPort:(NSString *)Port settings:(NSString *)Settings handshake:(BOOL)Handshake {
  if (_usb == nil) {
    _usb = [[SerialPortClass alloc] init];
    [_usb setDelegate:self];
  }
  if ([_usb openPort:Port settings:Settings handshake:Handshake]) {
    // 2.0: which connection belongs to the letter of this XpressNet object
    [_Logfile Connection:[NSString stringWithFormat:@"%@, %@%@", Port, Settings, Handshake ? @", handshake" : @""]];
    [self eventInterfaceConnected];
  }
  else {
    _connectionType = none;
    [self eventTcpError:[NSString stringWithFormat:@"Could not open %@ (%@)", Port, [_usb lastError]]];
  }
}


// Close all connections, and bring the send and receive objects in a well-defined state
- (void)closeConnections {
  // 2.0: log off from the Z21 (Z21 spec 2.2), via the same socket (port) as used for the login
  if ((_connectionType == lanUdpPort) && [_udp isOpen]) {
    static const uint8_t logoff[4] = {0x04, 0x00, 0x30, 0x00};
    [_send transmitZ21:logoff length:4];
    [_Logfile DetailsSent:@"Z21: LAN_LOGOFF"];
  }
  _connectionType = none;
  [_tcp close];                                           // The resulting TCP_SOCKET_CLOSED (reason nil) is ignored
  [_udp close];
  [_usb close];
  [self resetFlowControl];
}


- (void)resetFlowControl {
  [_send stop];
  [_send emptyFifo];
  [_send setWaitForAck:NO];
  [_send setWaitForlanKeepAliveResponse:NO];
  [_receive emptyIndata];
  [_AccAuto stop];
  [NSObject cancelPreviousPerformRequestsWithTarget:self selector:@selector(requestZentraleInfo) object:nil];
}


// ****************************************************************************************************************
// ********************************************* METHODS TO GENERATE INTERFACE SPECIFIC MESSAGES ******************
// ****************************************************************************************************************
// V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
- (void)CmdInterfaceVersion {
  [_send message:1, 0xF0];
  [_Logfile DetailsSent:@"Versionsnummer des Interface 23151 feststellen"];
}

// V3.6: 1.7 (XpressNet Gerateadresse des Interface 23151 feststellen und andern)
- (void)CmdXpressNetAddress {[self CmdXpressNetAddress:0];}

- (void)CmdXpressNetAddress:(int)address {
  [_send message:3, 0xF2, 0x01, address];
  [_Logfile DetailsSent:@"XpressNet Gerateadresse des Interface 23151 feststellen"];
}

// V3.6: 2.1 (Interface Status Befehl - Keep Alive) - Need not be used by the Application
- (void)CmdInterfaceStatus {
  [_send message:2, 0xF1, 0x01];
}

// V3.6: 2.2 (XpressNet Version Befehl)
- (void)CmdXpressNetVersion {
  [_send message:2, 0xF1, 0x02];
  [_Logfile DetailsSent:@"XpressNet Version Befehl"];
}

// V3.6: 2.3 (Verfugbare Freie Verbindungen)
- (void)CmdInterfaceFreieVerbindungen {
  [_send message:2, 0xF1, 0x03];
  [_Logfile DetailsSent:@"Verfugbare Freie Verbindungen"];
}


// ****************************************************************************************************************
// ******************************** V3.6 & V4 METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ***************
// ****************************************************************************************************************
// The commands below will be forwarded by the Interface to the Command Station

// V3.6: 3.2.1 / V4: 3.2 (Alles An)
- (void)CmdAllesAn {
  [_send message:2, 0x21, 0x81];
  [_Logfile DetailsSent:@"Alles An"];
}

// V3.6: 3.2.2 / V4: 3.3 (Alles Aus - Notaus)
- (void)CmdAllesAus {
  [_send message:2, 0x21, 0x80];
  [_Logfile DetailsSent:@"Alles Aus - Notaus"];
}

// V3.6: 3.2.3 / V4: 3.4 (Alle Loks anhalten - Nothalt)
- (void)CmdAlleLoksAnhalten {
  [_send message:1, 0x80];
  [_Logfile DetailsSent:@"Alle Loks anhalten - Nothalt"];
}

// V3.6: 3.2.4 / V4: 3.6 & 3.7 (Eine Lok anhalten - Nothalt fur eine Lok)
- (void)CmdEineLokAnhalten:(int)address {
  [_send message:3, 0x92, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Eine lok anhalten. Adresse: %d", address]];
}

// V3.6: 3.2.5 / V4: 3.9
- (void)CmdLeseanfrageProgrammierenRegistermode:(int)Register {
  if ([self VelueOutOfRange:Register min:1 max:8 LogText:@"Leseanfrage Programmieren Registermode: Register"]) return;
  [_send message:3, 0x22, 0x11, Register];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Leseanfrage Programmieren Registermode. Register: %d", Register]];
}

// V3.6: 3.2.6-3.2.10 / V4: 3.10-3.14 (Leseanfrage Programmieren 4-Byte-Format / LAN_X_CV_READ)
- (void)CmdLeseanfrageProgrammieren:(int)CV {
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"Leseanfrage Programmieren: CV"]) return;
  if (_Z21Protocol) {
    CV = CV - 1;
    [_send message:4, 0x23, 0x11, (CV / 256), (CV % 256)];
    [_Logfile DetailsSent:[NSString stringWithFormat:@"LAN_X_CV_READ, CV: %d", CV + 1]];
  }
  else {
    if (_zentraleVersionNumber < 36) [_send message:3, 0x22, 0x15, (CV % 256)];
    else [_send message:3, 0x22, 0x18 + ((CV / 256) % 4), (CV % 256)];
    [_Logfile DetailsSent:[NSString stringWithFormat:@"Leseanfrage Programmieren 4-Byte-Format. CV: %d", CV]];
  }
}

// V3.6: 3.2.11 / V4: 3.15 (Leseanfrage Programmieren 3-Byte-Format - Pagemode)
- (void)CmdLeseanfrageProgrammierenPagemode:(int)CV {
  if ([self VelueOutOfRange:CV min:1 max:256 LogText:@"Leseanfrage Programmieren Pagemode: CV"]) return;
  [_send message:3, 0x22, 0x14, (CV % 256)];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Leseanfrage Programmieren Pagemode. CV: %d", CV]];
}

// V3.6: 3.2.12 / V4: 3.16
- (void)CmdProgrammierergebnisAnfordern {
  [_send message:2, 0x21, 0x10];
  [_Logfile DetailsSent:@"Programmier Ergebnis Anfordern"];
}

// V3.6: 3.2.13 / V4: 3.17 (Schreibbefehl Programmieren 3-Byte-Format (Register-Mode) / LAN_X_DCC_WRITE_REGISTER)
- (void)CmdSchreibbefehlProgrammierenRegisterMode:(int)Register daten:(int)daten {
  if ([self VelueOutOfRange:Register min:1 max:8 LogText:@"Schreibbefehl Programmieren Registermode: Register"]) return;
  [_send message:4, 0x23, 0x12, Register, daten];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Schreibbefehl Programmieren RegisterMode. Register: %d, Value: %d", Register, daten]];
}

// V3.6: 3.2.14-3.2.18 / V4: 3.18-3.22 (Schreibbefehl Programmieren 4-Byte-Format / LAN_X_CV_WRITE)
- (void)CmdSchreibbefehlProgrammieren:(int)CV Val:(int)Val {
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"Schreibbefehl Programmieren: CV"]) return;
  if ([self VelueOutOfRange:Val min:0 max:255 LogText:@"Schreibbefehl Programmieren: Daten: "]) return;
  if (_Z21Protocol) {
    CV = CV - 1;
    [_send message:5, 0x24, 0x12, (CV / 256), (CV % 256), Val];
    [_Logfile DetailsSent:[NSString stringWithFormat:@"LAN_X_CV_WRITE, CV: %d, Value: %d", CV + 1, Val]];
  }
  else {
    if (_zentraleVersionNumber < 36) [_send message:4, 0x23, 0x16, (CV % 256), Val];
    else [_send message:4, 0x23, 0x1C + ((CV / 256) % 4), (CV % 256), Val];
    [_Logfile DetailsSent:[NSString stringWithFormat:@"Schreibbefehl Programmieren 4-Byte-Format. CV: %d, Value: %d", CV, Val]];
  }
}

// V3.6: 3.2.19 / V4: 3.23 (Schreibbefehl Programmieren 3-Byte-Format - Page-Mode)
- (void)CmdSchreibbefehlProgrammierenPageMode:(int)CV daten:(int)daten {
  if ([self VelueOutOfRange:CV min:1 max:256 LogText:@"Schreibbefehl Programmieren Pagemode: CV"]) return;
  [_send message:4, 0x23, 0x17, (CV % 256), daten];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Schreibbefehl Programmieren Pagemode. CV: %d, Value: %d", CV, daten]];
}

// V3.6: 3.2.20 / V4: 3.24 (Softwareversion der Zentrale anfordern)
- (void)CmdSoftwareversionZentraleAnfordern {
  [_send message:2, 0x21, 0x21];
  [_Logfile DetailsSent:@"Softwareversion der Zentrale anfordern"];
}

// V3.6: 3.2.21 / V4: 3.26 (Status der Zentrale anfordern)
- (void)CmdStatusZentraleAnfordern {
  [_send message:2, 0x21, 0x24];
  // This command may also be internally used as UDP (Z21) Keep Alive Message
  // In that case we should not log this like commands requested by the application
  if (![_send waitForlanKeepAliveResponse]) [_Logfile DetailsSent:@"Status der Zentrale anfordern"];
}

// V3.6: 3.2.22 / V4: 3.27 (Zentralen-Startmode setzen)
- (void)CmdZentralenStartmodeSetzen:(BOOL)AutoStart {
  [_send message:3, 0x22, 0x22, AutoStart ? 2 : 0];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Zentralen-Startmode setzen. Automatischer Start: %@", AutoStart ? @"True" : @"False"]];
}

// V3.6: 3.2.23 / V4: 3.37 & 3.38 / Z21: 5.1 (Schaltinformationen anfordern / LAN_X_GET_TURNOUT_INFO)
// Note: this command is intended for switches.
// The (switch) address gets divided by 4 to obtain the decoder address
- (void)CmdSwitchInfoRequest:(int)address {
  // STEP 1: Check the address size against the ZentraleVersion and Protocol.
  if ((_zentraleVersionNumber < 38) && (!_Z21Protocol)) {
    if ([self VelueOutOfRange:address min:1 max:1024 LogText:@"Accessory: Address"]) return;
  }
  if ([self VelueOutOfRange:address min:1 max:2048 LogText:@"Accessory: Address"]) return;
  // Switch numbering on the wire starts at 0. The LH100 numbering starts at 1
  int wireAddress = address - 1;
  // We need to determine the nibble for this specific switch: 0 and 1 become 0; 2 and 3 become 1
  int nibble = ((wireAddress % 4) > 1) ? 1 : 0;
  // Send the command that fits for this ZentraleVersion and Protocol.
  if (_Z21Protocol) {
    // Note: the XpressNet length, as contained in the four low order bits of the header,
    // seem to be incorrect! However, this is what the Z21 spec says...
    [_send message:3, 0x43, [self AccAH:wireAddress], [self AccAL:wireAddress]];
  }
  else {
    // We should set the highest bit of the byte that contains the nibble
    if (_zentraleVersionNumber < 38) [_send message:3, 0x42, [self AccAL:wireAddress], nibble + 128];
    else [_send message:4, 0x43, [self AccAH:wireAddress], [self AccAL:wireAddress], nibble + 128];   // Lenz V4: Command station >= 3.8
  }
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Feedback Request - Switch: %d", address]];
}

// V3.6: 3.2.23 / V4: 3.37 & 3.38 (Schaltinformationen anfordern)
// This command is intended for feedback (RS-Bus) decoders
- (void)CmdFeedbackRequest:(int)address Nibble:(int)Nibble {
  if ([self VelueOutOfRange:address min:1 max:128 LogText:@"RS-Bus Feedback: Address"]) return;
  if (Nibble != 0) Nibble = 1;                            // Ensure the nibble is 0 or 1
  // Substract 1 from the address. The LH100 numbering starts at 1, the wire at 0. Also set the highest bit of the nibble
  [_send message:3, 0x42, address - 1, Nibble + 128];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Feedback Request for address: %d  Nibble: %d", address, Nibble]];
}

// V3.6: 3.2.24 / V4: 3.39 & 3.40 / Z21: 5.2 (Schaltbefehl / LAN_X_SET_TURNOUT)
- (void)CmdAccessory:(int)address position:(int)position Activate:(int)Activate {
  // 2.0: comment corrected. Position 0 = diverging (red, -), 1 = straight (green, +), according to RCN-213. Some
  // command stations interpret 0 and 1 the other way around (see the README of the AP_DCC_library)
  // Was: Position 0 = straight / 1 = turn
  // Check the address size against the ZentraleVersion.
  if (_zentraleVersionNumber < 38) {
    if ([self VelueOutOfRange:address min:1 max:1024 LogText:@"Accessory: Address"]) return;
  }
  else {
    if ([self VelueOutOfRange:address min:1 max:2048 LogText:@"Accessory: Address"]) return;
  }
  address = address - 1;                                  // The LH100 numbering starts at 1, the wire at 0
  int daten = 0;
  if (!_Z21Protocol) daten = (address % 4) * 2;           // The Z21 uses switch addressing, instead of decoder addressing
  daten = 128 + daten;                                    // Set the MSB
  if (Activate > 0) daten = daten + 8;                    // Set the D1 bit (Activation)
  if (position > 0) daten = daten + 1;                    // Set the D2 bit (Position)
  // Send the command that fits for this ZentraleVersion and Kennung.
  // The first form is used for the LZV 200 and the Z21, the second form for the LZV100
  if ((_zentraleVersionNumber > 36) || _Z21Protocol) [_send message:4, 0x53, [self AccAH:address], [self AccAL:address], daten];
  else [_send message:3, 0x52, [self AccAL:address], daten];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Accessory Command - Address: %d  Position: %d  Activate: %d", address + 1, position, Activate]];
}

// V3.6: 3.2.24 / V4: 3.39 & 3.40 (Schaltbefehl)
- (void)CmdAccessoryAuto:(int)Switch position:(int)position {
  // Implemented in the AccAutoClass, Since it is not a standard XpressNet command.
  [_AccAuto handle:Switch position:position];
}

// V3.6: 3.2.25 / V4: 3.41.3 / Z21: 4.1 (Lokinformationen anfordern)
- (void)CmdLokInfoAnfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Lok Info Anfordern: Address"]) return;
  if (_Z21Protocol) [_send message:4, 0xE3, 0xF0, [self AH:address], [self AL:address]];
  else [_send message:4, 0xE3, 0x00, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Lok Informationen anfordern. Address: %d", address]];
}

// V3.6: 3.2.25.1 / V4: 3.41.4 (Funktionsstatus anfordern F0 - F12)
- (void)CmdFunktionsStatusAnfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Funktionsstatus Anfordern: Address"]) return;
  [_send message:4, 0xE3, 0x07, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktionsstatus F0-F12 anfordern. Address: %d", address]];
}

// V3.6: 3.2.25.2 / V4: 3.41.5 (Funktionsstatus anfordern F13 - F28)
- (void)CmdFunktionsStatusF13F28Anfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Funktionsstatus F13-F28 Anfordern: Address"]) return;
  [_send message:4, 0xE3, 0x08, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktionsstatus F13-F28 anfordern. Address: %d", address]];
}

// V3.6: 3.2.25.3 / V4: 3.41.7 (Funktionszustand anfordern F13 - F28)
- (void)CmdFunktionsZustandF13F28Anfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Funktionszustand F13-F28 Anfordern: Address"]) return;
  [_send message:4, 0xE3, 0x09, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktionszustand F13-F28 anfordern. Address: %d", address]];
}

// V4: 3.41.6 (Funktionsstatus anfordern F29 - F68)
- (void)CmdFunktionsStatusF29F68Anfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Funktionsstatus F29-F68 Anfordern: Address"]) return;
  [_send message:4, 0xE3, 0x0A, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktionsstatus F29-F68 anfordern. Address: %d", address]];
}

// V4: 3.41.8 (Funktionszustand anfordern F29 - F68)
- (void)CmdFunktionsZustandF29F68Anfordern:(int)address {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Funktionszustand F29-F68 Anfordern: Address"]) return;
  [_send message:4, 0xE3, 0x0B, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktionszustand F29-F68 anfordern. Address: %d", address]];
}

// V3.6: 3.2.26.1 / V4: 3.42.3 / Z21: 4.2 (Lok Fahrbefehl / LAN_X_SET_LOCO_DRIVE)
- (void)CmdFahrbefehl:(int)address Speed:(int)Speed Steps:(int)Steps Forward:(BOOL)Forward {
  int RVVV = 0;                                           // The direction and speed byte
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"Fahrbefehl: Address"]) return;
  switch (Steps) {
    case 14:
      if ([self VelueOutOfRange:Speed min:0 max:14 LogText:@"Fahrbefehl: Speed"]) return;
      if (Speed > 0) RVVV = Speed + 1;
      if (Forward) RVVV = RVVV + 128;
      [_send message:5, 0xE4, 0x10, [self AH:address], [self AL:address], RVVV];
      break;
    case 27:
      if ([self VelueOutOfRange:Speed min:0 max:27 LogText:@"Fahrbefehl: Speed"]) return;
      RVVV = [self EncodeSpeed:Speed];
      if (Forward) RVVV = RVVV + 128;                     // The forward / reverse bit
      [_send message:5, 0xE4, 0x11, [self AH:address], [self AL:address], RVVV];
      break;
    case 28:
      if ([self VelueOutOfRange:Speed min:0 max:28 LogText:@"Fahrbefehl: Speed"]) return;
      RVVV = [self EncodeSpeed:Speed];
      if (Forward) RVVV = RVVV + 128;
      [_send message:5, 0xE4, 0x12, [self AH:address], [self AL:address], RVVV];
      break;
    case 128:
      if ([self VelueOutOfRange:Speed min:0 max:126 LogText:@"Fahrbefehl: Speed"]) return;
      RVVV = Speed;
      if (Speed > 0) RVVV = RVVV + 1;
      if (Forward) RVVV = RVVV + 128;
      [_send message:5, 0xE4, 0x13, [self AH:address], [self AL:address], RVVV];
      break;
    default:                                              // 2.0: not silently (1.2: nothing was sent, without an error)
      [_Logfile AppendWithHost:@"Error Fahrbefehl: Steps must be 14, 27, 28 or 128"];
      [self eventSendError:FehlerInDenBefehlsparametern description:[NSString stringWithFormat:@"Message not sent: Fahrbefehl: Steps must be 14, 27, 28 or 128 (value: %d)", Steps]];
      return;
  }
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Fahrbefehl - Address: %d, Speed: %d", address, Speed]];
}

// V3.6: 3.2.26.2 / V4: 3.42.4 (Lok Funktionsbefehl)
- (void)CmdLokFunktionBefehl:(int)address group:(int)group data:(int)data {
  static const int kennung[11] = {0, 0x20, 0x21, 0x22, 0x23, 0x28, 0x29, 0x2A, 0x2B, 0x50, 0x51};
  // 2.0: a group outside 1..10 gives the event SendError (1.2: nothing was sent, without an error)
  if ([self VelueOutOfRange:group min:1 max:10 LogText:@"Funktions befehl: Group"]) return;
  [_send message:5, 0xE4, kennung[group], [self AH:address], [self AL:address], data];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Funktions befehl. Address: %d, Befehl: %d", address, data]];
}

// V3.6: 3.2.26.4 / V4: 3.42.6 (Funktionsstatus setzen)
- (void)CmdLokFunktionStatus:(int)address group:(int)group data:(int)data {
  static const int kennung[11] = {0, 0x24, 0x25, 0x26, 0x27, 0x2C, 0x2D, 0x2E, 0x52, 0x53, 0x54};
  // 2.0: a group outside 1..10 gives the event SendError (1.2: nothing was sent, without an error)
  if ([self VelueOutOfRange:group min:1 max:10 LogText:@"Status befehl: Group"]) return;
  [_send message:5, 0xE4, kennung[group], [self AH:address], [self AL:address], data];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Status befehl. Address: %d, Status: %d", address, data]];
}

// V3.6: 3.2.28.1 / V4: 3.43.1 (Programming on Main Byte schreiben / LAN_X_CV_POM_WRITE_BYTE)
- (void)CmdPomWrite:(int)address CV:(int)CV value:(int)value {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"PoM write: Address"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"PoM write: CV"]) return;
  CV = CV - 1;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xEC;                           // 0xEC = PoM Schreiben
  [_send message:7, 0xE6, 0x30, [self AH:address], [self AL:address], cvhigh, cvlow, value];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Pom Write. Address: %d, CV: %d, Value: %d", address, CV + 1, value]];
}

// V3.6: 3.2.28.2 / V4: 3.43.2 (Programming on Main Byte lesen / LAN_X_CV_POM_READ_BYTE)
- (void)CmdPomRead:(int)address CV:(int)CV {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"PoM Read: Address"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"PoM Read: CV"]) return;
  CV = CV - 1;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xE4;                           // 0xE4 = PoM Lesen
  [_send message:7, 0xE6, 0x30, [self AH:address], [self AL:address], cvhigh, cvlow, 0];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Pom Read. Address: %d, CV: %d", address, CV + 1]];
}

// V3.6: 3.2.28.3 / V4: 3.43.3 (Programming on Main Bit schreiben / LAN_X_CV_POM_WRITE_BIT)
- (void)CmdPomBitWrite:(int)address CV:(int)CV bitValue:(BOOL)bitValue bitPosition:(int)bitPosition {
  if ([self VelueOutOfRange:address min:0 max:9999 LogText:@"PoM bit write: Address"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"PoM bit write: CV"]) return;
  if ([self VelueOutOfRange:bitPosition min:0 max:7 LogText:@"PoM bit write: bit position"]) return;
  CV = CV - 1;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xE8;                           // 0xE8 = PoM Bit Schreiben (V3.6: 3.2.28.3, Z21: 6.7)
  // 2026: VB6 used 0x7C for XpressNet, as in the V4 spec (3.43.3). The V4 spec seems to be wrong here
  int value = bitPosition + 0xF0;                         // highest order bits must all be 1
  if (bitValue) value = value + 8;                        // bit 4 represents the new bit value
  [_send message:7, 0xE6, 0x30, [self AH:address], [self AL:address], cvhigh, cvlow, value];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Pom Bit Write. Address: %d, CV: %d, Bit value: %@, Bit Position: %d",
                     address, CV + 1, bitValue ? @"True" : @"False", bitPosition]];
}

// V3.6 3.2.30.3 / V4: 3.44.3 (Adressanfrage Lok in Zentralenstack)
- (void)CmdAdresAnfrageLokInStack:(int)address Vorwarts:(BOOL)Vorwarts {
  [_send message:4, 0xE3, Vorwarts ? 5 : 6, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Adressanfrage Lok in Zentralenstack. Address: %d, Vorwarts : %@", address, Vorwarts ? @"True" : @"False"]];
}

// V3.6: 3.2.31 / V4: 3.44.4 (Lok aus Stack loeschen / LAN_X_PURGE_LOCO)
- (void)CmdLokAusStack:(int)address {
  [_send message:4, 0xE3, 0x44, [self AH:address], [self AL:address]];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Lok aus Stack loeschen. Address: %d", address]];
}


// ****************************************************************************************************************
// ******************************** V4 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES *****************
// ****************************************************************************************************************
// V4: 3.25 (Erweiterte Zentralen Versionsinformation)
- (void)CmdErweiterteZentralenVersionsinformation {
  [_send message:2, 0x21, 0x23];
  [_Logfile DetailsSent:@"Erweiterte Zentralen Versionsinformation"];
}

// V4: 3.30 (PoM-Ergebnis anfordern)
- (void)CmdPoMErgebnisAnfordern {
  [_send message:2, 0x21, 0x27];
  [_Logfile DetailsSent:@"PoM Programmier Ergebnis Anfordern"];
}

// V4: 3.31 (Zentralen Reset auslosen)
- (void)CmdZentralenResetAuslosen {
  [_send message:2, 0x21, 0x28];
  [_Logfile DetailsSent:@"Zentralen Reset auslosen"];
}

// V4: 3.33 (Modellzeit anfordern)
- (void)CmdModellzeitAnfordern {
  [_send message:2, 0x21, 0x2A];
  [_Logfile DetailsSent:@"Modellzeit anfordern"];
}

// V4: 3.34 (Modellzeit stellen)
- (void)CmdModellzeitStellen:(int)day hour:(int)hour minute:(int)minute factor:(int)factor {
  if ([self VelueOutOfRange:day min:0 max:6 LogText:@"Modellzeit: Day"]) return;
  if ([self VelueOutOfRange:hour min:0 max:23 LogText:@"Modellzeit: Hour"]) return;
  if ([self VelueOutOfRange:minute min:0 max:59 LogText:@"Modellzeit: Minute"]) return;
  if ([self VelueOutOfRange:factor min:0 max:31 LogText:@"Modellzeit: Factor"]) return;
  int daten = day * 32 + hour;
  [_send message:4, 0x24, 0x2B, daten, minute, factor];
  [_Logfile DetailsSent:@"Modellzeit stellen"];
}

// V4: 3.35 (Modellzeit anhalten - V4.0)
- (void)CmdModellzeitAnhalten {
  [_send message:2, 0x21, 0x2D];
  [_Logfile DetailsSent:@"Modellzeit anhalten"];
}

// V4: 3.36 (Modellzeit starten - V4.0)
- (void)CmdModellzeitStarten {
  [_send message:2, 0x21, 0x2C];
  [_Logfile DetailsSent:@"Modellzeit starten"];
}


// ****************************************************************************************************************
// ******************************** Z21 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ****************
// ****************************************************************************************************************
// Note: these Z21 commands could not be tested, and may therefore not work as intended

// Z21: 2.16 (LAN_SET_BROADCASTFLAGS)
// 2.0: new. The flags determine which broadcasts the Z21 sends to this client, for example:
// 0x00000001: driving and switching (track power, programming mode, short circuit, stop, LAN_X_TURNOUT_INFO,
//             LAN_X_LOCO_INFO for the locomotives requested by this client)
// 0x00000002: R-Bus feedback (LAN_RMBUS_DATACHANGED)    - not decoded by the library
// 0x08000000: LocoNet detector (LAN_LOCONET_DETECTOR)   - not decoded by the library
// For the other flags, see the Z21 specification. CmdInitUdp sets the flags to Z21BroadcastFlagsDefault.
// This is not an XpressNet message: DataLen 0x08 0x00, Header 0x50 0x00, Flags (32 bits, little endian)
- (void)CmdZ21SetBroadcastFlags:(int)Flags {
  uint8_t z21Message[8];
  z21Message[0] = 0x08;
  z21Message[1] = 0x00;
  z21Message[2] = 0x50;
  z21Message[3] = 0x00;
  z21Message[4] = (uint8_t)(Flags & 0xFF);
  z21Message[5] = (uint8_t)((Flags >> 8) & 0xFF);
  z21Message[6] = (uint8_t)((Flags >> 16) & 0xFF);
  z21Message[7] = (uint8_t)((Flags >> 24) & 0xFF);
  [_send transmitZ21:z21Message length:8];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Z21: CmdZ21SetBroadcastFlags. Flags: 0x%08X", (unsigned int)Flags]];
}

// Z21: 2.15 (LAN_X_GET_FIRMWARE_VERSION)
- (void)CmdZ21GetFirmwareVersion {
  [_send message:2, 0xF1, 0x0A];
  [_Logfile DetailsSent:@"Z21: CmdZ21GetFirmwareVersion"];
}

// Z21: 4.3.1 (LAN_X_SET_LOCO_FUNCTION)
- (void)CmdZ21SetLocoFunction:(int)address CV:(int)CV TT:(int)TT {
  // TT Umschalttyp: 00=aus, 01=ein, 10=umschalten,11=nicht erlaubt
  // CV: 0x00=F0 (Licht), 0x01=F1 usw. bis F31
  if ([self VelueOutOfRange:CV min:0 max:31 LogText:@"LAN_X_SET_LOCO_FUNCTION: CV"]) return;
  if ([self VelueOutOfRange:TT min:0 max:2 LogText:@"LAN_X_SET_LOCO_FUNCTION: TT"]) return;
  int TTNNNNNN = TT * 64 + CV;
  [_send message:5, 0xE4, 0xF8, [self AH:address], [self AL:address], TTNNNNNN];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Z21: CmdZ21SetLocoFunction. CV: %d, TT: %d", CV, TT]];
}

// Z21: 5.4 LAN_X_SET_EXT_ACCESSORY
- (void)CmdZ21SetExtAccessory:(int)address value:(int)value {
  if ([self VelueOutOfRange:address min:1 max:2048 LogText:@"Extended Accessory: Address"]) return;
  // The first "user address" is 1, which is represented as 4 in the Xpressnet message
  int rawAddress = address + 3;
  if (rawAddress >= 2048) rawAddress = rawAddress - 2048;
  [_send message:5, 0x54, (rawAddress / 256), (rawAddress % 256), value, 0x00];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Z21: CmdZ21SetExtAccessory. Address: %d, Value: %d", address, value]];
}

// Z21: 5.5 LAN_X_GET_EXT_ACCESSORY_INFO
- (void)CmdZ21GetExtAccessory:(int)address {
  if ([self VelueOutOfRange:address min:1 max:2048 LogText:@"Extended Accessory: Address"]) return;
  // The first "user address" is 1, which is represented as 4 in the Xpressnet message
  int rawAddress = address + 3;
  if (rawAddress >= 2048) rawAddress = rawAddress - 2048;
  // 2026: as in the Z21 spec (DataLen 0x09): X-Header 0x44, Adr_MSB, Adr_LSB, 0x00. VB6 added an extra 0x00 byte
  [_send message:4, 0x44, (rawAddress / 256), (rawAddress % 256), 0x00];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Z21: CmdZ21GetExtAccessory. Address: %d", address]];
}

// Z21: 6.9 (LAN_X_CV_POM_ACCESSORY_WRITE_BYTE)
- (void)CmdZ21AccPomWrite:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV value:(int)value {
  if ([self VelueOutOfRange:decoderAddress min:0 max:511 LogText:@"Accessory PoM write: Decoder Address"]) return;
  if ([self VelueOutOfRange:CDDD min:0 max:15 LogText:@"Accessory PoM write: CDDD"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"Accessory PoM write: CV"]) return;
  CV = CV - 1;
  // Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
  // Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
  // Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
  int aaaaa = decoderAddress / 16;
  int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xEC;                           // 0xEC = Acc PoM Schreiben
  [_send message:7, 0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, value];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Accessory PoM Write. Decoder address: %d, CDDD: %d, CV: %d, Value: %d", decoderAddress, CDDD, CV + 1, value]];
}

// Z21: 6.10 (LAN_X_CV_POM_ACCESSORY_WRITE_BIT)
- (void)CmdZ21AccPomBitWrite:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV bitValue:(BOOL)bitValue bitPosition:(int)bitPosition {
  if ([self VelueOutOfRange:decoderAddress min:0 max:511 LogText:@"Accessory PoM Bit write: Decoder Address"]) return;
  if ([self VelueOutOfRange:CDDD min:0 max:15 LogText:@"Accessory PoM Bit write: CDDD"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"PoM Bit write: CV"]) return;
  CV = CV - 1;
  int aaaaa = decoderAddress / 16;
  int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xE8;                           // 0xE8 = Acc PoM Bit Schreiben
  // DB5: 0000VPPP
  int VPPP = bitPosition;
  if (bitValue) VPPP = VPPP + 8;                          // bit 4 represents the new bit value
  [_send message:7, 0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, VPPP];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Accessory PoM Bit Write. Decoder address: %d, CDDD: %d, CV: %d, bitValue: %@, bitPosition: %d",
                     decoderAddress, CDDD, CV + 1, bitValue ? @"True" : @"False", bitPosition]];
}

// Z21: 6.11 (LAN_X_CV_POM_ACCESSORY_READ_BYTE)
- (void)CmdZ21AccPomRead:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV {
  if ([self VelueOutOfRange:decoderAddress min:0 max:511 LogText:@"Accessory PoM read: Decoder Address"]) return;
  if ([self VelueOutOfRange:CDDD min:0 max:15 LogText:@"Accessory PoM read: CDDD"]) return;
  if ([self VelueOutOfRange:CV min:1 max:1024 LogText:@"Accessory PoM read: CV"]) return;
  CV = CV - 1;
  int aaaaa = decoderAddress / 16;
  int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
  int cvlow = CV % 256;                                   // 2026: VB6 had CV Mod 255
  int cvhigh = CV / 256 + 0xE4;                           // 0xE4 = Acc PoM Lesen
  [_send message:7, 0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, 0x00];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"Accessory PoM Read. Decoder address: %d, CDDD: %d, CV: %d", decoderAddress, CDDD, CV + 1]];
}

// Z21: 6.12 (LAN_X_MM_WRITE_BYTE)
- (void)CmdZ21MmWriteByte:(int)RegAdr value:(int)value {
  [_send message:5, 0x24, 0xFF, 0x00, RegAdr, value];
  [_Logfile DetailsSent:@"Z21: CmdZ21MmWriteByte"];
}


// ****************************************************************************************************************
// ******************************** OpenDCC ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ************
// ****************************************************************************************************************
// Note: these commands could not be tested, and may therefore not work as intended
- (void)CmdOpenDCCSetExtAccessory:(int)address value:(int)value {
  if ([self VelueOutOfRange:address min:1 max:2048 LogText:@"Ext. Accessory: Address"]) return;
  if ([self VelueOutOfRange:value min:0 max:31 LogText:@"Ext. Accessory: Value"]) return;
  int byte1 = (value * 8) + (address / 256);              // Shift value 3 bits left, plus the 3 highest address bits
  int byte2 = address % 256;                              // 8 lowest address bits
  [_send message:4, 0x13, 0x01, byte1, byte2];
  [_Logfile DetailsSent:[NSString stringWithFormat:@"OPENDCC: Extended Accessory Command. Address = %d, Value = %d", address, value]];
}


// ****************************************************************************************************************
// ************************** RAISING EVENTS: TCP, USB AND XPRESSNET FOR ERRORS, STATUS CHANGES AND FEEDBACK INFO *
// ****************************************************************************************************************
// SHOULD NOT BE CALLED BY THE LIBRARY USER!
// The methods below forward events from lower level objects to the library user.
// In addition Logfile messages are written, if needed.
- (void)eventInterfaceConnected {
  if (DELEGATE_RESPONDS(InterfaceConnected)) [_delegate InterfaceConnected];
  [_Logfile Details:@"Interface connection established"];
  // In case of a serial connection, it seems we must wait 100ms before the interface is stable
  if (_connectionType == serialPort) [self performSelector:@selector(requestZentraleInfo) withObject:nil afterDelay:SERIAL_START_DELAY];
  else [self requestZentraleInfo];
}

// Retrieve version and status info regarding the Command Station
- (void)requestZentraleInfo {
  [self CmdSoftwareversionZentraleAnfordern];
  [self CmdStatusZentraleAnfordern];
}

- (void)eventTcpError:(NSString *)Description {
  if (DELEGATE_RESPONDS(TcpError:)) [_delegate TcpError:Description];
  [_Logfile Errors:[NSString stringWithFormat:@"TCP Error: %@", Description]];
}

- (void)eventTcpClosedByRemoteHost {
  [self closeConnections];
  if (DELEGATE_RESPONDS(TcpClosedByRemoteHost)) [_delegate TcpClosedByRemoteHost];
  [_Logfile Errors:@"TCP Connection closed by remote host"];
}


// ------------------------------------------------------------------------------------------------------------------
// Interface specific events
// V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
- (void)eventInterfaceVersion:(int)daten1 daten2:(int)daten2 {
  NSString *Version = [NSString stringWithFormat:@"%d.%d", daten1 / 16, daten1 % 16];
  NSString *code = [NSString stringWithFormat:@"%d.%d", daten2 / 16, daten2 % 16];
  if (DELEGATE_RESPONDS(InterfaceVersion:code:)) [_delegate InterfaceVersion:Version code:code];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Interface - Version: %@  Code: %@", Version, code]];
}

// V3.6: 1.7 (XpressNet Address) / V3.6: 2.1 (Interface Status Befehl)
- (void)eventXpressNetAddress:(int)address {
  // This event is triggered as response to two commands:
  // XpressNet Address (CmdXpressNetAddress - V3.6: 1.7) and Interface Status Befehl (CmdInterfaceStatus - V3.6: 2.1).
  // CmdInterfaceStatus will be called internally by the library, whenever the keep alive timer expires.
  // In such case, the response to the keep alive message should not be returned (as event)
  // to the application program, but instead be kept within the library.
  // To distinguish between both cases, the keep alive timer sets the waitForlanKeepAliveResponse flag
  if ([_send waitForlanKeepAliveResponse]) {
    [_send setWaitForlanKeepAliveResponse:NO];
    [_Logfile DetailsReceived:@"TCP Keep Alive Response"];
  }
  else {
    if (DELEGATE_RESPONDS(XpressNetAddress:)) [_delegate XpressNetAddress:address];
    [_Logfile DetailsReceived:[NSString stringWithFormat:@"XpressNet - Address: %d", address]];
  }
}

// V3.6: 2.2 (XpressNet Version)
- (void)eventXpressNetVersion:(int)daten {
  NSString *Version = [NSString stringWithFormat:@"%d.%d", daten / 16, daten % 16];
  if (DELEGATE_RESPONDS(XpressNetVersion:)) [_delegate XpressNetVersion:Version];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"XpressNet - Version: %@", Version]];
}

// V3.6: 2.3 (Freie Verbindungen)
- (void)eventInterfaceFreieVerbindungen:(int)number {
  if (DELEGATE_RESPONDS(InterfaceFreieVerbindungen:)) [_delegate InterfaceFreieVerbindungen:number];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Interface - Free TCP Sockets: %d", number]];
}


// ------------------------------------------------------------------------------------------------------------------
// Xpressnet specific events
// V3.6: 3.1.1.1, V4: 2.4.1, Z21: 2.8 (LAN_X_BC_TRACK_POWER_ON)
- (void)eventAllesAn {
  _ZentraleNotAus = NO;
  _ZentraleNotHalt = NO;
  _ZentraleProgrammierMode = NO;
  if (DELEGATE_RESPONDS(AllesAn)) [_delegate AllesAn];
  [_Logfile DetailsReceived:@"Broadcast: Alles An"];
}

// V3.6: 3.1.1.2, V4: 2.4.2, Z21: 2.7 (LAN_X_BC_TRACK_POWER_OFF)
- (void)eventNotAus {
  _ZentraleNotAus = YES;
  if (DELEGATE_RESPONDS(NotAus)) [_delegate NotAus];
  [_Logfile DetailsReceived:@"Broadcast: Alles Aus (Notaus)"];
}

// V3.6: 3.1.1.3, V4: 2.4.3, Z21: 2.14 (LAN_X_BC_STOPPED)
- (void)eventNotHalt {
  _ZentraleNotHalt = YES;
  if (DELEGATE_RESPONDS(NotHalt)) [_delegate NotHalt];
  [_Logfile DetailsReceived:@"Broadcast: Alle Loks Aus - Nothalt"];
}

// V3.6: 3.1.1.4, V4: 2.4.4, Z21: 2.9 (LAN_X_BC_PROGRAMMING_MODE)
- (void)eventProgrammierMode {
  _ZentraleProgrammierMode = YES;
  if (DELEGATE_RESPONDS(ProgrammierMode)) [_delegate ProgrammierMode];
  [_Logfile DetailsReceived:@"Broadcast: Programmiermode"];
}

// To simplify the application program, we generate different feedback events for switches and for
// feedback (occupancy) decoders. The difference between both can be determined by the TT bits
// V3.6: 3.1.1.5, V4: 2.4.5, V4: 2.4.6 (BC RuckMeldung) / V3.6: 3.1.8, V4: 2.15, V4: 2.16 (Schaltinformation)
- (void)eventFeedback:(int)address Nibble:(int)Nibble ttBits:(int)ttBits iBit:(int)iBit data:(int)data {
  if (ttBits == 2) {                                      // normal feedback decoder
    // We add 1 to the RS-Bus address, since addresses on the wire range from 0..127. We display these as 1..128
    address = address + 1;
    if (DELEGATE_RESPONDS(Feedback:Nibble:data:)) [_delegate Feedback:address Nibble:Nibble data:data];
    [_Logfile DetailsReceived:[NSString stringWithFormat:@"Feedback Message received for RS-Bus address: %d  Data: %d  Nibble: %d", address, data, Nibble]];
  }
  else {                                                  // switch decoder
    // Address is the decoder address, but we will return a switch address pair
    int switchAddress;
    if (!_Z21Protocol) switchAddress = (address * 4) + (Nibble * 2) + 1;
    else switchAddress = address + 1;
    if (DELEGATE_RESPONDS(SwitchFeedback:ttBits:iBit:data:)) [_delegate SwitchFeedback:switchAddress ttBits:ttBits iBit:iBit data:data];
    [_Logfile DetailsReceived:[NSString stringWithFormat:@"Feedback Message received for switch addresses: %d/%d  Data: %d  TT-Bits: %d I-Bit: %d",
                       switchAddress, switchAddress + 1, data, ttBits, iBit]];
  }
}

// V3.6: 3.1.2.1, V4: 2.5.1, Z21: 2.10 (Programmierinfo Kurzschluss / LAN_X_BC_TRACK_SHORT_CIRCUIT)
- (void)eventProgrammierinfoKurzschluss {
  if (DELEGATE_RESPONDS(ProgrammierinfoKurzschluss)) [_delegate ProgrammierinfoKurzschluss];
  [_Logfile DetailsReceived:@"Programmierinfo Kurzschluss"];
}

// V3.6: 3.1.2.2, V4: 2.5.2, Z21: 6.4 (Programmierinfo Daten nicht gefunden / LAN_X_CV_NACK)
- (void)eventProgrammierinfoDatenNichtGefunden {
  if (DELEGATE_RESPONDS(ProgrammierinfoDatenNichtGefunden)) [_delegate ProgrammierinfoDatenNichtGefunden];
  [_Logfile DetailsReceived:@"Programmierinfo Daten nicht gefunden"];
}

// V3.6: 3.1.2.3, V4: 2.5.3 (Programmierinfo Zentrale Busy)
- (void)eventProgrammierinfoZentraleBusy {
  if (DELEGATE_RESPONDS(ProgrammierinfoZentraleBusy)) [_delegate ProgrammierinfoZentraleBusy];
  [_Logfile DetailsReceived:@"Programmierinfo Zentrale Busy"];
}

// V3.6: 3.1.2.4, V4: 2.5.4 (Programmierinfo Zentrale Bereit)
- (void)eventProgrammierinfoZentraleBereit {
  if (DELEGATE_RESPONDS(ProgrammierinfoZentraleBereit)) [_delegate ProgrammierinfoZentraleBereit];
  [_Logfile DetailsReceived:@"Programmierinfo Zentrale Bereit"];
}

// V3.6 3.1.2.5 / V4: 2.5.5 (Programmierinfo Daten 3-Byte-Format)
- (void)eventProgInfoPageRegister:(int)EEPromAddress value:(int)value {
  if (DELEGATE_RESPONDS(ProgrammierinfoPageRegister:value:)) [_delegate ProgrammierinfoPageRegister:EEPromAddress value:value];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Programmierinfo Page / Register. EEPromAddress: %d, Value: %d", EEPromAddress, value]];
}

// V3.6: 3.1.2.6 - 3.1.2.9, V4: 2.5.6 - 2.5.9 (Programmierinfo CV1..1024)
- (void)eventProgInfo:(int)CvRange daten1:(int)daten1 value:(int)value {
  int CV = daten1;
  switch (CvRange) {
    case 1: if (CV == 0) CV = (_zentraleVersionNumber < 36) ? 256 : 1024; break;
    case 2: CV = CV + 256; break;
    case 3: CV = CV + 512; break;
    case 4: CV = CV + 768; break;
    default: break;
  }
  if (DELEGATE_RESPONDS(ProgrammierinfoCV:value:)) [_delegate ProgrammierinfoCV:CV value:value];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Programmierinfo. CV: %d, Value: %d", CV, value]];
}

// Z21: 6.5 (LAN_X_CV_RESULT)
- (void)eventProgInfoZ21:(int)CVAdr_MSB CVAdr_LSB:(int)CVAdr_LSB value:(int)value {
  int CV = (CVAdr_MSB * 256) + CVAdr_LSB + 1;
  if (DELEGATE_RESPONDS(ProgrammierinfoCV:value:)) [_delegate ProgrammierinfoCV:CV value:value];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Programmierinfo Z21. CV: %d, Value: %d", CV, value]];
}

// V4: 2.6 (Service Variable melden)
- (void)eventServiceVariableMelden:(int)SvAddress SvValue:(int)SvValue {
  if (DELEGATE_RESPONDS(ServiceVariable:SvValue:)) [_delegate ServiceVariable:SvAddress SvValue:SvValue];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Service Variable melden. SV Address: %d, SV Value: %d", SvAddress, SvValue]];
}

// V3.6: 3.1.3, V4: 2.7.2, Z21: 2.3 (Softwareversion Zentrale / LAN_X_GET_VERSION)
- (void)eventSoftwareVersionZentrale:(int)daten2 daten3:(int)daten3 {
  NSString *Version = [NSString stringWithFormat:@"%d.%d", daten2 / 16, daten2 % 16];
  [self setZentraleVersion:Version];
  _zentraleVersionNumber = (daten2 / 16) * 10 + (daten2 % 16);
  _ZentraleKennung = daten3;
  if (DELEGATE_RESPONDS(ZentraleVersion:)) [_delegate ZentraleVersion:Version];
  // 2.0: with the name of the command station
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Version Zentrale: %@, Kennung: %d (%@)", Version, _ZentraleKennung, [self zentraleName:_ZentraleKennung]]];
}

// 2.0: The name of the command station, for the logfile (Zentralen-Kennung, V3.6: 3.1.3)
- (NSString *)zentraleName:(int)kennung {
  switch (kennung) {
    case KENNUNG_LZ100: return @"LZ 100";
    case KENNUNG_LH200: return @"LH 200";
    case KENNUNG_DPC: return @"DPC (Compact / Commander)";
    case KENNUNG_CNTRPLUS: return @"Control Plus";
    case KENNUNG_LOKMAUS2: return @"LokMaus II";
    case KENNUNG_LDH: return @"LDH";
    case KENNUNG_MULTIMAUS: return @"multiMAUS";
    case KENNUNG_Z21: return @"Z21 / Yamorc";
    case KENNUNG_OPENDCC: return @"OpenDCC";
    case KENNUNG_OPENDCCXP: return @"OpenDccXp";
    default: return @"unknown";
  }
}

// V3.6: 3.1.4, V4: 2.8, Z21: 2.12 (Status Zentrale / LAN_X_STATUS_CHANGED)
- (void)eventStatusZentrale:(int)data {
  // With Z21, we internally use the StatusZentrale command as "Keep Alive" Message
  // If we use the StatusZentrale command in this way, we need not inform the application
  // about the result. In that case it is sufficient to tell the remaining program that
  // we no longer wait for the Keep Alive Response. Just log it that we received the response, and return
  if ((_connectionType == lanUdpPort) && [_send waitForlanKeepAliveResponse]) {
    [_send setWaitForlanKeepAliveResponse:NO];
    [_Logfile DetailsReceived:@"UDP Keep Alive Response"];
    return;
  }
  // If we are here, this is not a Z21 Keep Alive, but a response to a normal request.
  _ZentraleNotHalt = bitValue(data, 0);
  _ZentraleNotAus = bitValue(data, 1);
  if (!_Z21Protocol) _ZentraleStartmodeAuto = bitValue(data, 2);
  if (!_Z21Protocol) _ZentraleProgrammierMode = bitValue(data, 3);
  if (_Z21Protocol) _ZentraleKurzSchluss = bitValue(data, 2);
  if (_Z21Protocol) _ZentraleProgrammierMode = bitValue(data, 5);
  _ZentraleKaltstart = bitValue(data, 6);
  _ZentraleRamCheckFehler = bitValue(data, 7);
  if (_ZentraleNotHalt) {
    if (DELEGATE_RESPONDS(NotHalt)) [_delegate NotHalt];
    [_Logfile DetailsReceived:@"Startup: NotHalt"];
  }
  if (_ZentraleNotAus) {
    if (DELEGATE_RESPONDS(NotAus)) [_delegate NotAus];
    [_Logfile DetailsReceived:@"Startup: NotAus"];
  }
  if (_ZentraleProgrammierMode) {
    if (DELEGATE_RESPONDS(ProgrammierMode)) [_delegate ProgrammierMode];
    [_Logfile DetailsReceived:@"Startup: ProgrammierMode"];
  }
  if (_ZentraleRamCheckFehler) [_Logfile Errors:@"Startup: Zentrale - Ram Check Fehler"];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Status Zentrale: %d", data]];
}

// V4: 2.10 (PoM Ergebnis melden)
- (void)eventPomErgebnisMelden:(int)AdrH AdrL:(int)AdrL CvValue:(int)CvValue {
  int LokAddress = AdrH * 256 + AdrL;
  if (DELEGATE_RESPONDS(PomErgebnis:CvValue:)) [_delegate PomErgebnis:LokAddress CvValue:CvValue];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"PomErgebnisMelden. Lok Address:: %d, CV Value:%d", LokAddress, CvValue]];
}

// V3.6: 3.1.9.1, V4: 2.19.1 (Lok Info Normale Lok)
- (void)eventNormaleLokinfo:(int)Kennung Speed:(int)Speed F0_F4:(int)F0_F4 F5_F12:(int)F5_F12 {
  BOOL Besetzt = bitValue(Kennung, 3);
  int Fahrstufen = [self speedStepsForKennung:Kennung];
  BOOL Forwarts = bitValue(Speed, 7);
  int Geschwindigkeit = [self speedForSpeedByte:Speed speedSteps:Fahrstufen];
  int F5_F8 = F5_F12 % 16;                                // Gruppe 2 - take the lowest 4 bits
  int F9_F12 = F5_F12 / 16;                               // Gruppe 3 - take the highest 4 bits
  if (DELEGATE_RESPONDS(NormaleLokinfo:Speedsteps:Speed:Forward:F0_F4:F5_F8:F9_F12:))
    [_delegate NormaleLokinfo:Besetzt Speedsteps:Fahrstufen Speed:Geschwindigkeit Forward:Forwarts F0_F4:F0_F4 F5_F8:F5_F8 F9_F12:F9_F12];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Lokinformation (normal) - Besetzt: %@, Speed: %d, Stufen: %d, Forwarts: %@, F0..F4 %d, F5..F8 %d, F9..F12 %d",
                     Besetzt ? @"True" : @"False", Geschwindigkeit, Fahrstufen, Forwarts ? @"True" : @"False", F0_F4, F5_F8, F9_F12]];
}

// Z21: 4.4 (LAN_X_LOCO_INFO)
- (void)eventZ21Lokinfo:(int)Adr_MSB Adr_LSB:(int)Adr_LSB Kennung:(int)Kennung Speed:(int)Speed F0_F4:(int)F0_F4 F5_F12:(int)F5_F12 F13_F20:(int)F13_F20 F21_F28:(int)F21_F28 {
  int LokAddress = (Adr_MSB & 0x3F) * 256 + Adr_LSB;
  BOOL Besetzt = bitValue(Kennung, 3);
  int Fahrstufen = [self speedStepsForKennung:Kennung];
  BOOL Forwarts = bitValue(Speed, 7);
  int Geschwindigkeit = [self speedForSpeedByte:Speed speedSteps:Fahrstufen];
  int F5_F8 = F5_F12 % 16;                                // Gruppe 2 - take the lowest 4 bits
  int F9_F12 = F5_F12 / 16;                               // Gruppe 3 - take the highest 4 bits
  if (DELEGATE_RESPONDS(Z21Lokinfo:Besetzt:Speedsteps:Speed:Forward:F0_F4:F5_F8:F9_F12:F13_F20:F21_F28:))
    [_delegate Z21Lokinfo:LokAddress Besetzt:Besetzt Speedsteps:Fahrstufen Speed:Geschwindigkeit Forward:Forwarts
                    F0_F4:F0_F4 F5_F8:F5_F8 F9_F12:F9_F12 F13_F20:F13_F20 F21_F28:F21_F28];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Lokinformation (Z21) - Address: %d, Besetzt: %@, Speed: %d, Stufen: %d, Forwarts: %@, F0..F4 %d, F5..F8 %d, F9..F12 %d, F13..F20 %d, F21..F28 %d",
                     LokAddress, Besetzt ? @"True" : @"False", Geschwindigkeit, Fahrstufen, Forwarts ? @"True" : @"False", F0_F4, F5_F8, F9_F12, F13_F20, F21_F28]];
}

// V3.6: 3.1.9.2, V4: 2.19.2 (Funktionszustand F13 bis F28)
- (void)eventFunctionsZustandF13F20:(int)F13_F20 F21_F28:(int)F21_F28 {
  if (DELEGATE_RESPONDS(FunctionsZustandF13F20:F21_F28:)) [_delegate FunctionsZustandF13F20:F13_F20 F21_F28:F21_F28];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Funktionszustand - F13 bis F20: %d, F21 bis F28: %d", F13_F20, F21_F28]];
}

// V4: 2.19.3 (Funktionszustand F29 bis F68)
- (void)eventFunctionsZustandF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68 {
  if (DELEGATE_RESPONDS(FunctionsZustandF29F68:F37_F44:F45_F52:F53_F60:F61_F68:))
    [_delegate FunctionsZustandF29F68:F29_F36 F37_F44:F37_F44 F45_F52:F45_F52 F53_F60:F53_F60 F61_F68:F61_F68];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Funktionszustand - F29 bis F68: %d, %d, %d, %d, %d", F29_F36, F37_F44, F45_F52, F53_F60, F61_F68]];
}

// V3.6: 3.1.11, V4: 2.19.8 (Funktionsstatus F0 bis F12 der angefragten Lok)
- (void)eventFunctionStatusF0F12:(int)F0_F4 F5_F12:(int)F5_F12 {
  int F5_F8 = F5_F12 % 16;                                // Gruppe 2 - take the lowest 4 bits
  int F9_F12 = F5_F12 / 16;                               // Gruppe 3 - take the highest 4 bits
  if (DELEGATE_RESPONDS(FunctionStatusF0F12:F5_F8:F9_F12:)) [_delegate FunctionStatusF0F12:F0_F4 F5_F8:F5_F8 F9_F12:F9_F12];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Funktionsstatus - F0..F4 %d, F5..F8 %d, F9..F12 %d", F0_F4, F5_F8, F9_F12]];
}

// V3.6: 3.1.12, V4: 2.20 (Funktionsstatus F13 bis F28)
- (void)eventFunctionStatusF13F20:(int)F13_F20 F21_F28:(int)F21_F28 RefreshMode:(int)RefreshMode {
  if (DELEGATE_RESPONDS(FunctionStatusF13F20:F21_F28:RefreshMode:)) [_delegate FunctionStatusF13F20:F13_F20 F21_F28:F21_F28 RefreshMode:RefreshMode];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Funktionsstatus - F13 bis F20: %d, F21 bis F28: %d, Refresh Mode: %d", F13_F20, F21_F28, RefreshMode]];
}

// V4: 2.21 (Funktionsstatus F29 bis F68)
- (void)eventFunctionStatusF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68 {
  if (DELEGATE_RESPONDS(FunctionStatusF29F68:F37_F44:F45_F52:F53_F60:F61_F68:))
    [_delegate FunctionStatusF29F68:F29_F36 F37_F44:F37_F44 F45_F52:F45_F52 F53_F60:F53_F60 F61_F68:F61_F68];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Funktionsstatus - F29 bis F68: %d, %d, %d, %d, %d", F29_F36, F37_F44, F45_F52, F53_F60, F61_F68]];
}

// V3.6: 3.1.14, V4: 2.24 (Fehlermeldung)
- (void)eventFehlermeldung:(int)data {
  if (bitValue(data, 7)) {                                // Highest order bit must have been set
    int value = data % 16;                                // Take the four lowest order bits
    if (DELEGATE_RESPONDS(Fehlermeldung:)) [_delegate Fehlermeldung:value];
    if (_LogWarnings) [_Logfile AppendWithHost:[NSString stringWithFormat:@"<= Fehlermeldung: %d", value]];   // 2.0: <=
  }
}

// V4: 2.4.7 / V4: 2.11 (BC: Modellzeit / Modellzeit melden). factor = 255 if the message has no factor
- (void)eventModellZeit:(int)daten1 daten2:(int)daten2 factor:(int)factor {
  int day = daten1 / 32;
  int hour = daten1 % 32;
  int minute = daten2 % 64;
  if (DELEGATE_RESPONDS(Modellzeit:hour:minute:factor:)) [_delegate Modellzeit:day hour:hour minute:minute factor:factor];
  // We do not log Modellzeit messages in the logfile
}

// V4: 2.9 (Erweiterte Versionsinformation melden)
- (void)eventErweiterteVersionZentrale:(int)ZBldH ZBldL:(int)ZBldL RMVer:(int)RMVer RMBldH:(int)RMBldH RMBldL:(int)RMBldL BIvers:(int)BIvers {
  NSString *BuildZentrale = [NSString stringWithFormat:@"%d", ZBldH * 256 + ZBldL];
  NSString *VersionRuckMelder = [NSString stringWithFormat:@"%d.%d", RMVer / 16, RMVer % 16];
  NSString *BuildRuckMelder = [NSString stringWithFormat:@"%d", RMBldH * 256 + RMBldL];
  NSString *VersionBootloader = [NSString stringWithFormat:@"%d.%d", BIvers / 16, BIvers % 16];
  if (DELEGATE_RESPONDS(ErweiterteZentraleVersion:VersionRuckMelder:BuildRuckMelder:VersionBootloader:))
    [_delegate ErweiterteZentraleVersion:BuildZentrale VersionRuckMelder:VersionRuckMelder BuildRuckMelder:BuildRuckMelder VersionBootloader:VersionBootloader];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Erweiterte Versionsinformation - BuildZentrale = %@, VersionRuckMelder = %@, BuildRuckMelder = %@, VersionBootloader = %@",
                     BuildZentrale, VersionRuckMelder, BuildRuckMelder, VersionBootloader]];
}

// Z21: 2.15 (Antwort auf LAN_X_GET_FIRMWARE_VERSION). The version is BCD coded: 0x01 0x23 = V1.23
- (void)eventz21FirmwareVersion:(int)V_MSB V_LSB:(int)V_LSB {
  NSString *Version = [NSString stringWithFormat:@"%X.%02X", V_MSB, V_LSB];    // 2026: VB6 showed the bytes as decimal values
  if (DELEGATE_RESPONDS(z21FirmwareVersion:)) [_delegate z21FirmwareVersion:Version];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Z21 Firmware - Version = %@", Version]];
}

// Z21: 5.6 (LAN_X_EXT_ACCESSORY_INFO)
- (void)eventExtendedAccessoryInfo:(int)Adr_MSB Adr_LSB:(int)Adr_LSB DDDDDDDD:(int)DDDDDDDD Status:(int)Status {
  int rawAddress = (Adr_MSB * 256) + Adr_LSB;
  if (rawAddress <= 3) rawAddress = rawAddress + 2048;
  int address = rawAddress - 3;                           // The first (XpressNet) address is 4, and presented to the user as 1
  if (DELEGATE_RESPONDS(ExtendedAccessoryInfo:value:Status:)) [_delegate ExtendedAccessoryInfo:address value:DDDDDDDD Status:Status];
  [_Logfile DetailsReceived:[NSString stringWithFormat:@"Extended Accessory Info - Address: %d, Value: %d, Status: %d", address, DDDDDDDD, Status]];
}

// ------------------------------------------------------------------------------------------------------------------
// 2.0: Sending events (see XpressNetSend.m). The logfile is written by the caller
- (void)eventSendError:(SendErrorReasonEnum)reason description:(NSString *)description {
  if (DELEGATE_RESPONDS(SendError:Description:)) [_delegate SendError:(int)reason Description:description];
}

- (void)eventInterfaceNichtAdressiert {
  if (DELEGATE_RESPONDS(InterfaceNichtAdressiert)) [_delegate InterfaceNichtAdressiert];
}

- (void)eventInterfaceWiederAdressiert {
  if (DELEGATE_RESPONDS(InterfaceWiederAdressiert)) [_delegate InterfaceWiederAdressiert];
}


// ****************************************************************************************************************
// ********************************* TCP, UDP AND SERIAL PORT DELEGATE METHODS ************************************
// ****************************************************************************************************************
// These methods are called on the main thread
- (void)tcpSocket:(TCPSocketClass *)socket didChangeState:(tcpSocketState_t)state reason:(NSString *)reason {
  switch (state) {
    case TCP_SOCKET_CONNECTING: break;
    case TCP_SOCKET_CONNECTED:
      // 2.0: which connection belongs to the letter of this XpressNet object
      [_Logfile Connection:[NSString stringWithFormat:@"%@, port %@, local port %d", [_Logfile myHost], _connectionPort, [socket localPort]]];
      [self eventInterfaceConnected];
      break;
    case TCP_SOCKET_FAILED:
      [self closeConnections];
      if ([socket lastErrorCode] == ETIMEDOUT) [self eventTcpError:[NSString stringWithFormat:@"No response from %@", [_Logfile myHost]]];
      else [self eventTcpError:[NSString stringWithFormat:@"Could not connect to %@ (%@)", [_Logfile myHost], reason]];
      break;
    case TCP_SOCKET_LOST:
      [self closeConnections];
      [self eventTcpError:[NSString stringWithFormat:@"Connection to %@ lost (%@)", [_Logfile myHost], reason]];
      break;
    case TCP_SOCKET_CLOSED:
      if (reason == nil) break;                           // We closed the connection ourselves
      [self eventTcpClosedByRemoteHost];                  // The other side closed the connection
      break;
  }
}

- (void)tcpSocket:(TCPSocketClass *)socket didReceiveBytes:(const uint8_t *)bytes length:(NSUInteger)length {
  [_receive addToIndata:bytes length:(int)length];
}

- (void)udpSocket:(UDPSocketClass *)socket didChangeState:(udpSocketState_t)state reason:(NSString *)reason {
  // 2.0: which connection belongs to the letter of this XpressNet object
  if (state == UDP_SOCKET_READY) {
    [_Logfile Connection:[NSString stringWithFormat:@"%@, port %@, local port %d", [_Logfile myHost], _connectionPort, [socket localPort]]];
  }
  if (state == UDP_SOCKET_FAILED) {
    [self closeConnections];
    [self eventTcpError:[NSString stringWithFormat:@"UDP connection to %@ failed (%@)", [_Logfile myHost], reason]];
  }
}

- (void)udpSocket:(UDPSocketClass *)socket didReceiveBytes:(const uint8_t *)bytes length:(NSUInteger)length {
  // Note: the Yamorc YD7001D retransmits broadcast messages three times
  [_receiveZ21 addToIndata:bytes length:(int)length];
}

- (void)serialPort:(SerialPortClass *)port didReceiveBytes:(const uint8_t *)bytes length:(NSUInteger)length {
  [_receive addToIndata:bytes length:(int)length];
}

- (void)serialPort:(SerialPortClass *)port didCloseWithReason:(NSString *)reason {
  [self closeConnections];
  [self eventTcpError:[NSString stringWithFormat:@"Connection to %@ lost (%@)", [_Logfile myHost], reason]];
}


// ****************************************************************************************************************
// ****************************** Methods for internal library use. Should not be called by the library user *****
// ****************************************************************************************************************
// 2.0: do we have a connection on which we can send? While the TCP or UDP connection is being established, messages
// are kept by the connection, and sent as soon as it is ready (see XpressNet.h)
- (BOOL)isConnected {
  switch (_connectionType) {
    case lanTcpPort: return [_tcp isOpen];
    case lanUdpPort: return [_udp isOpen];
    case usbPort:
    case serialPort: return [_usb isOpen];
    default: return NO;
  }
}

// 2.0: the port of the TCP or UDP connection, for the logfile
- (void)setConnectionPort:(NSString *)port {
  [_connectionPort release];
  _connectionPort = [port copy];
}

- (BOOL)VelueOutOfRange:(int)value min:(int)min max:(int)max LogText:(NSString *)LogText {
  // support function to check is Value is between some predefined bounds
  // 2.0: a value out of range gives the event SendError (the message is not sent)
  if (value < min) {
    [_Logfile AppendWithHost:[NSString stringWithFormat:@"Error %@ must be >= %d", LogText, min]];
    [self eventSendError:FehlerInDenBefehlsparametern description:[NSString stringWithFormat:@"Message not sent: %@ must be >= %d (value: %d)", LogText, min, value]];
    return YES;
  }
  if (value > max) {
    [_Logfile AppendWithHost:[NSString stringWithFormat:@"Error %@ must be <= %d", LogText, max]];
    [self eventSendError:FehlerInDenBefehlsparametern description:[NSString stringWithFormat:@"Message not sent: %@ must be <= %d (value: %d)", LogText, max, value]];
    return YES;
  }
  return NO;
}

// Functions for Accessory addresses
- (int)AccAL:(int)address {
  // Calculation of the low byte of a accessory address
  if (_Z21Protocol) return address % 256;                 // Select the last 8-bits
  if (address >= 1024) return (address - 1024) / 4;       // We already assured that the command station >= V3.8
  return address / 4;
}

- (int)AccAH:(int)address {
  // Calculation of the high byte of a accessory address. Lenz uses decoder addressing
  if (_Z21Protocol) return address / 256;                 // Remove the last 8-bits (shift right)
  return address / 1024;                                  // Remove the last 10-bits (shift right)
}

// Functions for lok addresses and speed
- (int)AL:(int)address {
  // Calculation of the low byte of a loco address
  return address % 256;
}

- (int)AH:(int)address {
  // Calculation of the high byte of a loco address
  if (address < 100) return 0;
  return address / 256 + 0xC0;
}

- (int)EncodeSpeed:(int)Speed {
  // Calculates the speed byte for decoders with 27/28 speed steps
  int result = 0;
  if (Speed > 0) {
    Speed = Speed + 3;                                    // RVVV expects an offset
    result = Speed / 2;                                   // neglect the last bit, thus all uneven speed steps
    if (Speed & 1) result = result + 16;                  // If the speed was even, set RVVV bit 4 to 1
  }
  return result;
}

- (int)DecodeSpeed:(int)Speed {
  // Derive from the speed byte an integer value (for decoders with 27/28 speed steps)
  int result = 0;                                         // Default speed
  int Bit4 = bitValue(Speed, 4) ? 1 : 0;                  // take bit 4
  int Bit0_3 = Speed % 16;                                // take the lowest 4 bits
  if (Bit0_3 >= 2) result = (Bit0_3 - 2) * 2 + 1 + Bit4;
  return result;
}

// The speed steps, from the bits 0..2 of the Kennung byte
- (int)speedStepsForKennung:(int)Kennung {
  switch (Kennung % 8) {
    case 0: return 14;
    case 1: return 27;
    case 2: return 28;
    case 4: return 128;
    default: return 0;
  }
}

// 2026: the speed byte is decoded according to the speed steps (VB6: always as 27/28 speed steps).
// 14 and 128 speed steps: 0 = stop, 1 = emergency stop (shown as 0), 2.. = speed step 1..
- (int)speedForSpeedByte:(int)Speed speedSteps:(int)speedSteps {
  int value;
  switch (speedSteps) {
    case 14:
      value = Speed % 16;
      return (value >= 2) ? value - 1 : 0;
    case 128:
      value = Speed % 128;
      return (value >= 2) ? value - 1 : 0;
    default:
      return [self DecodeSpeed:Speed];
  }
}

@end
