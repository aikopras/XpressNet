//
//  XpressNet.h
//  XpressNet Test
//
//  Created by Aiko Pras on 30-10-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  macOS version of the VB6 XpressNet DLL (XpressNetClass). Version 2.0, November 2026 (1.2: October 2026).
//  The library supports XpressNet over TCP (Lenz LAN/USB Interface 23151), over USB and over a serial interface,
//  as well as the XpressNet tunnel of the Z21 LAN protocol (UDP).
//
//  The various methods contain references to the V3.6, V4 and Z21 specifications:
//  - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
//    Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
//  - Dokumentation XpressNet Version 4.0, 02/2022
//  - Z21 LAN Protocol Specification, Version 1.13, 06.11.2023
//
//  Convention (as in the DLL): everything intended for the library user starts with a capital letter.
//  Internal objects, methods and attributes start with lowercase.
//  - Methods: same names and parameters as in the DLL. VB Integer and Byte become int, Boolean becomes BOOL,
//    and String becomes NSString. An Optional parameter results in two methods: with and without that parameter.
//  - Events: the methods of the XpressNetDelegate protocol, with the same names and parameters as the DLL events.
//    All events are optional: the application implements only the events it needs (as with VB6 WithEvents).
//    Events are called on the main thread.
//  - Attributes: properties.
//
//  Differences with the VB6 DLL:
//  - CmdConnectUsb and CmdConnectSerial: Port is the name of a serial device (for example /dev/cu.usbserial-A1B2C3),
//    instead of a COM port number. If the port can not be opened, or is lost, the event TcpError is raised
//    (VB6: a MsgBox).
//  - Opening a new connection first closes an existing connection.
//  - The response timeout (5 sec, or 90 sec in programming mode) is used for TCP, USB and serial connections (VB6: TCP only).
//    The Z21 (UDP) has no flow control: the Z21 does not answer every command (Z21 spec, section 1.1).
//    Therefore Z21 messages are sent immediately (VB6: after the next received message, without timeout).
//  - Corrected errors of the VB6 version:
//    - PoM commands: the low byte of the CV number was CV Mod 255, instead of CV Mod 256
//    - PoM Bit schreiben: 0xE8 + C, as in the V3.6 spec (VB6: 0x7C + C, as in the V4 spec, which seems to be wrong)
//    - Lokinfo: the speed is decoded according to the speed steps (VB6: always as 27/28 speed steps)
//    - Z21 receive: the XpressNet message is taken from the received Z21 message (VB6: from the shifted input buffer),
//      and a message with a wrong XOR is ignored
//    - Z21 firmware version: BCD coded, as in the Z21 spec (VB6: decimal)
//    - CmdZ21GetExtAccessory: 3 data bytes, as in the Z21 spec (VB6: 4)
//    - Version comparisons are numeric (VB6: partly as strings)
//  - 2.0: messages that are offered while the TCP or UDP connection is being established, are kept by the connection
//    (TCPSocketClass, UDPSocketClass) and sent as soon as the connection is ready (VB6 and .NET: SendError
//    KeineVerbindung)
//
//  New in version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
//  - Retransmission of a message after an error answer of the interface, and a pause as long as the command
//    station does not address the interface (01 05 / 01 07). See XpressNetSend.m
//  - The events SendError, InterfaceNichtAdressiert and InterfaceWiederAdressiert. SendError is raised for every
//    message that is not sent (not silently dropped). The reasons are in SendErrorReasonEnum
//  - CmdConnectUsb:Settings:Handshake: (hardware handshake RTS/CTS, default off)
//  - The names of more command stations in the logfile (LokMaus II, LDH, multiMAUS)
//  - The default directory of the logfile is ~/Library/Logs/TMC-Logfiles (see XpressNetLogFile.m)
//  - In the logfile every line starts with the letter of the XpressNet object (A, B, ...), and the description of a
//    message with => (sent) or <= (received)
//  - CmdAccessoryAuto switches off the activated switch if the queue is emptied (see XpressNetAccAuto.m)
//  - After a received message with an XOR error, the next framing bytes are searched (see XpressNetReceive.m)
//  - Z21: CmdInitUdp first sets the broadcast flags (LAN_SET_BROADCASTFLAGS); new command CmdZ21SetBroadcastFlags:.
//    A new connection and the end of the object (dealloc) first send LAN_LOGOFF
//

#import <Foundation/Foundation.h>
@class SendClass;
@class ReceiveClass;
@class ReceiveClassZ21;
@class AccAutoClass;
@class LogFileClass;
@class TCPSocketClass;
@class UDPSocketClass;
@class SerialPortClass;


// ****************************************************************************************************************
// ****************************************************** EVENTS **************************************************
// ****************************************************************************************************************
@protocol XpressNetDelegate <NSObject>
@optional
// GENERATED EVENTS - Interface: TCP as well as USB
- (void)InterfaceConnected;
- (void)TcpClosedByRemoteHost;
- (void)TcpError:(NSString *)Description;
// GENERATED EVENTS - USB (declared, but not raised; as in the DLL)
- (void)UsbConnected;
// GENERATED EVENTS - XPRESSNET
- (void)AllesAn;
- (void)NotAus;
- (void)NotHalt;
- (void)ProgrammierMode;
- (void)XpressNetAddress:(int)address;
- (void)XpressNetVersion:(NSString *)Version;
- (void)InterfaceVersion:(NSString *)Version code:(NSString *)code;
- (void)InterfaceFreieVerbindungen:(int)number;
- (void)ZentraleVersion:(NSString *)Version;
- (void)ErweiterteZentraleVersion:(NSString *)BuildZentrale VersionRuckMelder:(NSString *)VersionRuckMelder BuildRuckMelder:(NSString *)BuildRuckMelder VersionBootloader:(NSString *)VersionBootloader;
- (void)Fehlermeldung:(int)value;
- (void)ServiceVariable:(int)SvAddress SvValue:(int)SvValue;
// Feedback events
- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data;
- (void)SwitchFeedback:(int)address ttBits:(int)ttBits iBit:(int)iBit data:(int)data;
- (void)ExtendedAccessoryInfo:(int)address value:(int)value Status:(int)Status;
// Lok events
- (void)NormaleLokinfo:(BOOL)Besetzt Speedsteps:(int)Speedsteps Speed:(int)Speed Forward:(BOOL)Forward F0_F4:(int)F0_F4 F5_F8:(int)F5_F8 F9_F12:(int)F9_F12;
- (void)FunctionsZustandF13F20:(int)F13_F20 F21_F28:(int)F21_F28;
- (void)FunctionsZustandF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68;
- (void)FunctionStatusF0F12:(int)F0_F4 F5_F8:(int)F5_F8 F9_F12:(int)F9_F12;
- (void)FunctionStatusF13F20:(int)F13_F20 F21_F28:(int)F21_F28 RefreshMode:(int)RefreshMode;
- (void)FunctionStatusF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68;
// Programmier events - Service Mode
- (void)ProgrammierinfoKurzschluss;
- (void)ProgrammierinfoDatenNichtGefunden;
- (void)ProgrammierinfoZentraleBusy;
- (void)ProgrammierinfoZentraleBereit;
- (void)ProgrammierinfoCV:(int)CV value:(int)value;
- (void)ProgrammierinfoPageRegister:(int)EEPromAddress value:(int)value;
// Programmier events - Programming on Main (PoM)
- (void)PomErgebnis:(int)LokAddress CvValue:(int)CvValue;
// Modellzeit events
- (void)Modellzeit:(int)day hour:(int)hour minute:(int)minute factor:(int)factor;
// Z21 specific events
- (void)z21FirmwareVersion:(NSString *)Version;
- (void)Z21Lokinfo:(int)address Besetzt:(BOOL)Besetzt Speedsteps:(int)Speedsteps Speed:(int)Speed Forward:(BOOL)Forward F0_F4:(int)F0_F4 F5_F8:(int)F5_F8 F9_F12:(int)F9_F12 F13_F20:(int)F13_F20 F21_F28:(int)F21_F28;
// 2.0: Sending (see XpressNetSend.m)
- (void)SendError:(int)Reason Description:(NSString *)Description;
- (void)InterfaceNichtAdressiert;
- (void)InterfaceWiederAdressiert;
@end


// Enumeration to determine the type of connection
typedef enum
{ none = 0,
  lanTcpPort = 1,
  lanUdpPort = 2,
  usbPort = 3,
  serialPort = 4,
} connectionTypeEnum;

// 2.0: The values of the parameter Reason of the event SendError. The description (parameter Description) contains
// the bytes of the message (including the XOR byte, without the framing bytes)
// The names are those of the Lenz specification (Interface 23151, V1.2), where the specification has a name
typedef enum
{ WiederholungErfolglos = 1,                    // Given up after 3 retransmissions (61 80, 61 81, 01 01, 01 02, 01 03,
                                                // 01 06, 01 0A). Not a name of the specification
  BefehlNichtVorhanden = 2,                     // Answer 61 82 (V3.6: 3.1.7)
  Timeout = 3,                                  // No answer within 5 sec (90 sec in programming mode; V3.6: 1.3).
                                                // All queues are emptied
  FifoVoll = 4,                                 // The FIFO (128 messages) or the CmdAccessoryAuto queue (32 commands) is
                                                // full. Not a name of the specification
  KeineVerbindung = 5,                          // There is no connection. Not a name of the specification
  FehlerInDenBefehlsparametern = 6,             // Answer 01 09 (V3.6: 1.5), or a parameter value out of range (checked
                                                // by the library; the message is not sent)
  ZentraleAdressiertDasInterfaceNichtMehr = 7,  // No 01 07 within 5 sec (90 sec in programming mode) after 01 05 / 01 08
                                                // (V3.6: 1.5). All queues are emptied
} SendErrorReasonEnum;

// 2.0: Z21: the broadcast flags that CmdInitUdp sets (see CmdZ21SetBroadcastFlags:)
static const int Z21BroadcastFlagsDefault = 0x00000001;


@interface XpressNetClass : NSObject

// The object that receives the events. Not retained
@property (assign) id <XpressNetDelegate> delegate;

// ATTRIBUTES THAT THE LIBRARY USER MAY CHANGE. THE DEFAULT VALUES ARE SET IN init
@property (assign) BOOL Z21Protocol;                    // Is the library used for Z21 or XpressNet??
@property (assign) BOOL LogMessages;                    // Example: connection requested to host X
@property (assign) BOOL LogWarnings;                    // May indicate a software problem
@property (assign) BOOL LogErrors;                      // Indicates a hardware or software problem
@property (assign) BOOL LogHexMessages;                 // Hex dump of transmitted / received messages
@property (assign) BOOL LogDetails;                     // Description of what has been transmitted / received
@property (assign) int AccessoryInterval;               // Interval (ms) between two accessory commands

// ATTRIBUTES THAT THE LIBRARY USER MAY READ
@property (assign) BOOL ZentraleNotAus;
@property (assign) BOOL ZentraleNotHalt;
@property (assign) BOOL ZentraleProgrammierMode;
@property (copy)   NSString *ZentraleVersion;
@property (assign) int ZentraleKennung;
@property (assign) BOOL ZentraleKurzSchluss;            // Only for Z21
@property (assign) BOOL ZentraleStartmodeAuto;
@property (assign) BOOL ZentraleKaltstart;
@property (assign) BOOL ZentraleRamCheckFehler;
@property (assign) long NumberOfExpressNetMessagesReceived;
@property (assign) long NumberOfExpressNetMessagesTransmitted;

// OBJECTS THAT THE LIBRARY USER MAY USE
@property (retain) LogFileClass *Logfile;

// Objects for internal use. Although these are public, they should not be used outside the library
@property (retain) SendClass *send;                     // Everything related to sending XpressNet messages
@property (retain) ReceiveClass *receive;               // Everything related to receiving XpressNet messages
@property (retain) ReceiveClassZ21 *receiveZ21;         // To receive messages transported over a Z21 XpressNet tunnel
@property (retain) TCPSocketClass *tcp;                 // The TCP connection
@property (retain) UDPSocketClass *udp;                 // The UDP "connection" for the Z21
@property (retain) SerialPortClass *usb;                // The USB or serial connection
@property (retain) AccAutoClass *AccAuto;               // Implements the AccessoryAuto commands
@property (assign) connectionTypeEnum connectionType;
@property (assign) int zentraleVersionNumber;           // ZentraleVersion as number: version 3.6 => 36
@property (readonly) NSString *instanceId;              // 2.0: the letter of this object in the logfile (A, B, ...)


// ****************************************************************************************************************
// ********************************************* METHODS FOR ESTABLISHING CONNECTIONS *****************************
// ****************************************************************************************************************
- (void)CmdConnectTcp:(NSString *)Host;                                         // Port 5550
- (void)CmdConnectTcp:(NSString *)Host Port:(NSString *)Port;
- (void)CmdInitUdp:(NSString *)Host;                                            // Port 21106
- (void)CmdInitUdp:(NSString *)Host Port:(NSString *)Port;
- (void)CmdConnectUsb:(NSString *)Port;                                         // Settings "115200,N,8,1"
- (void)CmdConnectUsb:(NSString *)Port Settings:(NSString *)Settings;           // Lenz uses 57600, Yamorc 115200
- (void)CmdConnectUsb:(NSString *)Port Settings:(NSString *)Settings Handshake:(BOOL)Handshake;   // 2.0
- (void)CmdConnectSerial:(NSString *)Port Settings:(NSString *)Settings Handshake:(BOOL)Handshake;

// ****************************************************************************************************************
// ********************************************* METHODS TO GENERATE INTERFACE SPECIFIC MESSAGES ******************
// ****************************************************************************************************************
- (void)CmdInterfaceVersion;                                                    // V3.6: 1.6
- (void)CmdXpressNetAddress;                                                    // V3.6: 1.7 (address = 0)
- (void)CmdXpressNetAddress:(int)address;
- (void)CmdInterfaceStatus;                                                     // V3.6: 2.1 (keep alive)
- (void)CmdXpressNetVersion;                                                    // V3.6: 2.2
- (void)CmdInterfaceFreieVerbindungen;                                          // V3.6: 2.3

// ****************************************************************************************************************
// ******************************** V3.6 & V4 METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ***************
// ****************************************************************************************************************
- (void)CmdAllesAn;
- (void)CmdAllesAus;
- (void)CmdAlleLoksAnhalten;
- (void)CmdEineLokAnhalten:(int)address;
- (void)CmdLeseanfrageProgrammierenRegistermode:(int)Register;
- (void)CmdLeseanfrageProgrammieren:(int)CV;
- (void)CmdLeseanfrageProgrammierenPagemode:(int)CV;
- (void)CmdProgrammierergebnisAnfordern;
- (void)CmdSchreibbefehlProgrammierenRegisterMode:(int)Register daten:(int)daten;
- (void)CmdSchreibbefehlProgrammieren:(int)CV Val:(int)Val;
- (void)CmdSchreibbefehlProgrammierenPageMode:(int)CV daten:(int)daten;
- (void)CmdSoftwareversionZentraleAnfordern;
- (void)CmdStatusZentraleAnfordern;
- (void)CmdZentralenStartmodeSetzen:(BOOL)AutoStart;
- (void)CmdSwitchInfoRequest:(int)address;                                      // Intended for switches
- (void)CmdFeedbackRequest:(int)address Nibble:(int)Nibble;                     // Intended for RS-Bus decoders
- (void)CmdAccessory:(int)address position:(int)position Activate:(int)Activate;
- (void)CmdAccessoryAuto:(int)Switch position:(int)position;
- (void)CmdLokInfoAnfordern:(int)address;
- (void)CmdFunktionsStatusAnfordern:(int)address;
- (void)CmdFunktionsStatusF13F28Anfordern:(int)address;
- (void)CmdFunktionsZustandF13F28Anfordern:(int)address;
- (void)CmdFunktionsStatusF29F68Anfordern:(int)address;
- (void)CmdFunktionsZustandF29F68Anfordern:(int)address;
- (void)CmdFahrbefehl:(int)address Speed:(int)Speed Steps:(int)Steps Forward:(BOOL)Forward;
- (void)CmdLokFunktionBefehl:(int)address group:(int)group data:(int)data;
- (void)CmdLokFunktionStatus:(int)address group:(int)group data:(int)data;
- (void)CmdPomWrite:(int)address CV:(int)CV value:(int)value;
- (void)CmdPomRead:(int)address CV:(int)CV;
- (void)CmdPomBitWrite:(int)address CV:(int)CV bitValue:(BOOL)bitValue bitPosition:(int)bitPosition;
- (void)CmdAdresAnfrageLokInStack:(int)address Vorwarts:(BOOL)Vorwarts;
- (void)CmdLokAusStack:(int)address;

// ****************************************************************************************************************
// ******************************** V4 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES *****************
// ****************************************************************************************************************
- (void)CmdErweiterteZentralenVersionsinformation;
- (void)CmdPoMErgebnisAnfordern;
- (void)CmdZentralenResetAuslosen;
- (void)CmdModellzeitAnfordern;
- (void)CmdModellzeitStellen:(int)day hour:(int)hour minute:(int)minute factor:(int)factor;
- (void)CmdModellzeitAnhalten;
- (void)CmdModellzeitStarten;

// ****************************************************************************************************************
// ******************************** Z21 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ****************
// ****************************************************************************************************************
- (void)CmdZ21SetBroadcastFlags:(int)Flags;                                     // 2.0
- (void)CmdZ21GetFirmwareVersion;
- (void)CmdZ21SetLocoFunction:(int)address CV:(int)CV TT:(int)TT;
- (void)CmdZ21SetExtAccessory:(int)address value:(int)value;
- (void)CmdZ21GetExtAccessory:(int)address;
- (void)CmdZ21AccPomWrite:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV value:(int)value;
- (void)CmdZ21AccPomBitWrite:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV bitValue:(BOOL)bitValue bitPosition:(int)bitPosition;
- (void)CmdZ21AccPomRead:(int)decoderAddress CDDD:(int)CDDD CV:(int)CV;
- (void)CmdZ21MmWriteByte:(int)RegAdr value:(int)value;

// ****************************************************************************************************************
// ******************************** OpenDCC ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES ************
// ****************************************************************************************************************
- (void)CmdOpenDCCSetExtAccessory:(int)address value:(int)value;

// Derive from the speed byte an integer value (for decoders with 27/28 speed steps)
- (int)DecodeSpeed:(int)Speed;


// ****************************************************************************************************************
// ************************** RAISING EVENTS: SHOULD NOT BE CALLED BY THE LIBRARY USER! ***************************
// ****************************************************************************************************************
// The methods below forward events from lower level objects to the library user.
// In addition Logfile messages are written, if needed.
- (void)eventInterfaceConnected;
- (void)eventTcpError:(NSString *)Description;
- (void)eventSendError:(SendErrorReasonEnum)reason description:(NSString *)description;   // 2.0
- (void)eventInterfaceNichtAdressiert;                                          // 2.0
- (void)eventInterfaceWiederAdressiert;                                         // 2.0
- (BOOL)isConnected;                                                            // 2.0: can we send?
- (void)eventTcpClosedByRemoteHost;
- (void)eventInterfaceVersion:(int)daten1 daten2:(int)daten2;
- (void)eventXpressNetAddress:(int)address;
- (void)eventXpressNetVersion:(int)daten;
- (void)eventInterfaceFreieVerbindungen:(int)number;
- (void)eventAllesAn;
- (void)eventNotAus;
- (void)eventNotHalt;
- (void)eventProgrammierMode;
- (void)eventFeedback:(int)address Nibble:(int)Nibble ttBits:(int)ttBits iBit:(int)iBit data:(int)data;
- (void)eventProgrammierinfoKurzschluss;
- (void)eventProgrammierinfoDatenNichtGefunden;
- (void)eventProgrammierinfoZentraleBusy;
- (void)eventProgrammierinfoZentraleBereit;
- (void)eventProgInfoPageRegister:(int)EEPromAddress value:(int)value;
- (void)eventProgInfo:(int)CvRange daten1:(int)daten1 value:(int)value;
- (void)eventProgInfoZ21:(int)CVAdr_MSB CVAdr_LSB:(int)CVAdr_LSB value:(int)value;
- (void)eventServiceVariableMelden:(int)SvAddress SvValue:(int)SvValue;
- (void)eventSoftwareVersionZentrale:(int)daten2 daten3:(int)daten3;
- (void)eventStatusZentrale:(int)data;
- (void)eventPomErgebnisMelden:(int)AdrH AdrL:(int)AdrL CvValue:(int)CvValue;
- (void)eventNormaleLokinfo:(int)Kennung Speed:(int)Speed F0_F4:(int)F0_F4 F5_F12:(int)F5_F12;
- (void)eventZ21Lokinfo:(int)Adr_MSB Adr_LSB:(int)Adr_LSB Kennung:(int)Kennung Speed:(int)Speed F0_F4:(int)F0_F4 F5_F12:(int)F5_F12 F13_F20:(int)F13_F20 F21_F28:(int)F21_F28;
- (void)eventFunctionsZustandF13F20:(int)F13_F20 F21_F28:(int)F21_F28;
- (void)eventFunctionsZustandF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68;
- (void)eventFunctionStatusF0F12:(int)F0_F4 F5_F12:(int)F5_F12;
- (void)eventFunctionStatusF13F20:(int)F13_F20 F21_F28:(int)F21_F28 RefreshMode:(int)RefreshMode;
- (void)eventFunctionStatusF29F68:(int)F29_F36 F37_F44:(int)F37_F44 F45_F52:(int)F45_F52 F53_F60:(int)F53_F60 F61_F68:(int)F61_F68;
- (void)eventFehlermeldung:(int)data;
- (void)eventModellZeit:(int)daten1 daten2:(int)daten2 factor:(int)factor;
- (void)eventErweiterteVersionZentrale:(int)ZBldH ZBldL:(int)ZBldL RMVer:(int)RMVer RMBldH:(int)RMBldH RMBldL:(int)RMBldL BIvers:(int)BIvers;
- (void)eventz21FirmwareVersion:(int)V_MSB V_LSB:(int)V_LSB;
- (void)eventExtendedAccessoryInfo:(int)Adr_MSB Adr_LSB:(int)Adr_LSB DDDDDDDD:(int)DDDDDDDD Status:(int)Status;

@end
