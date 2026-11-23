//
//  XpressNet.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  .NET version of the VB6 XpressNet DLL (XpressNetClass). Version 2.0, November 2026.
//  The library supports XpressNet over TCP (Lenz LAN/USB Interface 23151), over USB and over a serial interface,
//  as well as the XpressNet tunnel of the Z21 LAN protocol (UDP).
//
//  The various methods contain references to the V3.6, V4 and Z21 specifications:
//  - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
//    Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
//    https://www.lenz-elektronik.de/src/pdf/23151_Info.zip
//  - Dokumentation XpressNet Version 4.0, 02/2022:
//    https://www.lenz-elektronik.de/src/pdf/Lenz_XpressNet_Doku.pdf
//  - Z21 LAN Protocol Specification, Version 1.13, 06.11.2023
//
//  Convention (as in the DLL): everything intended for the library user starts with a capital letter.
//  Internal objects, methods and attributes start with lowercase.
//  - Methods, events and attributes: same names and parameters as in the DLL. VB6 Integer and Long become int,
//    Byte becomes byte, Boolean becomes bool and String becomes string. An Optional parameter keeps its default value.
//    For VB.NET applications this means that the calls and event handlers of a VB6 or VB.NET application with
//    the COM DLL can stay as they are.
//  - Events are raised on the main thread (the thread that created the XpressNetClass object). See XpressNetMainThread.cs
//
//  Differences with the VB6 DLL:
//  - The attribute ZentraleVersion is called ZentraleSoftwareVersion. In .NET an event and an attribute can not have
//    the same name; the event ZentraleVersion keeps its name.
//  - CmdConnectUsb and CmdConnectSerial: the parameter Settings is ByVal (VB6: ByRef). For VB.NET callers this
//    makes no difference.
//  - Opening a new connection first closes an existing connection.
//  - As long as the connection is not (yet) established, messages are dropped. In VB6 a message during the setup of
//    the TCP connection closed the connection (event TcpClosedByRemoteHost).
//  - Dispose closes the connection and stops the timers (VB6: Set XpressNet = Nothing).
//  - TcpClosedByRemoteHost: the connection is cleaned up before the event is raised (VB6: after), so the event
//    handler of the application may open a new connection.
//  - A parameter value that does not fit in a byte is not sent, but written to the logfile (VB6: run time error
//    "Overflow").
//  - An error (exception) in an event handler of the application does not disturb the library: the error is written
//    to the logfile, and passed back to the application after the library has finished handling the received message.
//  - Calls from another thread are passed to the main thread (see XpressNetMainThread.cs).
//  - The objects and subroutines for internal use (send, receive, tcp, eventFeedback, ...) are not visible outside
//    the library.
//
//  New in version 2.0 (November 2026), the same in the VB6, .NET and macOS versions:
//  - Retransmission of a message after an error answer of the interface, and a pause as long as the command
//    station does not address the interface (01 05 / 01 07). See XpressNetSend.cs
//  - The events SendError, InterfaceNichtAdressiert and InterfaceWiederAdressiert. SendError is raised for every
//    message that is not sent (not silently dropped). The reasons are in SendErrorReasonEnum
//  - CmdConnectUsb has the optional parameter Handshake (hardware handshake RTS/CTS, default off)
//  - The names of more command stations in the logfile (LokMaus II, LDH, multiMAUS)
//  - The default directory of the logfile is %PUBLIC%\TMC-Logfiles (see XpressNetLogFile.cs)
//  - CmdAccessoryAuto switches off the activated switch if the queue is emptied (see XpressNetAccAuto.cs)
//  - After a received message with an XOR error, the next framing bytes are searched (see XpressNetReceive.cs)
//

using System;
using System.Globalization;
using System.Runtime.ExceptionServices;

namespace XpressNet
{
  // ****************************************************************************************************************
  // ************************************************ EVENT TYPES ***************************************************
  // ****************************************************************************************************************
  // Interface: TCP as well as USB
  public delegate void InterfaceConnectedEventHandler();
  public delegate void TcpClosedByRemoteHostEventHandler();
  public delegate void TcpErrorEventHandler(ref string Description);
  // USB
  public delegate void UsbConnectedEventHandler();
  // XpressNet
  public delegate void AllesAnEventHandler();
  public delegate void NotAusEventHandler();
  public delegate void NotHaltEventHandler();
  public delegate void ProgrammierModeEventHandler();
  public delegate void XpressNetAddressEventHandler(byte address);
  public delegate void XpressNetVersionEventHandler(string Version);
  public delegate void InterfaceVersionEventHandler(string Version, string code);
  public delegate void InterfaceFreieVerbindungenEventHandler(byte number);
  public delegate void ZentraleVersionEventHandler(string Version);
  public delegate void ErweiterteZentraleVersionEventHandler(string BuildZentrale, string VersionRuckMelder, string BuildRuckMelder, string VersionBootloader);
  public delegate void FehlermeldungEventHandler(byte value);
  public delegate void ServiceVariableEventHandler(byte SvAddress, byte SvValue);
  // Feedback events
  public delegate void FeedbackEventHandler(int address, byte Nibble, byte data);
  public delegate void SwitchFeedbackEventHandler(int address, byte ttBits, byte iBit, byte data);
  public delegate void ExtendedAccessoryInfoEventHandler(int address, byte value, byte Status);
  // Lok events
  public delegate void NormaleLokinfoEventHandler(bool Besetzt, int Speedsteps, int Speed, bool Forward, byte F0_F4, byte F5_F8, byte F9_F12);
  public delegate void FunctionsZustandF13F20EventHandler(byte F13_F20, byte F21_F28);
  public delegate void FunctionsZustandF29F68EventHandler(byte F29_F36, byte F37_F44, byte F45_F52, byte F53_F60, byte F61_F68);
  public delegate void FunctionStatusF0F12EventHandler(byte F0_F4, byte F5_F8, byte F9_F12);
  public delegate void FunctionStatusF13F20EventHandler(byte F13_F20, byte F21_F28, byte RefreshMode);
  public delegate void FunctionStatusF29F68EventHandler(byte F29_F36, byte F37_F44, byte F45_F52, byte F53_F60, byte F61_F68);
  // Programmier events - Service Mode
  public delegate void ProgrammierinfoKurzschlussEventHandler();
  public delegate void ProgrammierinfoDatenNichtGefundenEventHandler();
  public delegate void ProgrammierinfoZentraleBusyEventHandler();
  public delegate void ProgrammierinfoZentraleBereitEventHandler();
  public delegate void ProgrammierinfoCVEventHandler(int CV, byte value);
  public delegate void ProgrammierinfoPageRegisterEventHandler(byte EEPromAddress, byte value);
  // Programmier events - Programming on Main (PoM)
  public delegate void PomErgebnisEventHandler(int LokAddress, byte CvValue);
  // Modellzeit events
  public delegate void ModellzeitEventHandler(byte day, byte hour, byte minute, byte factor);
  // Z21 specific events
  public delegate void z21FirmwareVersionEventHandler(string Version);
  public delegate void Z21LokinfoEventHandler(int address, bool Besetzt, int Speedsteps, int Speed, bool Forward, byte F0_F4, byte F5_F8, byte F9_F12, byte F13_F20, byte F21_F28);
  // 2.0: Sending
  public delegate void SendErrorEventHandler(byte Reason, string Description);
  public delegate void InterfaceNichtAdressiertEventHandler();
  public delegate void InterfaceWiederAdressiertEventHandler();


  // Enumeration to determine the type of connection
  public enum connectionTypeEnum
  {
    none,
    lanTcpPort,
    lanUdpPort,
    usbPort,
    serialPort
  }

  // 2.0: The values of the parameter Reason of the event SendError. The description (parameter Description) contains
  // the bytes of the message (including the XOR byte, without the framing bytes)
  // The names are those of the Lenz specification (Interface 23151, V1.2), where the specification has a name
  public enum SendErrorReasonEnum : byte
  {
    WiederholungErfolglos = 1,                    // Given up after 3 retransmissions (61 80, 61 81, 01 01, 01 02, 01 03,
                                                  // 01 06, 01 0A). Not a name of the specification
    BefehlNichtVorhanden = 2,                     // Answer 61 82 (V3.6: 3.1.7)
    Timeout = 3,                                  // No answer within 5 sec (90 sec in programming mode; V3.6: 1.3).
                                                  // All queues are emptied
    FifoVoll = 4,                                 // The FIFO (128 messages) or the CmdAccessoryAuto queue (32 commands) is
                                                  // full. Not a name of the specification
    KeineVerbindung = 5,                          // There is no connection, or the connection has not (yet) been
                                                  // established. Not a name of the specification
    FehlerInDenBefehlsparametern = 6,             // Answer 01 09 (V3.6: 1.5), or a parameter value out of range (checked
                                                  // by the library; the message is not sent)
    ZentraleAdressiertDasInterfaceNichtMehr = 7   // No 01 07 within 5 sec (90 sec in programming mode) after 01 05 / 01 08
                                                  // (V3.6: 1.5). All queues are emptied
  }


  public class XpressNetClass : IDisposable
  {
    // **************************************************************************************************************
    // *************************************************** EVENTS ***************************************************
    // **************************************************************************************************************
    // GENERATED EVENTS - Interface: TCP as well as USB
    public event InterfaceConnectedEventHandler InterfaceConnected;
    public event TcpClosedByRemoteHostEventHandler TcpClosedByRemoteHost;
    public event TcpErrorEventHandler TcpError;

    // GENERATED EVENTS - USB
#pragma warning disable 0067
    public event UsbConnectedEventHandler UsbConnected;   // As in the VB6 DLL: declared, but never raised
#pragma warning restore 0067

    // GENERATED EVENTS - XPRESSNET
    public event AllesAnEventHandler AllesAn;
    public event NotAusEventHandler NotAus;
    public event NotHaltEventHandler NotHalt;
    public event ProgrammierModeEventHandler ProgrammierMode;
    public event XpressNetAddressEventHandler XpressNetAddress;
    public event XpressNetVersionEventHandler XpressNetVersion;
    public event InterfaceVersionEventHandler InterfaceVersion;
    public event InterfaceFreieVerbindungenEventHandler InterfaceFreieVerbindungen;
    public event ZentraleVersionEventHandler ZentraleVersion;
    public event ErweiterteZentraleVersionEventHandler ErweiterteZentraleVersion;
    public event FehlermeldungEventHandler Fehlermeldung;
    public event ServiceVariableEventHandler ServiceVariable;
    // Feedback events
    public event FeedbackEventHandler Feedback;
    public event SwitchFeedbackEventHandler SwitchFeedback;
    public event ExtendedAccessoryInfoEventHandler ExtendedAccessoryInfo;
    // Lok events
    public event NormaleLokinfoEventHandler NormaleLokinfo;
    public event FunctionsZustandF13F20EventHandler FunctionsZustandF13F20;
    public event FunctionsZustandF29F68EventHandler FunctionsZustandF29F68;
    public event FunctionStatusF0F12EventHandler FunctionStatusF0F12;
    public event FunctionStatusF13F20EventHandler FunctionStatusF13F20;
    public event FunctionStatusF29F68EventHandler FunctionStatusF29F68;
    // Programmier events - Service Mode
    public event ProgrammierinfoKurzschlussEventHandler ProgrammierinfoKurzschluss;
    public event ProgrammierinfoDatenNichtGefundenEventHandler ProgrammierinfoDatenNichtGefunden;
    public event ProgrammierinfoZentraleBusyEventHandler ProgrammierinfoZentraleBusy;
    public event ProgrammierinfoZentraleBereitEventHandler ProgrammierinfoZentraleBereit;
    public event ProgrammierinfoCVEventHandler ProgrammierinfoCV;
    public event ProgrammierinfoPageRegisterEventHandler ProgrammierinfoPageRegister;
    // Programmier events - Programming on Main (PoM)
    public event PomErgebnisEventHandler PomErgebnis;
    // Modellzeit events
    public event ModellzeitEventHandler Modellzeit;
    // Z21 specific events
    public event z21FirmwareVersionEventHandler z21FirmwareVersion;
    public event Z21LokinfoEventHandler Z21Lokinfo;
    // 2.0: Sending (see XpressNetSend.cs)
    public event SendErrorEventHandler SendError;
    public event InterfaceNichtAdressiertEventHandler InterfaceNichtAdressiert;
    public event InterfaceWiederAdressiertEventHandler InterfaceWiederAdressiert;


    // **************************************************************************************************************
    // ************************************************* ATTRIBUTES *************************************************
    // **************************************************************************************************************
    // ATTRIBUTES THAT THE DLL USER MAY CHANGE. THE DEFAULT VALUES ARE SET IN THE CONSTRUCTOR
    public bool Z21Protocol { get; set; }               // Is the library used for Z21 or XpressNet??
    public bool LogMessages { get; set; }               // Example: connection requested to host X
    public bool LogWarnings { get; set; }               // May indicate a software problem
    public bool LogErrors { get; set; }                 // Indicates a hardware or software problem
    public bool LogHexMessages { get; set; }            // Hex dump of transmitted / received messages
    public bool LogDetails { get; set; }                // Description of what has been transmitted / received
    public int AccessoryInterval { get; set; }          // Interval between two accesory commands

    // ATTRIBUTES THAT THE DLL USER MAY READ
    public bool ZentraleNotAus { get; set; }
    public bool ZentraleNotHalt { get; set; }
    public bool ZentraleProgrammierMode { get; set; }
    public string ZentraleSoftwareVersion { get; set; } // VB6: ZentraleVersion (see the header of this file)
    public byte ZentraleKennung { get; set; }
    public bool ZentraleKurzSchluss { get; set; }       // Only for Z21
    public bool ZentraleStartmodeAuto { get; set; }
    public bool ZentraleKaltstart { get; set; }
    public bool ZentraleRamCheckFehler { get; set; }
    public int NumberOfExpressNetMessagesReceived { get; set; }
    public int NumberOfExpressNetMessagesTransmitted { get; set; }

    // OBJECTS THAT THE DLL USER MAY USE
    public LogFileClass Logfile { get; private set; }

    // The type of connection
    public connectionTypeEnum connectionType { get; set; }

    // Objects for internal use
    internal SendClass send;                            // Everything related to sending XpressNet messages
    internal ReceiveClass receive;                      // Everything related to receiving XpressNet messages
    internal ReceiveClassZ21 receiveZ21;                // To receive messages transported over a Z21 XpressNet tunnel
    internal TCPSocketClass tcp;                        // The TCP connection (Lenz LAN interface)
    internal UDPSocketClass udp;                        // The UDP connection (Z21)
    internal SerialPortClass usb;                       // The USB or serial connection
    internal AccAutoClass AccAuto;                      // Implements the AccessoryAuto commands
    internal MainThreadClass mainThread;                // The thread model: see XpressNetMainThread.cs

    // Known values for the Zentralen-Kennung
    private const byte LZ100 = 0x00;                    // LZ 100
    private const byte LH200 = 0x01;                    // LH 200
    private const byte DPC = 0x02;                      // DPC - Zentrale (Compact und Commander)
    private const byte CntrPlus = 0x03;                 // Control Plus - Zentrale
    private const byte LokMausII = 0x04;                // 2.0: Roco LokMaus II (source: JMRI XNetConstants)
    private const byte LDH = 0x07;                      // 2.0: LDH (source: JMRI XNetConstants)
    private const byte multiMAUS = 0x10;                // 2.0: Roco multiMAUS (source: JMRI XNetConstants)
    private const byte Z21 = 0x12;                      // Z21 / Yamorc
    private const byte OpenDcc = 0x20;                  // See: https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html
    private const byte OpenDccXp = 0x22;                // See: https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html

    private bool _disposed = false;


    // **************************************************************************************************************
    // ******************************************* CONSTRUCTOR AND DISPOSE ******************************************
    // **************************************************************************************************************
    // VB6: Class_Initialize
    // 2.0: the letter of this object in the logfile: A for the first XpressNet object of the program, B for the second,
    // and so on (after Z: AA, AB, ...). Two programs that write in the same logfile both start with A
    private static int instanceCounter = 0;
    internal readonly string instanceId;

    private static string letters(int n)
    {
      string s = "";
      while (n > 0) {
        n = n - 1;
        s = (char)('A' + (n % 26)) + s;
        n = n / 26;
      }
      return s;
    }

    public XpressNetClass()
    {
      instanceId = letters(System.Threading.Interlocked.Increment(ref instanceCounter));
      mainThread = new MainThreadClass();
      connectionType = connectionTypeEnum.none;           // We start without any connection
      // Logfile specific initialization. Logfile is created first, since the other objects may use it
      Logfile = new LogFileClass(this);
      send = new SendClass(this);
      receive = new ReceiveClass(this);
      receiveZ21 = new ReceiveClassZ21(this);
      AccAuto = new AccAutoClass(this);
      Logfile.OpenLogfile();                              // Check Logfile size, and rotate if needed
      LogMessages = true;                                 // Set the defaults for what messages will be logged.
      LogErrors = true;
      LogWarnings = true;
      LogHexMessages = true;
      LogDetails = true;
      // Attributes the DLL user may change
      Z21Protocol = false;
      AccessoryInterval = 125;                            // Time in msec between two accessory commands
      // Attributes the DLL user may read
      ZentraleNotAus = false;
      ZentraleNotHalt = false;
      ZentraleKurzSchluss = false;                        // Only for Z21
      ZentraleProgrammierMode = false;
      ZentraleSoftwareVersion = "0.0";                    // Give initial version, to avoid program crash if interface is not open
      ZentraleKennung = 0;
      ZentraleStartmodeAuto = false;
      ZentraleKaltstart = false;
      ZentraleRamCheckFehler = false;
      NumberOfExpressNetMessagesReceived = 0;
      NumberOfExpressNetMessagesTransmitted = 0;
    }

    // VB6: Class_Terminate (Set XpressNet = Nothing)
    // Closes the connection and stops the timers. Afterwards the object can not be used anymore
    public void Dispose()
    {
      if (otherThread(() => Dispose())) return;
      if (_disposed) return;
      _disposed = true;
      Logfile.Details("Close XpressNet interface");
      closeConnections();                                 // Before leaving, make sure we close any open connections
      send.close();                                       // Stop the timers
      AccAuto.close();
      mainThread.Closed = true;                           // Actions that are still underway will be ignored
    }


    // ==============================================================================================================
    // METHODS FOR ESTABLSHING CONNECTIONS
    // Note: opening a new connection first closes an existing connection
    public void CmdConnectTcp(string Host, string Port = "5550")
    {
      if (otherThread(() => CmdConnectTcp(Host, Port))) return;
      if (_disposed) return;
      closeConnections();
      Logfile.myHost = Host;                              // Save the host (IP address) to include in the Logfile
      Logfile.Details("TCP connection requested to Host: " + Host + " at Port: " + Port);
      connectionType = connectionTypeEnum.lanTcpPort;     // We connect via TCP over a LAN
      tcp = new TCPSocketClass(this);                     // Instantiated a new TCP class object
      tcp.connect(Host, Port);                            // Send a connection request
    }

    public void CmdInitUdp(string Host, string Port = "21106")
    {
      if (otherThread(() => CmdInitUdp(Host, Port))) return;
      if (_disposed) return;
      closeConnections();
      Logfile.myHost = Host;                              // Save the host (IP address) to include in the Logfile
      Logfile.Details("UDP initialised: Host: " + Host + ", Port: " + Port);
      connectionType = connectionTypeEnum.lanUdpPort;     // We connect via UDP over a LAN
      udp = new UDPSocketClass(this);                     // Instantiated a new UDP class object
      Z21Protocol = true;
      udp.init(Host, Port);                               // Store the address and port
      // 2.0: the first message is LAN_SET_BROADCASTFLAGS (Z21 spec 2.16). Without these flags the Z21 sends no
      // broadcasts, such as LAN_X_BC_TRACK_POWER_ON/OFF (events AllesAn / NotAus) and LAN_X_TURNOUT_INFO.
      // The flags must be set again after every login; the login is done with the first command (Z21 spec 2.2)
      CmdZ21SetBroadcastFlags(Z21BroadcastFlagsDefault);
      CmdSoftwareversionZentraleAnfordern();
      CmdStatusZentraleAnfordern();
      CmdZ21GetFirmwareVersion();
    }

    // 2.0: the optional parameter Handshake (default: off). The Lenz documentation of the LI-USB and the 23151 says
    // no handshake; JMRI uses hardware handshake (RTS/CTS) for the LI-USB 23150. See SerialPort.cs
    public void CmdConnectUsb(byte Port, string Settings = "115200,N,8,1", bool Handshake = false)
    {
      // Lenz uses 57600, Yamorc 115200
      if (otherThread(() => CmdConnectUsb(Port, Settings, Handshake))) return;
      if (_disposed) return;
      closeConnections();
      Logfile.myHost = "COM" + Port;                      // Save the Com port to include in the Logfile
      Logfile.Details("USB connect attempt using Port=" + Port + ", Settings=" + Settings + " and Handshake=" + Handshake);
      connectionType = connectionTypeEnum.usbPort;        // We connect via USB
      usb = new SerialPortClass(this);
      usb.openPort(Port, Settings, Handshake);            // Send a connection request
    }

    public void CmdConnectSerial(byte Port, string Settings, bool Handshake)
    {
      if (otherThread(() => CmdConnectSerial(Port, Settings, Handshake))) return;
      if (_disposed) return;
      closeConnections();
      Logfile.myHost = "COM" + Port;                      // Save the Com port to include in the Logfile
      Logfile.Details("Serial connect attempt using Port=" + Port + ", Settings=" + Settings + " and Handshake=" + Handshake);
      connectionType = connectionTypeEnum.serialPort;     // We connect via a serial (RS232) connection
      usb = new SerialPortClass(this);                    // Note that usb and serial use the same class and object
      usb.openPort(Port, Settings, Handshake);            // Send a connection request
    }


    // ==============================================================================================================
    // METHODS TO GENERATE INTERFACE SPECIFIC MESSAGES
    // The various subroutines below contain references to the V3.6 and V4 XpressNet specifications:
    // - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
    //   Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
    //   https://www.lenz-elektronik.de/src/pdf/23151_Info.zip
    // - Dokumentation XpressNet Version 4.0, 02/2022:
    //   https://www.lenz-elektronik.de/src/pdf/Lenz_XpressNet_Doku.pdf
    // For details, see these documents

    // V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
    public void CmdInterfaceVersion()
    {
      send.message(0xF0);
      Logfile.DetailsSent("Versionsnummer des Interface 23151 feststellen");
    }

    // V3.6: 1.7 (XpressNet Gerateadresse des Interface 23151 feststellen und andern)
    public void CmdXpressNetAddress(byte address = 0)
    {
      send.message(0xF2, 0x01, address);
      Logfile.DetailsSent("XpressNet Gerateadresse des Interface 23151 feststellen");
    }

    // V3.6: 2.1 (Interface Status Befehl - Keep Alive) - Need not be used by the Application
    public void CmdInterfaceStatus()
    {
      send.message(0xF1, 0x01);
    }

    // V3.6: 2.2 (XpressNet Version Befehl)
    public void CmdXpressNetVersion()
    {
      send.message(0xF1, 0x02);
      Logfile.DetailsSent("XpressNet Version Befehl");
    }

    // V3.6: 2.3 (Verfugbare Freie Verbindungen)
    public void CmdInterfaceFreieVerbindungen()
    {
      send.message(0xF1, 0x03);
      Logfile.DetailsSent("Verfugbare Freie Verbindungen");
    }


    // ==============================================================================================================
    // V3.6 & V4 METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
    // The commands below will be forwarded by the Interface to the Command Station

    // V3.6: 3.2.1 / V4: 3.2 (Alles An)
    public void CmdAllesAn()
    {
      send.message(0x21, 0x81);
      Logfile.DetailsSent("Alles An");
    }

    // V3.6: 3.2.2 / V4: 3.3 (Alles Aus - Notaus)
    public void CmdAllesAus()
    {
      send.message(0x21, 0x80);
      Logfile.DetailsSent("Alles Aus - Notaus");
    }

    // V3.6: 3.2.3 / V4: 3.4 (Alle Loks anhalten - Nothalt)
    public void CmdAlleLoksAnhalten()
    {
      send.message(0x80);
      Logfile.DetailsSent("Alle Loks anhalten - Nothalt");
    }

    // V3.6: 3.2.4 / V4: 3.6 & 3.7 (Eine Lok anhalten - Nothalt fur eine Lok)
    public void CmdEineLokAnhalten(int address)
    {
      send.message(0x92, AH(address), AL(address));
      Logfile.DetailsSent("Eine lok anhalten. Adresse: " + address);
    }

    // V3.6: 3.2.5 / V4: 3.9
    public void CmdLeseanfrageProgrammierenRegistermode(byte Register)
    {
      if (VelueOutOfRange(Register, 1, 8, "Leseanfrage Programmieren Registermode: Register")) return;
      send.message(0x22, 0x11, Register);
      Logfile.DetailsSent("Leseanfrage Programmieren Registermode. Register: " + Register);
    }

    // V3.6: 3.2.6-3.2.10 / V4: 3.10-3.14 (Leseanfrage Programmieren 4-Byte-Format / LAN_X_CV_READ )
    public void CmdLeseanfrageProgrammieren(int CV)
    {
      if (VelueOutOfRange(CV, 1, 1024, "Leseanfrage Programmieren: CV")) return;
      if (Z21Protocol) {
        CV = CV - 1;
        send.message(0x23, 0x11, (CV / 256), (CV % 256));
        Logfile.DetailsSent("LAN_X_CV_READ, CV: " + (CV + 1));
      }
      else {
        // The version is compared as a number (see Val)
        if (Val(ZentraleSoftwareVersion) < 3.6) {
          send.message(0x22, 0x15, (CV % 256));
        }
        else {
          int daten1 = 0x18 + ((CV / 256) % 4);
          send.message(0x22, daten1, (CV % 256));
        }
        Logfile.DetailsSent("Leseanfrage Programmieren 4-Byte-Format. CV: " + CV);
      }
    }

    // V3.6: 3.2.11 / V4: 3.15 (Leseanfrage Programmieren 3-Byte-Format - Pagemode)
    public void CmdLeseanfrageProgrammierenPagemode(int CV)
    {
      if (VelueOutOfRange(CV, 1, 256, "Leseanfrage Programmieren Pagemode: CV")) return;
      send.message(0x22, 0x14, (CV % 256));
      Logfile.DetailsSent("Leseanfrage Programmieren Pagemode. CV: " + CV);
    }

    // V3.6: 3.2.12 / V4: 3.16
    public void CmdProgrammierergebnisAnfordern()
    {
      send.message(0x21, 0x10);
      Logfile.DetailsSent("Programmier Ergebnis Anfordern");
    }

    // V3.6: 3.2.13 / V4: 3.17 (Schreibbefehl Programmieren 3-Byte-Format (Register-Mode) / LAN_X_DCC_WRITE_REGISTER)
    public void CmdSchreibbefehlProgrammierenRegisterMode(byte Register, int daten)
    {
      if (VelueOutOfRange(Register, 1, 8, "Schreibbefehl Programmieren Registermode: Register")) return;
      send.message(0x23, 0x12, Register, daten);
      Logfile.DetailsSent("Schreibbefehl Programmieren RegisterMode. Register: " + Register + ", Value: " + daten);
    }

    // V3.6: 3.2.14-3.2.18 / V4: 3.18-3.22 (Schreibbefehl Programmieren 4-Byte-Format / LAN_X_CV_WRITE)
    public void CmdSchreibbefehlProgrammieren(int CV, byte Val)
    {
      if (VelueOutOfRange(CV, 1, 1024, "Schreibbefehl Programmieren: CV")) return;
      if (VelueOutOfRange(Val, 0, 255, "Schreibbefehl Programmieren: Daten: ")) return;
      if (Z21Protocol) {
        CV = CV - 1;
        send.message(0x24, 0x12, (CV / 256), (CV % 256), Val);
        Logfile.DetailsSent("LAN_X_CV_WRITE, CV: " + (CV + 1) + ", Value: " + Val);
      }
      else {
        // The version is compared as a number (see Val)
        if (XpressNetClass.Val(ZentraleSoftwareVersion) < 3.6) {
          send.message(0x23, 0x16, (CV % 256), Val);
        }
        else {
          int daten1 = 0x1C + ((CV / 256) % 4);
          send.message(0x23, daten1, (CV % 256), Val);
        }
        Logfile.DetailsSent("Schreibbefehl Programmieren 4-Byte-Format. CV: " + CV + ", Value: " + Val);
      }
    }

    // V3.6: 3.2.19 / V4: 3.23 (Schreibbefehl Programmieren 3-Byte-Format - Page-Mode)
    public void CmdSchreibbefehlProgrammierenPageMode(int CV, int daten)
    {
      if (VelueOutOfRange(CV, 1, 256, "Schreibbefehl Programmieren Pagemode: CV")) return;
      send.message(0x23, 0x17, (CV % 256), daten);
      Logfile.DetailsSent("Schreibbefehl Programmieren Pagemode. CV: " + CV + ", Value: " + daten);
    }

    // V3.6: 3.2.20 / V4: 3.24 (Softwareversion der Zentrale anfordern)
    public void CmdSoftwareversionZentraleAnfordern()
    {
      send.message(0x21, 0x21);
      Logfile.DetailsSent("Softwareversion der Zentrale anfordern");
    }

    // V3.6: 3.2.21 / V4: 3.26 (Status der Zentrale anfordern)
    public void CmdStatusZentraleAnfordern()
    {
      send.message(0x21, 0x24);
      // This command may also be internally used as UDP (Z21) Keep Alive Message
      // In that case we should not log this like commands requested by the application
      if (send.waitForlanKeepAliveResponse == false) {
        Logfile.DetailsSent("Status der Zentrale anfordern");
      }
    }

    // V3.6: 3.2.22 / V4: 3.27 (Zentralen-Startmode setzen)
    public void CmdZentralenStartmodeSetzen(bool AutoStart)
    {
      byte value;
      if (AutoStart) value = 2;
      else value = 0;
      send.message(0x22, 0x22, value);
      Logfile.DetailsSent("Zentralen-Startmode setzen. Automatischer Start: " + AutoStart);
    }

    // V3.6: 3.2.23 / V4: 3.37 & 3.38 / Z21: 5.1 (Schaltinformationen anfordern / LAN_X_GET_TURNOUT_INFO)
    // Note: this command is intended for switches.
    // The (switch) address gets divided by 4 to obtain the decoder address
    public void CmdSwitchInfoRequest(int address)
    {
      // STEP 1: Check the address size against the ZentraleVersion and Protocol (compared as a number, see Val)
      if ((Val(ZentraleSoftwareVersion) < 3.8) && (!Z21Protocol)) {
        if (VelueOutOfRange(address, 1, 1024, "Accessory: Address")) return;
      }
      if (VelueOutOfRange(address, 1, 2048, "Accessory: Address")) return;
      // Switch numbering on the wire starts at 0. The LH100 numbering starts at 1
      int wireAddress = address - 1;
      // We need to determine the nibble for this specific switch.
      // We use temp variable to calculate the offset of the switch in this decoder
      int temp;                                           // a value between 0 and 3
      byte Nibble = 0;
      temp = wireAddress % 4;                             // the result should be 0..3
      if (temp > 1) Nibble = 1;                           // 0 and 1 become 0; 2 and 3 become 1
      // Send the command that fits for this ZentraleVersion and Protocol.
      if (Z21Protocol) {
        // Note: the XpressNet length, as contained in the four low order bits of the header,
        // seem to be incorrect! However, this is what the Z21 spec says...
        send.message(0x43, AccAH(wireAddress), AccAL(wireAddress));
      }
      else {
        // We should set the highest bit of the byte that contains the nibble
        if (Val(ZentraleSoftwareVersion) < 3.8) {
          send.message(0x42, AccAL(wireAddress), Nibble + 128);
        }
        else {                                            // Lenz V4 spec says Command station >= 3.8 SHOULD use the new form
          send.message(0x43, AccAH(wireAddress), AccAL(wireAddress), Nibble + 128);
        }
      }
      // STEP 5: Write the command details to the logfile
      Logfile.DetailsSent("Feedback Request - Switch: " + address);
    }

    // V3.6: 3.2.23 / V4: 3.37 & 3.38 (Schaltinformationen anfordern)
    // This command is intended for feedback (RS-Bus) decoders
    public void CmdFeedbackRequest(int address, byte Nibble)
    {
      // STEP 1: Check the address size.
      if (VelueOutOfRange(address, 1, 128, "RS-Bus Feedback: Address")) return;
      if (Nibble != 0) Nibble = 1;                        // Ensure the nibble is 0 or 1
      // STEP 2: Call the associated ExpressNet command
      // Substract 1 from the address. The LH100 numbering starts at 1, the wire at 0
      // Also set the highest bit of the nibble
      send.message(0x42, address - 1, Nibble + 128);
      // STEP 3: Write the command details to the logfile
      Logfile.DetailsSent("Feedback Request for address: " + address + "  Nibble: " + Nibble);
    }

    // V3.6: 3.2.24 / V4: 3.39 & 3.40 / Z21: 5.2 (Schaltbefehl / LAN_X_SET_TURNOUT)
    public void CmdAccessory(int address, byte position, byte Activate)
    {
      // 2.0: comment corrected. Position 0 = diverging (red, -), 1 = straight (green, +), according to RCN-213. Some
      // command stations interpret 0 and 1 the other way around (see the README of the AP_DCC_library)
      // Was: Position 0 = straight / 1 = turn
      // Check the address size against the ZentraleVersion (compared as a number, see Val)
      if (Val(ZentraleSoftwareVersion) < 3.8) {
        if (VelueOutOfRange(address, 1, 1024, "Accessory: Address")) return;
      }
      else {
        if (VelueOutOfRange(address, 1, 2048, "Accessory: Address")) return;
      }
      address = address - 1;                              // The LH100 numbering starts at 1, the wire at 0
      int daten = 0;                                      // Fill the Daten byte
      if (!Z21Protocol) {                                 // The Z21 uses switch addressing, instead of decoder addressing
        daten = (address % 4) * 2;                        // Store the 2 address bits that belong to this decoder in daten
      }
      daten = 128 + daten;                                // Set the MSB
      if (Activate > 0) daten = daten + 8;                // Set the D1 bit (Activation)
      if (position > 0) daten = daten + 1;                // Set the D2 bit (Position)
      // Send the command that fits for this ZentraleVersion and Kennung.
      // The IF clause is used for the LZV 200 and the Z21
      // The Else clause is used for LZV100
      if ((Val(ZentraleSoftwareVersion) > 3.6) || Z21Protocol) {
        send.message(0x53, AccAH(address), AccAL(address), daten);
      }
      else {
        send.message(0x52, AccAL(address), daten);
      }
      // Write the command details to the logfile
      Logfile.DetailsSent("Accessory Command - Address: " + (address + 1) + "  Position: " + position + "  Activate: " + Activate);
    }

    // V3.6: 3.2.24 / V4: 3.39 & 3.40 (Schaltbefehl)
    public void CmdAccessoryAuto(int Switch, byte position)
    {
      // Implemented in the AccAutoClass, Since it is not a standard XpressNet command.
      AccAuto.handle(Switch, position);
    }

    // V3.6: 3.2.25 / V4: 3.41.3 / Z21: 4.1 (Lokinformationen anfordern)
    public void CmdLokInfoAnfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Lok Info Anfordern: Address")) return;
      if (Z21Protocol) {
        send.message(0xE3, 0xF0, AH(address), AL(address));
      }
      else {
        send.message(0xE3, 0x00, AH(address), AL(address));
      }
      Logfile.DetailsSent("Lok Informationen anfordern. Address: " + address);
    }

    // V3.6: 3.2.25.1 / V4: 3.41.4 (Funktionsstatus anfordern F0 - F12)
    public void CmdFunktionsStatusAnfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Funktionsstatus Anfordern: Address")) return;
      send.message(0xE3, 0x07, AH(address), AL(address));
      Logfile.DetailsSent("Funktionsstatus F0-F12 anfordern. Address: " + address);
    }

    // V3.6: 3.2.25.2 / V4: 3.41.5 (Funktionsstatus anfordern F13 - F28)
    public void CmdFunktionsStatusF13F28Anfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Funktionsstatus F13-F28 Anfordern: Address")) return;
      send.message(0xE3, 0x08, AH(address), AL(address));
      Logfile.DetailsSent("Funktionsstatus F13-F28 anfordern. Address: " + address);
    }

    // V3.6: 3.2.25.3 / V4: 3.41.7 (Funktionszustand anfordern F13 - F28)
    public void CmdFunktionsZustandF13F28Anfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Funktionszustand F13-F28 Anfordern: Address")) return;
      send.message(0xE3, 0x09, AH(address), AL(address));
      Logfile.DetailsSent("Funktionszustand F13-F28 anfordern. Address: " + address);
    }

    // V4: 3.41.6 (Funktionsstatus anfordern F29 - F68)
    public void CmdFunktionsStatusF29F68Anfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Funktionsstatus F29-F68 Anfordern: Address")) return;
      send.message(0xE3, 0x0A, AH(address), AL(address));
      Logfile.DetailsSent("Funktionsstatus F29-F68 anfordern. Address: " + address);
    }

    // V4: 3.41.8 (Funktionszustand anfordern F29 - F68)
    public void CmdFunktionsZustandF29F68Anfordern(int address)
    {
      if (VelueOutOfRange(address, 0, 9999, "Funktionszustand F29-F68 Anfordern: Address")) return;
      send.message(0xE3, 0x0B, AH(address), AL(address));
      Logfile.DetailsSent("Funktionszustand F29-F68 anfordern. Address: " + address);
    }

    // V3.6: 3.2.26.1 / V4: 3.42.3 / Z21: 4.2 (Lok Fahrbefehl / LAN_X_SET_LOCO_DRIVE )
    public void CmdFahrbefehl(int address, int Speed, int Steps, bool Forward)
    {
      int RVVV = 0;                                       // The direction and speed byte
      if (VelueOutOfRange(address, 0, 9999, "Fahrbefehl: Address")) return;
      switch (Steps) {
        case 14:
          if (VelueOutOfRange(Speed, 0, 14, "Fahrbefehl: Speed")) return;
          if (Speed > 0) RVVV = Speed + 1;
          if (Forward) RVVV = RVVV + 128;
          send.message(0xE4, 0x10, AH(address), AL(address), RVVV);
          break;
        case 27:
          if (VelueOutOfRange(Speed, 0, 27, "Fahrbefehl: Speed")) return;
          RVVV = EncodeSpeed(Speed);
          if (Forward) RVVV = RVVV + 128;                 // The forward / reverse bit
          send.message(0xE4, 0x11, AH(address), AL(address), RVVV);
          break;
        case 28:
          if (VelueOutOfRange(Speed, 0, 28, "Fahrbefehl: Speed")) return;
          RVVV = EncodeSpeed(Speed);
          if (Forward) RVVV = RVVV + 128;
          send.message(0xE4, 0x12, AH(address), AL(address), RVVV);
          break;
        case 128:
          if (VelueOutOfRange(Speed, 0, 126, "Fahrbefehl: Speed")) return;
          RVVV = Speed;
          if (Speed > 0) RVVV = RVVV + 1;
          if (Forward) RVVV = RVVV + 128;
          send.message(0xE4, 0x13, AH(address), AL(address), RVVV);
          break;
        default:                                          // 2.0: not silently (1.3: nothing was sent, without an error)
          Logfile.AppendWithHost("Error Fahrbefehl: Steps must be 14, 27, 28 or 128");
          eventSendError(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Message not sent: Fahrbefehl: Steps must be 14, 27, 28 or 128 (value: " + Steps + ")");
          return;
      }
      Logfile.DetailsSent("Fahrbefehl - Address: " + address + ", Speed: " + Speed);
    }

    // V3.6: 3.2.26.2 / V4: 3.42.4 (Lok Funktionsbefehl)
    public void CmdLokFunktionBefehl(int address, byte group, byte data)
    {
      // 2.0: a group outside 1..10 gives the event SendError (1.3: nothing was sent, without an error)
      if (VelueOutOfRange(group, 1, 10, "Funktions befehl: Group")) return;
      switch (group) {
        case 1: send.message(0xE4, 0x20, AH(address), AL(address), data); break;
        case 2: send.message(0xE4, 0x21, AH(address), AL(address), data); break;
        case 3: send.message(0xE4, 0x22, AH(address), AL(address), data); break;
        case 4: send.message(0xE4, 0x23, AH(address), AL(address), data); break;
        case 5: send.message(0xE4, 0x28, AH(address), AL(address), data); break;
        case 6: send.message(0xE4, 0x29, AH(address), AL(address), data); break;
        case 7: send.message(0xE4, 0x2A, AH(address), AL(address), data); break;
        case 8: send.message(0xE4, 0x2B, AH(address), AL(address), data); break;
        case 9: send.message(0xE4, 0x50, AH(address), AL(address), data); break;
        case 10: send.message(0xE4, 0x51, AH(address), AL(address), data); break;
      }
      Logfile.DetailsSent("Funktions befehl. Address: " + address + ", Befehl: " + data);
    }

    // V3.6: 3.2.26.4 / V4: 3.42.6 (Funktionsstatus setzen)
    public void CmdLokFunktionStatus(int address, byte group, byte data)
    {
      // 2.0: a group outside 1..10 gives the event SendError (1.3: nothing was sent, without an error)
      if (VelueOutOfRange(group, 1, 10, "Status befehl: Group")) return;
      switch (group) {
        case 1: send.message(0xE4, 0x24, AH(address), AL(address), data); break;
        case 2: send.message(0xE4, 0x25, AH(address), AL(address), data); break;
        case 3: send.message(0xE4, 0x26, AH(address), AL(address), data); break;
        case 4: send.message(0xE4, 0x27, AH(address), AL(address), data); break;
        case 5: send.message(0xE4, 0x2C, AH(address), AL(address), data); break;
        case 6: send.message(0xE4, 0x2D, AH(address), AL(address), data); break;
        case 7: send.message(0xE4, 0x2E, AH(address), AL(address), data); break;
        case 8: send.message(0xE4, 0x52, AH(address), AL(address), data); break;
        case 9: send.message(0xE4, 0x53, AH(address), AL(address), data); break;
        case 10: send.message(0xE4, 0x54, AH(address), AL(address), data); break;
      }
      Logfile.DetailsSent("Status befehl. Address: " + address + ", Status: " + data);
    }

    // V3.6: 3.2.28.1 / V4: 3.43.1 (Programming on Main Byte schreiben / LAN_X_CV_POM_WRITE_BYTE)
    public void CmdPomWrite(int address, int CV, byte value)
    {
      if (VelueOutOfRange(address, 0, 9999, "PoM write: Address")) return;
      if (VelueOutOfRange(CV, 1, 1024, "PoM write: CV")) return;
      CV = CV - 1;
      int cvlow = CV % 256;                               // The low byte of the CV number
      int cvhigh = CV / 256 + 0xEC;                       // 0xEC = PoM Schreiben
      send.message(0xE6, 0x30, AH(address), AL(address), cvhigh, cvlow, value);
      Logfile.DetailsSent("Pom Write. Address: " + address + ", CV: " + (CV + 1) + ", Value: " + value);
    }

    // V3.6: 3.2.28.2 / V4: 3.43.2 (Programming on Main Byte lesen / LAN_X_CV_POM_READ_BYTE)
    public void CmdPomRead(int address, int CV)
    {
      if (VelueOutOfRange(address, 0, 9999, "PoM Read: Address")) return;
      if (VelueOutOfRange(CV, 1, 1024, "PoM Read: CV")) return;
      CV = CV - 1;
      int cvlow = CV % 256;                               // The low byte of the CV number
      int cvhigh = CV / 256 + 0xE4;                       // 0xE4 = PoM Lesen
      send.message(0xE6, 0x30, AH(address), AL(address), cvhigh, cvlow, 0);
      Logfile.DetailsSent("Pom Read. Address: " + address + ", CV: " + (CV + 1));
    }

    // V3.6: 3.2.28.3 / V4: 3.43.3 (Programming on Main Bit schreiben / LAN_X_CV_POM_WRITE_BIT)
    public void CmdPomBitWrite(int address, int CV, bool bitValue, byte bitPosition)
    {
      if (VelueOutOfRange(address, 0, 9999, "PoM bit write: Address")) return;
      if (VelueOutOfRange(CV, 1, 1024, "PoM bit write: CV")) return;
      if (VelueOutOfRange(bitPosition, 0, 7, "PoM bit write: bit position")) return;
      CV = CV - 1;
      int cvlow = CV % 256;                               // The low byte of the CV number
      // 0xE8 = PoM Bit Schreiben, as in the V3.6 spec (3.2.28.3) and the Z21 spec. The V4 spec (3.43.3) says 0x7C,
      // which seems to be an error in the V4 spec
      int cvhigh = CV / 256 + 0xE8;
      int value = bitPosition + 0xF0;                     // highest order bits must all be 1
      if (bitValue) value = value + 8;                    // bit 4 represents the new bit value
      send.message(0xE6, 0x30, AH(address), AL(address), cvhigh, cvlow, value);
      Logfile.DetailsSent("Pom Bit Write. Address: " + address + ", CV: " + (CV + 1) + ", Bit value: " + bitValue + ", Bit Position: " + bitPosition);
    }

    // V3.6 3.2.30.3 / V4: 3.44.3 (Adressanfrage Lok in Zentralenstack )
    public void CmdAdresAnfrageLokInStack(int address, bool Vorwarts)
    {
      byte Kennung;
      if (Vorwarts) Kennung = 5;
      else Kennung = 6;
      send.message(0xE3, Kennung, AH(address), AL(address));
      Logfile.DetailsSent("Adressanfrage Lok in Zentralenstack. Address: " + address + ", Vorwarts : " + Vorwarts);
    }

    // V3.6: 3.2.31 / V4: 3.44.4 (Lok aus Stack loeschen / LAN_X_PURGE_LOCO)
    public void CmdLokAusStack(int address)
    {
      send.message(0xE3, 0x44, AH(address), AL(address));
      Logfile.DetailsSent("Lok aus Stack loeschen. Address: " + address);
    }


    // ==============================================================================================================
    // V4 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
    // The commands below will be forwarded by the Interface to the Command Station

    // V4: 3.25 (Erweiterte Zentralen Versionsinformation)
    public void CmdErweiterteZentralenVersionsinformation()
    {
      send.message(0x21, 0x23);
      Logfile.DetailsSent("Erweiterte Zentralen Versionsinformation");
    }

    // V4: 3.30 (PoM-Ergebnis anfordern)
    public void CmdPoMErgebnisAnfordern()
    {
      send.message(0x21, 0x27);
      Logfile.DetailsSent("PoM Programmier Ergebnis Anfordern");
    }

    // V4: 3.31 (Zentralen Reset auslosen)
    public void CmdZentralenResetAuslosen()
    {
      send.message(0x21, 0x28);
      Logfile.DetailsSent("Zentralen Reset auslosen");
    }

    // V4: 3.33 (Modellzeit anfordern)
    public void CmdModellzeitAnfordern()
    {
      send.message(0x21, 0x2A);
      Logfile.DetailsSent("Modellzeit anfordern");
    }

    // V4: 3.34 (Modellzeit stellen)
    public void CmdModellzeitStellen(byte day, byte hour, byte minute, byte factor)
    {
      if (VelueOutOfRange(day, 0, 6, "Modellzeit: Day")) return;
      if (VelueOutOfRange(hour, 0, 23, "Modellzeit: Hour")) return;
      if (VelueOutOfRange(minute, 0, 59, "Modellzeit: Minute")) return;
      if (VelueOutOfRange(factor, 0, 31, "Modellzeit: Factor")) return;
      int daten = day * 32 + hour;
      send.message(0x24, 0x2B, daten, minute, factor);
      Logfile.DetailsSent("Modellzeit stellen");
    }

    // V4: 3.35 (Modellzeit anhalten - V4.0)
    public void CmdModellzeitAnhalten()
    {
      send.message(0x21, 0x2D);
      Logfile.DetailsSent("Modellzeit anhalten");
    }

    // V4: 3.36 (Modellzeit starten - V4.0)
    public void CmdModellzeitStarten()
    {
      send.message(0x21, 0x2C);
      Logfile.DetailsSent("Modellzeit starten");
    }


    // ==============================================================================================================
    // Z21 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
    // The commands below will be forwarded by the Interface to the Command Station
    // Note: these Z21 could not be tested, and may therefore not work as intended

    // Z21: 2.16 (LAN_SET_BROADCASTFLAGS)
    // 2.0: new. The flags determine which broadcasts the Z21 sends to this client, for example:
    // 0x00000001: driving and switching (track power, programming mode, short circuit, stop, LAN_X_TURNOUT_INFO,
    //             LAN_X_LOCO_INFO for the locomotives requested by this client)
    // 0x00000002: R-Bus feedback (LAN_RMBUS_DATACHANGED)    - not decoded by the library
    // 0x08000000: LocoNet detector (LAN_LOCONET_DETECTOR)   - not decoded by the library
    // For the other flags, see the Z21 specification. CmdInitUdp sets the flags to Z21BroadcastFlagsDefault.
    // This is not an XpressNet message: DataLen 0x08 0x00, Header 0x50 0x00, Flags (32 bits, little endian)
    public const int Z21BroadcastFlagsDefault = 0x00000001;

    public void CmdZ21SetBroadcastFlags(int Flags)
    {
      if (otherThread(() => CmdZ21SetBroadcastFlags(Flags))) return;
      byte[] z21Message = new byte[8];
      z21Message[0] = 0x08;
      z21Message[1] = 0x00;
      z21Message[2] = 0x50;
      z21Message[3] = 0x00;
      z21Message[4] = (byte)(Flags & 0xFF);
      z21Message[5] = (byte)((Flags >> 8) & 0xFF);
      z21Message[6] = (byte)((Flags >> 16) & 0xFF);
      z21Message[7] = (byte)((Flags >> 24) & 0xFF);
      send.transmitZ21(z21Message);
      Logfile.DetailsSent("Z21: CmdZ21SetBroadcastFlags. Flags: 0x" + Flags.ToString("X8"));
    }

    // Z21: 2.15 (LAN_X_GET_FIRMWARE_VERSION)
    public void CmdZ21GetFirmwareVersion()
    {
      send.message(0xF1, 0x0A);
      Logfile.DetailsSent("Z21: CmdZ21GetFirmwareVersion");
    }

    // Z21: 4.3.1 (LAN_X_SET_LOCO_FUNCTION)
    public void CmdZ21SetLocoFunction(int address, byte CV, byte TT)
    {
      // TT Umschalttyp: 00=aus, 01=ein, 10=umschalten,11=nicht erlaubt
      // CV: 0x00=F0 (Licht), 0x01=F1 usw. bis F31
      if (VelueOutOfRange(CV, 0, 31, "LAN_X_SET_LOCO_FUNCTION: CV")) return;
      if (VelueOutOfRange(TT, 0, 2, "LAN_X_SET_LOCO_FUNCTION: TT")) return;
      int TTNNNNNN = TT * 64 + CV;
      send.message(0xE4, 0xF8, AH(address), AL(address), TTNNNNNN);
      Logfile.DetailsSent("Z21: CmdZ21SetLocoFunction. CV: " + CV + ", TT: " + TT);
    }

    // Z21: 5.4 LAN_X_SET_EXT_ACCESSORY
    public void CmdZ21SetExtAccessory(int address, byte value)
    {
      if (VelueOutOfRange(address, 1, 2048, "Extended Accessory: Address")) return;
      // The first "user address" is 1, which is represented as 4 in the Xpressnet message
      int rawAddress = address + 3;
      if (rawAddress >= 2048) rawAddress = rawAddress - 2048;
      send.message(0x54, (rawAddress / 256), (rawAddress % 256), value, 0x00);
      Logfile.DetailsSent("Z21: CmdZ21SetExtAccessory. Address: " + address + ", Value: " + value);
    }

    // Z21: 5.5 LAN_X_GET_EXT_ACCESSORY_INFO
    public void CmdZ21GetExtAccessory(int address)
    {
      if (VelueOutOfRange(address, 1, 2048, "Extended Accessory: Address")) return;
      // The first "user address" is 1, which is represented as 4 in the Xpressnet message
      int rawAddress = address + 3;
      if (rawAddress >= 2048) rawAddress = rawAddress - 2048;
      // As opposed to normal accesories, no address correction takes place
      // As in the Z21 spec (5.5, DataLen 0x09): X-Header 0x44, Adr_MSB, Adr_LSB, 0x00
      send.message(0x44, (rawAddress / 256), (rawAddress % 256), 0x00);
      Logfile.DetailsSent("Z21: CmdZ21SetExtAccessory. Address: " + address);
    }

    // Z21: 6.9 (LAN_X_CV_POM_ACCESSORY_WRITE_BYTE)
    public void CmdZ21AccPomWrite(int decoderAddress, byte CDDD, int CV, byte value)
    {
      if (VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM write: Decoder Address")) return;
      if (VelueOutOfRange(CDDD, 0, 15, "Accessory PoM write: CDDD")) return;
      if (VelueOutOfRange(CV, 1, 1024, "Accessory PoM write: CV")) return;
      CV = CV - 1;
      // Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
      // Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
      // Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
      int aaaaa = decoderAddress / 16;
      int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
      // CV bytes
      int cvlow = CV % 256;                               // The low byte of the CV number
      int cvhigh = CV / 256 + 0xEC;                       // 0xEC = Acc PoM Schreiben
      send.message(0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, value);
      Logfile.DetailsSent("Accessory PoM Write. Decoder address: " + decoderAddress + ", CDDD: " + CDDD + ", CV: " + (CV + 1) + ", Value: " + value);
    }

    // Z21: 6.10 (LAN_X_CV_POM_ ACCESSORY_WRITE_BIT )
    public void CmdZ21AccPomBitWrite(int decoderAddress, byte CDDD, int CV, bool bitValue, byte bitPosition)
    {
      if (VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM Bit write: Decoder Address")) return;
      if (VelueOutOfRange(CDDD, 0, 15, "Accessory PoM Bit write: CDDD")) return;
      if (VelueOutOfRange(CV, 1, 1024, "PoM Bit write: CV")) return;
      CV = CV - 1;
      // Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
      // Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
      // Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
      int aaaaa = decoderAddress / 16;
      int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
      // CV bytes
      int cvlow = CV % 256;                               // The low byte of the CV number
      int cvhigh = CV / 256 + 0xE8;                       // 0xE8 = Acc PoM Bit Schreiben
      // DB5: 0000VPPP
      int VPPP = bitPosition;
      if (bitValue) VPPP = VPPP + 8;                      // bit 4 represents the new bit value
      send.message(0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, VPPP);
      Logfile.DetailsSent("Accessory PoM Bit Write. Decoder address: " + decoderAddress + ", CDDD: " + CDDD + ", CV: " + (CV + 1) + ", bitValue: " + bitValue + ", bitPosition: " + bitPosition);
    }

    // Z21: 6.11 (LAN_X_CV_POM_ACCESSORY_READ_BYTE)
    public void CmdZ21AccPomRead(int decoderAddress, byte CDDD, int CV)
    {
      if (VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM read: Decoder Address")) return;
      if (VelueOutOfRange(CDDD, 0, 15, "Accessory PoM read: CDDD")) return;
      if (VelueOutOfRange(CV, 1, 1024, "Accessory PoM read: CV")) return;
      CV = CV - 1;
      // Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
      // Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
      // Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
      int aaaaa = decoderAddress / 16;
      int AAAACDDD = (decoderAddress % 16) * 16 + CDDD;
      // CV bytes
      int cvlow = CV % 256;                               // The low byte of the CV number
      int cvhigh = CV / 256 + 0xE4;                       // 0xE4 = Acc PoM Lesen
      send.message(0xE6, 0x31, aaaaa, AAAACDDD, cvhigh, cvlow, 0x00);
      Logfile.DetailsSent("Accessory PoM Read. Decoder address: " + decoderAddress + ", CDDD: " + CDDD + ", CV: " + (CV + 1));
    }

    // Z21: 6.12 (LAN_X_MM_WRITE_BYTE )
    public void CmdZ21MmWriteByte(byte RegAdr, byte value)
    {
      send.message(0x24, 0xFF, 0x00, RegAdr, value);
      Logfile.DetailsSent("Z21: CmdZ21MmWriteByte");
    }


    // ==============================================================================================================
    // OpenDCC ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
    // The commands below will be forwarded by the Interface to the Command Station
    // Note: these commands could not be tested, and may therefore not work as intended
    public void CmdOpenDCCSetExtAccessory(int address, byte value)
    {
      if (VelueOutOfRange(address, 1, 2048, "Ext. Accessory: Address")) return;
      if (VelueOutOfRange(value, 0, 31, "Ext. Accessory: Value")) return;
      int byte1 = value * 8;                              // Shift value 3 bits left
      byte1 = byte1 + address / 256;                      // 3 highest address bits
      int byte2 = address % 256;                          // 8 lowest address bits
      send.message(0x13, 0x01, byte1, byte2);
      Logfile.DetailsSent("OPENDCC: Extended Accessory Command. Address = " + address + ", Value = " + value);
    }


    // **************************************************************************************************************
    // ***************************************************** EVENTS *************************************************
    // **************************************************************************************************************
    // RAISING EVENTS: TCP, USB AND XPRESSNET FOR ERRORS, STATUS CHANGES AND FEEDBACK INFO
    // The subroutines below forward events from lower level objects to the DLL user
    // In addition Logfile messages are written, if needed.
    internal void eventInterfaceConnected()
    {
      raise(() => { if (InterfaceConnected != null) InterfaceConnected(); });
      Logfile.Details("Interface connection established");
      // In case of a serial connection, it seems we must wait 100ms before the interface is stable
      if (connectionType == connectionTypeEnum.serialPort) System.Threading.Thread.Sleep(100);
      // Retrieve version and status info regarding the Command Station
      CmdSoftwareversionZentraleAnfordern();
      CmdStatusZentraleAnfordern();
    }

    // 2.0: logText: the complete description for the logfile (default: Description)
    internal void eventTcpError(string Description, string logText = null)
    {
      string text = Description;
      raise(() => { if (TcpError != null) TcpError(ref text); });
      Logfile.Errors("TCP Error: " + (logText ?? Description));
    }

    internal void eventTcpClosedByRemoteHost()
    {
      // The connection is cleaned up before the event is raised (VB6: after). The event handler of the application
      // may then open a new connection (for example CmdConnectTcp), which should not be closed afterwards
      if (tcp != null) tcp.closeConnection();
      tcp = null;
      connectionType = connectionTypeEnum.none;
      raise(() => { if (TcpClosedByRemoteHost != null) TcpClosedByRemoteHost(); });
      Logfile.Errors("TCP Connection closed by remote host");
    }


    // --------------------------------------------------------------------------------------------------------------
    // Interface specific events
    // V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
    internal void eventInterfaceVersion(byte daten1, byte daten2)
    {
      string Version = (daten1 / 16) + "." + (daten1 % 16);
      string code = (daten2 / 16) + "." + (daten2 % 16);
      raise(() => { if (InterfaceVersion != null) InterfaceVersion(Version, code); });
      Logfile.DetailsReceived("Interface - Version: " + Version + "  Code: " + code);
    }

    // V3.6: 1.7 (XpressNet Address) / V3.6: 2.1 (Interface Status Befehl)
    internal void eventXpressNetAddress(byte address)
    {
      // This event is triggered as response to two commands:
      // XpressNet Address (CmdXpressNetAddress - V3.6: 1.7) and
      // Interface Status Befehl (CmdInterfaceStatus - V3.6: 2.1)
      // CmdInterfaceStatus will be called internally by the DLL, whenever the
      // keep alive timer expires (send.lanKeepAliveTimer_Timer). In such case,
      // the response to the keep alive message should not be returned (as event)
      // to the application program, but instead be kept within the DLL.
      // To distinguish between both cases, the send.lanKeepAliveTimer_Timer sets the
      // waitForlanKeepAliveResponse flag
      if (send.waitForlanKeepAliveResponse) {
        send.waitForlanKeepAliveResponse = false;
        Logfile.DetailsReceived("TCP Keep Alive Response");
      }
      else {
        raise(() => { if (XpressNetAddress != null) XpressNetAddress(address); });
        Logfile.DetailsReceived("XpressNet - Address: " + address);
      }
    }

    // V3.6: 2.2 (XpressNet Version)
    internal void eventXpressNetVersion(byte daten)
    {
      string Version = (daten / 16) + "." + (daten % 16);
      raise(() => { if (XpressNetVersion != null) XpressNetVersion(Version); });
      Logfile.DetailsReceived("XpressNet - Version: " + Version);
    }

    // V3.6: 2.3 (Freie Verbindungen)
    internal void eventInterfaceFreieVerbindungen(byte number)
    {
      raise(() => { if (InterfaceFreieVerbindungen != null) InterfaceFreieVerbindungen(number); });
      Logfile.DetailsReceived("Interface - Free TCP Sockets: " + number);
    }

    // --------------------------------------------------------------------------------------------------------------
    // Xpressnet specific events
    // V3.6: 3.1.1.1, V4: 2.4.1, Z21: 2.8 (LAN_X_BC_TRACK_POWER_ON)
    internal void eventAllesAn()
    {
      ZentraleNotAus = false;
      ZentraleNotHalt = false;
      ZentraleProgrammierMode = false;
      raise(() => { if (AllesAn != null) AllesAn(); });
      Logfile.DetailsReceived("Broadcast: Alles An");
    }

    // V3.6: 3.1.1.2, V4: 2.4.2, Z21: 2.7 (LAN_X_BC_TRACK_POWER_OFF)
    internal void eventNotAus()
    {
      ZentraleNotAus = true;
      raise(() => { if (NotAus != null) NotAus(); });
      Logfile.DetailsReceived("Broadcast: Alles Aus (Notaus)");
    }

    // V3.6: 3.1.1.3, V4: 2.4.3, Z21: 2.14 (LAN_X_BC_STOPPED)
    internal void eventNotHalt()
    {
      ZentraleNotHalt = true;
      raise(() => { if (NotHalt != null) NotHalt(); });
      Logfile.DetailsReceived("Broadcast: Alle Loks Aus - Nothalt");
    }

    // V3.6: 3.1.1.4, V4: 2.4.4, Z21: 2.9 (LAN_X_BC_PROGRAMMING_MODE)
    internal void eventProgrammierMode()
    {
      ZentraleProgrammierMode = true;
      raise(() => { if (ProgrammierMode != null) ProgrammierMode(); });
      Logfile.DetailsReceived("Broadcast: Programmiermode");
    }

    // To simplify the application program, we generate different feedback events for switches and for
    // feedback (occupancy) decoders. The difference between both can be determined by the TT bits
    // V3.6: 3.1.1.5, V4: 2.4.5, V4: 2.4.6 (BC RuckMeldung) / V3.6: 3.1.8, V4: 2.15, V4: 2.16 (Schaltinformation)
    internal void eventFeedback(int address, byte Nibble, byte ttBits, byte iBit, byte data)
    {
      if (ttBits == 2) {                                  // normal feedback decoder
        // We add 1 to the RS-Bus address, since addresses on the wire range from 0..127. We display these as 1..128
        address = address + 1;
        int rsBusAddress = address;
        raise(() => { if (Feedback != null) Feedback(rsBusAddress, Nibble, data); });
        Logfile.DetailsReceived("Feedback Message received for RS-Bus address: " + address + "  Data: " + data + "  Nibble: " + Nibble);
      }
      else {                                              // switch decoder
        // Adress is the decoder address, but we will return a switch address pair
        int switchAddress;
        if (!Z21Protocol) {
          switchAddress = (address * 4) + (Nibble * 2) + 1;
        }
        else {
          switchAddress = address + 1;
        }
        raise(() => { if (SwitchFeedback != null) SwitchFeedback(switchAddress, ttBits, iBit, data); });
        Logfile.DetailsReceived("Feedback Message received for switch addresses: " + switchAddress + "/" + (switchAddress + 1) + "  Data: " + data + "  TT-Bits: " + ttBits + " I-Bit: " + iBit);
      }
    }

    // V3.6: 3.1.2.1, V4: 2.5.1, Z21: 2.10 (Programmierinfo Kurzschluss / LAN_X_BC_TRACK_SHORT_CIRCUIT)
    internal void eventProgrammierinfoKurzschluss()
    {
      raise(() => { if (ProgrammierinfoKurzschluss != null) ProgrammierinfoKurzschluss(); });
      Logfile.DetailsReceived("Programmierinfo Kurzschluss");
    }

    // V3.6: 3.1.2.2, V4: 2.5.2, Z21: 6.4 (Programmierinfo Daten nicht gefunden / LAN_X_CV_NACK)
    internal void eventProgrammierinfoDatenNichtGefunden()
    {
      raise(() => { if (ProgrammierinfoDatenNichtGefunden != null) ProgrammierinfoDatenNichtGefunden(); });
      Logfile.DetailsReceived("Programmierinfo Daten nicht gefunden");
    }

    // V3.6: 3.1.2.3, V4: 2.5.3 (Programmierinfo Zentrale Busy)
    internal void eventProgrammierinfoZentraleBusy()
    {
      raise(() => { if (ProgrammierinfoZentraleBusy != null) ProgrammierinfoZentraleBusy(); });
      Logfile.DetailsReceived("Programmierinfo Zentrale Busy");
    }

    // V3.6: 3.1.2.4, V4: 2.5.4 (Programmierinfo Zentrale Bereit)
    internal void eventProgrammierinfoZentraleBereit()
    {
      raise(() => { if (ProgrammierinfoZentraleBereit != null) ProgrammierinfoZentraleBereit(); });
      Logfile.DetailsReceived("Programmierinfo Zentrale Bereit");
    }

    // V3.6 3.1.2.5 / V4: 2.5.5 (Programmierinfo Daten 3-Byte-Format)
    internal void eventProgInfoPageRegister(byte EEPromAddress, byte value)
    {
      raise(() => { if (ProgrammierinfoPageRegister != null) ProgrammierinfoPageRegister(EEPromAddress, value); });
      Logfile.DetailsReceived("Programmierinfo Page / Register. EEPromAddress: " + EEPromAddress + ", Value: " + value);
    }

    // V3.6: 3.1.2.6 - 3.1.2.9, V4: 2.5.6 - 2.5.9 (Programmierinfo CV1..1024)
    internal void eventProgInfo(byte CvRange, byte daten1, byte value)
    {
      int CV = daten1;
      switch (CvRange) {
        case 1:
          if (CV == 0) {
            // The version is compared as a number (see Val)
            if (Val(ZentraleSoftwareVersion) < 3.6) CV = 256;
            else CV = 1024;
          }
          break;
        case 2: CV = CV + 256; break;
        case 3: CV = CV + 512; break;
        case 4: CV = CV + 768; break;
      }
      raise(() => { if (ProgrammierinfoCV != null) ProgrammierinfoCV(CV, value); });
      Logfile.DetailsReceived("Programmierinfo. CV: " + CV + ", Value: " + value);
    }

    // Z21: 6.5 (LAN_X_CV_RESULT)
    internal void eventProgInfoZ21(byte CVAdr_MSB, byte CVAdr_LSB, byte value)
    {
      int CV = (CVAdr_MSB * 256) + CVAdr_LSB + 1;
      raise(() => { if (ProgrammierinfoCV != null) ProgrammierinfoCV(CV, value); });
      Logfile.DetailsReceived("Programmierinfo Z21. CV: " + CV + ", Value: " + value);
    }

    // V4: 2.6 (Service Variable melden)
    internal void eventServiceVariableMelden(byte SvAddress, byte SvValue)
    {
      raise(() => { if (ServiceVariable != null) ServiceVariable(SvAddress, SvValue); });
      Logfile.DetailsReceived("Service Variable melden. SV Address: " + SvAddress + ", SV Value: " + SvValue);
    }

    // V3.6: 3.1.3, V4: 2.7.2, Z21: 2.3 (Softwareversion Zentrale / LAN_X_GET_VERSION)
    internal void eventSoftwareVersionZentrale(byte daten2, byte daten3)
    {
      string Version = (daten2 / 16) + "." + (daten2 % 16);
      ZentraleSoftwareVersion = Version;
      ZentraleKennung = daten3;
      raise(() => { if (ZentraleVersion != null) ZentraleVersion(Version); });
      // 2.0: with the name of the command station
      Logfile.DetailsReceived("Version Zentrale: " + Version + ", Kennung: " + ZentraleKennung + " (" + zentraleName(ZentraleKennung) + ")");
    }

    // 2.0: The name of the command station, for the logfile (Zentralen-Kennung, V3.6: 3.1.3)
    private static string zentraleName(byte kennung)
    {
      switch (kennung) {
        case LZ100: return "LZ 100";
        case LH200: return "LH 200";
        case DPC: return "DPC (Compact / Commander)";
        case CntrPlus: return "Control Plus";
        case LokMausII: return "LokMaus II";
        case LDH: return "LDH";
        case multiMAUS: return "multiMAUS";
        case Z21: return "Z21 / Yamorc";
        case OpenDcc: return "OpenDCC";
        case OpenDccXp: return "OpenDccXp";
        default: return "unknown";
      }
    }

    // V3.6: 3.1.4, V4: 2.8, Z21: 2.12 (Status Zentrale / LAN_X_STATUS_CHANGED)
    internal void eventStatusZentrale(byte data)
    {
      // With Z21, we internally use the StatusZentrale command as "Keep Alive" Message
      // If we use the StatusZentrale command in this way, we need not inform the application
      // about the result. In that case it is sufficient to tell the remaining program that
      // we no longer wait for the Keep Alive Response. Just log it that we received
      // the response, and exit
      if ((connectionType == connectionTypeEnum.lanUdpPort) && send.waitForlanKeepAliveResponse) {
        send.waitForlanKeepAliveResponse = false;
        Logfile.DetailsReceived("UDP Keep Alive Response");
        return;
      }
      // If we are here, this is not a Z21 Keep Alive, but a response to a normal request.
      ZentraleNotHalt = bitValue(data, 0);
      ZentraleNotAus = bitValue(data, 1);
      if (!Z21Protocol) ZentraleStartmodeAuto = bitValue(data, 2);
      if (!Z21Protocol) ZentraleProgrammierMode = bitValue(data, 3);
      if (Z21Protocol) ZentraleKurzSchluss = bitValue(data, 2);
      if (Z21Protocol) ZentraleProgrammierMode = bitValue(data, 5);
      ZentraleKaltstart = bitValue(data, 6);
      ZentraleRamCheckFehler = bitValue(data, 7);
      if (ZentraleNotHalt) {
        raise(() => { if (NotHalt != null) NotHalt(); });
        Logfile.DetailsReceived("Startup: NotHalt");
      }
      if (ZentraleNotAus) {
        raise(() => { if (NotAus != null) NotAus(); });
        Logfile.DetailsReceived("Startup: NotAus");
      }
      if (ZentraleProgrammierMode) {
        raise(() => { if (ProgrammierMode != null) ProgrammierMode(); });
        Logfile.DetailsReceived("Startup: ProgrammierMode");
      }
      if (ZentraleRamCheckFehler) {
        Logfile.Errors("Startup: Zentrale - Ram Check Fehler");
      }
      Logfile.DetailsReceived("Status Zentrale: " + data);
    }

    // V4: 2.10 (PoM Ergebnis melden)
    internal void eventPomErgebnisMelden(byte AdrH, byte AdrL, byte CvValue)
    {
      int LokAddress = AdrH * 256 + AdrL;
      raise(() => { if (PomErgebnis != null) PomErgebnis(LokAddress, CvValue); });
      Logfile.DetailsReceived("PomErgebnisMelden. Lok Address:: " + LokAddress + ", CV Value:" + CvValue);
    }

    // V3.6: 3.1.9.1, V4: 2.19.1 (Lok Info Normale Lok)
    internal void eventNormaleLokinfo(byte Kennung, byte Speed, byte F0_F4, byte F5_F12)
    {
      bool Besetzt = bitValue(Kennung, 3);
      int Fahrstufen = 0;
      switch (Kennung % 8) {                              // Analyse bits 0..2
        case 0: Fahrstufen = 14; break;
        case 1: Fahrstufen = 27; break;
        case 2: Fahrstufen = 28; break;
        case 4: Fahrstufen = 128; break;
      }
      bool Forwarts = bitValue(Speed, 7);
      // The speed byte is decoded according to the speed steps
      int Geschwindigkeit = DecodeSpeedSteps(Speed, Fahrstufen);
      byte F5_F8 = (byte)(F5_F12 % 16);                   // Gruppe 2 - take the lowest 4 bits
      byte F9_F12 = (byte)(F5_F12 / 16);                  // Gruppe 3 - take the highest 4 bits
      raise(() => { if (NormaleLokinfo != null) NormaleLokinfo(Besetzt, Fahrstufen, Geschwindigkeit, Forwarts, F0_F4, F5_F8, F9_F12); });
      Logfile.DetailsReceived("Lokinformation (normal) - Besetzt: " + Besetzt + ", Speed: " + Geschwindigkeit + ", Stufen: " + Fahrstufen + ", Forwarts: " + Forwarts + ", F0..F4 " + F0_F4 + ", F5..F8 " + F5_F8 + ", F9..F12 " + F9_F12);
    }

    // Z21: 4.4 (LAN_X_LOCO_INFO)
    internal void eventZ21Lokinfo(byte Adr_MSB, byte Adr_LSB, byte Kennung, byte Speed, byte F0_F4, byte F5_F12, byte F13_F20, byte F21_F28)
    {
      int LokAddress = (Adr_MSB & 0x3F) * 256 + Adr_LSB;
      bool Besetzt = bitValue(Kennung, 3);
      int Fahrstufen = 0;
      switch (Kennung % 8) {                              // Analyse bits 0..2
        case 0: Fahrstufen = 14; break;
        case 1: Fahrstufen = 27; break;
        case 2: Fahrstufen = 28; break;
        case 4: Fahrstufen = 128; break;
      }
      bool Forwarts = bitValue(Speed, 7);
      // The speed byte is decoded according to the speed steps
      int Geschwindigkeit = DecodeSpeedSteps(Speed, Fahrstufen);
      byte F5_F8 = (byte)(F5_F12 % 16);                   // Gruppe 2 - take the lowest 4 bits
      byte F9_F12 = (byte)(F5_F12 / 16);                  // Gruppe 3 - take the highest 4 bits
      raise(() => { if (Z21Lokinfo != null) Z21Lokinfo(LokAddress, Besetzt, Fahrstufen, Geschwindigkeit, Forwarts, F0_F4, F5_F8, F9_F12, F13_F20, F21_F28); });
      Logfile.DetailsReceived("Lokinformation (Z21) - Address: " + LokAddress + ", Besetzt: " + Besetzt + ", Speed: " + Geschwindigkeit + ", Stufen: " + Fahrstufen + ", Forwarts: " + Forwarts + ", F0..F4 " + F0_F4 + ", F5..F8 " + F5_F8 + ", F9..F12 " + F9_F12 + ", F13..F20 " + F13_F20 + ", F21..F28 " + F21_F28);
    }

    // V3.6: 3.1.9.2, V4: 2.19.2 (Funktionszustand F13 bis F28)
    internal void eventFunctionsZustandF13F20(byte F13_F20, byte F21_F28)
    {
      raise(() => { if (FunctionsZustandF13F20 != null) FunctionsZustandF13F20(F13_F20, F21_F28); });
      Logfile.DetailsReceived("Funktionszustand - F13 bis F20: " + F13_F20 + ", F21 bis F28: " + F21_F28);
    }

    // V4: 2.19.3 (Funktionszustand F29 bis F68)
    internal void eventFunctionsZustandF29F68(byte F29_F36, byte F37_F44, byte F45_F52, byte F53_F60, byte F61_F68)
    {
      raise(() => { if (FunctionsZustandF29F68 != null) FunctionsZustandF29F68(F29_F36, F37_F44, F45_F52, F53_F60, F61_F68); });
      Logfile.DetailsReceived("Funktionszustand - F29 bis F68: " + F29_F36 + ", " + F37_F44 + ", " + F45_F52 + ", " + F53_F60 + ", " + F61_F68);
    }

    // V3.6: 3.1.11, V4: 2.19.8 (Funktionsstatus F0 bis F12 der angefragten Lok)
    internal void eventFunctionStatusF0F12(byte F0_F4, byte F5_F12)
    {
      byte F5_F8 = (byte)(F5_F12 % 16);                   // Gruppe 2 - take the lowest 4 bits
      byte F9_F12 = (byte)(F5_F12 / 16);                  // Gruppe 3 - take the highest 4 bits
      raise(() => { if (FunctionStatusF0F12 != null) FunctionStatusF0F12(F0_F4, F5_F8, F9_F12); });
      Logfile.DetailsReceived("Funktionsstatus - F0..F4 " + F0_F4 + ", F5..F8 " + F5_F8 + ", F9..F12 " + F9_F12);
    }

    // V3.6: 3.1.12, V4: 2.20 (Funktionsstatus F13 bis F28)
    internal void eventFunctionStatusF13F20(byte F13_F20, byte F21_F28, byte RefreshMode)
    {
      raise(() => { if (FunctionStatusF13F20 != null) FunctionStatusF13F20(F13_F20, F21_F28, RefreshMode); });
      Logfile.DetailsReceived("Funktionsstatus - F13 bis F20: " + F13_F20 + ", F21 bis F28: " + F21_F28 + ", Refresh Mode: " + RefreshMode);
    }

    // V4: 2.21 (Funktionsstatus F29 bis F68)
    internal void eventFunctionStatusF29F68(byte F29_F36, byte F37_F44, byte F45_F52, byte F53_F60, byte F61_F68)
    {
      raise(() => { if (FunctionStatusF29F68 != null) FunctionStatusF29F68(F29_F36, F37_F44, F45_F52, F53_F60, F61_F68); });
      Logfile.DetailsReceived("Funktionsstatus - F29 bis F68: " + F29_F36 + ", " + F37_F44 + ", " + F45_F52 + ", " + F53_F60 + ", " + F61_F68);
    }

    // V3.6: 3.1.14, V4: 2.24 (Fehlermeldung)
    internal void eventFehlermeldung(byte data)
    {
      if (bitValue(data, 7)) {                            // Highest order bit must have been set
        byte value = (byte)(data % 16);                   // Take the four lowest order bits
        raise(() => { if (Fehlermeldung != null) Fehlermeldung(value); });
        if (LogWarnings) {
          Logfile.AppendWithHost("<= Fehlermeldung: " + value);
        }
      }
    }

    // V4: 2.4.7 / V4: 2.11 (BC: Modellzeit / Modellzeit melden)
    internal void eventModellZeit(byte daten1, byte daten2, byte factor = 255)
    {
      byte day = (byte)(daten1 / 32);
      byte hour = (byte)(daten1 % 32);
      byte minute = (byte)(daten2 % 64);
      // run = daten2 \ 128                       ' Seems Run is always zero.... Error in LZV200??
      raise(() => { if (Modellzeit != null) Modellzeit(day, hour, minute, factor); });
      // We do not log Modellzeit messages in the logfile
    }

    // V4: 2.9 (Erweiterte Versionsinformation melden)
    internal void eventErweiterteVersionZentrale(byte ZBldH, byte ZBldL, byte RMVer, byte RMBldH, byte RMBldL, byte BIvers)
    {
      string BuildZentrale = (ZBldH * 256 + ZBldL).ToString();
      string VersionRuckMelder = (RMVer / 16) + "." + (RMVer % 16);
      string BuildRuckMelder = (RMBldH * 256 + RMBldL).ToString();
      string VersionBootloader = (BIvers / 16) + "." + (BIvers % 16);
      raise(() => { if (ErweiterteZentraleVersion != null) ErweiterteZentraleVersion(BuildZentrale, VersionRuckMelder, BuildRuckMelder, VersionBootloader); });
      Logfile.DetailsReceived("Erweiterte Versionsinformation - BuildZentrale = " + BuildZentrale + ", VersionRuckMelder = " + VersionRuckMelder + ", BuildRuckMelder = " + BuildRuckMelder + ", VersionBootloader = " + VersionBootloader);
    }

    // Z21: 2.15 (Antwort auf LAN_X_GET_FIRMWARE_VERSION)
    internal void eventz21FirmwareVersion(byte V_MSB, byte V_LSB)
    {
      // The version is BCD coded (Z21 spec 2.15: 0x01 0x23 = V1.23)
      string Version = V_MSB.ToString("X") + "." + V_LSB.ToString("X2");
      raise(() => { if (z21FirmwareVersion != null) z21FirmwareVersion(Version); });
      Logfile.DetailsReceived("Z21 Firmware - Version = " + Version);
    }

    // Z21: 5.6 (ExtendedAccessoryInfo(ByVal address As Integer, ByVal Value As Byte, ByVal Valid As Boolean))
    internal void eventExtendedAccessoryInfo(byte Adr_MSB, byte Adr_LSB, byte DDDDDDDD, byte Status)
    {
      int rawAddress = (Adr_MSB * 256) + Adr_LSB;
      if (rawAddress <= 3) rawAddress = rawAddress + 2048;
      int address = rawAddress - 3;                       // The first (XpressNet) address is 4, and presented to the user as 1
      raise(() => { if (ExtendedAccessoryInfo != null) ExtendedAccessoryInfo(address, DDDDDDDD, Status); });
      Logfile.DetailsReceived("Extended Accessory Info - Address: " + address + ", Value: " + DDDDDDDD + ", Status: " + Status);
    }

    // --------------------------------------------------------------------------------------------------------------
    // 2.0: Sending events (see XpressNetSend.cs). The logfile is written by the caller
    internal void eventSendError(SendErrorReasonEnum reason, string description)
    {
      byte Reason = (byte)reason;
      raise(() => { if (SendError != null) SendError(Reason, description); });
    }

    internal void eventInterfaceNichtAdressiert()
    {
      raise(() => { if (InterfaceNichtAdressiert != null) InterfaceNichtAdressiert(); });
    }

    internal void eventInterfaceWiederAdressiert()
    {
      raise(() => { if (InterfaceWiederAdressiert != null) InterfaceWiederAdressiert(); });
    }


    // **************************************************************************************************************
    // ************************************************* INTERNAL USE ***********************************************
    // **************************************************************************************************************
    // Raise an event. An error (exception) in the event handler of the application should not disturb the
    // library (for example: the next message in the FIFO must still be sent). Therefore the error is written
    // to the logfile, and passed back to the application (via the main thread) after the library has finished.
    // The application then gets the error in the usual way (Visual Studio debugger, or the error dialog).
    private void raise(Action raiseEvent)
    {
      try {
        raiseEvent();
      }
      catch (Exception e) {
        Logfile.Errors("Error in an event handler of the application: " + e.Message);
        ExceptionDispatchInfo error = ExceptionDispatchInfo.Capture(e);
        mainThread.PostAlways(() => error.Throw());
      }
    }

    // Is the library called from another thread than the main thread? If yes, the call is passed to the
    // main thread, and the calling method should return immediately
    internal bool otherThread(Action call)
    {
      if (mainThread.IsMainThread) return false;
      Logfile.Warnings("XpressNet called from another thread. The call is passed to the main thread");
      mainThread.Post(call);
      return true;
    }

    // Do we have a connection on which we can send?
    internal bool isConnected()
    {
      switch (connectionType) {
        case connectionTypeEnum.lanTcpPort: return (tcp != null) && tcp.connected;
        case connectionTypeEnum.lanUdpPort: return (udp != null) && udp.ready;
        case connectionTypeEnum.usbPort: return (usb != null) && usb.isOpen;
        case connectionTypeEnum.serialPort: return (usb != null) && usb.isOpen;
        default: return false;
      }
    }

    // Close all connections
    private void closeConnections()
    {
      if (tcp != null) tcp.closeConnection();
      // 2.0: log off from the Z21 (Z21 spec 2.2), via the same socket (port) as used for the login
      if ((udp != null) && (connectionType == connectionTypeEnum.lanUdpPort)) {
        send.transmitZ21(new byte[] { 0x04, 0x00, 0x30, 0x00 });
        Logfile.DetailsSent("Z21: LAN_LOGOFF");
      }
      if (udp != null) udp.close();
      if (usb != null) usb.closePort();
      tcp = null;
      udp = null;
      usb = null;
      connectionType = connectionTypeEnum.none;
    }

    // Some bit manipulation functions
    private static bool bitValue(int data, int index)
    {
      return ((data >> index) & 1) == 1;
    }

    // support function to check is Value is between some predefined bounds
    // 2.0: a value out of range gives the event SendError (the message is not sent)
    private bool VelueOutOfRange(int value, int min, int max, string LogText)
    {
      bool outOfRange = false;
      if (value < min) {
        Logfile.AppendWithHost("Error " + LogText + " must be >= " + min);
        eventSendError(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Message not sent: " + LogText + " must be >= " + min + " (value: " + value + ")");
        outOfRange = true;
      }
      if (value > max) {
        Logfile.AppendWithHost("Error " + LogText + " must be <= " + max);
        eventSendError(SendErrorReasonEnum.FehlerInDenBefehlsparametern, "Message not sent: " + LogText + " must be <= " + max + " (value: " + value + ")");
        outOfRange = true;
      }
      return outOfRange;
    }

    // Functions for Accessory addresses
    private int AccAL(int address)
    {
      // Calculation of the low byte of a accessory address
      if (Z21Protocol) {
        return address % 256;                             // Select the last 8-bits
      }
      else {
        if (address >= 1024) {                            // We already assured that the command station >= V3.8
          return (address - 1024) / 4;                    // Integer division. Result fits in an 8-bit byte
        }
        else {
          return address / 4;                             // Integer division
        }
      }
    }

    private int AccAH(int address)
    {
      // Calculation of the high byte of a accessory address
      if (Z21Protocol) {
        return address / 256;                             // Remove the last 8-bits (shift right)
      }
      else {                                              // Lenz uses decoder addressing
        return address / 1024;                            // Remove the last 10-bits (shift right)
      }
    }

    // Functions for lok addresses and speed
    private static int AL(int address)
    {
      // Calculation of the low byte of a loco address
      return address % 256;
    }

    private static int AH(int address)
    {
      // Calculation of the high byte of a loco address
      if (address < 100) return 0;
      else return address / 256 + 0xC0;
    }

    private static int EncodeSpeed(int Speed)
    {
      // Calculates the speed byte for decoders with 27/28 speed steps
      int result = 0;
      if (Speed > 0) {
        Speed = Speed + 3;                                // RVVV expects an offset
        result = Speed / 2;                               // neglect the last bit, thus all uneven speed steps
        bool even = (Speed & 1) == 1;                     // mask / take the last speed bit
        if (even) result = result + 16;                   // If the speed was even, set RVVV bit 4 to 1
      }
      return result;
    }

    public int DecodeSpeed(int Speed)
    {
      // Derive from the speed byte an integer value (for decoders with 27/28 speed steps)
      int result = 0;                                     // Default speed
      int Bit4 = bitValue(Speed, 4) ? 1 : 0;              // take bit 4
      int Bit0_3 = Speed % 16;                            // take the lowest 4 bits
      if (Bit0_3 >= 2) {
        result = (Bit0_3 - 2) * 2 + 1 + Bit4;
      }
      return result;
    }

    // Derive from the speed byte an integer value, according to the speed steps
    // 14 and 128 speed steps: 0 = stop, 1 = emergency stop (returned as 0), 2.. = speed step 1..
    // 27 and 28 speed steps: DecodeSpeed
    private int DecodeSpeedSteps(int Speed, int Speedsteps)
    {
      int value;
      switch (Speedsteps) {
        case 14:
          value = Speed % 16;                             // take the lowest 4 bits
          if (value >= 2) return value - 1; else return 0;
        case 128:
          value = Speed % 128;                            // take the lowest 7 bits
          if (value >= 2) return value - 1; else return 0;
        default:
          return DecodeSpeed(Speed);
      }
    }

    // The VB6 Val function: the number at the start of the text, with a point as decimal separator
    // (independent of the Windows regional settings). Returns 0 if the text does not start with a number.
    // The version of the command station is compared as a number: a string comparison does not work for
    // versions >= 10.0
    internal static double Val(string text)
    {
      if (text == null) return 0;
      string number = "";
      bool point = false;
      foreach (char c in text.Trim()) {
        if (char.IsDigit(c)) number += c;
        else if ((c == '.') && !point) { number += c; point = true; }
        else break;
      }
      double result;
      if (double.TryParse(number, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out result)) return result;
      return 0;
    }
  }
}
