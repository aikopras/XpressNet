VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XpressNetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The line below is needed to use the Sleep function from the kernel.
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Convention: everything intended for the DLL user starts with a capitol letter.
' Internal objects, methods and attributes (variables) start with lowercase

' GENERATED EVENTS - Interface: TCP as well as USB
Event InterfaceConnected()
Event TcpClosedByRemoteHost()
Event TcpError(ByRef Description As String)

' GENERATED EVENTS - USB
Event UsbConnected()

' GENERATED EVENTS - XPRESSNET
Event AllesAn()
Event NotAus()
Event NotHalt()
Event ProgrammierMode()
Event XpressNetAddress(ByVal address As Byte)
Event XpressNetVersion(ByVal Version As String)
Event InterfaceVersion(ByVal Version As String, ByVal code As String)
Event InterfaceFreieVerbindungen(ByVal number As Byte)
Event ZentraleVersion(ByVal Version As String)
Event ErweiterteZentraleVersion(ByVal BuildZentrale As String, ByVal VersionRuckMelder As String, ByVal BuildRuckMelder As String, ByVal VersionBootloader As String)
Event Fehlermeldung(ByVal value As Byte)
Event ServiceVariable(ByVal SvAddress As Byte, ByVal SvValue As Byte)
' Feedback events
Event Feedback(ByVal address As Integer, ByVal Nibble As Byte, ByVal data As Byte)
Event SwitchFeedback(ByVal address As Integer, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte)
Event ExtendedAccessoryInfo(ByVal address As Integer, ByVal value As Byte, ByVal Status As Byte)
' Lok events
Event NormaleLokinfo(ByVal Besetzt As Boolean, ByVal Speedsteps As Integer, ByVal Speed As Integer, ByVal Forward As Boolean, ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte)
Event FunctionsZustandF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)
Event FunctionsZustandF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte)
Event FunctionStatusF0F12(ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte)
Event FunctionStatusF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte, ByVal RefreshMode As Byte)
Event FunctionStatusF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte)
' Programmier events - Service Mode
Event ProgrammierinfoKurzschluss()
Event ProgrammierinfoDatenNichtGefunden()
Event ProgrammierinfoZentraleBusy()
Event ProgrammierinfoZentraleBereit()
Event ProgrammierinfoCV(ByVal CV As Integer, ByVal value As Byte)
Event ProgrammierinfoPageRegister(ByVal EEPromAddress As Byte, ByVal value As Byte)
' Programmier events - Programming on Main (PoM)
Event PomErgebnis(ByVal LokAddress As Integer, ByVal CvValue As Byte)
' Modellzeit events
Event Modellzeit(ByVal day As Byte, ByVal hour As Byte, ByVal minute As Byte, ByVal factor As Byte)
' Z21 specific events
Event z21FirmwareVersion(ByVal Version As String)
Event Z21Lokinfo(ByVal address As Integer, ByVal Besetzt As Boolean, ByVal Speedsteps As Integer, ByVal Speed As Integer, ByVal Forward As Boolean, ByVal F0_F4 As Byte, ByVal F5_F8 As Byte, ByVal F9_F12 As Byte, ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)

' ATTRIBUTES THAT THE DLL USER MAY CHANGE. THE DEFAULT VALUES ARE SET IN Class_Initialize()
Public Z21Protocol As Boolean          ' Is the library used for Z21 or XpressNet??
Public LogMessages As Boolean          ' Example: connection requested to host X
Public LogWarnings As Boolean          ' May indicate a software problem
Public LogErrors As Boolean            ' Indicates a hardware or software problem
Public LogHexMessages As Boolean       ' Hex dump of transmitted / received messages
Public LogDetails As Boolean           ' Description of what has been transmitted / received
Public AccessoryInterval As Long       ' Interval between two accesory commands
    
' ATTRIBUTES THAT THE DLL USER MAY READ
Public ZentraleNotAus As Boolean
Public ZentraleNotHalt As Boolean
Public ZentraleProgrammierMode As Boolean
Public ZentraleVersion As String
Public ZentraleKennung As Byte
Public ZentraleKurzSchluss As Boolean  ' Only for Z21
Public ZentraleStartmodeAuto As Boolean
Public ZentraleKaltstart As Boolean
Public ZentraleRamCheckFehler As Boolean
Public NumberOfExpressNetMessagesReceived As Long
Public NumberOfExpressNetMessagesTransmitted As Long

' OBJECTS THAT THE DLL USER MAY USE
Public Logfile As LogFileClass

' Objects for DLL internal use. Although these are Public, they should not be used outside the DLL
Public send As SendClass               ' Everything related to sending XpressNet messages
Public receive As ReceiveClass         ' Everything related to receiving XpressNet messages
Public receiveZ21 As ReceiveClassZ21   ' To receive messages transported over a Z21 XpressNet tunnel
Public tcp As TcpClass                 ' A wrapper around the Form with the WinSock control
Public udp As UdpClass                 ' A wrapper around the Form with the WinSock control
Public usb As UsbClass                 ' A wrapper around the Form with the MSComm control
Public AccAuto As AccAutoClass         ' Implements the AccessoryAuto commands

' Enumeration to determine the type of connection
Public Enum connectionTypeEnum
  none
  lanTcpPort
  lanUdpPort
  usbPort
  serialPort
End Enum
Public connectionType As connectionTypeEnum

' Known values for the Zentralen-Kennung
Private Const LZ100 = &H0               ' LZ 100
Private Const LH200 = &H1               ' LH 200
Private Const DPC = &H2                 ' DPC – Zentrale (Compact und Commander)
Private Const CntrPlus = &H3            ' Control Plus – Zentrale
Private Const Z21 = &H12                ' Z21 / Yamorc
Private Const OpenDcc = &H20            ' See: https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html
Private Const OpenDccXp = &H22          ' See: https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html



'==================================================================================================================
' METHODS FOR ESTABLSHING CONNECTIONS
Public Sub CmdConnectTcp(ByVal Host As String, Optional ByVal Port As String = "5550")
   Logfile.myHost = Host           ' Save the host (IP address) to include in the Logfile
   Logfile.Details "TCP connection requested to Host: " & Host & " at Port: " & Port
   connectionType = lanTcpPort        ' We connect via TCP over a LAN
   Set tcp = New TcpClass          ' Instantiated a new TCP class object
   Set tcp.Parent = Me             ' Give the new object a pointer ro its parent
   Call tcp.connect(Host, Port)    ' Send a connection request
End Sub

Public Sub CmdInitUdp(ByVal Host As String, Optional ByVal Port As String = "21106")
   Logfile.myHost = Host           ' Save the host (IP address) to include in the Logfile
   Logfile.Details "UDP initialised: Host: " & Host & ", Port: " & Port
   connectionType = lanUdpPort     ' We connect via TCP over a LAN
   Set udp = New UdpClass          ' Instantiated a new TCP class object
   Set udp.Parent = Me             ' Give the new object a pointer ro its parent
   Z21Protocol = True
   Call udp.init(Host, Port)       ' Store the address and port
   Call CmdSoftwareversionZentraleAnfordern
   Call CmdStatusZentraleAnfordern
   Call CmdZ21GetFirmwareVersion
End Sub

Public Sub CmdConnectUsb(ByVal Port As Byte, Optional ByRef Settings As String = "115200,N,8,1")
   ' Lenz uses 57600, Yamorc 115200
   Logfile.myHost = "COM" & Port   ' Save the Com port to include in the Logfile
   Logfile.Details "USB connect attempt using Port: " & Port
   connectionType = usbPort        ' We connect via USB
   Set usb = New UsbClass
   Set usb.Parent = Me
   Call usb.connectUsb(Port, Settings)    ' Send a connection request
End Sub

Public Sub CmdConnectSerial(ByVal Port As Byte, ByRef Settings As String, ByVal Handshake As Boolean)
   Logfile.myHost = "COM" & Port   ' Save the Com port to include in the Logfile
   Logfile.Details "Serial connect attempt using Port=" & Port & ", Settings=" & Settings & " and Handshake=" & Handshake
   connectionType = serialPort     ' We connect via a serial (RS232) connection
   Set usb = New UsbClass          ' Note that usb and serial use the same class and object
   Set usb.Parent = Me
   Call usb.connectSerial(Port, Settings, Handshake)    ' Send a connection request
End Sub


'==================================================================================================================
' METHODS TO GENERATE INTERFACE SPECIFIC MESSAGES
' The various subroutines below contain references to the V3.6 and V4 XpressNet specifications:
' - Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151
'   Kommunikation auf den Schnittstellen, Doku Version 1.2 05/2019
'   https://www.lenz-elektronik.de/src/pdf/23151_Info.zip
' - Dokumentation XpressNet Version 4.0, 02/2022:
'   https://www.lenz-elektronik.de/src/pdf/Lenz_XpressNet_Doku.pdf
' For details, see these documents

' V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
Public Sub CmdInterfaceVersion()
  Call send.message(&HF0)
  Logfile.Details "Versionsnummer des Interface 23151 feststellen"
End Sub

' V3.6: 1.7 (XpressNet Gerateadresse des Interface 23151 feststellen und andern)
Public Sub CmdXpressNetAddress(Optional ByVal address As Byte = 0)
  Call send.message(&HF2, &H1, address)
  Logfile.Details "XpressNet Gerateadresse des Interface 23151 feststellen"
End Sub

' V3.6: 2.1 (Interface Status Befehl - Keep Alive) - Need not be used by the Application
Public Sub CmdInterfaceStatus()
  Call send.message(&HF1, &H1)
End Sub

' V3.6: 2.2 (XpressNet Version Befehl)
Public Sub CmdXpressNetVersion()
  Call send.message(&HF1, &H2)
  Logfile.Details "XpressNet Version Befehl"
End Sub

' V3.6: 2.3 (Verfugbare Freie Verbindungen)
Public Sub CmdInterfaceFreieVerbindungen()
  Call send.message(&HF1, &H3)
  Logfile.Details "Verfugbare Freie Verbindungen"
End Sub


'==================================================================================================================
' V3.6 & V4 METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
' The commands below will be forwarded by the Interface to the Command Station

' V3.6: 3.2.1 / V4: 3.2 (Alles An)
Public Sub CmdAllesAn()
  Call send.message(&H21, &H81)
  Logfile.Details "Alles An"
End Sub

' V3.6: 3.2.2 / V4: 3.3 (Alles Aus - Notaus)
Public Sub CmdAllesAus()
  Call send.message(&H21, &H80)
  Logfile.Details "Alles Aus - Notaus"
End Sub

' V3.6: 3.2.3 / V4: 3.4 (Alle Loks anhalten - Nothalt)
Public Sub CmdAlleLoksAnhalten()
  Call send.message(&H80)
  Logfile.Details "Alle Loks anhalten - Nothalt"
End Sub

' V3.6: 3.2.4 / V4: 3.6 & 3.7 (Eine Lok anhalten - Nothalt für eine Lok)
Public Sub CmdEineLokAnhalten(ByVal address As Integer)
  Call send.message(&H92, AH(address), AL(address))
  Logfile.Details "Eine lok anhalten. Adresse: " & address
End Sub

' TODO TESTEN
' V3.6: 3.2.5 / V4: 3.9
Public Sub CmdLeseanfrageProgrammierenRegistermode(ByVal Register As Byte)
  If VelueOutOfRange(Register, 1, 8, "Leseanfrage Programmieren Registermode: Register") Then Exit Sub
  Call send.message(&H22, &H11, Register)
  Logfile.Details "Leseanfrage Programmieren Registermode. Register: " & Register
End Sub

' TODO TESTEN
' V3.6: 3.2.6-3.2.10 / V4: 3.10-3.14 (Leseanfrage Programmieren 4-Byte-Format / LAN_X_CV_READ )
Public Sub CmdLeseanfrageProgrammieren(ByVal CV As Integer)
  If VelueOutOfRange(CV, 1, 1024, "Leseanfrage Programmieren: CV") Then Exit Sub
  If (Z21Protocol) Then
    CV = CV - 1
    Call send.message(&H23, &H11, (CV \ 256), (CV Mod 256))
    Logfile.Details "LAN_X_CV_READ, CV: " & (CV + 1)
  Else
    If (ZentraleVersion < "3.6") Then
      Call send.message(&H22, &H15, (CV Mod 256))
    Else
      Dim daten1 As Byte
      daten1 = &H18 + ((CV \ 256) Mod 4)
      Call send.message(&H22, daten1, (CV Mod 256))
    End If
    Logfile.Details "Leseanfrage Programmieren 4-Byte-Format. CV: " & CV
  End If
End Sub

' TODO TESTEN
' V3.6: 3.2.11 / V4: 3.15 (Leseanfrage Programmieren 3-Byte-Format - Pagemode)
Public Sub CmdLeseanfrageProgrammierenPagemode(ByVal CV As Integer)
  If VelueOutOfRange(CV, 1, 256, "Leseanfrage Programmieren Pagemode: CV") Then Exit Sub
  Call send.message(&H22, &H14, (CV Mod 256))
  Logfile.Details "Leseanfrage Programmieren Pagemode. CV: " & CV
End Sub

' TODO TESTEN
' V3.6: 3.2.12 / V4: 3.16
Public Sub CmdProgrammierergebnisAnfordern()
  Call send.message(&H21, &H10)
  Logfile.Details "Programmier Ergebnis Anfordern"
End Sub

' TODO TESTEN
' V3.6: 3.2.13 / V4: 3.17 (Schreibbefehl Programmieren 3-Byte-Format (Register-Mode) / LAN_X_DCC_WRITE_REGISTER)
Public Sub CmdSchreibbefehlProgrammierenRegisterMode(ByVal Register As Byte, ByVal daten As Integer)
  If VelueOutOfRange(Register, 1, 8, "Schreibbefehl Programmieren Registermode: Register") Then Exit Sub
  Call send.message(&H23, &H12, Register, daten)
  Logfile.Details "Schreibbefehl Programmieren RegisterMode. Register: " & Register & ", Value: " & daten
End Sub

' TODO TESTEN
' V3.6: 3.2.14-3.2.18 / V4: 3.18-3.22 (Schreibbefehl Programmieren 4-Byte-Format / LAN_X_CV_WRITE)
Public Sub CmdSchreibbefehlProgrammieren(ByVal CV As Integer, ByVal Val As Byte)
  If VelueOutOfRange(CV, 1, 1024, "Schreibbefehl Programmieren: CV") Then Exit Sub
  If VelueOutOfRange(Val, 0, 255, "Schreibbefehl Programmieren: Daten: ") Then Exit Sub
  If (Z21Protocol) Then
    CV = CV - 1
    Call send.message(&H24, &H12, (CV \ 256), (CV Mod 256), Val)
    Logfile.Details "LAN_X_CV_WRITE, CV: " & (CV + 1) & ", Value: " & Val
  Else
    If (ZentraleVersion < "3.6") Then
      Call send.message(&H23, &H16, (CV Mod 256), Val)
    Else
      Dim daten1 As Byte
      daten1 = &H1C + ((CV \ 256) Mod 4)
      Call send.message(&H23, daten1, (CV Mod 256), Val)
    End If
    Logfile.Details "Schreibbefehl Programmieren 4-Byte-Format. CV: " & CV & ", Value: " & Val
  End If
End Sub

' TODO TESTEN
' V3.6: 3.2.19 / V4: 3.23 (Schreibbefehl Programmieren 3-Byte-Format - Page-Mode)
Public Sub CmdSchreibbefehlProgrammierenPageMode(ByVal CV As Integer, ByVal daten As Integer)
  If VelueOutOfRange(CV, 1, 256, "Schreibbefehl Programmieren Pagemode: CV") Then Exit Sub
  Call send.message(&H23, &H17, (CV Mod 256), daten)
  Logfile.Details "Schreibbefehl Programmieren Pagemode. CV: " & CV & ", Value: " & daten
End Sub

' V3.6: 3.2.20 / V4: 3.24 (Softwareversion der Zentrale anfordern)
Public Sub CmdSoftwareversionZentraleAnfordern()
  Call send.message(&H21, &H21)
  Logfile.Details "Softwareversion der Zentrale anfordern"
End Sub

' V3.6: 3.2.21 / V4: 3.26 (Status der Zentrale anfordern)
Public Sub CmdStatusZentraleAnfordern()
  Call send.message(&H21, &H24)
  ' This command may also be internally used as UDP (Z21) Keep Alive Message
  ' In that case we should not log this like commands requested by the application
  If send.waitForlanKeepAliveResponse = False Then
    Logfile.Details "Status der Zentrale anfordern"
  End If
End Sub

' V3.6: 3.2.22 / V4: 3.27 (Zentralen-Startmode setzen)
Public Sub CmdZentralenStartmodeSetzen(ByVal AutoStart As Boolean)
  Dim value As Byte
  If AutoStart Then
    value = 2
    Else: value = 0
  End If
  Call send.message(&H22, &H22, value)
  Logfile.Details "Zentralen-Startmode setzen. Automatischer Start: " & AutoStart
End Sub

' V3.6: 3.2.23 / V4: 3.37 & 3.38 / Z21: 5.1 (Schaltinformationen anfordern / LAN_X_GET_TURNOUT_INFO)
Public Sub CmdFeedbackRequest(ByVal address As Integer, ByVal Nibble As Byte)
  ' STEP 1: Check the address size against the ZentraleVersion and Protocol.
  If (ZentraleVersion < "3.8") And (Not Z21Protocol) Then
    If VelueOutOfRange(address, 1, 1024, "Accessory: Address") Then Exit Sub
  End If
  If VelueOutOfRange(address, 1, 2048, "Accessory: Address") Then Exit Sub
  address = address - 1                      ' The LH100 numbering starts at 1, the wire at 0
  If Nibble <> 0 Then Nibble = 1             ' Ensure the nibble is 0 or 1
  Nibble = Nibble + 128                      ' The highest bit must be set
  ' Send the command that fits for this ZentraleVersion and Protocol.
  If Z21Protocol Then
    ' Note: the XpressNet length, as contained in the four low order bits of the header,
    ' seem to be incorrect! However, this is what the Z21 spec says...
    Call send.message(&H43, AccAH(address), AccAL(address))
  Else
    If (ZentraleVersion < "3.8") Then
      Call send.message(&H42, AccAL(address), Nibble)
    Else                                       ' Lenz V4 spec says Command station >= 3.8 SHOULD use the new form
      Call send.message(&H43, AccAH(address), AccAL(address), Nibble)
    End If
  End If
  ' STEP 5: Write the command details to the logfile
  Logfile.Details "Feedback Request - Switch: " & Switch + 1 & "  Nibble: " & Nibble
End Sub

' V3.6: 3.2.24 / V4: 3.39 & 3.40 / Z21: 5.2 (Schaltbefehl / LAN_X_SET_TURNOUT)
Public Sub CmdAccessory(ByVal address As Integer, ByVal position As Byte, ByVal Activate As Byte)
  'Position 0 = straight / 1 = turn
  ' Check the address size against the ZentraleVersion.
  If ZentraleVersion < 3.8 Then
    If VelueOutOfRange(address, 1, 1024, "Accessory: Address") Then Exit Sub
  Else
    If VelueOutOfRange(address, 1, 2048, "Accessory: Address") Then Exit Sub
  End If
  address = address - 1                      ' The LH100 numbering starts at 1, the wire at 0
  Dim daten As Byte                          ' Fill the Daten byte
  If Not Z21Protocol Then                    ' The Z21 uses switch addressing, instead of decoder addressing
    daten = (address Mod 4) * 2              ' Store the 2 address bits that belong to this decoder in daten
  End If
  daten = 128 + daten                        ' Set the MSB
  If Activate > 0 Then daten = daten + 8     ' Set the D1 bit (Activation)
  If position > 0 Then daten = daten + 1     ' Set the D2 bit (Position)
  ' Send the command that fits for this ZentraleVersion and Kennung.
  ' The IF clause is used for the LZV 200 and the Z21
  ' The Else clause is used for LZV100
  If (ZentraleVersion > "3.6") Or Z21Protocol Then
    Call send.message(&H53, AccAH(address), AccAL(address), daten)
  Else
    Call send.message(&H52, AccAL(address), daten)
  End If
  ' Write the command details to the logfile
  Logfile.Details "Accessory Command - Address: " & (address + 1) & "  Position: " & position & "  Activate: " & Activate
End Sub

' V3.6: 3.2.24 / V4: 3.39 & 3.40 (Schaltbefehl)
Public Sub CmdAccessoryAuto(ByVal Switch As Integer, ByVal position As Byte)
  ' Implemented in the AccAutoClass, Since it is not a standard XpressNet command.
  Call AccAuto.handle(Switch, position)
End Sub

' V3.6: 3.2.25 / V4: 3.41.3 / Z21: 4.1 (Lokinformationen anfordern)
Public Sub CmdLokInfoAnfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Lok Info Anfordern: Address") Then Exit Sub
  If Z21Protocol Then
    Call send.message(&HE3, &HF0, AH(address), AL(address))
  Else
    Call send.message(&HE3, &H0, AH(address), AL(address))
  End If
  Logfile.Details "Lok Informationen anfordern. Address: " & address
End Sub

' V3.6: 3.2.25.1 / V4: 3.41.4 (Funktionsstatus anfordern F0 – F12)
Public Sub CmdFunktionsStatusAnfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Funktionsstatus Anfordern: Address") Then Exit Sub
  Call send.message(&HE3, &H7, AH(address), AL(address))
  Logfile.Details "Funktionsstatus F0-F12 anfordern. Address: " & address
End Sub

' V3.6: 3.2.25.2 / V4: 3.41.5 (Funktionsstatus anfordern F13 – F28)
Public Sub CmdFunktionsStatusF13F28Anfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Funktionsstatus F13-F28 Anfordern: Address") Then Exit Sub
  Call send.message(&HE3, &H8, AH(address), AL(address))
  Logfile.Details "Funktionsstatus F13-F28 anfordern. Address: " & address
End Sub

' V3.6: 3.2.25.3 / V4: 3.41.7 (Funktionszustand anfordern F13 – F28)
Public Sub CmdFunktionsZustandF13F28Anfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Funktionszustand F13-F28 Anfordern: Address") Then Exit Sub
  Call send.message(&HE3, &H9, AH(address), AL(address))
  Logfile.Details "Funktionszustand F13-F28 anfordern. Address: " & address
End Sub

' V4: 3.41.6 (Funktionsstatus anfordern F29 – F68)
Public Sub CmdFunktionsStatusF29F68Anfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Funktionsstatus F29-F68 Anfordern: Address") Then Exit Sub
  Call send.message(&HE3, &HA, AH(address), AL(address))
  Logfile.Details "Funktionsstatus F29-F68 anfordern. Address: " & address
End Sub

' V4: 3.41.8 (Funktionszustand anfordern F29 – F68)
Public Sub CmdFunktionsZustandF29F68Anfordern(ByVal address As Integer)
  If VelueOutOfRange(address, 0, 9999, "Funktionszustand F29-F68 Anfordern: Address") Then Exit Sub
  Call send.message(&HE3, &HB, AH(address), AL(address))
  Logfile.Details "Funktionszustand F29-F68 anfordern. Address: " & address
End Sub

' V3.6: 3.2.26.1 / V4: 3.42.3 / Z21: 4.2 (Lok Fahrbefehl / LAN_X_SET_LOCO_DRIVE )
Public Sub CmdFahrbefehl(ByVal address As Integer, ByVal Speed As Integer, ByVal Steps As Integer, ByVal Forward As Boolean)
  Dim RVVV As Byte       ' The direction and speed byte
  Dim even As Boolean    ' To calculate Bit 4 of the 27/28 speed step
  If VelueOutOfRange(address, 0, 9999, "Fahrbefehl: Address") Then Exit Sub
  Select Case Steps
    Case 14:
      If VelueOutOfRange(Speed, 0, 14, "Fahrbefehl: Speed") Then Exit Sub
      If Speed > 0 Then RVVV = Speed + 1
      If Forward Then RVVV = RVVV + 128
      Call send.message(&HE4, &H10, AH(address), AL(address), RVVV)
    Case 27:
      If VelueOutOfRange(Speed, 0, 27, "Fahrbefehl: Speed") Then Exit Sub
      RVVV = EncodeSpeed(Speed)
      If Forward Then RVVV = RVVV + 128   ' The forward / reverse bit
      Call send.message(&HE4, &H11, AH(address), AL(address), RVVV)
    Case 28:
      If VelueOutOfRange(Speed, 0, 28, "Fahrbefehl: Speed") Then Exit Sub
      RVVV = EncodeSpeed(Speed)
      If Forward Then RVVV = RVVV + 128
      Call send.message(&HE4, &H12, AH(address), AL(address), RVVV)
    Case 128:
      If VelueOutOfRange(Speed, 0, 126, "Fahrbefehl: Speed") Then Exit Sub
      RVVV = Speed
      If Speed > 0 Then RVVV = RVVV + 1
      If Forward Then RVVV = RVVV + 128
      Call send.message(&HE4, &H13, AH(address), AL(address), RVVV)
  End Select
  Logfile.Details "Fahrbefehl - Address: " & address & ", Speed: " & Speed
End Sub

' V3.6: 3.2.26.2 / V4: 3.42.4 (Lok Funktionsbefehl)
Public Sub CmdLokFunktionBefehl(ByVal address As Integer, ByVal group As Byte, ByVal data As Byte)
  Select Case group
    Case 1: Call send.message(&HE4, &H20, AH(address), AL(address), data)
    Case 2: Call send.message(&HE4, &H21, AH(address), AL(address), data)
    Case 3: Call send.message(&HE4, &H22, AH(address), AL(address), data)
    Case 4: Call send.message(&HE4, &H23, AH(address), AL(address), data)
    Case 5: Call send.message(&HE4, &H28, AH(address), AL(address), data)
    Case 6: Call send.message(&HE4, &H29, AH(address), AL(address), data)
    Case 7: Call send.message(&HE4, &H2A, AH(address), AL(address), data)
    Case 8: Call send.message(&HE4, &H2B, AH(address), AL(address), data)
    Case 9: Call send.message(&HE4, &H50, AH(address), AL(address), data)
    Case 10: Call send.message(&HE4, &H51, AH(address), AL(address), data)
  End Select
 Logfile.Details "Funktions befehl. Address: " & address & ", Befehl: " & data
End Sub

' V3.6: 3.2.26.4 / V4: 3.42.6 (Funktionsstatus setzen)
Public Sub CmdLokFunktionStatus(ByVal address As Integer, ByVal group As Byte, ByVal data As Byte)
  Select Case group
    Case 1: Call send.message(&HE4, &H24, AH(address), AL(address), data)
    Case 2: Call send.message(&HE4, &H25, AH(address), AL(address), data)
    Case 3: Call send.message(&HE4, &H26, AH(address), AL(address), data)
    Case 4: Call send.message(&HE4, &H27, AH(address), AL(address), data)
    Case 5: Call send.message(&HE4, &H2C, AH(address), AL(address), data)
    Case 6: Call send.message(&HE4, &H2D, AH(address), AL(address), data)
    Case 7: Call send.message(&HE4, &H2E, AH(address), AL(address), data)
    Case 8: Call send.message(&HE4, &H52, AH(address), AL(address), data)
    Case 9: Call send.message(&HE4, &H53, AH(address), AL(address), data)
    Case 10: Call send.message(&HE4, &H54, AH(address), AL(address), data)
  End Select
 Logfile.Details "Status befehl. Address: " & address & ", Status: " & data
End Sub

' V3.6: 3.2.28.1 / V4: 3.43.1 (Programming on Main Byte schreiben / LAN_X_CV_POM_WRITE_BYTE)
Public Sub CmdPomWrite(ByVal address As Integer, ByVal CV As Integer, ByVal value As Byte)
  If VelueOutOfRange(address, 0, 9999, "PoM write: Address") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "PoM write: CV") Then Exit Sub
  CV = CV - 1
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  cvhigh = CV \ 256 + &HEC    ' 0xEC = PoM Schreiben
  Call send.message(&HE6, &H30, AH(address), AL(address), cvhigh, cvlow, value)
  Logfile.Details "Pom Write. Address: " & address & ", CV: " & CV + 1 & ", Value: " & value
End Sub

' V3.6: 3.2.28.2 / V4: 3.43.2 (Programming on Main Byte lesen / LAN_X_CV_POM_READ_BYTE)
Public Sub CmdPomRead(ByVal address As Integer, ByVal CV As Integer)
  If VelueOutOfRange(address, 0, 9999, "PoM Read: Address") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "PoM Read: CV") Then Exit Sub
  CV = CV - 1
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  cvhigh = CV \ 256 + &HE4        ' 0xE4 = PoM Lesen
  Call send.message(&HE6, &H30, AH(address), AL(address), cvhigh, cvlow, 0)
  Logfile.Details "Pom Read. Address: " & address & ", CV: " & CV + 1
End Sub

' V3.6: 3.2.28.3 / V4: 3.43.3 (Programming on Main Bit schreiben / LAN_X_CV_POM_WRITE_BIT)
Public Sub CmdPomBitWrite(ByVal address As Integer, ByVal CV As Integer, ByVal bitValue As Boolean, ByVal bitPosition As Byte)
  If VelueOutOfRange(address, 0, 9999, "PoM bit write: Address") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "PoM bit write: CV") Then Exit Sub
  If VelueOutOfRange(bitPosition, 0, 7, "PoM bit write: bit position") Then Exit Sub
  CV = CV - 1
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  If Z21Protocol Then
    cvhigh = CV \ 256 + &HE8               ' 0xE8 = PoM Bit Schreiben (23151 spec says this as well
  Else
    cvhigh = CV \ 256 + &H7C               ' 0x7C = PoM Bit Schreiben: according to V4 spec
  End If
  Dim value As Byte
  value = bitPosition + &HF0               ' highest order bits must all be 1
  If bitValue Then value = value + 8       ' bit 4 represents the new bit value
  Call send.message(&HE6, &H30, AH(address), AL(address), cvhigh, cvlow, value)
  Logfile.Details "Pom Bit Write. Address: " & address & ", CV: " & CV + 1 & ", Bit value: " & bitValue & ", Bit Position: " & bitPosition
End Sub

' V3.6 3.2.30.3 / V4: 3.44.3 (Adressanfrage Lok in Zentralenstack )
Public Sub CmdAdresAnfrageLokInStack(ByVal address As Integer, ByVal Vorwarts As Boolean)
  Dim Kennung As Byte
  If Vorwarts Then
    Kennung = 5
  Else
    Kennung = 6
  End If
  Call send.message(&HE3, Kennung, AH(address), AL(address))
  Logfile.Details "Adressanfrage Lok in Zentralenstack. Address: " & address & ", Vorwarts : " & Vorwarts
End Sub

' V3.6: 3.2.31 / V4: 3.44.4 (Lok aus Stack loeschen / LAN_X_PURGE_LOCO)
Public Sub CmdLokAusStack(ByVal address As Integer)
  Call send.message(&HE3, &H44, AH(address), AL(address))
  Logfile.Details "Lok aus Stack loeschen. Address: " & address
End Sub


'==================================================================================================================
' V4 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
' The commands below will be forwarded by the Interface to the Command Station

' V4: 3.25 (Erweiterte Zentralen Versionsinformation)
Public Sub CmdErweiterteZentralenVersionsinformation()
  Call send.message(&H21, &H23)
  Logfile.Details "Erweiterte Zentralen Versionsinformation"
End Sub

' V4: 3.30 (PoM-Ergebnis anfordern)
Public Sub CmdPoMErgebnisAnfordern()
  Call send.message(&H21, &H27)
  Logfile.Details "PoM Programmier Ergebnis Anfordern"
End Sub

' V4: 3.31 (Zentralen Reset auslosen)
Public Sub CmdZentralenResetAuslosen()
  Call send.message(&H21, &H28)
  Logfile.Details "Zentralen Reset auslosen"
End Sub

' V4: 3.33 (Modellzeit anfordern)
Public Sub CmdModellzeitAnfordern()
  Call send.message(&H21, &H2A)
  Logfile.Details "Modellzeit anfordern"
End Sub

' V4: 3.34 (Modellzeit stellen)
Public Sub CmdModellzeitStellen(ByVal day As Byte, ByVal hour As Byte, ByVal minute As Byte, ByVal factor As Byte)
  If VelueOutOfRange(day, 0, 6, "Modellzeit: Day") Then Exit Sub
  If VelueOutOfRange(hour, 0, 23, "Modellzeit: Hour") Then Exit Sub
  If VelueOutOfRange(minute, 0, 59, "Modellzeit: Minute") Then Exit Sub
  If VelueOutOfRange(factor, 0, 31, "Modellzeit: Factor") Then Exit Sub
  Dim daten As Byte
  daten = day * 32 + hour
  Call send.message(&H24, &H2B, daten, minute, factor)
  Logfile.Details "Modellzeit stellen"
End Sub

' V4: 3.35 (Modellzeit anhalten - V4.0)
Public Sub CmdModellzeitAnhalten()
  Call send.message(&H21, &H2D)
  Logfile.Details "Modellzeit anhalten"
End Sub

' V4: 3.36 (Modellzeit starten - V4.0)
Public Sub CmdModellzeitStarten()
  Call send.message(&H21, &H2C)
  Logfile.Details "Modellzeit starten"
End Sub


'==================================================================================================================
' Z21 ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
' The commands below will be forwarded by the Interface to the Command Station
' Note: these Z21 could not be tested, and may therefore not work as intended

' Z21: 2.15 (LAN_X_GET_FIRMWARE_VERSION)
Public Sub CmdZ21GetFirmwareVersion()
  Call send.message(&HF1, &HA)
  Logfile.Details "Z21: CmdZ21GetFirmwareVersion"
End Sub

' Z21: 4.3.1 (LAN_X_SET_LOCO_FUNCTION)
Public Sub CmdZ21SetLocoFunction(ByVal address As Integer, ByVal CV As Byte, ByVal TT As Byte)
  ' TT Umschalttyp: 00=aus, 01=ein, 10=umschalten,11=nicht erlaubt
  ' CV: 0x00=F0 (Licht), 0x01=F1 usw. bis F31
  If VelueOutOfRange(CV, 0, 31, "LAN_X_SET_LOCO_FUNCTION: CV") Then Exit Sub
  If VelueOutOfRange(TT, 0, 2, "LAN_X_SET_LOCO_FUNCTION: TT") Then Exit Sub
  Dim TTNNNNNN As Byte
  TTNNNNNN = TT * 64 + CV
  Call send.message(&HE4, &HF8, AH(address), AL(address), TTNNNNNN)
  Logfile.Details "Z21: CmdZ21SetLocoFunction. CV: " & CV & ", TT: " & TT
End Sub

' Z21: 5.4 LAN_X_SET_EXT_ACCESSORY
Public Sub CmdZ21SetExtAccessory(ByVal address As Integer, ByVal value As Byte)
  If VelueOutOfRange(address, 1, 2048, "Extended Accessory: Address") Then Exit Sub
  ' The first "user address" is 1, which is represented as 4 in the Xpressnet message
  Dim rawAddress As Integer
  rawAddress = address + 3
  If rawAddress >= 2048 Then rawAddress = rawAddress - 2048
  Call send.message(&H54, (rawAddress \ 256), (rawAddress Mod 256), value, &H0)
  Logfile.Details "Z21: CmdZ21SetExtAccessory. Address: " & address & ", Value: " & value
End Sub

' Z21: 5.5 LAN_X_GET_EXT_ACCESSORY_INFO
Public Sub CmdZ21GetExtAccessory(ByVal address As Integer)
  If VelueOutOfRange(address, 1, 2048, "Extended Accessory: Address") Then Exit Sub
  ' The first "user address" is 1, which is represented as 4 in the Xpressnet message
  Dim rawAddress As Integer
  rawAddress = address + 3
  If rawAddress >= 2048 Then rawAddress = rawAddress - 2048
  ' As opposed to normal accesories, no address correction takes place
  ' Note: It seems the Z21 specification has an error here: length in the header (4) doesn't match actual length
  ' Either the header should have been &H43, or an extra byte should have been added.
  ' Since we are not able to test this, we decided to add an extra byte. This may be incorrect, however.
  Call send.message(&H44, (rawAddress \ 256), (rawAddress Mod 256), &H0, &H0)
  Logfile.Details "Z21: CmdZ21SetExtAccessory. Address: " & address
End Sub

' Z21: 6.9 (LAN_X_CV_POM_ACCESSORY_WRITE_BYTE)
Public Sub CmdZ21AccPomWrite(ByVal decoderAddress As Integer, ByVal CDDD As Byte, ByVal CV As Integer, ByVal value As Byte)
  If VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM write: Decoder Address") Then Exit Sub
  If VelueOutOfRange(CDDD, 0, 15, "Accessory PoM write: CDDD") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "Accessory PoM write: CV") Then Exit Sub
  CV = CV - 1
  ' Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
  ' Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
  ' Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
  Dim aaaaa As Byte
  Dim AAAACDDD As Byte
  aaaaa = decoderAddress \ 16
  AAAACDDD = (decoderAddress Mod 16) * 16 + CDDD
  ' CV bytes
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  cvhigh = CV \ 256 + &HEC    ' 0xEC = Acc PoM Schreiben
  Call send.message(&HE6, &H31, aaaaa, AAAACDDD, cvhigh, cvlow, value)
  Logfile.Details "Accessory PoM Write. Decoder address: " & decoderAddress & ", CDDD: " & CDDD & ", CV: " & CV + 1 & ", Value: " & value
End Sub

' Z21: 6.10 (LAN_X_CV_POM_ ACCESSORY_WRITE_BIT )
Public Sub CmdZ21AccPomBitWrite(ByVal decoderAddress As Integer, ByVal CDDD As Byte, ByVal CV As Integer, ByVal bitValue As Boolean, ByVal bitPosition As Byte)
  If VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM Bit write: Decoder Address") Then Exit Sub
  If VelueOutOfRange(CDDD, 0, 15, "Accessory PoM Bit write: CDDD") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "PoM Bit write: CV") Then Exit Sub
  CV = CV - 1
  ' Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
  ' Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
  ' Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
  Dim aaaaa As Byte
  Dim AAAACDDD As Byte
  aaaaa = decoderAddress \ 16
  AAAACDDD = (decoderAddress Mod 16) * 16 + CDDD
  ' CV bytes
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  cvhigh = CV \ 256 + &HE8    ' 0xEC = Acc PoM Bit Schreiben
  ' DB5: 0000VPPP
  Dim VPPP
  VPPP = bitPosition
  If bitValue Then VPPP = VPPP + 8       ' bit 4 represents the new bit value
  Call send.message(&HE6, &H31, aaaaa, AAAACDDD, cvhigh, cvlow, VPPP)
  Logfile.Details "Accessory PoM Bit Write. Decoder address: " & decoderAddress & ", CDDD: " & CDDD & ", CV: " & CV + 1 & ", bitValue: " & bitValue & ", bitPosition: " & bitPosition
End Sub

' Z21: 6.11 (LAN_X_CV_POM_ACCESSORY_READ_BYTE)
Public Sub CmdZ21AccPomRead(ByVal decoderAddress As Integer, ByVal CDDD As Byte, ByVal CV As Integer)
  If VelueOutOfRange(decoderAddress, 0, 511, "Accessory PoM read: Decoder Address") Then Exit Sub
  If VelueOutOfRange(CDDD, 0, 15, "Accessory PoM read: CDDD") Then Exit Sub
  If VelueOutOfRange(CV, 1, 1024, "Accessory PoM read: CV") Then Exit Sub
  CV = CV - 1
  ' Es gilt: aaaaaAAAACDDD = (Decoder Addresse & 0x1FF) << 4) | CDDD;
  ' Falls CDDD=0000, dann bezieht sich die CV auf den ganzen Decoder.
  ' Falls C=1, so ist DD die Nummer des zu programmierenden Ausgangs.
  Dim aaaaa As Byte
  Dim AAAACDDD As Byte
  aaaaa = decoderAddress \ 16
  AAAACDDD = (decoderAddress Mod 16) * 16 + CDDD
  ' CV bytes
  Dim cvlow As Byte
  cvlow = CV Mod 255
  Dim cvhigh As Byte
  cvhigh = CV \ 256 + &HE4    ' 0xE4 = Acc PoM Lesen
  Call send.message(&HE6, &H31, aaaaa, AAAACDDD, cvhigh, cvlow, &H0)
  Logfile.Details "Accessory PoM Read. Decoder address: " & decoderAddress & ", CDDD: " & CDDD & ", CV: " & CV + 1
End Sub

' Z21: 6.12 (LAN_X_MM_WRITE_BYTE )
Public Sub CmdZ21MmWriteByte(ByVal RegAdr As Byte, ByVal value As Byte)
  Call send.message(&H24, &HFF, &H0, RegAdr, value)
  Logfile.Details "Z21: CmdZ21MmWriteByte"
End Sub


'==================================================================================================================
' OpenDCC ONLY METHODS TO GENERATE COMMAND STATION SPECIFIC MESSAGES
' The commands below will be forwarded by the Interface to the Command Station
' Note: these commands could not be tested, and may therefore not work as intended
Public Sub CmdOpenDCCSetExtAccessory(ByVal address As Integer, ByVal value As Byte)
  If VelueOutOfRange(address, 1, 2048, "Ext. Accessory: Address") Then Exit Sub
  If VelueOutOfRange(value, 0, 31, "Ext. Accessory: Value") Then Exit Sub
  Dim byte1 As Byte
  Dim byte2 As Byte
  byte1 = value * 8                   ' Shift value 3 bits left
  byte1 = byte1 + address \ 256       ' 3 highest address bits
  byte2 = address Mod 256             ' 8 lowest address bits
  Call send.message(&H13, &H1, byte1, byte2)
  Logfile.Details "OPENDCC: Extended Accessory Command. Address = " & address & ", Value = " & value
End Sub

'==================================================================================================================
'==================================================================================================================
'==================================================================================================================
' RAISING EVENTS: TCP, USB AND XPRESSNET FOR ERRORS, STATUS CHANGES AND FEEDBACK INFO
' SHOULD NOT BE CALLED BY THE EVENT USER!
' The subroutines below forward events from lower level objects to the DLL user
' In addition Logfile messages are written, if needed.
Public Sub eventInterfaceConnected()
   RaiseEvent InterfaceConnected
   Logfile.Details "Interface connection established"
   ' In case of a serial connection, it seems we must wait 100ms before the interface is stable
   If connectionType = serialPort Then Sleep (100)
   ' Retrieve version and status info regarding the Command Station
   Call CmdSoftwareversionZentraleAnfordern
   Call CmdStatusZentraleAnfordern
End Sub

Public Sub eventTcpError(ByRef Description As String)
   RaiseEvent TcpError(Description)
   Logfile.Errors "TCP Error: " & Description
End Sub

Public Sub eventTcpClosedByRemoteHost()
   RaiseEvent TcpClosedByRemoteHost
   Logfile.Errors "TCP Connection closed by remote host"
   Set tcp.Parent = Nothing
   Set tcp = Nothing
   connectionType = none
End Sub


'------------------------------------------------------------------------------------------------------------------
' Interface specific events
' V3.6: 1.6 (Versionsnummer des Interface 23151 feststellen)
Public Sub eventInterfaceVersion(ByVal daten1 As Byte, ByVal daten2 As Byte)
  Dim Version As String
  Dim code As String
  Version = (daten1 \ 16) & "." & (daten1 Mod 16)
  code = (daten2 \ 16) & "." & (daten2 Mod 16)
  RaiseEvent InterfaceVersion(Version, code)
  Logfile.Details "Interface - Version: " & Version & "  Code: " & code
End Sub
  
' V3.6: 1.7 (XpressNet Address) / V3.6: 2.1 (Interface Status Befehl)
Public Sub eventXpressNetAddress(ByVal address As Byte)
  ' This event is triggered as response to two commands:
  ' XpressNet Address (CmdXpressNetAddress - V3.6: 1.7) and
  ' Interface Status Befehl (CmdInterfaceStatus - V3.6: 2.1)
  ' CmdInterfaceStatus will be called internally by the DLL, whenever the
  ' keep alive timer expires (send.lanKeepAliveTimer_Timer). In such case,
  ' the response to the keep alive message should not be returned (as event)
  ' to the application program, but instead be kept within the DLL.
  ' To distinguish between both cases, the send.lanKeepAliveTimer_Timer sets the
  ' waitForlanKeepAliveResponse flag
  If send.waitForlanKeepAliveResponse Then
    send.waitForlanKeepAliveResponse = False
    Logfile.Details "TCP Keep Alive Response"
  Else
    RaiseEvent XpressNetAddress(address)
    Logfile.Details "XpressNet - Address: " & address
  End If
End Sub

' V3.6: 2.2 (XpressNet Version)
Public Sub eventXpressNetVersion(ByVal daten As Byte)
  Dim Version As String
  Version = (daten \ 16) & "." & (daten Mod 16)
  RaiseEvent XpressNetVersion(Version)
  Logfile.Details "XpressNet - Version: " & Version
End Sub

' V3.6: 2.3 (Freie Verbindungen)
Public Sub eventInterfaceFreieVerbindungen(ByVal number As Byte)
  RaiseEvent InterfaceFreieVerbindungen(number)
  Logfile.Details "Interface - Free TCP Sockets: " & number
End Sub

'------------------------------------------------------------------------------------------------------------------
' Xpressnet specific events
' V3.6: 3.1.1.1, V4: 2.4.1, Z21: 2.8 (LAN_X_BC_TRACK_POWER_ON)
Public Sub eventAllesAn()
   ZentraleNotAus = False
   ZentraleNotHalt = False
   ZentraleProgrammierMode = False
   RaiseEvent AllesAn
   Logfile.Details "Broadcast: Alles An"
End Sub
 
' V3.6: 3.1.1.2, V4: 2.4.2, Z21: 2.7 (LAN_X_BC_TRACK_POWER_OFF)
Public Sub eventNotAus()
   ZentraleNotAus = True
   RaiseEvent NotAus
   Logfile.Details "Broadcast: Alles Aus (Notaus)"
End Sub
 
' V3.6: 3.1.1.3, V4: 2.4.3, Z21: 2.14 (LAN_X_BC_STOPPED)
Public Sub eventNotHalt()
   ZentraleNotHalt = True
   RaiseEvent NotHalt
   Logfile.Details "Broadcast: Alle Loks Aus - Nothalt"
End Sub
 
' V3.6: 3.1.1.4, V4: 2.4.4, Z21: 2.9 (LAN_X_BC_PROGRAMMING_MODE)
Public Sub eventProgrammierMode()
   ZentraleProgrammierMode = True
   RaiseEvent ProgrammierMode
   Logfile.Details "Broadcast: Programmiermode"
End Sub

' To simplify the application program, we generate different feedback events for switches and for
' feedback (occupancy) decoders. The difference between both can be determined by the TT bits
' V3.6: 3.1.1.5, V4: 2.4.5, V4: 2.4.6 (BC RuckMeldung) / V3.6: 3.1.8, V4: 2.15, V4: 2.16 (Schaltinformation)
Public Sub eventFeedback(ByVal address As Integer, ByVal Nibble As Byte, ByVal ttBits As Byte, ByVal iBit As Byte, ByVal data As Byte)
  If ttBits = 2 Then                      ' normal feedback decoder
    ' We add 1 to the RS-Bus address, since addresses on the wire range from 0..127. We display these as 1..128
    RaiseEvent Feedback(address + 1, Nibble, data)
    Logfile.Details "Feedback Message received for RS-Bus address: " & address & "  Data: " & data & "  Nibble: " & Nibble
  Else                                    ' switch decoder
    ' Adress is the decoder address, but we will return a switch address
    Dim switchAddress As Integer
    If Not Z21Protocol Then
      switchAddress = (address * 4) + (Nibble * 2) + 1
    Else
      switchAddress = address + 1
    End If
    RaiseEvent SwitchFeedback(switchAddress, ttBits, iBit, data)
    Logfile.Details "Feedback Message received for switch address: " & switchAddress & "  Data: " & data & "  TT-Bits: " & ttBits & " I-Bit: " & iBit
  End If
End Sub

' V3.6: 3.1.2.1, V4: 2.5.1, Z21: 2.10 (Programmierinfo Kurzschluss / LAN_X_BC_TRACK_SHORT_CIRCUIT)
Public Sub eventProgrammierinfoKurzschluss()
  RaiseEvent ProgrammierinfoKurzschluss
  Logfile.Details "Programmierinfo Kurzschluss"
End Sub

' V3.6: 3.1.2.2, V4: 2.5.2, Z21: 6.4 (Programmierinfo Daten nicht gefunden / LAN_X_CV_NACK)
Public Sub eventProgrammierinfoDatenNichtGefunden()
  RaiseEvent ProgrammierinfoDatenNichtGefunden
  Logfile.Details "Programmierinfo Daten nicht gefunden"
End Sub

' V3.6: 3.1.2.3, V4: 2.5.3 (Programmierinfo Zentrale Busy)
Public Sub eventProgrammierinfoZentraleBusy()
  RaiseEvent ProgrammierinfoZentraleBusy
  Logfile.Details "Programmierinfo Zentrale Busy"
End Sub

' V3.6: 3.1.2.4, V4: 2.5.4 (Programmierinfo Zentrale Bereit)
Public Sub eventProgrammierinfoZentraleBereit()
  RaiseEvent ProgrammierinfoZentraleBereit
  Logfile.Details "Programmierinfo Zentrale Bereit"
End Sub

' V3.6 3.1.2.5 / V4: 2.5.5 (Programmierinfo Daten 3-Byte-Format)
' ProgrammierinfoPageRegister(ByVal EEPromAddress As Byte, ByVal Daten As Byte)
Public Sub eventProgInfoPageRegister(ByVal EEPromAddress As Byte, ByVal value As Byte)
  RaiseEvent ProgrammierinfoPageRegister(EEPromAddress, value)
  Logfile.Details "Programmierinfo Page / Register. EEPromAddress: " & EEPromAddress & ", Value: " & value
End Sub

' V3.6: 3.1.2.6 - 3.1.2.9, V4: 2.5.6 - 2.5.9 (Programmierinfo CV1..1024)
Public Sub eventProgInfo(ByVal CvRange As Byte, ByVal daten1 As Byte, ByVal value As Byte)
  Dim CV As Integer
  CV = daten1
  Select Case CvRange
    Case 1:
      If CV = 0 Then
        If ZentraleVersion < 3.6 Then
          CV = 256
        Else
          CV = 1024
        End If
      End If
    Case 2: CV = CV + 256
    Case 3: CV = CV + 512
    Case 4: CV = CV + 768
  End Select
  RaiseEvent ProgrammierinfoCV(CV, value)
  Logfile.Details "Programmierinfo. CV: " & CV & ", Value: " & value
End Sub

' Z21: 6.5 (LAN_X_CV_RESULT)
Public Sub eventProgInfoZ21(CVAdr_MSB, CVAdr_LSB, value)
  Dim CV As Integer
  CV = (CVAdr_MSB * 256) + CVAdr_LSB + 1
  RaiseEvent ProgrammierinfoCV(CV, value)
  Logfile.Details "Programmierinfo Z21. CV: " & CV & ", Value: " & value
End Sub

' V4: 2.6 (Service Variable melden)
Public Sub eventServiceVariableMelden(ByVal SvAddress As Byte, ByVal SvValue As Byte)
  RaiseEvent ServiceVariable(SvAddress, SvValue)
  Logfile.Details "Service Variable melden. SV Address: " & SvAddress & ", SV Value: " & SvValue
End Sub

' V3.6: 3.1.3, V4: 2.7.2, Z21: 2.3 (Softwareversion Zentrale / LAN_X_GET_VERSION)
Public Sub eventSoftwareVersionZentrale(ByVal daten2 As Byte, ByVal daten3 As Byte)
  Dim Version As String
  Version = (daten2 \ 16) & "." & (daten2 Mod 16)
  ZentraleVersion = Version
  ZentraleKennung = daten3
  RaiseEvent ZentraleVersion(Version)
  Logfile.Details "Version Zentrale: " & Version & ", Kennung: " & ZentraleKennung
End Sub

' V3.6: 3.1.4, V4: 2.8, Z21: 2.12 (Status Zentrale / LAN_X_STATUS_CHANGED)
Public Sub eventStatusZentrale(ByVal data As Byte)
  ' With Z21, we internally use the StatusZentrale command as "Keep Alive" Message
  ' If we use the StatusZentrale command in this way, we need not inform the application
  ' about the result. In that case it is sufficient to tell the remaining program that
  ' we no longer wait for the Keep Alive Response. Just log it that we received
  ' the response, and exit
  If connectionType = lanUdpPort And send.waitForlanKeepAliveResponse Then
    send.waitForlanKeepAliveResponse = False
    Logfile.Details "UDP Keep Alive Response"
    Exit Sub
  End If
  ' If we are here, this is not a Z21 Keep Alive, but a response to a normal request.
  ZentraleNotHalt = bitValue(data, 0)
  ZentraleNotAus = bitValue(data, 1)
  If Not Z21Protocol Then ZentraleStartmodeAuto = bitValue(data, 2)
  If Not Z21Protocol Then ZentraleProgrammierMode = bitValue(data, 3)
  If Z21Protocol Then ZentraleKurzSchluss = bitValue(data, 2)
  If Z21Protocol Then ZentraleProgrammierMode = bitValue(data, 5)
  ZentraleKaltstart = bitValue(data, 6)
  ZentraleRamCheckFehler = bitValue(data, 7)
  If ZentraleNotHalt Then
    RaiseEvent NotHalt
    Logfile.Details "Startup: NotHalt"
  End If
  If ZentraleNotAus Then
    RaiseEvent NotAus
    Logfile.Details "Startup: NotAus"
  End If
  If ZentraleProgrammierMode Then
    RaiseEvent ProgrammierMode
    Logfile.Details "Startup: ProgrammierMode"
  End If
  If ZentraleRamCheckFehler Then
    Logfile.Errors "Startup: Zentrale - Ram Check Fehler"
  End If
  Logfile.Details "Status Zentrale: " & data
End Sub

' V4: 2.10 (PoM Ergebnis melden)
Public Sub eventPomErgebnisMelden(ByVal AdrH As Byte, ByVal AdrL As Byte, ByVal CvValue As Byte)
  Dim LokAddress As Integer
  LokAddress = AdrH * 256 + AdrL
  RaiseEvent PomErgebnis(LokAddress, CvValue)
  Logfile.Details "PomErgebnisMelden. Lok Address:: " & LokAddress & ", CV Value:" & CvValue
End Sub

' V3.6: 3.1.9.1, V4: 2.19.1 (Lok Info Normale Lok)
Public Sub eventNormaleLokinfo(ByVal Kennung As Byte, ByVal Speed As Byte, ByVal F0_F4 As Byte, ByVal F5_F12 As Byte)
  Dim Besetzt As Boolean
  Dim Fahrstufen As Byte
  Dim Forwarts As Boolean
  Dim Geschwindigkeit As Integer
  Dim F5_F8 As Byte
  Dim F9_F12 As Byte
  Besetzt = bitValue(Kennung, 3)
  Fahrstufen = 0
  Select Case (Kennung Mod 8)       ' Analyse bits 0..2
    Case 0: Fahrstufen = 14
    Case 1: Fahrstufen = 27
    Case 2: Fahrstufen = 28
    Case 4: Fahrstufen = 128
  End Select
  Forwarts = bitValue(Speed, 7)
  Geschwindigkeit = DecodeSpeed(Speed)
  F5_F8 = F5_F12 Mod 16      ' Gruppe 2 - take the lowest 4 bits
  F9_F12 = F5_F12 \ 16       ' Gruppe 3 - take the highest 4 bits
  RaiseEvent NormaleLokinfo(Besetzt, Fahrstufen, Geschwindigkeit, Forwarts, F0_F4, F5_F8, F9_F12)
  Logfile.Details "Lokinformation (normal) - Besetzt: " & Besetzt & ", Speed: " & Geschwindigkeit & ", Stufen: " & Fahrstufen & ", Forwarts: " & Forwarts & ", F0..F4 " & F0_F4 & ", F5..F8 " & F5_F8 & ", F9..F12 " & F9_F12
End Sub

' Z21: 4.4 (LAN_X_LOCO_INFO)
Public Sub eventZ21Lokinfo(ByVal Adr_MSB As Byte, ByVal Adr_LSB As Byte, ByVal Kennung As Byte, ByVal Speed As Byte, ByVal F0_F4 As Byte, ByVal F5_F12 As Byte, ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)
  Dim LokAddress As Integer
  Dim Besetzt As Boolean
  Dim Fahrstufen As Byte
  Dim Forwarts As Boolean
  Dim Geschwindigkeit As Integer
  Dim F5_F8 As Byte
  Dim F9_F12 As Byte
'  LokAddress = (message(1) And &H3F) * 256 + message(2)
  LokAddress = (Adr_MSB And &H3F) * 256 + Adr_LSB
  Besetzt = bitValue(Kennung, 3)
  Fahrstufen = 0
  Select Case (Kennung Mod 8)       ' Analyse bits 0..2
    Case 0: Fahrstufen = 14
    Case 1: Fahrstufen = 27
    Case 2: Fahrstufen = 28
    Case 4: Fahrstufen = 128
  End Select
  Forwarts = bitValue(Speed, 7)
  Geschwindigkeit = DecodeSpeed(Speed)
  F5_F8 = F5_F12 Mod 16      ' Gruppe 2 - take the lowest 4 bits
  F9_F12 = F5_F12 \ 16       ' Gruppe 3 - take the highest 4 bits
  RaiseEvent Z21Lokinfo(LokAddress, Besetzt, Fahrstufen, Geschwindigkeit, Forwarts, F0_F4, F5_F8, F9_F12, F13_F20, F21_F28)
  Logfile.Details "Lokinformation (Z21) - Address: " & LokAddress & ", Besetzt: " & Besetzt & ", Speed: " & Geschwindigkeit & ", Stufen: " & Fahrstufen & ", Forwarts: " & Forwarts & ", F0..F4 " & F0_F4 & ", F5..F8 " & F5_F8 & ", F9..F12 " & F9_F12 & ", F13..F20 " & F13_F20 & ", F21..F28 " & F21_F28
End Sub

' V3.6: 3.1.9.2, V4: 2.19.2 (Funktionszustand F13 bis F28)
Public Sub eventFunctionsZustandF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte)
  RaiseEvent FunctionsZustandF13F20(F13_F20, F21_F28)
  Logfile.Details "Funktionszustand - F13 bis F20: " & F13_F20 & ", F21 bis F28: " & F21_F28
End Sub

' V4: 2.19.3 (Funktionszustand F29 bis F68)
Public Sub eventFunctionsZustandF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte)
  RaiseEvent FunctionsZustandF29F68(F29_F36, F37_F44, F45_F52, F53_F60, F61_F68)
  Logfile.Details "Funktionszustand - F29 bis F68: " & F29_F36 & ", " & F37_F44 & ", " & F45_F52 & ", " & F53_F60 & ", " & F61_F68
End Sub

' V3.6: 3.1.11, V4: 2.19.8 (Funktionsstatus F0 bis F12 der angefragten Lok)
Public Sub eventFunctionStatusF0F12(ByVal F0_F4 As Byte, ByVal F5_F12 As Byte)
  Dim F5_F8 As Byte
  Dim F9_F12 As Byte
  F5_F8 = F5_F12 Mod 16      ' Gruppe 2 - take the lowest 4 bits
  F9_F12 = F5_F12 \ 16       ' Gruppe 3 - take the highest 4 bits
  RaiseEvent FunctionStatusF0F12(F0_F4, F5_F8, F9_F12)
  Logfile.Details "Funktionsstatus - F0..F4 " & F0_F4 & ", F5..F8 " & F5_F8 & ", F9..F12 " & F9_F12
End Sub

' V3.6: 3.1.12, V4: 2.20 (Funktionsstatus F13 bis F28)
Public Sub eventFunctionStatusF13F20(ByVal F13_F20 As Byte, ByVal F21_F28 As Byte, ByVal RefreshMode As Byte)
  RaiseEvent FunctionStatusF13F20(F13_F20, F21_F28, RefreshMode)
  Logfile.Details "Funktionsstatus - F13 bis F20: " & F13_F20 & ", F21 bis F28: " & F21_F28 & ", Refresh Mode: " & RefreshMode
End Sub

' V4: 2.21 (Funktionsstatus F29 bis F68)
Public Sub eventFunctionStatusF29F68(ByVal F29_F36 As Byte, ByVal F37_F44 As Byte, ByVal F45_F52 As Byte, ByVal F53_F60 As Byte, ByVal F61_F68 As Byte)
  RaiseEvent FunctionStatusF29F68(F29_F36, F37_F44, F45_F52, F53_F60, F61_F68)
  Logfile.Details "Funktionsstatus - F29 bis F68: " & F29_F36 & ", " & F37_F44 & ", " & F45_F52 & ", " & F53_F60 & ", " & F61_F68
End Sub

' V3.6: 3.1.14, V4: 2.24 (Fehlermeldung)
Public Sub eventFehlermeldung(ByVal data As Byte)
  Dim value As Byte
  If bitValue(data, 7) Then                ' Highest order bit must have been set
    value = data Mod 16                   ' Take the four lowest order bits
    RaiseEvent Fehlermeldung(value)
    If LogWarnings Then
      Logfile.AppendWithHost "Fehlermeldung: " & value
    End If
  End If
End Sub

' V4: 2.4.7 / V4: 2.11 (BC: Modellzeit / Modellzeit melden)
Public Sub eventModellZeit(ByVal daten1 As Byte, ByVal daten2 As Byte, Optional ByVal factor As Byte = 255)
  Dim day As Byte
  Dim hour As Byte
  Dim minute As Byte
  day = daten1 \ 32
  hour = daten1 Mod 32
  minute = daten2 Mod 64
  'run = daten2 \ 128                       ' Seems Run is always zero.... Error in LZV200??
  RaiseEvent Modellzeit(day, hour, minute, factor)
  ' We do not log Modellzeit messages in the logfile
End Sub

' V4: 2.9 (Erweiterte Versionsinformation melden)
Public Sub eventErweiterteVersionZentrale(ZBldH As Byte, ZBldL As Byte, RMVer As Byte, RMBldH As Byte, RMBldL As Byte, BIvers As Byte)
  Dim BuildZentrale As String
  Dim VersionRuckMelder As String
  Dim BuildRuckMelder As String
  Dim VersionBootloader As String
  BuildZentrale = ZBldH * 256 + ZBldL
  VersionRuckMelder = (RMVer \ 16) & "." & (RMVer Mod 16)
  BuildRuckMelder = RMBldH * 256 + RMBldL
  VersionBootloader = (BIvers \ 16) & "." & (BIvers Mod 16)
  RaiseEvent ErweiterteZentraleVersion(BuildZentrale, VersionRuckMelder, BuildRuckMelder, VersionBootloader)
  Logfile.Details "Erweiterte Versionsinformation - BuildZentrale = " & BuildZentrale & ", VersionRuckMelder = " & VersionRuckMelder & ", BuildRuckMelder = " & BuildRuckMelder & ", VersionBootloader = " & VersionBootloader
End Sub

' Z21: 2.15 (Antwort auf LAN_X_GET_FIRMWARE_VERSION)
Public Sub eventz21FirmwareVersion(ByVal V_MSB As Byte, ByVal V_LSB As Byte)
  Dim Version As String
  Version = V_MSB & "." & V_LSB
  RaiseEvent z21FirmwareVersion(Version)
  Logfile.Details "Z21 Firmware - Version = " & Version
  End Sub

' Z21: 5.6 (ExtendedAccessoryInfo(ByVal address As Integer, ByVal Value As Byte, ByVal Valid As Boolean))
Public Sub eventExtendedAccessoryInfo(ByVal Adr_MSB As Byte, ByVal Adr_LSB As Byte, ByVal DDDDDDDD As Byte, ByVal Status As Byte)
  Dim address As Integer
  Dim rawAddress As Integer
  Dim Valid As Boolean
  rawAddress = (Adr_MSB * 256) + Adr_LSB
  If (rawAddress <= 3) Then rawAddress = rawAddress + 2048
  address = rawAddress - 3  ' The first (XpressNet) address is 4, and presented to the user as 1
  RaiseEvent ExtendedAccessoryInfo(address, DDDDDDDD, Status)
  Logfile.Details "Extended Accessory Info - Address: " & address & ", Value: " & DDDDDDDD & ", Status: " & Status
End Sub




'===============================================================================================================
' Methods for internal DLL use. Should not be called by the DLL user
' Some bit manipulation functions
Private Function bitValue(ByVal data As Byte, ByVal index As Byte) As Boolean
  index = 2 ^ index
  bitValue = (data And index) / index
End Function

Private Function bitSet(ByVal data As Byte, ByVal index As Byte) As Byte
  index = 2 ^ index
  bitSet = data Or index
End Function

Private Function bitClear(ByVal data As Byte, ByVal index As Byte) As Byte
  index = 2 ^ index              ' Set the bit at position index (other bist are 0)
  index = 255 - index            ' Flip all bits
  bitClear = data And index      ' mask the data
End Function

Private Function bitUpdate(ByVal data As Byte, ByVal index As Byte, ByVal bitIsSet As Byte) As Byte
  If bitIsSet Then
    bitUpdate = bitSet(data, index)
  Else
    bitUpdate = bitClear(data, index)
  End If
End Function

Private Sub writeXOR(ByRef byteArray() As Byte)
  ' This subroutine calculates the XOR value and writes this back as last byte of the array
  ' Should work for all possible Array sizes
  Dim xorPosition As Byte                    ' Where should the XOR value be stored in the XpressNet message array?
  xorPosition = UBound(byteArray)            ' XOR should be in the last byte of the Array (upperbound)
  Dim xorValue As Byte
  xorValue = 0
  Dim i As Integer
  For i = 2 To (xorPosition - 1)             ' i=2, since XOR calculation should ignore the Frame1 (FF) and Frame2 (FE) bytes
    xorValue = xorValue Xor byteArray(i)
  Next
  byteArray(xorPosition) = xorValue          ' byteArray is called by Reference (pointer), thus changes will be returned
End Sub

Private Function VelueOutOfRange(ByVal value As Integer, ByVal min As Integer, ByVal max As Integer, ByVal LogText As String) As Boolean
  ' support function to check is Value is between some predefined bounds
  VelueOutOfRange = False
  If value < min Then
    Logfile.AppendWithHost "Error " & LogText & " must be >= " & min
    VelueOutOfRange = True
  End If
  If value > max Then
    Logfile.AppendWithHost "Error " & LogText & " must be <= " & max
    VelueOutOfRange = True
  End If
End Function

' Functions for Accessory addresses
Private Function AccAL(ByVal address As Integer) As Byte
  ' Calculation of the low byte of a accessory address
  If Z21Protocol Then
    AccAL = address Mod 256                  ' Select the last 8-bits
  Else
    If address >= 1024 Then                  ' We already assured that the command station >= V3.8
      AccAL = (address - 1024) \ 4           ' Integer division. Result fits in an 8-bit byte
    Else
      AccAL = address \ 4                    ' Integer division
    End If
  End If
End Function

Private Function AccAH(ByVal address As Integer) As Byte
  ' Calculation of the high byte of a accessory address                                      ' Lenz uses decoder addressing
  If Z21Protocol Then
    AccAH = address \ 256                  ' Remove the last 8-bits (shift right)
  Else
    AccAH = address \ 1024                 ' Remove the last 10-bits (shift right)
  End If
End Function

' Functions for lok addresses and speed
Private Function AL(ByVal address As Integer) As Byte
  ' Calculation of the low byte of a loco address
  AL = address Mod 256
End Function

Private Function AH(ByVal address As Integer) As Byte
  ' Calculation of the high byte of a loco address
  If address < 100 Then
    AH = 0
  Else
    AH = address \ 256 + &HC0
  End If
End Function

Private Function EncodeSpeed(ByVal Speed As Integer) As Integer
  ' Calculates the speed byte for decoders with 27/28 speed steps
  Dim result As Integer
  Dim even As Boolean
  result = 0
  If Speed > 0 Then
    Speed = Speed + 3                   ' RVVV expects an offset
    result = Speed \ 2                  ' neglect the last bit, thus all uneven speed steps
    even = Speed And 1                  ' mask / take the last speed bit
    If even Then result = result + 16   ' If the speed was even, set RVVV bit 4 to 1
  End If
  EncodeSpeed = result
End Function

Public Function DecodeSpeed(ByVal Speed As Integer) As Integer
  ' Derive from the speed byte an integer value (for decoders with 27/28 speed steps)
  Dim result As Integer
  Dim Bit0_3 As Integer
  Dim Bit4 As Integer
  result = 0                            ' Default speed
  Bit4 = bitValue(Speed, 4)             ' take bit 4
  Bit0_3 = Speed Mod 16                 ' take the lowest 4 bits
  If Bit0_3 >= 2 Then
    result = (Bit0_3 - 2) * 2 + 1 + Bit4
  End If
  DecodeSpeed = result
End Function


Public Sub Class_Initialize()
   connectionType = none           ' We start without any connection
   Set send = New SendClass
   Set send.Parent = Me            ' Give the subclass a pointer to me, so it can call others
   Set receive = New ReceiveClass
   Set receive.Parent = Me
   Set receiveZ21 = New ReceiveClassZ21
   Set receiveZ21.Parent = Me
   Set AccAuto = New AccAutoClass
   Set AccAuto.Parent = Me
   ' Logfile specific initialization
   Set Logfile = New LogFileClass  ' Logfile does not call other objects, so needs no Parent
   Set Logfile.Parent = Me
   Logfile.OpenLogfile             ' Check Logfile size, and rotate if needed
   LogMessages = True              ' Set the defaults for what messages will be logged.
   LogErrors = True
   LogWarnings = True
   LogHexMessages = True
   LogDetails = True               ' TODO: MOET TZT OP FALSE WORDEN GEZET. TIJDELIJK HANDIG VOOR DEBUGGEN
   ' Attributes the DLL user may change
   Z21Protocol = False
   AccessoryInterval = 250         ' Time in msec between two accessory commands
   ' Attributes the DLL user may read
   ZentraleNotAus = False
   ZentraleNotHalt = False
   ZentraleKurzSchluss = False     ' Only for Z21
   ZentraleProgrammierMode = False
   ZentraleVersion = "0.0"         ' Give  initial version, to avoid program crash if interface is not open
   ZentraleKennung = 0
   ZentraleStartmodeAuto = False
   ZentraleKaltstart = False
   ZentraleRamCheckFehler = False
   NumberOfExpressNetMessagesReceived = 0
   NumberOfExpressNetMessagesTransmitted = 0
   ZentraleKennung = 0
End Sub

Private Sub Class_Terminate()
   Logfile.Details ("Close XpressNet interface")
   ' Before leaving, make sure we close any open connections
   If connectionType = lanTcpPort Then Call tcp.closeConnection
   ' Declare all objects as Nothing
   Set tcp = Nothing
   Set usb = Nothing
   Set send = Nothing
   Set receive = Nothing
   Set AccAuto = Nothing
   Set Logfile = Nothing
End Sub

