# XpressNet library – Reference manual

This reference describes all methods, events and attributes of the XpressNet library, version 2.0. It applies to the VB6 DLL, XpressNet.NET, the macOS version and the Linux version (XpressNetQt). The signatures are in VB syntax, which is the same for VB6 and VB.NET; all parameters are `ByVal`, unless stated otherwise. The few differences for .NET, macOS and Linux are in [section 13](#13-net-macos-and-linux). For the use of the library, see the user manual.

The columns V3.6, V4 and Z21 give the chapter in the specification:
- **V3.6**: Lenz, *Befehlsbeschreibung XpressNet V3.6 mit LAN/USB Interface 23151*, Doku Version 1.2 (05/2019). Chapters 1 and 2 describe the interface itself.
- **V4**: Lenz, *Dokumentation XpressNet Version 4.0* (02/2022)
- **Z21**: Roco, *Z21 LAN Protokoll Spezifikation* 1.13

The names in the column **XpressNet** are those of the specification. The column **From** gives the oldest command station version that supports the message. **Hex** gives the XpressNet message without the framing bytes and the XOR byte, in the notation of the specification. A **—** in the first column means that the message is not implemented.

## Contents
1. [Connection](#1-connection)
2. [Interface](#2-interface)
3. [Command station](#3-command-station)
4. [Accessories](#4-accessories)
5. [Feedback](#5-feedback)
6. [Locomotives](#6-locomotives)
7. [Programming (service mode)](#7-programming-service-mode)
8. [Programming on main (PoM)](#8-programming-on-main-pom)
9. [Model time](#9-model-time)
10. [Z21 and OpenDCC](#10-z21-and-opendcc)
11. [Sending and SendError](#11-sending-and-senderror)
12. [Attributes and logfile](#12-attributes-and-logfile)
13. [.NET, macOS and Linux](#13-net-macos-and-linux)

---

## 1. Connection

| Method | Notes |
|---|---|
| `CmdConnectTcp(Host As String, Optional Port As String = "5550")` | Lenz LAN interface 23151. Host: IP address or DNS name |
| `CmdInitUdp(Host As String, Optional Port As String = "21106")` | Z21. Sets `Z21Protocol` to True, sends LAN_SET_BROADCASTFLAGS with `Z21BroadcastFlagsDefault` (see [section 10](#10-z21-and-opendcc)), and requests 21 21, 21 24 and F1 0A |
| `CmdConnectUsb(Port As Byte, Optional ByRef Settings As String = "115200,N,8,1", Optional Handshake As Boolean = False)` | Handshake: hardware handshake RTS/CTS |
| `CmdConnectSerial(Port As Byte, ByRef Settings As String, Handshake As Boolean)` | LI100, LI100F, LI101F |

A new connection first closes an existing connection of the same object (Z21: with LAN_LOGOFF).

| Event | Raised when |
|---|---|
| `InterfaceConnected()` | The TCP, USB or serial connection is open. The library then requests the version (21 21) and the status (21 24) of the command station |
| `TcpClosedByRemoteHost()` | The interface has closed the TCP connection |
| `TcpError(ByRef Description As String)` | The TCP connection could not be established. .NET: a short description (for example `Connection refused`, `Connection timed out`); the complete description is in the logfile. VB6: the description of the Winsock control, or `Connection timed out`. A connection attempt stops after 5 seconds |
| `UsbConnected()` | Declared, but never raised |

| Interface | Settings | Handshake |
|---|---|---|
| LI100 | `"9600,N,8,1"` | True |
| LI100F, LI101F | `"19200,N,8,1"` | True |
| LI-USB, LAN/USB interface 23151 | `"57600,N,8,1"` | False |
| Yamorc YD7001 | `"115200,N,8,1"` or `"57600,N,8,1"` | False |
| Digikeijs DR5000 | `"115200,N,8,1"` | False |

---

## 2. Interface

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdInterfaceVersion()` | Versionsnummer des Interface 23151 feststellen | `F0` | 1.6 | | | |
| `CmdXpressNetAddress(Optional Address As Byte = 0)` | XpressNet Geräteadresse des Interface 23151 feststellen und ändern | `F2 01 ADR` | 1.7 | | | |
| `CmdInterfaceStatus()` | Interface Status Befehl | `F1 01` | 2.1 | | | |
| `CmdXpressNetVersion()` | XpressNet Version Befehl | `F1 02` | 2.2 | | | |
| `CmdInterfaceFreieVerbindungen()` | Verfügbare Freie Verbindungen | `F1 03` | 2.3 | | | |

- **Address**: 0 = request the address; 1..31 = set the address.
- **CmdInterfaceStatus**: the library sends this message every 30 seconds as keep alive (TCP). The application need not use it.

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `InterfaceVersion(Version As String, Code As String)` | Versionsnummer des Interface 23151 | `02 VV CC` | 1.6 | | | |
| `XpressNetAddress(Address As Byte)` | XpressNet Address | `F2 01 ADR` | 1.7 | | | |
| | Interface Status | `F2 01 Daten` | 2.1 | | | |
| `XpressNetVersion(Version As String)` | XpressNet Version | `F2 02 VV` | 2.2 | | | |
| `InterfaceFreieVerbindungen(Number As Byte)` | Verfügbare Freie Verbindungen | `F2 03 AA` | 2.3 | | | |
| see [section 11](#11-sending-and-senderror) | Antworten des Interface 23151 | `01 Meldung` | 1.5 | | | |

- **Version**, **Code**: as text, for example `"3.6"`.
- The answer to `CmdInterfaceStatus` has the same header as the XpressNet address (F2 01), and is also raised as `XpressNetAddress`; the answer to the keep alive message is not passed to the application.

---

## 3. Command station

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdAllesAn()` | Alles An / LAN_X_SET_TRACK_POWER_ON | `21 81` | 3.2.1 | 3.2 | 2.6 | V3 |
| `CmdAllesAus()` | Alles Aus (Notaus) / LAN_X_SET_TRACK_POWER_OFF | `21 80` | 3.2.2 | 3.3 | 2.5 | V3 |
| `CmdAlleLoksAnhalten()` | Alle Loks anhalten (Nothalt) / LAN_X_SET_STOP | `80` | 3.2.3 | 3.4 | 2.13 | V3 |
| `CmdSoftwareversionZentraleAnfordern()` | Softwareversion der Zentrale anfordern / LAN_X_GET_VERSION | `21 21` | 3.2.20 | 3.24 | 2.3 | V3 |
| `CmdErweiterteZentralenVersionsinformation()` | Erweiterte Zentralen Versionsinformation abfragen | `21 23` | | 3.25 | | V3.8 |
| `CmdStatusZentraleAnfordern()` | Status der Zentrale anfordern / LAN_X_GET_STATUS | `21 24` | 3.2.21 | 3.26 | 2.4 | V3 |
| `CmdZentralenStartmodeSetzen(AutoStart As Boolean)` | Zentralen-Startmode setzen | `22 22 00000M0` | 3.2.22 | 3.27 | | |
| `CmdZentralenResetAuslosen()` | Zentralen Reset auslösen | `21 28` | | 3.31 | | V3.8 |
| — | Quittierung | `20` | | 3.1 | | V3 |
| — | Service Variable Lesebefehl | `22 25 SV-ADR` | | 3.28 | | V3.8 |
| — | Service Variable Schreibbefehl | `22 26 SV-ADR Daten` | | 3.29 | | V3.8 |
| — | Geräteinfo senden | `24 29 SW Level Type` | | 3.32 | | V4.0 |

- **AutoStart**: True = the command station switches the track on after power on (M = 1).

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `AllesAn()` | BC Alles An / LAN_X_BC_TRACK_POWER_ON | `61 01` | 3.1.1.1 | 2.4.1 | 2.8 | V3 |
| `NotAus()` | BC Alles Aus (Notaus) / LAN_X_BC_TRACK_POWER_OFF | `61 00` | 3.1.1.2 | 2.4.2 | 2.7 | V3 |
| `NotHalt()` | BC Alle Loks Aus (Nothalt) / LAN_X_BC_STOPPED | `81 00` | 3.1.1.3 | 2.4.3 | 2.14 | V3 |
| `ProgrammierMode()` | BC Programmiermode / LAN_X_BC_PROGRAMMING_MODE | `61 02` | 3.1.1.4 | 2.4.4 | 2.9 | V3 |
| `ZentraleVersion(Version As String)` | Softwareversion / LAN_X_GET_VERSION | `63 21 Daten1 Daten2` | 3.1.3 | 2.7.2 | 2.3 | V3 |
| `NotHalt()`, `NotAus()`, `ProgrammierMode()` | Status Zentrale / LAN_X_STATUS_CHANGED | `62 22 Daten` | 3.1.4 | 2.8 | 2.12 | V3 |
| `ErweiterteZentraleVersion(BuildZentrale As String, VersionRuckMelder As String, BuildRuckMelder As String, VersionBootloader As String)` | Erw. Versionsinfo | `67 23 ZBldH ZBldL RMver RMBldH RMBldL BIvers` | | 2.9 | | V3.8 |
| `ServiceVariable(SvAddress As Byte, SvValue As Byte)` | Service Variable melden | `63 20 SV# SVval` | | 2.6 | | V3.8 |
| `Fehlermeldung(Value As Byte)` | Fehlermeldung | `E1 80+F` | 3.1.14 | 2.24 | | V3 |

- **AllesAn** sets `ZentraleNotAus`, `ZentraleNotHalt` and `ZentraleProgrammierMode` to False; **NotAus**, **NotHalt** and **ProgrammierMode** set the corresponding attribute to True.
- **ZentraleVersion** also sets the attributes `ZentraleVersion` and `ZentraleKennung`. The logfile contains the name of the command station (see `ZentraleKennung` in [section 12](#12-attributes-and-logfile)).
- **Status Zentrale** sets the attributes `ZentraleNotHalt`, `ZentraleNotAus`, `ZentraleStartmodeAuto`, `ZentraleProgrammierMode`, `ZentraleKaltstart` and `ZentraleRamCheckFehler`, and raises `NotHalt`, `NotAus` and `ProgrammierMode` for the bits that are set. Z21: bit 2 is `ZentraleKurzSchluss`, bit 5 `ZentraleProgrammierMode`.
- **Fehlermeldung**: Value is F (4 bits).

---

## 4. Accessories

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdAccessory(Address As Integer, Position As Byte, Activate As Byte)` | Schaltbefehl | `52 ADR DBBD` | 3.2.24 | 3.39 | | V3 |
| | Schaltbefehl ab Version 3.8 / LAN_X_SET_TURNOUT | `53 ADRH ADRL DBBD` | | 3.40 | 5.2 | V3.8 |
| `CmdAccessoryAuto(Switch As Integer, Position As Byte)` | Schaltbefehl: activate, deactivate after `AccessoryInterval` | as CmdAccessory | | | | |
| `CmdSwitchInfoRequest(Address As Integer)` | Schaltinformation anfordern | `42 ADR Nibble` | 3.2.23 | 3.37 | | V3 |
| | Schaltinformation anfordern, Adresse bis 2048 / LAN_X_GET_TURNOUT_INFO | `43 ADRH ADRL Nibble` | | 3.38 | 5.1 | V3.8 |
| `CmdFeedbackRequest(Address As Integer, Nibble As Byte)` | Schaltinformation anfordern (RS-Bus) | `42 ADR Nibble` | 3.2.23 | 3.37 | | V3 |

- **Address** (switch): 1..1024; 1..2048 if `ZentraleVersion` ≥ 3.8. `CmdSwitchInfoRequest` with the Z21: 1..2048. With the Z21 the limit follows the X-Bus version that the Z21 reports (firmware 1.43: 4.0, so 1..2048).
- **Address** (`CmdFeedbackRequest`): RS-Bus address 1..128.
- **Position**: 0 = diverging (red, −), 1 = straight (green, +), according to RCN-213. Some command stations interpret 0 and 1 the other way around. Values > 1 count as 1.
- **Activate**: 0 = deactivate, > 0 = activate.
- **Nibble**: 0 = inputs 1-4, 1 = inputs 5-8; values > 1 count as 1.
- **CmdAccessoryAuto**: queue of 32 commands. The output is activated for `AccessoryInterval` ms (default 125), followed by a pause of `AccessoryInterval` / 2 before the next switch. If the queue is emptied ([section 11](#11-sending-and-senderror)), the switch that is activated at that moment is deactivated first.
- With the Z21 the address is a switch address; with Lenz a decoder address with two address bits in DBBD. The library converts this.

---

## 5. Feedback

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `Feedback(Address As Integer, Nibble As Byte, Data As Byte)` | BC Rückmeldung | `40+N Adr Dat ...` | 3.1.1.5 | 2.4.5 | | V3 |
| | Schaltinformation | `42 Adr Dat` | 3.1.8 | 2.15 | | V3 |
| `SwitchFeedback(Address As Integer, ttBits As Byte, iBit As Byte, Data As Byte)` | BC Rückmeldung, Schaltinformation | `40+N ...`, `42 Adr Dat` | 3.1.1.5, 3.1.8 | 2.4.5, 2.15 | | V3 |
| | BC Rückmeldung ab W1025, Schaltinformation ab W1025 / LAN_X_TURNOUT_INFO | `43 AdrH AdrL Dat` | | 2.4.6, 2.16 | 5.3 | V3.8 |
| `ExtendedAccessoryInfo(...)` | LAN_X_EXT_ACCESSORY_INFO | see [section 10](#10-z21-and-opendcc) | | | 5.6 | |

- The TT bits decide the event: a feedback decoder (TT = 2) gives `Feedback`, a switch decoder gives `SwitchFeedback`. One message may contain several feedbacks; each gives its own event.
- **Feedback**: Address 1..128; Nibble 0 = inputs 1-4, 1 = inputs 5-8; Data: the 4 inputs (bits 0-3).
- **SwitchFeedback**: Address is the first switch of a pair (the second is Address + 1). ttBits: 0 = switch decoder without feedback, 1 = with feedback. iBit = 1: the switch has not yet reached its end position. Data: bits 0-1 first switch, bits 2-3 second switch. Z21: Address is the switch itself.

---

## 6. Locomotives

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdLokInfoAnfordern(Address As Integer)` | Lokinformationen anfordern ab Zentralen-Version 3.0 / LAN_X_GET_LOCO_INFO | `E3 00 AH AL` (Z21: `E3 F0 AH AL`) | 3.2.25 | 3.41.3 | 4.1 | V3 |
| `CmdFunktionsStatusAnfordern(Address As Integer)` | Funktionsstatus anfordern ab Zentralen-Version 3.0 | `E3 07 AH AL` | 3.2.25.1 | 3.41.4 | | V3 |
| `CmdFunktionsStatusF13F28Anfordern(Address As Integer)` | Funktionsstatus anfordern F13 – F28 | `E3 08 AH AL` | 3.2.25.2 | 3.41.5 | | V3.6 |
| `CmdFunktionsStatusF29F68Anfordern(Address As Integer)` | Funktionsstatus anfordern F29 – F68 | `E3 0A AH AL` | | 3.41.6 | | V4.0 |
| `CmdFunktionsZustandF13F28Anfordern(Address As Integer)` | Funktionszustand anfordern F13 – F28 | `E3 09 AH AL` | 3.2.25.3 | 3.41.7 | | V3.6 |
| `CmdFunktionsZustandF29F68Anfordern(Address As Integer)` | Funktionszustand anfordern F29 – F68 | `E3 0B AH AL` | | 3.41.8 | | V4.0 |
| `CmdFahrbefehl(Address As Integer, Speed As Integer, Steps As Integer, Forward As Boolean)` | Lok Fahrbefehl ab Zentralen-Version 3.0 / LAN_X_SET_LOCO_DRIVE | `E4 10..13 AH AL RV` | 3.2.26.1 | 3.42.3 | 4.2 | V3 |
| `CmdLokFunktionBefehl(Address As Integer, Group As Byte, Data As Byte)` | Lok Funktionsbefehl ab Zentralen-Version 3.0 / LAN_X_SET_LOCO_FUNCTION_GROUP | `E4 20..2B, 50, 51 AH AL Daten` | 3.2.26.2 | 3.42.4 | 4.3.2 | V3-V4 |
| `CmdLokFunktionStatus(Address As Integer, Group As Byte, Data As Byte)` | Funktionsstatus setzen | `E4 24..27, 2C..2E, 52..54 AH AL Daten` | 3.2.26.4 | 3.42.6 | | V3-V4 |
| `CmdEineLokAnhalten(Address As Integer)` | Eine Lok anhalten (Nothalt für eine Lok) / LAN_X_SET_LOCO_E_STOP | `92 AH AL` | 3.2.4 | 3.7 | 4.5 | V3 |
| `CmdAdresAnfrageLokInStack(Address As Integer, Vorwarts As Boolean)` | Adressanfrage Lok in Zentralenstack | `E3 05+R AH AL` | 3.2.30.3 | 3.44.3 | | V3 |
| `CmdLokAusStack(Address As Integer)` | Lok aus Stack löschen ab V3 / LAN_X_PURGE_LOCO | `E3 44 AH AL` | 3.2.31 | 3.44.4 | 4.6 | V3 |
| — | Funktionsbefehl „binary states“ - Short Form | `E4 5E AH AL FDDD` | | 3.42.5.1 | | V4 |
| — | Funktionsbefehl „binary states“ - Long Form / LAN_X_SET_LOCO_BINARY_STATE | `E5 5F AH AL FLLL HHH` | | 3.42.5.2 | 4.3.3 | V4 |
| — | Funktionsrefresh-Modus setzen | `E4 2F AH AL RF` | 3.2.26.5 | 3.42.7 | | V3.6 |
| — | Doppeltraktion montieren | `E5 43 AH1 AL1 AH2 AL2` | 3.2.27.1 | 3.42.9.1 | | V3 |
| — | Doppeltraktion auflösen | `E5 43 AH1 AL1 00 00` | 3.2.27.2 | 3.42.9.2 | | V3 |
| — | Lok zu einer Mehrfachtraktion hinzufügen | `E4 40+R AH AL MTR` | 3.2.29.1 | 3.42.10.1 | | V3 |
| — | Lok aus einer Mehrfachtraktion entfernen | `E4 42 AH AL MTR` | 3.2.29.2 | 3.42.10.2 | | V3 |
| — | Adressanfrage Mitglied einer Mehrfachtraktion | `E4 01+R MTR AH AL` | 3.2.30.1 | 3.44.1 | | V3 |
| — | Adressanfrage Mehrfachtraktion | `E2 03+R MTR` | 3.2.30.2 | 3.44.2 | | V3 |

- **Address**: 0..9999. From 100 the long address is used (AH = C0 + Address / 256).
- **Steps** and **Speed**: 14 (Speed 0..14), 27 (0..27), 28 (0..28) or 128 (0..126). Other values: the message is not sent (SendError 6).
- **Forward**: True = forward.
- **Group** and **Data**: Data contains the function bits of the group. Group 6..10: V4 only.

| Group | Functions | Bits in Data | Befehl | Status |
|---|---|---|---|---|
| 1 | F0-F4 | F0 = bit 4, F1..F4 = bits 0-3 | `20` | `24` |
| 2 | F5-F8 | bits 0-3 | `21` | `25` |
| 3 | F9-F12 | bits 0-3 | `22` | `26` |
| 4 | F13-F20 | bits 0-7 | `23` | `27` |
| 5 | F21-F28 | bits 0-7 | `28` | `2C` |
| 6 | F29-F36 | bits 0-7 | `29` | `2D` |
| 7 | F37-F44 | bits 0-7 | `2A` | `2E` |
| 8 | F45-F52 | bits 0-7 | `2B` | `52` |
| 9 | F53-F60 | bits 0-7 | `50` | `53` |
| 10 | F61-F68 | bits 0-7 | `51` | `54` |

- **Vorwarts**: True = search forward (05), False = backward (06).
- `DecodeSpeed(Speed As Integer) As Integer`: converts the speed byte (27 and 28 speed steps) into a speed step.

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `NormaleLokinfo(Besetzt As Boolean, Speedsteps As Integer, Speed As Integer, Forward As Boolean, F0_F4 As Byte, F5_F8 As Byte, F9_F12 As Byte)` | Normale Lokinfo | `E4 Kennung Speed FKT0 FKT1` | 3.1.9.1 | 2.19.1 | | V3 |
| `FunctionsZustandF13F20(F13_F20 As Byte, F21_F28 As Byte)` | F-Zustand F13 … F28 | `E3 52 F13-F20 F21-F28` | 3.1.9.2 | 2.19.2 | | V3.6 |
| `FunctionsZustandF29F68(F29_F36 As Byte, F37_F44 As Byte, F45_F52 As Byte, F53_F60 As Byte, F61_F68 As Byte)` | F-Zustand F29 … F68 | `E6 53 F29-F36 … F61-F68` | | 2.19.3 | | V4.0 |
| `FunctionStatusF0F12(F0_F4 As Byte, F5_F8 As Byte, F9_F12 As Byte)` | Funktionsstatus | `E3 50 Status0 Status1` | 3.1.11 | 2.19.8 | | V3 |
| `FunctionStatusF13F20(F13_F20 As Byte, F21_F28 As Byte, RefreshMode As Byte)` | F-status F13 bis F28 | `E4 51 Status2 Status3 RStatus` | 3.1.12 | 2.20 | | V3.6 |
| `FunctionStatusF29F68(F29_F36 As Byte, F37_F44 As Byte, F45_F52 As Byte, F53_F60 As Byte, F61_F68 As Byte)` | F-status F29 bis F68 | `E6 54 Status6 … Status10` | | 2.21 | | V4 |
| `Z21Lokinfo(...)` | LAN_X_LOCO_INFO | see [section 10](#10-z21-and-opendcc) | | | 4.4 | |
| — | MTR-Mitglied | `E5 Kennung Speed FKT0 FKT1 MTR` | 3.1.9.3 | 2.19.4 | | V3 |
| — | MTR-Basisadresse | `E2 Kennung Speed` | 3.1.9.4 | 2.19.5 | | V3 |
| — | Lok ist in DTR | `E6 Kennung Speed FKT0 FKT1 AdrHigh AdrLow` | 3.1.9.5 | 2.19.6 | | V3 |
| — | Lok besetzt | `E3 40 AdrHigh AdrLow` | 3.1.10 | 2.19.7 | | V3 |
| — | Lok-Suchergebnis | `E3 30+K AdrHigh AdrLow` | 3.1.13 | 2.22 | | V3 |

- **Besetzt**: True = the locomotive is controlled by another device.
- **Speedsteps**: 14, 27, 28 or 128. **Speed**: the speed step; 0 = stop (an emergency stop is also returned as 0).
- **F0_F4**: F0 = bit 4, F1..F4 = bits 0-3. The other parameters: one bit per function, the lowest function in bit 0.

---

## 7. Programming (service mode)

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdLeseanfrageProgrammierenRegistermode(Register As Byte)` | Leseanfrage Programmieren 3-Byte-Format (Registermode) / LAN_X_DCC_READ_REGISTER | `22 11 REG` | 3.2.5 | 3.9 | 6.13 | V3 |
| `CmdLeseanfrageProgrammieren(CV As Integer)` | Leseanfrage Programmieren 4-Byte-Format (CV-Mode, CV 1-256) | `22 15 CV` | 3.2.6 | 3.10 | | V3 |
| | Leseanfrage Programmieren 4-Byte-Format (CV 1-255 und CV1024, CV 256-511, CV 512-767, CV 768-1023) | `22 18..1B CV` | 3.2.7 - 3.2.10 | 3.11 - 3.14 | | V3.6 |
| | LAN_X_CV_READ | `23 11 CVAdr_MSB CVAdr_LSB` | | | 6.1 | |
| `CmdLeseanfrageProgrammierenPagemode(CV As Integer)` | Leseanfrage Programmieren 3-Byte-Format (Pagemode) | `22 14 CV` | 3.2.11 | 3.15 | | V3 |
| `CmdProgrammierergebnisAnfordern()` | Programmierergebnis anfordern | `21 10` | 3.2.12 | 3.16 | | V3 |
| `CmdSchreibbefehlProgrammierenRegisterMode(Register As Byte, Daten As Integer)` | Schreibbefehl Programmieren 3-Byte-Format (Register-Mode) / LAN_X_DCC_WRITE_REGISTER | `23 12 REG Daten` | 3.2.13 | 3.17 | 6.14 | V3 |
| `CmdSchreibbefehlProgrammieren(CV As Integer, Val As Byte)` | Schreibbefehl Programmieren 4-Byte-Format (CV-Mode, CV 1-256) | `23 16 CV Daten` | 3.2.14 | 3.18 | | V3 |
| | Schreibbefehl Programmieren 4-Byte-Format (CV 1-255 und CV1024, CV 256-511, CV 512-767, CV 768-1023) | `23 1C..1F CV Daten` | 3.2.15 - 3.2.18 | 3.19 - 3.22 | | V3.6 |
| | LAN_X_CV_WRITE | `24 12 CVAdr_MSB CVAdr_LSB Value` | | | 6.2 | |
| `CmdSchreibbefehlProgrammierenPageMode(CV As Integer, Daten As Integer)` | Schreibbefehl Programmieren 3-Byte-Format (Page-Mode) | `23 17 CV Daten` | 3.2.19 | 3.23 | | V3 |

- **Register**: 1..8. **CV**: 1..1024; Pagemode: 1..256. **Val**, **Daten**: 0..255.
- The library selects the message by `ZentraleVersion` (< 3.6: `22 15` / `23 16`) and by `Z21Protocol` (LAN_X_CV_READ / LAN_X_CV_WRITE).
- In programming mode the answer may take up to 90 seconds (V3.6: 1.3).

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `ProgrammierinfoKurzschluss()` | P-Info Kurzschluß / LAN_X_CV_NACK_SC | `61 12` | 3.1.2.1 | 2.5.1 | 6.3 | V3 |
| | LAN_X_BC_TRACK_SHORT_CIRCUIT | `61 08` | | | 2.10 | |
| `ProgrammierinfoDatenNichtGefunden()` | P-Info Keine Daten / LAN_X_CV_NACK | `61 13` | 3.1.2.2 | 2.5.2 | 6.4 | V3 |
| `ProgrammierinfoZentraleBusy()` | P-Info Busy | `61 1F` | 3.1.2.3 | 2.5.3 | | V3 |
| `ProgrammierinfoZentraleBereit()` | P-Info bereit | `61 11` | 3.1.2.4 | 2.5.4 | | V3 |
| `ProgrammierinfoPageRegister(EEPromAddress As Byte, Value As Byte)` | P-Info Daten 3 Byte | `63 10 EE Daten` | 3.1.2.5 | 2.5.5 | | V3 |
| `ProgrammierinfoCV(CV As Integer, Value As Byte)` | P-Info CV1-255 u. 1024, CV256 - 511, CV512 - 767, CV768 - 1023 | `63 14..17 CV Daten` | 3.1.2.6 - 3.1.2.9 | 2.5.6 - 2.5.9 | | V3.6 |
| | LAN_X_CV_RESULT | `64 14 CVAdr_MSB CVAdr_LSB Value` | | | 6.5 | |

- **CV**: 1..1024. CV 0 in the message is returned as 1024 (`ZentraleVersion` < 3.6: as 256).

---

## 8. Programming on main (PoM)

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdPomWrite(Address As Integer, CV As Integer, Value As Byte)` | Programming on Main Byte schreiben / LAN_X_CV_POM_WRITE_BYTE | `E6 30 AH AL EC+C CV Daten` | 3.2.28.1 | 3.43.1 | 6.6 | V3 |
| `CmdPomRead(Address As Integer, CV As Integer)` | Programming on Main Byte lesen / LAN_X_CV_POM_READ_BYTE | `E6 30 AH AL E4+C CV 00` | 3.2.28.2 | 3.43.2 | 6.8 | V3.6 |
| `CmdPomBitWrite(Address As Integer, CV As Integer, bitValue As Boolean, bitPosition As Byte)` | Programming on Main Bit schreiben / LAN_X_CV_POM_WRITE_BIT | `E6 30 AH AL E8+C CV 1111VBBB` | 3.2.28.3 | 3.43.3 | 6.7 | V3 |
| `CmdPoMErgebnisAnfordern()` | PoM-Ergebnis anfordern | `21 27` | | 3.30 | | V3.8 |

- **Address**: 0..9999. **CV**: 1..1024 (C contains bits 8-9 of CV - 1). **bitPosition**: 0..7.
- `CmdPomBitWrite` uses E8, as in V3.6. V4 (3.43.3) gives 7C, which seems to be an error in the V4 specification.

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `PomErgebnis(LokAddress As Integer, CvValue As Byte)` | PoM Ergebnis melden | `64 24 AdrH AdrL PomValue` | | 2.10 | | V3.8 |

---

## 9. Model time

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdModellzeitAnfordern()` | Modellzeit anfordern | `21 2A` | | 3.33 | | V3.8 |
| `CmdModellzeitStellen(Day As Byte, Hour As Byte, Minute As Byte, Factor As Byte)` | Modellzeit stellen | `24 2B DOW&h Min Factor` | | 3.34 | | V3.8 |
| `CmdModellzeitAnhalten()` | Modellzeit anhalten | `21 2D` | | 3.35 | | V4.0 |
| `CmdModellzeitStarten()` | Modellzeit starten | `21 2C` | | 3.36 | | V4.0 |

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `Modellzeit(Day As Byte, Hour As Byte, Minute As Byte, Factor As Byte)` | BC Modellzeit | `63 03 dddhhhhh Min&stop` | | 2.4.7 | | V3.8 |
| | Modellzeit melden | `64 25 H&DOW min Faktor` | | 2.11 | | V3.8 |

- **Day**: 0 = Monday … 6 = Sunday. **Hour**: 0..23. **Minute**: 0..59. **Factor**: 0..31 (0 = stopped). With the broadcast (63 03) Factor is 255: the message does not contain the factor.
- The library does not write model time messages to the logfile.

---

## 10. Z21 and OpenDCC

The Z21 sends broadcasts (such as track power on/off and the turnout info) only to a client that has set the broadcast flags (Z21: 2.16). The flags must be set after every login; the login is done with the first message. `CmdInitUdp` therefore first sends LAN_SET_BROADCASTFLAGS with `Z21BroadcastFlagsDefault`. A new connection and `Dispose` (VB6: `Set XpressNet = Nothing`) first send LAN_LOGOFF (Z21: 2.2). Every received Z21 message is logged (`LogHexMessages`), also the messages that the library does not decode.

After `CmdInitUdp` (`Z21Protocol` = True) the methods of the previous sections send the Z21 variant where it differs: `CmdLokInfoAnfordern` (`E3 F0`), `CmdLeseanfrageProgrammieren` (LAN_X_CV_READ), `CmdSchreibbefehlProgrammieren` (LAN_X_CV_WRITE), `CmdAccessory` (`53`, switch address) and `CmdSwitchInfoRequest` (`43 AH AL`).

| Method | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `CmdZ21SetBroadcastFlags(Flags As Integer)` (VB6: `Flags As Long`) | LAN_SET_BROADCASTFLAGS | Z21: `08 00 50 00 Flags` (not XpressNet) | | | 2.16 | |
| `CmdZ21GetFirmwareVersion()` | LAN_X_GET_FIRMWARE_VERSION | `F1 0A` | | | 2.15 | |
| `CmdZ21SetLocoFunction(Address As Integer, CV As Byte, TT As Byte)` | LAN_X_SET_LOCO_FUNCTION | `E4 F8 AH AL TTNNNNNN` | | | 4.3.1 | |
| `CmdZ21SetExtAccessory(Address As Integer, Value As Byte)` | LAN_X_SET_EXT_ACCESSORY | `54 AH AL Daten 00` | | | 5.4 | |
| `CmdZ21GetExtAccessory(Address As Integer)` | LAN_X_GET_EXT_ACCESSORY_INFO | `44 AH AL 00` | | | 5.5 | |
| `CmdZ21AccPomWrite(decoderAddress As Integer, CDDD As Byte, CV As Integer, Value As Byte)` | LAN_X_CV_POM_ACCESSORY_WRITE_BYTE | `E6 31 aaaaa AAAACDDD EC+C CV Daten` | | | 6.9 | |
| `CmdZ21AccPomBitWrite(decoderAddress As Integer, CDDD As Byte, CV As Integer, bitValue As Boolean, bitPosition As Byte)` | LAN_X_CV_POM_ACCESSORY_WRITE_BIT | `E6 31 aaaaa AAAACDDD E8+C CV 0000VPPP` | | | 6.10 | |
| `CmdZ21AccPomRead(decoderAddress As Integer, CDDD As Byte, CV As Integer)` | LAN_X_CV_POM_ACCESSORY_READ_BYTE | `E6 31 aaaaa AAAACDDD E4+C CV 00` | | | 6.11 | |
| `CmdZ21MmWriteByte(RegAdr As Byte, Value As Byte)` | LAN_X_MM_WRITE_BYTE | `24 FF 00 REG Daten` | | | 6.12 | |
| `CmdOpenDCCSetExtAccessory(Address As Integer, Value As Byte)` | OpenDCC: extended accessory | `13 01 B+AddrH AddrL` | | | | |

- **CmdZ21SetBroadcastFlags**: Flags as in the Z21 specification (32 bits). The constant `Z21BroadcastFlagsDefault` = `&H00000001`: driving and switching (track power, programming mode, short circuit, stop, LAN_X_TURNOUT_INFO, and LAN_X_LOCO_INFO of the locomotives requested by this client). The library does not decode the messages of the other flags, such as R-Bus (`&H00000002`) and LocoNet (`&H01000000` and up). Without a Z21 connection: SendError 5 (KeineVerbindung). Version 2.0.
- **CmdZ21SetLocoFunction**: CV = the function, 0..31 (0 = F0). TT: 0 = off, 1 = on, 2 = toggle.
- **Ext. accessory** (Z21): Address 1..2048; the library sends Address + 3 (the first user address is 1, in the message 4).
- **Accessory PoM**: decoderAddress 0..511; CDDD 0..15 (0000: the whole decoder; C = 1: DD is the output); CV 1..1024.
- **OpenDCC**: Address 1..2048, Value 0..31. See https://www.opendcc.de/elektronik/opendcc/xpressnet_commands_e.html

| Event | XpressNet | Hex | V3.6 | V4 | Z21 | From |
|---|---|---|---|---|---|---|
| `z21FirmwareVersion(Version As String)` | Antwort auf LAN_X_GET_FIRMWARE_VERSION | `F3 0A V_MSB V_LSB` | | | 2.15 | |
| `Z21Lokinfo(Address As Integer, Besetzt As Boolean, Speedsteps As Integer, Speed As Integer, Forward As Boolean, F0_F4 As Byte, F5_F8 As Byte, F9_F12 As Byte, F13_F20 As Byte, F21_F28 As Byte)` | LAN_X_LOCO_INFO | `EF Lokinfo` | | | 4.4 | |
| `ExtendedAccessoryInfo(Address As Integer, Value As Byte, Status As Byte)` | LAN_X_EXT_ACCESSORY_INFO | `44 FAdrMSB FAdrLSB DDDDDDDD Status` | | | 5.6 | |

- **z21FirmwareVersion**: BCD coded; `01 23` gives `"1.23"`.
- **Z21Lokinfo**: the parameters as in `NormaleLokinfo`, plus Address, F13_F20 and F21_F28.

---

## 11. Sending and SendError

The interface answers every message; the next message is sent after the answer (V3.6: 1.3). In the meantime messages wait in the FIFO (128 messages). The Z21 has no flow control: messages are sent immediately, and there is no retransmission and no pause.

| Situation | Behaviour |
|---|---|
| Answer 01 04: Befehl ist an Zentrale geschickt | The next message is sent |
| Answer 61 80, 61 81, 01 01, 01 02, 01 03, 01 06, 01 0A | Sent again: at most 3 times, after 100, 200 and 400 ms. Then given up (SendError 1), and the next message is sent |
| Answer 01 09 / 61 82 | Not sent again (SendError 6 / 2); the next message is sent |
| No answer within the timeout: 5 s, 90 s in programming mode | Both queues emptied; the activated switch is switched off (SendError 3). TCP: keep alive message |
| 01 05 or 01 08 | Pause (`InterfaceNichtAdressiert`). The FIFO is kept; CmdAccessoryAuto activates no new switch; messages for the interface itself (F0, F1, F2) are still sent. With 01 08 the last message is sent again after the pause |
| 01 07 | End of the pause (`InterfaceWiederAdressiert`): first the message of 01 08, then the FIFO. 01 07 without a pause is only logged |
| No 01 07 within 5 s (90 s in programming mode) | Both queues emptied; the activated switch is switched off (SendError 7) |
| Every 30 s without messages (TCP / Z21) | Keep alive message: `F1 01` (TCP), `21 24` (Z21) |

| Answer | Antworten des Interface 23151 (V3.6: 1.5) and Zentrale (V3.6: 3.1.5 - 3.1.7) |
|---|---|
| 01 01 | Die im Header angegebene Anzahl Bytes stimmt nicht mit der Anzahl der empfangenen Bytes überein |
| 01 02 | Fehler zwischen Interface und Zentrale (Timeout bei Datenübertragung des Interface an die Zentrale) |
| 01 03 | unbekannter Fehler (Zentrale adressierte das Interface mit Quittierungsaufforderung) |
| 01 04 | Befehl ist an Zentrale geschickt |
| 01 05 | Zentrale adressiert das Interface nicht mehr |
| 01 06 | Puffer-Überlauf im Interface 23151 |
| 01 07 | Zentrale adressiert das Interface wieder |
| 01 08 | Derzeit können keine Befehle an die Zentrale gesendet werden |
| 01 09 | Fehler in den Befehlsparametern |
| 01 0A | unbekannter Fehler (Zentrale lieferte nicht die erwartete Antwort) |
| 61 80 | Übertragungsfehler |
| 61 81 | Zentrale Busy |
| 61 82 | Befehl in Zentrale nicht vorhanden |

| Event | Raised when |
|---|---|
| `SendError(Reason As Byte, Description As String)` | A message is not sent. Description contains the bytes of the message, with the XOR byte |
| `InterfaceNichtAdressiert()` | 01 05 or 01 08: sending is paused |
| `InterfaceWiederAdressiert()` | 01 07 after a pause: sending continues |

| Reason | SendErrorReasonEnum | Cause |
|---|---|---|
| 1 | WiederholungErfolglos | Given up after 3 retransmissions |
| 2 | BefehlNichtVorhanden | Answer 61 82 (V3.6: 3.1.7) |
| 3 | Timeout | No answer within the timeout (V3.6: 1.3) |
| 4 | FifoVoll | FIFO (128) or CmdAccessoryAuto queue (32) full |
| 5 | KeineVerbindung | No connection, or the connection has not (yet) been established |
| 6 | FehlerInDenBefehlsparametern | Answer 01 09 (V3.6: 1.5), or a parameter out of range (checked by the library) |
| 7 | ZentraleAdressiertDasInterfaceNichtMehr | No 01 07 after 01 05 / 01 08 (V3.6: 1.5) |

Every case is also written to the logfile.

---

## 12. Attributes and logfile

| Attribute | Type | Default | Meaning |
|---|---|---|---|
| `Z21Protocol` | Boolean | False | The library uses the Z21 protocol. Set by `CmdInitUdp` |
| `AccessoryInterval` | Long | 125 | Activation time (ms) of CmdAccessoryAuto |
| `LogMessages` | Boolean | True | Not used by the library |
| `LogErrors` | Boolean | True | Log errors |
| `LogWarnings` | Boolean | True | Log warnings |
| `LogHexMessages` | Boolean | True | Log the bytes of every message |
| `LogDetails` | Boolean | True | Log a description of every message |
| `ZentraleVersion` | String | `"0.0"` | Software version of the command station (V3.6: 3.1.3). .NET: `ZentraleSoftwareVersion` |
| `ZentraleKennung` | Byte | 0 | The command station (V3.6: 3.1.3): 0x00 LZ 100, 0x01 LH 200, 0x02 DPC (Compact / Commander), 0x03 Control Plus, 0x04 LokMaus II, 0x07 LDH, 0x10 multiMAUS, 0x12 Z21 / Yamorc, 0x20 OpenDCC, 0x22 OpenDccXp |
| `ZentraleNotAus` | Boolean | False | Alles Aus (V3.6: 3.1.1.2, 3.1.4) |
| `ZentraleNotHalt` | Boolean | False | Alle Loks Aus (V3.6: 3.1.1.3, 3.1.4) |
| `ZentraleProgrammierMode` | Boolean | False | Programming mode (V3.6: 3.1.1.4, 3.1.4) |
| `ZentraleStartmodeAuto` | Boolean | False | Automatic start (V3.6: 3.1.4) |
| `ZentraleKaltstart` | Boolean | False | Cold start (V3.6: 3.1.4) |
| `ZentraleRamCheckFehler` | Boolean | False | RAM check error (V3.6: 3.1.4) |
| `ZentraleKurzSchluss` | Boolean | False | Short circuit (Z21 only, Z21: 2.12) |
| `NumberOfExpressNetMessagesReceived` | Long | 0 | Number of received XpressNet messages |
| `NumberOfExpressNetMessagesTransmitted` | Long | 0 | Number of transmitted XpressNet messages (including retransmissions) |
| `connectionType` | connectionTypeEnum | none | `none`, `lanTcpPort`, `lanUdpPort`, `usbPort`, `serialPort` |
| `Logfile` | LogFileClass | | The logfile (see below) |

**Logfile** (`XpressNet.Logfile`):

| Member | Meaning |
|---|---|
| `LogDirectory` (String) | Directory of the logfile. Default `%PUBLIC%\TMC-Logfiles` (macOS: `~/Library/Logs/TMC-Logfiles`, Linux: `~/TMC-Logfiles`); the library creates the directory |
| `Append(Text As String)` | Writes a line with date and time |
| `AppendWithHost(Text As String)` | As Append, preceded by the letter of the XpressNet object (see below) |
| `Details(Text)`, `Errors(Text)`, `Warnings(Text)` | As AppendWithHost, only if `LogDetails`, `LogErrors` or `LogWarnings` is True |

In the logfile, the bytes and the description of a sent message start with `=>`, those of a received message with `<=`. The description of a command is written when the command is given; with flow control (TCP, USB) the bytes follow when the command is sent.

- The name of the logfile is `TMC-LogFile.txt`. If the logfile is larger than 100000 bytes at the start, it is renamed to `TMC-LogFile-1.txt`; at most three old logfiles are kept.
- By default all XpressNet objects of an application write to the same logfile. Every line of an XpressNet object starts with its letter: A for the first object of the application, B for the second, and so on. When a connection is made, a line such as `A = 192.168.0.200, port 5550, local port 54321` (USB / serial: `A = COM3, 57600,N,8,1`) tells which connection belongs to the letter. Version 2.0 (1.3: the IP address or COM port at the end of the line).

---

## 13. .NET, macOS and Linux

**.NET** (XpressNet.NET): the same names and parameters, except:
- `ZentraleSoftwareVersion` instead of `ZentraleVersion` (in .NET an event and an attribute can not have the same name; the event `ZentraleVersion` keeps its name).
- `Dispose()` closes the connection and stops the timers (VB6: `Set XpressNet = Nothing`).
- `Settings` of `CmdConnectUsb` and `CmdConnectSerial` is ByVal.
- Events are raised on the thread that created the object. Calls from another thread are passed to that thread.
- C#: Integer and Long = `int`, Byte = `byte`, Boolean = `bool`, String = `string`; `TcpError(ref string Description)`.

**macOS**: the same names, in Objective-C syntax. Events are methods of the `XpressNetDelegate` protocol, with the same names. Parameters are `int` or `NSString *`; the USB port is the name of the device. Messages that are offered while the TCP or UDP connection is being established, are sent as soon as the connection is ready (VB6, .NET and Linux: SendError 5).
```objc
[xpressNet CmdAccessory:12 position:1 Activate:1];
- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data;   // XpressNetDelegate
```

**Linux** (XpressNetQt, C++ with Qt 6): the same names and parameters as in .NET (`ZentraleSoftwareVersion`). Integer and Long = `int`, Byte = `quint8`, Boolean = `bool`, String = `QString`; an Optional parameter has a default value. Events are Qt signals with the same names, emitted in the thread that created the object. `delete` closes the connection and stops the timers. The USB port is the name of the device (for example `/dev/ttyUSB0`); if the port can not be opened, `TcpError` is emitted. The logfile directory is set with `Logfile->setLogDirectory(...)`.
```cpp
xpressNet->CmdAccessory(12, 1, 1);
connect(xpressNet, &XpressNetClass::Feedback, this, &MainWindow::Feedback);   // void Feedback(int address, quint8 Nibble, quint8 data)
```
