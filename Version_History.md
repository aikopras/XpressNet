# Version history

## Version 2.0 (November 2026)

Up to and including version 1.3 the library was only available as VB6 DLL. From version 2.0 there is also a .NET version, a macOS version and a Linux version (Qt), with the same functionality. The repository is therefore divided into the directories `VB6`, `DotNet`, `macOS` and `Linux`.

**New**
- .NET version (XpressNet.NET, C#, .NET Framework 4.7.2) for VB.NET and C# applications. The names and parameters are the same as in the VB6 DLL; the DLL does not need to be registered. The examples are also available in VB.NET (except XpressNet Regfree Test and XpressNet Project Group Test).
- macOS version (Objective-C), with four Xcode examples.
- Linux version (XpressNetQt, C++17 with Qt 6, CMake), with five examples. The library uses only Qt, and should therefore also be usable on macOS and Windows.
- After an error answer of the interface the library sends a message again, and as long as the command station does not address the interface, sending is paused.
- New events `SendError`, `InterfaceNichtAdressiert` and `InterfaceWiederAdressiert`. A message that can not be sent no longer disappears without notice.
- `CmdConnectUsb` has the optional parameter `Handshake`.
- By default the logfile is in `%PUBLIC%\TMC-Logfiles` (macOS: `~/Library/Logs/TMC-Logfiles`, Linux: `~/TMC-Logfiles`). In the logfile the description of a message also starts with `=>` (sent) or `<=` (received).
- In the logfile every line starts with the letter of the XpressNet object (A, B, …). When a connection is made, a line such as `A = 192.168.0.200, port 5550, local port 54321` tells which connection belongs to the letter (1.3: the IP address or COM port at the end of the line).
- TCP: a connection attempt stops after 5 seconds (Windows: about 21 seconds).
- Z21: after `CmdInitUdp` the library sets the broadcast flags (LAN_SET_BROADCASTFLAGS), so that the Z21 sends, among others, Alles An / Alles Aus and the turnout info. New command `CmdZ21SetBroadcastFlags`. At a new connection and when closing, the library logs off (LAN_LOGOFF).
- New example RS-Bus Monitor, in all four versions: shows the feedback information of all 128 RS-Bus addresses. The VB6 and macOS versions existed already, with their own copy of the library code; they now use the library.
- The example programs have the same names in all four versions: XpressNet Basic Test, XpressNet Test, Z21 Basic Test, XpressNet Multiple Instances, XpressNet Regfree Test and XpressNet Project Group Test (VB6 only), and RS-Bus Monitor. The directories of the examples have these names as well (1.3: 1-Basic_test, 2-Extended_test, …).
- Documentation: a [Reference manual](Documentation/Reference/Reference.md) and the [user manual](Documentation/HowToUse/Handleiding.md) in Markdown (Dutch).

**Changes**
- `CmdAccessoryAuto`: if the queue is emptied, the switch that is activated at that moment first gets an OFF command.
- Receiving: after a message with an XOR error, the library searches for the start of the next message.
- A new connection first closes an existing connection of the same object.

**Note when switching**: applications must be built again (because of the extra parameter of `CmdConnectUsb`). In .NET the attribute `ZentraleVersion` is called `ZentraleSoftwareVersion`, and `Dispose()` closes the connection. The VB6 DLL gets new GUIDs when it is compiled: VB6 applications must set the reference again.

## Version 1.3 (October 2026)

VB6 DLL only.

- Corrections in the PoM commands (CV number and bit write), the comparison of the version of the command station (as a number), the decoding of the speed (14 and 128 speed steps) and of Z21 messages, and the input buffer with much received data.
- The response timeout (5 seconds) also applies to USB and serial; for the Z21 there is no flow control anymore.
- `CmdAccessoryAuto`: after a deactivation there is a short pause before the next switch is activated.
- The DLL has new GUIDs: applications must set the reference again and must be compiled again.

## Version 1.2 (August 2025)

- The application program can set the directory for the logfiles (`LogDirectory`). See `XpressNet Basic Test` for an example.
- The FIFO for XpressNet messages was enlarged from 32 to 128 entries.
- The time between two accessory commands (`AccessoryInterval`) was changed from 250 to 125 ms.

## Version 1.1 (December 2024)

- `CmdSwitchInfoRequest` was added, to request the switch information. `CmdFeedbackRequest` is now intended for (RS-Bus) feedback decoders.
- The examples request the correct feedback data, use the default Lenz and Roco IP addresses, and position 1 is now straight.

## Version 1.0 (November 2024)

- First version.
