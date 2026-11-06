# Using the macOS library

The macOS version of the XpressNet library is written in Objective-C. Version 2.0 has the same functionality as version 2.0 of the VB6 DLL, the .NET library and the Linux version. The source code consists of two directories:
- [XpressNet](../XpressNet): the XpressNet classes (XpressNetClass, Send, Receive, ReceiveZ21, AccAuto and LogFile)
- [Communication](../Communication): the TCP, UDP and serial (USB) connections

To use the library in an Xcode project, copy both directories to the directory of the project, and add them to the project (File → Add Files to "…"). The library uses manual reference counting (no ARC), as the example projects.

The methods, events and attributes have the same names as in the VB6 DLL. The events are the methods of the XpressNetDelegate protocol; the application implements only the events it needs. The events are called on the main thread. Example:

```objc
_XpressNet = [[XpressNetClass alloc] init];
[_XpressNet setDelegate:self];
[[_XpressNet Logfile] setLogDirectory:@"~/Library/Logs/TMC-Logfiles"];
[_XpressNet CmdConnectTcp:@"192.168.0.200" Port:@"5550"];

- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data {
  ...
}
```

See XpressNet.h for all methods and events, and for the differences with the VB6 DLL. The example projects (macOS/Examples) contain a copy of both directories.

## New in version 2.0
- After an error answer of the interface the library sends a message again; as long as the command station does not address the interface, sending is paused.
- New events (XpressNetDelegate): `SendError:Description:` for every message that can not be sent, and `InterfaceNichtAdressiert` / `InterfaceWiederAdressiert`. The values of Reason are in `SendErrorReasonEnum`.
- `CmdConnectUsb:Settings:Handshake:` (hardware handshake RTS/CTS); without Handshake it is off.
- By default the logfile is in `~/Library/Logs/TMC-Logfiles`. Every line starts with the letter of the XpressNet object (A, B, …).
- Z21: after `CmdInitUdp:Port:` the library sets the broadcast flags (new command `CmdZ21SetBroadcastFlags:`).
- Messages that are offered while the TCP or UDP connection is being established, are sent as soon as the connection is ready (VB6, .NET and Linux: event SendError).

The details are in the [Reference manual](../../../Documentation/Reference/Reference.md), sections 10, 11 and 13.
