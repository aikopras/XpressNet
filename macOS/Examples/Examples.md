# Examples (macOS)
This directory contains four Xcode projects (Objective-C) that use the macOS version of the XpressNet library. They are macOS versions of the VB6 examples. Each project contains a copy of the library (the directories XpressNet and Communication).

The compiled programs are in [Examples-Compiled/macOS_Examples.zip](../Examples-Compiled/macOS_Examples.zip) (universal: Apple silicon and Intel, macOS 12.4 or newer). The program names are the same in all versions of the library (VB6, .NET, macOS and Linux). The programs are not notarized by Apple. The first time, macOS refuses to open a program: then choose System Settings → Privacy & Security → Open Anyway (up to macOS 14: right click on the program → Open). After that it starts normally.

## XpressNet Test ##
macOS version of the VB6 example XpressNet Test: a large number of XpressNet commands, and the responses of the command station. Messages of the library (SendError, interface nicht / wieder adressiert) are shown in the status line.

## XpressNet Basic Test ##
macOS version of the VB6 example XpressNet Basic Test: connecting via TCP/IP or USB (with the option Handshake), version information, switching accessories, switch and RS-Bus feedback, and locomotives and their functions. The logfiles are written to ~/Library/Logs/TMC-Logfiles. Messages of the library (SendError, interface nicht / wieder adressiert) are shown in the status line.

## Z21 Basic Test ##
macOS version of the VB6 example Z21 Basic Test: a limited number of Z21 commands, and the responses of the command station. The event SendError is shown in the status line.

## RS-Bus Monitor ##
Shows the feedback information of all 128 RS-Bus addresses: 4 blocks of 32 rows, each row with 8 LEDs (two nibbles of 4 bits) and a small square for the type of the last feedback message (blue: feedback decoder, green: switch decoder with feedback, red: information from the command station). "Get RS-Bus data" requests both nibbles of all 128 addresses, in two parts of 128 requests (the FIFO of the library holds 128 messages); the second part is requested as soon as the answer on the last request of the first part has been received. The program also converts an RS-Bus address into switch addresses, and vice versa. The IP address is set in the preferences (RS-Bus Monitor → Preferences…), and is kept for the next start. From version 2.0 the program uses the library, instead of its own TCP code.
