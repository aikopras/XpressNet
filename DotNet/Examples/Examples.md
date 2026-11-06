# Examples (.NET)
This directory contains example applications in VB.NET (Visual Studio 2019, .NET Framework 4.7.2) that use the .NET library XpressNet.NET.dll. They are the same examples as for [VB6](../../VB6/Examples/Examples.md): the layout of the forms is taken from the VB6 forms (somewhat wider where needed, since the texts need more room under .NET), and the code follows the VB6 code.

The projects refer to the DLL in the directory DotNet/Code (Project → References → XpressNet.NET). After downloading the complete repository, the examples can be built and started directly; registering the DLL is not needed.

The compiled programs (.exe and .exe.config) are, together with the DLL, in [Examples-Compiled/DotNet_Examples.zip](../Examples-Compiled/DotNet_Examples.zip). The program names are the same in all versions of the library (VB6, .NET, macOS and Linux). The program, its .exe.config and the DLL must be put in the same directory on the production machine, after which the application can be started. .NET Framework 4.7.2 or newer is needed (standard in current versions of Windows 10 and 11).

## XpressNet Basic Test ##
The basic use of the library: connecting via TCP/IP or USB, requesting version information, switching accessories (with automatic deactivation), requesting switch and RS-Bus feedback, and controlling locomotives and their functions (F0 to F68). It also shows how the application can set the directory for the logfiles. Messages of the library (SendError, interface nicht / wieder adressiert) are shown in the status line.

## XpressNet Test ##
With this test program a large number of XpressNet commands can be tested (also serial connections, model time, programming and the erweiterte Versionsinformation), and the responses of the command station can be viewed.

## Z21 Basic Test ##
With this test program a limited number of Z21 commands can be tested, and the responses of the command station can be viewed. After "Set" the library sets the broadcast flags of the Z21; this makes the status LEDs (NotAus, NotHalt, Programmier Mode) and the turnout info work.

## XpressNet Multiple Instances ##
An application can create multiple XpressNet objects. Every click on "Add XpressNet" or "Add Z21" opens a new form (XpressNet Test or Z21 Basic Test) with its own XpressNet object. In the logfile every XpressNet object has its own letter (A, B, …).

## RS-Bus Monitor ##
Shows the feedback information of all 128 RS-Bus addresses: 4 blocks of 32 rows, each row with 8 LEDs (two nibbles of 4 bits) and a small square for the type of the last feedback message (blue: feedback decoder, green: switch decoder with feedback, red: information from the command station). "Get RS-Bus data" requests both nibbles of all 128 addresses, in two parts of 128 requests (the FIFO of the library holds 128 messages); the second part is requested as soon as the answer on the last request of the first part has been received. The program also converts an RS-Bus address into switch addresses, and vice versa. The IP address is kept for the next start, with SaveSetting / GetSetting (the same setting as the VB6 program).

## XpressNet Regfree Test and XpressNet Project Group Test ##
These VB6 examples have no .NET version: the .NET library does not need to be registered, and the library project (DotNet/Code/XpressNet.NET/XpressNet.NET.csproj) can be added to the solution of the application (see the [installation manual](../Code/Install/Install-Library.md)).
