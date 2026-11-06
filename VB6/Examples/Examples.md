# Examples
This directory contains a number of example applications that demonstrate certain elements of using the XpressNet library.

Except for the last tests, all other test programs assume that the XpressNet DLL is installed in the windows registry. See the installation manual on how to do this registration. Don't forget to provide a Reference to the DLL!

The examples are for version 2.0 of the DLL. The messages of the library (`SendError`, and in the XpressNet examples also interface nicht / wieder adressiert) are shown in the status line; the Z21 examples have a new status line for this. For USB there is a check box Handshake. The logfiles are in `C:\Users\Public\TMC-Logfiles`.

The compiled programs are in [Examples-Compiled/VB6_Examples.zip](../Examples-Compiled/VB6_Examples.zip), together with the DLL. The program names are the same in all versions of the library (VB6, .NET, macOS and Linux). The examples use the registered DLL, except XpressNet Regfree Test (the DLL in the same directory) and XpressNet Project Group Test (its own DLL); see below.

## XpressNet Basic Test ##
This test program shows the basic use of the XpressNet DLL: connecting via TCP/IP or USB, requesting version information, switching accessories (with automatic deactivation), requesting switch and RS-Bus feedback, and controlling locomotives and their functions (F0 to F68). It also shows how the application can set the directory for the logfiles.

## XpressNet Test ##
This test program can be used to test a large number of XpressNet commands, and see which responses are sent by the command station.
<br/><img src="XpressNet%20Test/A_Screenshot.png" alt="XpressNet Test" width="700"/>
<br/><br/>

## Z21 Basic Test ##
This test program can be used to test a limited number of Z21 commands and see which responses are sent by the command station. After "Set" the library sets the broadcast flags of the Z21; this makes the status LEDs (NotAus, NotHalt, Programmier Mode) and the turnout info work.
<br/><img src="Z21%20Basic%20Test/A_Screenshot.png" alt="Z21 Basic Test" width="600"/>
<br/><br/>

## XpressNet Multiple Instances ##
It is possible to instantiate multiple XpressNet objects within one application. This test program demonstrates this. In the logfile every XpressNet object has its own letter (A, B, …).
<br/><img src="XpressNet%20Multiple%20Instances/A_Screenshot.png" alt="XpressNet Multiple Instances" width="700"/>
<br/><br/>

## XpressNet Regfree Test ##
Most of the example programs assume that the XpressNet DLL is registered, not only on the development, but also on the production machines (in the Windows Registry). To do this registration, administrator privileges are required on each machine. On production machines, it is sometimes impossible to obtain administrator privileges. In such cases, the application can also be written “Reg-Free”.

To do this, the XpressNet DLL must be registered on the development environment (i.e. administrator privileges are required on this machine only). The DLL file must also be added to the directory containing the source files as well as a special class file (XPNRegFree.cls). This class file should be added to the own program. See the image below.
<br/><img src="XpressNet%20Regfree%20Test/A_Screenshot-1.png" alt="Kopieer files" width="700"/>
<br/>

In addition, one line in the source file must be modified. See image below for details. This technique is described in [this entry in the VB6 forum](https://www.vbforums.com/showthread.php?908980-VB6-and-DLLs-Runtime-Error-429-ActiveX-Component-Can-t-Create-Object); with thanks to VanGoghGaming.
<br/><img src="XpressNet%20Regfree%20Test/A_Screenshot-2.png" alt="Kopieer files" width="700"/>
<br/>


In VB6_Examples.zip the compiled program (XpressNet Regfree Test.exe) and the DLL are in the same directory. These two files must be put on the production machine (in the same directory), after which the application can be started (without the DLL being registered).

## XpressNet Project Group Test ##
This example shows how the application and the XpressNet code can be taken together in one project group. The advantage of this approach is that the XpressNet code is integrated within the application. This may be useful during application testing, since the DLL need not yet be registered.

The group compiles the XpressNet code into its own DLL, in the subdirectory DLL (Project Compatibility with `DLL\XpressNet.dll`, so that the GUIDs stay the same). The compiled program (in VB6_Examples.zip the directory XpressNet Project Group Test) uses this DLL, and not the DLL in C:\XpressNetDll. On another machine this DLL must therefore be registered first (regsvr32).
<br/><img src="XpressNet%20Project%20Group%20Test/A_Screenshot.png" alt="XpressNet Project Group Test" width="300"/>

## RS-Bus Monitor ##
Shows the feedback information of all 128 RS-Bus addresses: 4 blocks of 32 rows, each row with 8 LEDs (two nibbles of 4 bits) and a small square for the type of the last feedback message (blue: feedback decoder, green: switch decoder with feedback, red: information from the command station). "Get RS-Bus data" requests both nibbles of all 128 addresses, in two parts of 128 requests (the FIFO of the library holds 128 messages); the second part is requested as soon as the answer on the last request of the first part has been received. The program also converts an RS-Bus address into switch addresses, and vice versa. The IP address is kept for the next start (SaveSetting / GetSetting). Before version 2.0 this program contained a copy of the classes of the library; it now uses the registered DLL.

## VB.NET (Visual Studio 2019)
The examples (except XpressNet Regfree Test and XpressNet Project Group Test) are also available in VB.NET, with the .NET library XpressNet.NET.dll: see the [.NET examples](../../DotNet/Examples/Examples.md). A VB.NET application can still use the VB6 DLL via COM (Project → Add Reference…, COM), but the .NET library is easier to install: it does not need to be registered.
