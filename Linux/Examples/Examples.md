# Examples (Linux, Qt)
This directory contains example applications in C++ (Qt 6 Widgets) that use the Linux library XpressNetQt. They are the same examples as for [.NET](../../DotNet/Examples/Examples.md): the layout of the forms (the .ui files) is converted from the .NET forms, and the code follows the .NET code. The .ui files can be changed with Qt Designer (or Qt Creator).

There are no compiled programs: the examples are built from the source code, together with the library. See the [installation manual](../Code/Install/Install-Library.md) for the packages that are needed. In the directory Linux:
```
cmake -S . -B build
cmake --build build
"./build/Examples/XpressNet Basic Test/XpressNet Basic Test"
```
The program names are the same in all versions of the library (VB6, .NET, macOS and Linux); because of the spaces, a name must be put between quotes in the terminal. Every example can also be built on its own, with the CMakeLists.txt in its directory (for example `cmake -S "Examples/RS-Bus Monitor" -B build-rs`); the library is then built as well.

Differences with the .NET examples:
- The USB port (and the serial port) is the name of the serial device, for example /dev/ttyUSB0. The list contains the serial ports that Qt finds; another name can be typed.
- The logfiles are written to ~/TMC-Logfiles.
- The LEDs are a small widget (Led.h / Led.cpp, in every example). A Led can also be created in the code, for example for a grid of many LEDs.

## XpressNet Basic Test ##
The basic use of the library: connecting via TCP/IP or USB, requesting version information, switching accessories (with automatic deactivation), requesting switch and RS-Bus feedback, and controlling locomotives and their functions (F0 to F68). It also shows how the application can set the directory for the logfiles. Messages of the library (SendError, interface nicht / wieder adressiert) are shown in the status line.

## XpressNet Test ##
With this test program a large number of XpressNet commands can be tested (also serial connections, model time, programming and the erweiterte Versionsinformation), and the responses of the command station can be viewed.

## Z21 Basic Test ##
With this test program a limited number of Z21 commands can be tested, and the responses of the command station can be viewed. After "Set" the library sets the broadcast flags of the Z21; this makes the status LEDs (NotAus, NotHalt, Programmier Mode) and the turnout info work.

## XpressNet Multiple Instances ##
An application can create multiple XpressNet objects. Every click on "Add XpressNet" or "Add Z21" opens a new form (XpressNet Test or Z21 Basic Test) with its own XpressNet object. In the logfile every XpressNet object has its own letter (A, B, …). Closing the start form closes the other forms as well.

## RS-Bus Monitor ##
Shows the feedback information of all 128 RS-Bus addresses: 4 blocks of 32 rows, each row with 8 LEDs (two nibbles of 4 bits) and a small square for the type of the last feedback message (blue: feedback decoder, green: switch decoder with feedback, red: information from the command station). "Get RS-Bus data" requests both nibbles of all 128 addresses, in two parts of 128 requests (the FIFO of the library holds 128 messages); the second part is requested as soon as the answer on the last request of the first part has been received. The program also converts an RS-Bus address into switch addresses, and vice versa. The matrix is drawn by PanelManager (PanelManager.cpp, as in VB6 and .NET). The IP address is kept for the next start, in ~/.config/RS-Bus Monitor/Settings.conf (QSettings).
