# Examples (Qt: Linux)
This directory contains example applications in C++ (Qt 6 Widgets) that use the library XpressNetQt. They are the same examples as for [.NET](../DotNet/Examples.md): the layout of the forms (the .ui files) is converted from the .NET forms, and the code follows the .NET code. The .ui files can be changed with Qt Designer (or Qt Creator). The same source code is used for Windows, see [Examples-Windows.md](Examples-Windows.md).

The compiled programs for Linux (x86-64, Qt 6) are in [Examples/compiled/Linux-Qt](../../compiled/Linux-Qt), one zip file per example; Qt 6 must be installed (see the [installation manual](../../../Library/src/Qt/Install/Install-Library.md)). The program names are the same in all versions of the library (VB6, .NET, macOS, Windows and Linux); because of the spaces, a name must be put between quotes in the terminal.

## Building ##
Needed: Qt 6 (Widgets, Network, SerialPort), CMake and a C++17 compiler; see the [installation manual](../../../Library/src/Qt/Install/Install-Library.md) for the packages. For the all-in-one script also `zip` (`sudo apt install zip`).

Unzip (or clone) the complete repository and keep its directory structure. Copying only the directory of the examples does not work: the examples take the library source from `Library/src/Qt/XpressNetQt`, and the RS-Bus Decoder Scanner takes its CV overview from `Examples/src/macOS`. Both are found relative to the directory of the examples.

There are two ways to build:

| | All in one | Per example |
|---|---|---|
| How | [build-linux.sh](build-linux.sh) in a terminal | commands in a terminal |
| What | the library and all six examples, always from scratch | one example (and the library) |
| Result | a zip file per example in `Examples/compiled/Linux-Qt` | the program in the directory of the example |
| Use for | (re)building everything, for example for a new release | working on one example, or finding out which step fails |

### All in one: build-linux.sh ###
In the directory `Examples/src/Qt`:

    ./build-linux.sh

(or `sh build-linux.sh`, if the file is not executable after unzipping). The script checks that the library source, `cmake` and `zip` are found, removes the directory `build`, builds the library and all examples (Release), and makes `Examples/compiled/Linux-Qt/<example>.zip`. The programs can also be started from `build/<example>/`, for example `"./build/RS-Bus Decoder Scanner/RS-Bus Decoder Scanner"`.

### Per example ###
In the directory of the example, for example:

    cd "Examples/src/Qt/RS-Bus Decoder Scanner"
    cmake -S . -B build
    cmake --build build
    "./build/RS-Bus Decoder Scanner"

The library is built as well. All examples can also be built together, without zip files, in the directory `Examples/src/Qt`, with the same `cmake` commands; the programs are then in `build/<example>/`. If CMake reports that `../../../../Library/src/Qt/XpressNetQt` is not an existing directory, the directory structure is not complete.

## Differences with the .NET examples ##
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

## RS-Bus Decoder Scanner ##
Qt version of the macOS example: finds the decoders with RS-Bus feedback on up to 5 command stations (Lenz LAN/USB interfaces), and reads via PoM their type, software version and other CVs, which can be chosen in the preferences (File → Preferences…). The results are shown in a table, that can be saved as CSV file. The lift decoder, whose PoM address does not follow from its RS-Bus address, can be included via the preferences. This example shows how multiple XpressNet objects (one per command station) are used: the signals of every XpressNet object are connected to lambda functions that add the command station (DecoderScanner.cpp). The CV overview (Help → CV overview…) is the same HTML file as in the macOS version. The settings are kept with QSettings, in `~/.config/RS-Bus Decoder Scanner/Settings.conf`.
