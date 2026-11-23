# Examples (Qt: Windows)
This directory contains example applications in C++ (Qt 6 Widgets) that use the library XpressNetQt. They are the same examples as for [.NET](../DotNet/Examples.md): the layout of the forms (the .ui files) is converted from the .NET forms, and the code follows the .NET code. The .ui files can be changed with Qt Designer (or Qt Creator). The same source code is used for Linux, see [Examples-Linux.md](Examples-Linux.md).

The compiled programs for Windows (x64) are in [Examples/compiled/Windows-Qt](../../compiled/Windows-Qt), one zip file per example, with the Qt DLLs; unzip and start the .exe. The programs need the Microsoft Visual C++ Redistributable (x64), which is installed on most computers.

## Building ##
Needed: Visual Studio 2019 (workload "Desktop development with C++", with the component "C++ CMake tools for Windows") and Qt 6.7 for MSVC 2019 (Qt Online Installer: Qt 6.7.3 → MSVC 2019 64-bit and Additional Libraries → Qt Serial Port, in `C:\Qt`). With Visual Studio 2022, Qt 6.8 or later can be used in the same way (`msvc2022_64` instead of `msvc2019_64`).

Unzip the complete repository, for example to `C:\Users\<name>\Desktop\XpressNet`, and keep its directory structure. Copying only the directory of the examples does not work: the examples take the library source from `Library\src\Qt\XpressNetQt`, and the RS-Bus Decoder Scanner takes its CV overview from `Examples\src\macOS`. Both are found relative to the directory of the examples.

There are two ways to build:

| | All in one | Per example |
|---|---|---|
| How | double click [build-windows.bat](build-windows.bat) | commands in a command prompt |
| What | the library and all six examples, always from scratch | one example (and the library) |
| Result | a zip file per example in `Examples\compiled\Windows-Qt` | the program in the directory of the example |
| Use for | (re)building everything, for example for a new release | working on one example, or finding out which step fails |

### All in one: build-windows.bat ###
Double click `build-windows.bat` in `Examples\src\Qt`. The batch file:
1. checks that the library source, Qt, Visual Studio and its CMake are found (if not, it says what is missing);
2. removes the directories `build` and `deploy`, and builds the library and all examples (Release);
3. copies every program to `deploy\<example>`, adds the Qt DLLs with `windeployqt`, and makes `Examples\compiled\Windows-Qt\<example>.zip`.

The window stays open at the end, so that the messages can be read. The programs can also be started from `deploy\<example>`. The paths of Qt and Visual Studio are set in three lines at the top of the batch file (`QTDIR`, `GENERATOR` and `VSVERSION`); change these for another version of Qt or Visual Studio.

### Per example ###
Open "x64 Native Tools Command Prompt for VS 2019" (Start menu, Visual Studio 2019), and go to the directory of the example, for example:

    cd "C:\Users\<name>\Desktop\XpressNet\Examples\src\Qt\RS-Bus Decoder Scanner"
    dir ..\..\..\..\Library\src\Qt\XpressNetQt

The `dir` command must show `CMakeLists.txt`; if it reports "The system cannot find the path specified", the directory structure is not complete. Build the example (the library is built as well):

    cmake -S . -B build -G "Visual Studio 16 2019" -A x64 -DCMAKE_PREFIX_PATH=C:\Qt\6.7.3\msvc2019_64
    cmake --build build --config Release

The message "Could NOT find WrapVulkanHeaders" can be ignored (Vulkan is optional in Qt and not used). If the first `cmake` command failed, remove the `build` directory (`rmdir /s /q build`) before trying again.

The program is now in `build\Release`, for example `build\Release\RS-Bus Decoder Scanner.exe`. It needs the Qt DLLs; `windeployqt` copies these next to a copy of the program, in `deploy`:

    mkdir "deploy\RS-Bus Decoder Scanner"
    copy /y "build\Release\RS-Bus Decoder Scanner.exe" "deploy\RS-Bus Decoder Scanner"
    C:\Qt\6.7.3\msvc2019_64\bin\windeployqt.exe --release --no-translations --no-opengl-sw --no-system-d3d-compiler --no-compiler-runtime "deploy\RS-Bus Decoder Scanner\RS-Bus Decoder Scanner.exe"

The program can then be started from `deploy\RS-Bus Decoder Scanner`. To make the zip file (as the batch file does; `del` and `mkdir` may report that a file was not found or already exists, which can be ignored):

    del "deploy\RS-Bus Decoder Scanner\dxcompiler.dll" "deploy\RS-Bus Decoder Scanner\dxil.dll"
    mkdir ..\..\..\compiled\Windows-Qt
    tar -a -c -f "..\..\..\compiled\Windows-Qt\RS-Bus Decoder Scanner.zip" -C deploy "RS-Bus Decoder Scanner"

The DirectX shader compilers (dxcompiler.dll, dxil.dll) that windeployqt may add are not used by the programs. For the other examples the commands are the same, with their own name.

## Differences with the .NET examples ##
- The USB port (and the serial port) is the name of the COM port, for example COM3. The list contains the serial ports that Qt finds; another name can be typed.
- The logfiles are written to `C:\Users\<name>\TMC-Logfiles` (the examples set this directory; the default directory of the library is `C:\Users\Public\TMC-Logfiles`, as for .NET).
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
Shows the feedback information of all 128 RS-Bus addresses: 4 blocks of 32 rows, each row with 8 LEDs (two nibbles of 4 bits) and a small square for the type of the last feedback message (blue: feedback decoder, green: switch decoder with feedback, red: information from the command station). "Get RS-Bus data" requests both nibbles of all 128 addresses, in two parts of 128 requests (the FIFO of the library holds 128 messages); the second part is requested as soon as the answer on the last request of the first part has been received. The program also converts an RS-Bus address into switch addresses, and vice versa. The matrix is drawn by PanelManager (PanelManager.cpp, as in VB6 and .NET). The IP address is kept for the next start, in the registry, `HKEY_CURRENT_USER\Software\RS-Bus Monitor\Settings` (QSettings).

## RS-Bus Decoder Scanner ##
Qt version of the macOS example: finds the decoders with RS-Bus feedback on up to 5 command stations (Lenz LAN/USB interfaces), and reads via PoM their type, software version and other CVs, which can be chosen in the preferences (File → Preferences…). The results are shown in a table, that can be saved as CSV file. The lift decoder, whose PoM address does not follow from its RS-Bus address, can be included via the preferences. This example shows how multiple XpressNet objects (one per command station) are used: the signals of every XpressNet object are connected to lambda functions that add the command station (DecoderScanner.cpp). The CV overview (Help → CV overview…) is the same HTML file as in the macOS version. The settings are kept with QSettings, in the registry: `HKEY_CURRENT_USER\Software\RS-Bus Decoder Scanner\Settings`.
