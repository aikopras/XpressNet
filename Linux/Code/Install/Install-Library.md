# Installing the Linux library (XpressNetQt)

XpressNetQt is the Linux version of the XpressNet library. It is written in C++17 with Qt 6, and is a port of the .NET version: version 2.0 has the same functionality as version 2.0 of the VB6 DLL, the .NET library and the macOS version. The library uses only Qt (QtCore, QtNetwork and QtSerialPort) and no functions that are specific for Linux; the source code should therefore also be usable on macOS and Windows (see below). It was written and tested on Ubuntu 24.04 with Qt 6.4.2.

The source code consists of two directories:
- [XpressNet](../XpressNetQt/XpressNet): the XpressNet classes (XpressNetClass, Send, Receive, ReceiveZ21, AccAuto and LogFile)
- [Communication](../XpressNetQt/Communication): the TCP, UDP and serial (USB) connections

[CMakeLists.txt](../XpressNetQt/CMakeLists.txt) builds a static library, that is linked into the application. There is no precompiled version: Qt differs per distribution, so the library is built together with the application.

## Requirements
On Debian and Ubuntu (and derivatives such as Raspberry Pi OS and Linux Mint):
```
sudo apt install build-essential cmake qt6-base-dev qt6-serialport-dev
```
Other distributions have the same packages under other names (Fedora: `qt6-qtbase-devel qt6-qtserialport-devel cmake gcc-c++`). Qt 6 and CMake 3.16 or newer are needed. Qt Creator (`sudo apt install qtcreator`) is optional: it opens a CMakeLists.txt directly, and contains Qt Designer for the .ui files of the examples.

For a USB or serial connection, the user must be allowed to use the serial device. On most distributions this means: member of the group `dialout` (log out and in again afterwards):
```
sudo usermod -aG dialout $USER
```

## Building the library and the examples
In the directory Linux:
```
cmake -S . -B build
cmake --build build
```
This builds the library and the five [examples](../../Examples/Examples.md). The programs are then in build/Examples/…, for example `"./build/Examples/XpressNet Basic Test/XpressNet Basic Test"` (with quotes, because of the spaces).

## Using the library in your own application
Copy the directory XpressNetQt to your project (or refer to it), and add two lines to the CMakeLists.txt of the application:
```cmake
add_subdirectory(XpressNetQt)
target_link_libraries(MyApplication PRIVATE XpressNetQt)
```
The methods, signals (events) and attributes have the same names and parameters as in the VB6 DLL and the .NET library. The events are Qt signals; the application connects the signals it needs:

```cpp
#include "XpressNet.h"

MainWindow::MainWindow(QWidget *parent) : QWidget(parent)
{
  XpressNet = new XpressNetClass(this);
  connect(XpressNet, &XpressNetClass::InterfaceConnected, this, &MainWindow::InterfaceConnected);
  connect(XpressNet, &XpressNetClass::Feedback, this, &MainWindow::Feedback);
  XpressNet->CmdConnectTcp("192.168.0.200", "5550");
}

MainWindow::~MainWindow()
{
  delete XpressNet;                                       // Closes the connection
}

void MainWindow::InterfaceConnected()
{
  XpressNet->CmdAccessoryAuto(12, 1);
}

void MainWindow::Feedback(int address, quint8 nibble, quint8 data)
{
  label->setText(QString::number(address) + ": " + QString::number(data));
}
```

Everything runs in the thread that created the XpressNetClass object (normally the main thread of the application, with the Qt event loop). The signals are emitted in that thread, so a slot may directly change the widgets of the window.

See [XpressNet.h](../XpressNetQt/XpressNet/XpressNet.h) for all methods and signals, and the [Reference manual](../../../Documentation/Reference/Reference.md) for the XpressNet messages.

## Differences with the .NET library
- VB Integer and Long are `int`, Byte is `quint8`, Boolean is `bool` and String is `QString`. An Optional parameter has a default value.
- The attribute ZentraleVersion is called `ZentraleSoftwareVersion` (as in .NET).
- `CmdConnectUsb` and `CmdConnectSerial`: Port is the name of the serial device, for example `/dev/ttyUSB0` (FTDI, such as the LI-USB) or `/dev/ttyACM0`. If the port can not be opened, the signal `TcpError` is emitted (as in the macOS version; .NET: a MessageBox).
- `delete` closes the connection and stops the timers (.NET: `Dispose()`, VB6: `Set XpressNet = Nothing`).
- The logfile directory is set with `XpressNet->Logfile->setLogDirectory(...)` and read with `LogDirectory()`.
- The sockets never use a proxy of the system: the interface and the Z21 are in the local network.

## New in version 2.0 (the same as in the other versions)
- After an error answer of the interface the library sends a message again; as long as the command station does not address the interface, sending is paused.
- The signals `SendError(Reason, Description)`, `InterfaceNichtAdressiert()` and `InterfaceWiederAdressiert()`. The values of Reason are in `SendErrorReasonEnum`.
- `CmdConnectUsb` has the optional parameter `Handshake` (default off).
- In the logfile every line starts with the letter of the XpressNet object (A, B, …).
- Z21: after `CmdInitUdp` the library sets the broadcast flags (`CmdZ21SetBroadcastFlags`).

## Logfile
The logfile is the same as for the other versions: TMC-LogFile.txt (with at most three old logfiles TMC-LogFile-1..3.txt). The default directory is `~/TMC-Logfiles`; the library creates the directory.

## macOS and Windows
The library should also work on macOS and Windows, with Qt 6 for that operating system (for example installed with the Qt online installer, or on macOS with `brew install qt cmake`). This has not been tested. Differences:
- The serial port: on macOS for example `/dev/cu.usbserial-A1B2C3`, on Windows for example `COM3`.
- The default directory of the logfile: on macOS `~/Library/Logs/TMC-Logfiles` (as the macOS version), on Windows `%PUBLIC%\TMC-Logfiles` (as the VB6 and .NET versions).

For Windows applications the [.NET library](../../../DotNet/Code/Install/Install-Library.md) remains the obvious choice, and for macOS applications in Objective-C the [macOS version](../../../macOS/Code/Install/Install-Library.md).
