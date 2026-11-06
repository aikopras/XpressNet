# Installeren van de Linux library (XpressNetQt)

XpressNetQt is de Linux versie van de XpressNet library. De library is geschreven in C++17 met Qt 6, en is een omzetting van de .NET versie: versie 2.0 heeft dezelfde functionaliteit als versie 2.0 van de VB6 DLL, de .NET library en de macOS versie. De library gebruikt alleen Qt (QtCore, QtNetwork en QtSerialPort) en geen functies die specifiek zijn voor Linux; de source code zou daarom ook op macOS en Windows bruikbaar moeten zijn (zie hieronder). De library is geschreven en getest op Ubuntu 24.04 met Qt 6.4.2.

De source code bestaat uit twee directories:
- [XpressNet](../XpressNetQt/XpressNet): de XpressNet classes (XpressNetClass, Send, Receive, ReceiveZ21, AccAuto en LogFile)
- [Communication](../XpressNetQt/Communication): de TCP, UDP en seriële (USB) verbindingen

[CMakeLists.txt](../XpressNetQt/CMakeLists.txt) bouwt een statische library, die in de applicatie wordt meegelinkt. Er is geen gecompileerde versie: Qt verschilt per distributie, daarom wordt de library samen met de applicatie gebouwd.

## Benodigd
Op Debian en Ubuntu (en afgeleiden zoals Raspberry Pi OS en Linux Mint):
```
sudo apt install build-essential cmake qt6-base-dev qt6-serialport-dev
```
Andere distributies hebben dezelfde pakketten onder een andere naam (Fedora: `qt6-qtbase-devel qt6-qtserialport-devel cmake gcc-c++`). Nodig zijn Qt 6 en CMake 3.16 of nieuwer. Qt Creator (`sudo apt install qtcreator`) is niet nodig, maar handig: het opent een CMakeLists.txt direct, en bevat Qt Designer voor de .ui files van de voorbeelden.

Voor een USB of seriële verbinding moet de gebruiker het seriële device mogen gebruiken. Bij de meeste distributies betekent dit: lid van de groep `dialout` (daarna opnieuw inloggen):
```
sudo usermod -aG dialout $USER
```

## De library en de voorbeelden bouwen
In de directory Linux:
```
cmake -S . -B build
cmake --build build
```
Dit bouwt de library en de vijf [voorbeelden](../../Examples/Voorbeelden.md). De programma's staan daarna in build/Examples/…, bijvoorbeeld `"./build/Examples/XpressNet Basic Test/XpressNet Basic Test"` (met aanhalingstekens, vanwege de spaties).

## De library in een eigen applicatie gebruiken
Kopieer de directory XpressNetQt naar het project (of verwijs ernaar), en voeg twee regels toe aan de CMakeLists.txt van de applicatie:
```cmake
add_subdirectory(XpressNetQt)
target_link_libraries(MijnApplicatie PRIVATE XpressNetQt)
```
De methods, signals (events) en attributen hebben dezelfde namen en parameters als in de VB6 DLL en de .NET library. De events zijn Qt signals; de applicatie koppelt de signals die nodig zijn:

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
  delete XpressNet;                                       // Sluit de verbinding
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

Alles draait in de thread die het XpressNetClass object heeft gemaakt (normaal de main thread van de applicatie, met de Qt event loop). De signals komen in die thread, dus een slot mag direct de widgets van het venster aanpassen.

Zie [XpressNet.h](../XpressNetQt/XpressNet/XpressNet.h) voor alle methods en signals, en de [Reference manual](../../../Documentation/Reference/Reference.md) (Engels) voor de XpressNet berichten.

## Verschillen met de .NET library
- VB Integer en Long zijn `int`, Byte is `quint8`, Boolean is `bool` en String is `QString`. Een Optional parameter heeft een default waarde.
- Het attribuut ZentraleVersion heet `ZentraleSoftwareVersion` (net als bij .NET).
- `CmdConnectUsb` en `CmdConnectSerial`: Port is de naam van het seriële device, bijvoorbeeld `/dev/ttyUSB0` (FTDI, zoals de LI-USB) of `/dev/ttyACM0`. Als de poort niet geopend kan worden, komt het signal `TcpError` (net als bij de macOS versie; .NET: een MessageBox).
- `delete` sluit de verbinding en stopt de timers (.NET: `Dispose()`, VB6: `Set XpressNet = Nothing`).
- De directory van de logfile wordt ingesteld met `XpressNet->Logfile->setLogDirectory(...)` en gelezen met `LogDirectory()`.
- De sockets gebruiken nooit een proxy van het systeem: het interface en de Z21 staan in het lokale netwerk.

## Nieuw in versie 2.0 (hetzelfde als bij de andere versies)
- Na een foutmelding van het interface verstuurt de library een bericht opnieuw; zolang de centrale het interface niet adresseert, wordt het versturen gepauzeerd.
- De signals `SendError(Reason, Description)`, `InterfaceNichtAdressiert()` en `InterfaceWiederAdressiert()`. De waarden van Reason staan in `SendErrorReasonEnum`.
- `CmdConnectUsb` heeft de optionele parameter `Handshake` (standaard uit).
- In de logfile begint iedere regel met de letter van het XpressNet object (A, B, …).
- Z21: na `CmdInitUdp` zet de library de broadcast flags (`CmdZ21SetBroadcastFlags`).

## Logfile
De logfile is dezelfde als bij de andere versies: TMC-LogFile.txt (met maximaal drie oude logfiles TMC-LogFile-1..3.txt). De standaard directory is `~/TMC-Logfiles`; de library maakt deze directory zelf aan.

## macOS en Windows
De library zou ook op macOS en Windows moeten werken, met Qt 6 voor dat besturingssysteem (bijvoorbeeld geïnstalleerd met de Qt online installer, of op macOS met `brew install qt cmake`). Dit is niet getest. Verschillen:
- De seriële poort: op macOS bijvoorbeeld `/dev/cu.usbserial-A1B2C3`, op Windows bijvoorbeeld `COM3`.
- De standaard directory van de logfile: op macOS `~/Library/Logs/TMC-Logfiles` (net als de macOS versie), op Windows `%PUBLIC%\TMC-Logfiles` (net als de VB6 en .NET versies).

Voor Windows applicaties blijft de [.NET library](../../../DotNet/Code/Install/Installeren-Library.md) de voor de hand liggende keuze, en voor macOS applicaties in Objective-C de [macOS versie](../../../macOS/Code/Install/Installeren-Library.md).
