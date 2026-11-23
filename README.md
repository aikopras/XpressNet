[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet
This repository contains an XpressNet library, and example programs for Windows, Mac and Linux that use this library. ***Most XpressNet commands (V3.6 and V4) are implemented, as well as some Z21 commands***. The original motivation for developing this library, was to allow existing VB6 programs to communicate via XpressNet (or Z21) with existing DCC command stations. Version 2 of the library is a major rewrite, and has been extended for use with modern .NET Windows programs, as well as macOS and Linux.

The repository has two parts:
- **[Library](Library/README.md)**: the source code of the library, the compiled library (where applicable) and the documentation.
- **[Examples](Examples/README.md)**: example programs, from small test programs to complete tools such as the RS-Bus Monitor and the RS-Bus Decoder Scanner. Per example the source code, and the compiled programs.

Both parts are divided per code base:

| Code base | Language | Use with | Library | Examples |
|---|---|---|---|---|
| macOS | Objective-C | macOS applications (Xcode) | [src](Library/src/macOS) · [install](Library/src/macOS/Install/Install-Library.md) | [src](Examples/src/macOS) · [compiled](Examples/compiled/macOS) |
| Qt | C++ (Qt 6) | Linux and Windows applications (Qt, CMake) | [src](Library/src/Qt) · [install](Library/src/Qt/Install/Install-Library.md) | [src](Examples/src/Qt) · [compiled Windows](Examples/compiled/Windows-Qt) · [compiled Linux](Examples/compiled/Linux-Qt) |
| .NET | C# (.NET Framework 4.7.2) | VB.NET and C# applications (Visual Studio 2019 and newer) | [src](Library/src/DotNet) · [install](Library/src/DotNet/Install/Install-Library.md) | [src](Examples/src/DotNet) · [compiled](Examples/compiled/Windows-DotNet) |
| VB6 | VB6 (COM DLL) | VB6 applications. Via COM also other languages (32 and 64 bit) | [src](Library/src/VB6) · [install](Library/src/VB6/Install/Install-Library.md) | [src](Examples/src/VB6) · [compiled](Examples/compiled/Windows-VB6) |

### Background
The [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) originally developed a number of VB6 programs to control its large model train layout; currently a switch is made towards VB.NET. The programs mainly control switches, signals and occupancy detectors (RS-Bus). Despite this focus, most XpressNet commands are implemented, although not all commands have been thoroughly tested. Since the XpressNet and Z21 protocols are quite similar, a subset of the Z21 commands is implemented as well.

### Documentation
- [Reference manual](Library/Documentation/Reference/Reference.md): all methods, events and attributes, with the XpressNet messages and the chapters in the specifications.
- [PDF files in Library/Documentation/Commands](Library/Documentation/Commands): which commands are implemented and how well they were tested.
- [User manual](Library/Documentation/HowToUse/Handleiding.md) (Dutch only): the use of the library from VB.NET and VB6.
- [Version history](Version_History.md).
