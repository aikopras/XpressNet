[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet examples
The example programs show how the library is used; some are complete tools (RS-Bus Monitor, RS-Bus Decoder Scanner). The program names are the same in all code bases.

| Example | macOS | Qt | .NET | VB6 |
|---|:---:|:---:|:---:|:---:|
| XpressNet Basic Test | ✓ | ✓ | ✓ | ✓ |
| XpressNet Test | ✓ | ✓ | ✓ | ✓ |
| Z21 Basic Test | ✓ | ✓ | ✓ | ✓ |
| XpressNet Multiple Instances |  | ✓ | ✓ | ✓ |
| XpressNet Project Group Test |  |  |  | ✓ |
| XpressNet Regfree Test |  |  |  | ✓ |
| RS-Bus Monitor | ✓ | ✓ | ✓ | ✓ |
| RS-Bus Decoder Scanner | ✓ | ✓ |  |  |

Per code base:

| Code base | Source and description | Compiled |
|---|---|---|
| macOS (Xcode) | [src/macOS](src/macOS/Examples.md) | [compiled/macOS](compiled/macOS) |
| Qt (CMake) | [src/Qt](src/Qt) ([Windows](src/Qt/Examples-Windows.md), [Linux](src/Qt/Examples-Linux.md)) | [compiled/Windows-Qt](compiled/Windows-Qt), [compiled/Linux-Qt](compiled/Linux-Qt) |
| .NET (VB.NET) | [src/DotNet](src/DotNet/Examples.md) | [compiled/Windows-DotNet](compiled/Windows-DotNet) |
| VB6 | [src/VB6](src/VB6/Examples.md) | [compiled/Windows-VB6](compiled/Windows-VB6) |

The Qt examples are built with the library in [Library/src/Qt](../Library/src/Qt), the .NET examples use the DLL in [Library/compiled/DotNet](../Library/compiled/DotNet), and the VB6 examples the registered DLL (see the [installation manual](../Library/src/VB6/Install/Install-Library.md)). The macOS examples are different: every Xcode project contains its own copy of the library (the directories XpressNet and Communication), so that it can be built on its own.
