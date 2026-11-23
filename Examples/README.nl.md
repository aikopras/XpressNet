[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet voorbeelden
De voorbeeld programma's laten zien hoe de library gebruikt wordt; sommige zijn complete hulpmiddelen (RS-Bus Monitor, RS-Bus Decoder Scanner). De namen van de programma's zijn in alle code bases gelijk.

| Voorbeeld | macOS | Qt | .NET | VB6 |
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

| Code base | Source en beschrijving | Gecompileerd |
|---|---|---|
| macOS (Xcode) | [src/macOS](src/macOS/Voorbeelden.md) | [compiled/macOS](compiled/macOS) |
| Qt (CMake) | [src/Qt](src/Qt) ([Windows](src/Qt/Voorbeelden-Windows.md), [Linux](src/Qt/Voorbeelden-Linux.md)) | [compiled/Windows-Qt](compiled/Windows-Qt), [compiled/Linux-Qt](compiled/Linux-Qt) |
| .NET (VB.NET) | [src/DotNet](src/DotNet/Voorbeelden.md) | [compiled/Windows-DotNet](compiled/Windows-DotNet) |
| VB6 | [src/VB6](src/VB6/Voorbeelden.md) | [compiled/Windows-VB6](compiled/Windows-VB6) |

De Qt voorbeelden worden gebouwd met de library in [Library/src/Qt](../Library/src/Qt), de .NET voorbeelden gebruiken de DLL in [Library/compiled/DotNet](../Library/compiled/DotNet), en de VB6 voorbeelden de geregistreerde DLL (zie de [installatie handleiding](../Library/src/VB6/Install/Installeren-Library.md)). De macOS voorbeelden wijken af: ieder Xcode project bevat een eigen kopie van de library (de directories XpressNet en Communication), zodat het op zichzelf gebouwd kan worden.
