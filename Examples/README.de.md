[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet-Beispiele
Die Beispielprogramme zeigen, wie die Bibliothek verwendet wird; einige sind vollständige Werkzeuge (RS-Bus Monitor, RS-Bus Decoder Scanner). Die Programmnamen sind in allen Codebasen gleich.

| Beispiel | macOS | Qt | .NET | VB6 |
|---|:---:|:---:|:---:|:---:|
| XpressNet Basic Test | ✓ | ✓ | ✓ | ✓ |
| XpressNet Test | ✓ | ✓ | ✓ | ✓ |
| Z21 Basic Test | ✓ | ✓ | ✓ | ✓ |
| XpressNet Multiple Instances |  | ✓ | ✓ | ✓ |
| XpressNet Project Group Test |  |  |  | ✓ |
| XpressNet Regfree Test |  |  |  | ✓ |
| RS-Bus Monitor | ✓ | ✓ | ✓ | ✓ |
| RS-Bus Decoder Scanner | ✓ | ✓ |  |  |

Pro Codebasis (die Beschreibungen gibt es auf Englisch und Niederländisch):

| Codebasis | Quellcode und Beschreibung | Kompiliert |
|---|---|---|
| macOS (Xcode) | [src/macOS](src/macOS/Examples.md) | [compiled/macOS](compiled/macOS) |
| Qt (CMake) | [src/Qt](src/Qt) ([Windows](src/Qt/Examples-Windows.md), [Linux](src/Qt/Examples-Linux.md)) | [compiled/Windows-Qt](compiled/Windows-Qt), [compiled/Linux-Qt](compiled/Linux-Qt) |
| .NET (VB.NET) | [src/DotNet](src/DotNet/Examples.md) | [compiled/Windows-DotNet](compiled/Windows-DotNet) |
| VB6 | [src/VB6](src/VB6/Examples.md) | [compiled/Windows-VB6](compiled/Windows-VB6) |

Die Qt-Beispiele werden mit der Bibliothek in [Library/src/Qt](../Library/src/Qt) gebaut, die .NET-Beispiele verwenden die DLL in [Library/compiled/DotNet](../Library/compiled/DotNet) und die VB6-Beispiele die registrierte DLL (siehe die [Installationsanleitung](../Library/src/VB6/Install/Install-Library.md)). Die macOS-Beispiele sind anders: jedes Xcode-Projekt enthält eine eigene Kopie der Bibliothek (die Verzeichnisse XpressNet und Communication), damit es für sich allein gebaut werden kann.
