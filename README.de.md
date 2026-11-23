[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet
Dieses Repository enthält eine XpressNet-Bibliothek und Beispielprogramme für Windows, Mac und Linux, die diese Bibliothek verwenden. ***Die meisten XpressNet-Befehle (V3.6 und V4) sind implementiert, sowie einige Z21-Befehle***. Ursprünglich wurde die Bibliothek entwickelt, damit bestehende VB6-Programme über XpressNet (oder Z21) mit DCC-Zentralen kommunizieren können. Version 2 der Bibliothek ist eine vollständige Neufassung und wurde für moderne .NET-Windows-Programme sowie für macOS und Linux erweitert.

Das Repository besteht aus zwei Teilen:
- **[Library](Library/README.de.md)**: der Quellcode der Bibliothek, die kompilierte Bibliothek (soweit zutreffend) und die Dokumentation.
- **[Examples](Examples/README.de.md)**: Beispielprogramme, von kleinen Testprogrammen bis zu vollständigen Werkzeugen wie dem RS-Bus Monitor und dem RS-Bus Decoder Scanner. Pro Beispiel der Quellcode und die kompilierten Programme.

Beide Teile sind nach Codebasis gegliedert:

| Codebasis | Sprache | Verwendung mit | Library | Beispiele |
|---|---|---|---|---|
| macOS | Objective-C | macOS-Anwendungen (Xcode) | [src](Library/src/macOS) · [Installation](Library/src/macOS/Install/Install-Library.md) | [src](Examples/src/macOS) · [compiled](Examples/compiled/macOS) |
| Qt | C++ (Qt 6) | Linux- und Windows-Anwendungen (Qt, CMake) | [src](Library/src/Qt) · [Installation](Library/src/Qt/Install/Install-Library.md) | [src](Examples/src/Qt) · [compiled Windows](Examples/compiled/Windows-Qt) · [compiled Linux](Examples/compiled/Linux-Qt) |
| .NET | C# (.NET Framework 4.7.2) | VB.NET- und C#-Anwendungen (Visual Studio 2019 und neuer) | [src](Library/src/DotNet) · [Installation](Library/src/DotNet/Install/Install-Library.md) | [src](Examples/src/DotNet) · [compiled](Examples/compiled/Windows-DotNet) |
| VB6 | VB6 (COM-DLL) | VB6-Anwendungen. Über COM auch andere Sprachen (32 und 64 Bit) | [src](Library/src/VB6) · [Installation](Library/src/VB6/Install/Install-Library.md) | [src](Examples/src/VB6) · [compiled](Examples/compiled/Windows-VB6) |

Die Installationsanleitungen gibt es auf Englisch und Niederländisch.

### Hintergrund
Der [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) hat ursprünglich eine Reihe von VB6-Programmen zur Steuerung seiner großen Modellbahnanlage entwickelt; derzeit werden diese nach VB.NET umgestellt. Die Programme steuern vor allem Weichen, Signale und Besetztmelder (RS-Bus). Trotz dieses Schwerpunkts sind die meisten XpressNet-Befehle implementiert, auch wenn nicht alle gründlich getestet werden konnten. Da sich die Protokolle XpressNet und Z21 stark ähneln, ist auch ein Teil der Z21-Befehle implementiert.

### Dokumentation
- [Reference manual](Library/Documentation/Reference/Reference.md) (Englisch): alle Methoden, Events und Attribute, mit den XpressNet-Nachrichten und den Kapiteln der Spezifikationen.
- [PDF-Dateien in Library/Documentation/Commands](Library/Documentation/Commands): welche Befehle implementiert sind und wie gut sie getestet wurden.
- [Benutzerhandbuch](Library/Documentation/HowToUse/Handleiding.md) (nur Niederländisch): die Verwendung der Bibliothek aus VB.NET und VB6.
- [Version history](Version_History.md) (Englisch).
