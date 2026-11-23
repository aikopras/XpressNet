[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet
Deze repository bevat een XpressNet library, en voorbeeld programma's voor Windows, Mac en Linux die deze library gebruiken. ***Bijna alle XpressNet commando's (V3.6 en V4) zijn geïmplementeerd, alsmede een aantal Z21 commando's***. De initiële reden om deze library te ontwikkelen, was om bestaande VB6 programma's de mogelijkheid te geven via XpressNet (of Z21) met DCC centrales te communiceren. Versie 2 van de library is compleet herschreven, en is uitgebreid voor gebruik met moderne .NET Windows programma's, alsmede macOS en Linux.

De repository bestaat uit twee delen:
- **[Library](Library/README.nl.md)**: de source code van de library, de gecompileerde library (waar van toepassing) en de documentatie.
- **[Examples](Examples/README.nl.md)**: voorbeeld programma's, van kleine test programma's tot complete hulpmiddelen zoals de RS-Bus Monitor en de RS-Bus Decoder Scanner. Per voorbeeld de source code, en de gecompileerde programma's.

Beide delen zijn ingedeeld per code base:

| Code base | Taal | Te gebruiken met | Library | Voorbeelden |
|---|---|---|---|---|
| macOS | Objective-C | macOS applicaties (Xcode) | [src](Library/src/macOS) · [installatie](Library/src/macOS/Install/Installeren-Library.md) | [src](Examples/src/macOS) · [compiled](Examples/compiled/macOS) |
| Qt | C++ (Qt 6) | Linux en Windows applicaties (Qt, CMake) | [src](Library/src/Qt) · [installatie](Library/src/Qt/Install/Installeren-Library.md) | [src](Examples/src/Qt) · [compiled Windows](Examples/compiled/Windows-Qt) · [compiled Linux](Examples/compiled/Linux-Qt) |
| .NET | C# (.NET Framework 4.7.2) | VB.NET en C# applicaties (Visual Studio 2019 en nieuwer) | [src](Library/src/DotNet) · [installatie](Library/src/DotNet/Install/Installeren-Library.md) | [src](Examples/src/DotNet) · [compiled](Examples/compiled/Windows-DotNet) |
| VB6 | VB6 (COM DLL) | VB6 applicaties. Via COM ook andere talen (32 en 64 bit) | [src](Library/src/VB6) · [installatie](Library/src/VB6/Install/Installeren-Library.md) | [src](Examples/src/VB6) · [compiled](Examples/compiled/Windows-VB6) |

### Achtergrond
De [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) heeft een aantal VB6 programma's ontwikkeld waarmee de stations Enschede, Hengelo en Oldenzaal bestuurd worden; op dit moment worden deze programma's omgezet naar VB.NET. De besturing beperkt zich vooral tot wissels, seinen en bezetmelders (RS-Bus). Ondanks deze focus zijn de meeste XpressNet commando's geïmplementeerd, al kon niet alles even grondig getest worden. Vanwege de grote overeenkomst tussen het XpressNet en Z21 protocol is ook een subset van het Z21 protocol geïmplementeerd.

### Documentatie
- [Reference manual](Library/Documentation/Reference/Reference.md) (Engels): alle methoden, events en attributen, met de XpressNet berichten en de hoofdstukken in de specificaties.
- [PDF files in Library/Documentation/Commands](Library/Documentation/Commands): welke commando's geïmplementeerd zijn en hoe goed ze getest zijn.
- [Gebruikers handleiding](Library/Documentation/HowToUse/Handleiding.md): het gebruik van de library vanuit VB.NET en VB6.
- [Version history](Version_History.md) (Engels).
