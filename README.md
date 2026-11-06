# XpressNet (EN)
This repository contains an XpressNet library, and example programs for Windows, Mac and Linux that use this library. ***Most XpressNet commands (V3.6 and V4) are implemented, as well as some Z21 commands***. The original motivation for developing this XpressNet library, was to allow existing VB6 programs to communicate via XpressNet (or Z21) with existing DCC command stations. The current version 2 of the library is a major rewrite, and has been extended for use with modern .Net Windows programs, as well as macOS and Linux.

| Version | Language | Use with | Installation | Examples |
|---|---|---|---|---|
| [VB6](VB6) | VB6 (COM DLL) | VB6 applications. Via COM also other languages (32 and 64 bit) | [manual](VB6/Code/Install/Install-Library.md) | [examples](VB6/Examples/Examples.md) |
| [.NET](DotNet) | C# (.NET Framework 4.7.2) | VB.NET and C# applications (Visual Studio 2019 and newer) | [manual](DotNet/Code/Install/Install-Library.md) | [examples](DotNet/Examples/Examples.md) |
| [macOS](macOS) | Objective-C | macOS applications (Xcode) | [manual](macOS/Code/Install/Install-Library.md) | [examples](macOS/Examples/Examples.md) |
| [Linux](Linux) | C++ (Qt 6) | Linux applications (Qt, CMake). Should also be usable on macOS and Windows | [manual](Linux/Code/Install/Install-Library.md) | [examples](Linux/Examples/Examples.md) |

### Background
The [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) originally developed a number of VB6 programs to control their large modeltrain layout, and currently a switch is made towards VB.NET. For VB.NET applications the .NET library is the easiest choice, since it does not need to be registered in the windows registry.
Our programs mainly control switches, signals and occupancy detectors (RS-Bus).

Despite TMC's focus on accessory commands and feedbacks, this library has implemented most XpressNet commands, although not all commands have been thoroughly tested. Since there are quite some similarities between the XpressNet and Z21 protocols, this library implements a subset of the Z21 commands as well.

All methods, events and attributes, with the XpressNet messages and the chapters in the specifications, are described in the [Reference manual](Documentation/Reference/Reference.md).

The various [PDF files in Documentation/Commands](Documentation/Commands) provide further details. The first column shows the VB6 stub or method, the events and logfile messages that are associated with each specific command; if empty the command is not implemented. The various colours indicate how well a command was tested.

The changes per version are described in the [Version history](Version_History.md).

The [User manual](Documentation/HowToUse/Handleiding.md) (only available in Dutch) describes how to use the library from VB.NET and VB6. The old PDF version of the manual (version 1.3, VB6 only) is still in the Documentation/HowToUse directory.

---

# XpressNet (NL)
Deze repository bevat een XpressNet library, alsmede een aantal voorbeeld programma's die tonen hoe de library kan worden gebruikt. Bijna alle XpressNet commando's zijn geïmplementeerd, alsmede een aantal Z21 commando's. De initiële reden om deze DLL te ontwikkelen, was om bestaande VB6 programma's de mogelijkheid te geven via XpressNet (of Z21) met DCC centrales te communiceren. De huidige versie 2 van deze library is compleet herschreven, en is uitgebreid om ook gebruikt te kunnen worden met moderne .Net Windows programma's, alsmede macOS en Linux.


| Versie | Taal | Te gebruiken met | Installatie | Voorbeelden |
|---|---|---|---|---|
| [VB6](VB6) | VB6 (COM DLL) | VB6 applicaties. Via COM ook andere talen (32 en 64 bit) | [handleiding](VB6/Code/Install/Installeren-Library.md) | [voorbeelden](VB6/Examples/Voorbeelden.md) |
| [.NET](DotNet) | C# (.NET Framework 4.7.2) | VB.NET en C# applicaties (Visual Studio 2019 en nieuwer) | [handleiding](DotNet/Code/Install/Installeren-Library.md) | [voorbeelden](DotNet/Examples/Voorbeelden.md) |
| [macOS](macOS) | Objective-C | macOS applicaties (Xcode) | [handleiding](macOS/Code/Install/Installeren-Library.md) | [voorbeelden](macOS/Examples/Voorbeelden.md) |
| [Linux](Linux) | C++ (Qt 6) | Linux applicaties (Qt, CMake). Zou ook op macOS en Windows bruikbaar moeten zijn | [handleiding](Linux/Code/Install/Installeren-Library.md) | [voorbeelden](Linux/Examples/Voorbeelden.md) |

## Achtergrond
De [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) heeft een aantal VB6 programma's ontwikkeld waarmee de stations Enschede, Hengelo en Oldenzaal bestuurd worden. Op dit moment worden deze VB6 programma's omgezet naar .NET. Voor VB.NET applicaties is de .NET library de eenvoudigste keuze, omdat die niet geregistreerd hoeft te worden in de Windows Registry. De besturing beperkt zich vooral tot wissels, seinen, en bezetmelders (RS-Bus).

Ondanks deze focus op wissels, seinen en terugmelders, zijn de meeste XpressNet commando's gewoon geïmplementeerd. Niet alles kon echter even rigoureus getest worden, dus fouten zijn niet uitgesloten. Vanwege de grote overeenkomst tussen het XpressNet en Z21 protocol, is ook een subset van het Z21 protocol geïmplementeerd.

Alle methoden, events en attributen, met de XpressNet berichten en de hoofdstukken in de specificaties, staan in de [Reference manual](Documentation/Reference/Reference.md) (Engels).

Om precies te weten welke commando's geïmplementeerd zijn, download de [PDF files in Documentation/Commands](Documentation/Commands) voor details. De eerste kolom bevat de VB6 stub / event / logfile bericht dat bij het specifieke commando behoort; indien deze leeg is, is het betreffende commando niet geïmplementeerd. De kleuren geven de graad van testen aan.

De wijzigingen per versie staan in de [Versiegeschiedenis](Versiegeschiedenis.md).

De [Gebruikers handleiding](Documentation/HowToUse/Handleiding.md) beschrijft het gebruik van de library vanuit VB.NET en VB6. De oude PDF versie van de handleiding (versie 1.3, alleen VB6) staat nog in de directory Documentation/HowToUse.
