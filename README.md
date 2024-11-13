# XpressNet (EN)
This repository contains an XpressNet Library (DLL) and some example programs that use the DLL. Most XpressNet commands are implemented, as well as some Z21 commands. The code is written in VB6, but it is believed that the DLL can also be used form other languages.

To download and install the DLL, see the [manual](Code/Install/Install-Library.md) for details.

### Background
The motivation for developing this XpressNet DLL, was to allow existing VB6 programs to communicate via XpressNet (or Z21) with existing DCC command stations.

The [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) uses such VB6 programs to control their large modeltrain layout.
Since control is (currently) limited to switches, signals and feedbacks, the DLL focuses on Accessory Commands and (RS-Bus) Feedback signals.

Despite TMC's focus on accessory commands and feedbacks, this library has implemented most XpressNet commands, although not all commands have been thoroughly tested. Since there are quite some similarities between the XpressNet and Z21 protocols, this DLL implements a subset of the Z21 commands as well.

Download the various [PDF files in Documentation/Commands](Documentation/Commands) for details. The first column shows the VB6 stub / event / logfile message that is associated with each specific command; if empty the command is not implemented. The various colors indicate how well a command was tested.

[Examples](Examples/Examples.md) can be found under the "Examples" branch of this repository.

A PDF version of the [User manual](Documentation/HowToUse/XpressNet-DLL_NL.md) (currently only available in Dutch) can be found under the Documentation/Howto branch of this repository.

---

# XpressNet (NL)
Deze repository bevat een XpressNet Library (DLL), alsmede een aantal voorbeeld programma's die tonen hoe de DLL kan worden gebruikt. Bijna alle XpressNet commando's zijn geïmplementeerd, alsmede een aantal Z21 commando's. De code is geschreven in VB6, maar de verwachting is dat de DLL ook in combinatie met (modernere) programmeertalen gebruikt kan worden.

Om de DLL te downloaden en installeren, zie de [handleiding](Code/Install/Installeren-Library.md) voor verdere details.

## Achtergrond
De motivatie om deze DLL te ontwikkelen, is om bestaande VB6 programma's de mogelijkheid te geven via XpressNet (of Z21) met DCC centrales te communiceren.

De [Twentse Modelspoorwegclub](https://twentsemodelspoorweg.club) (TMC) gebruikt momenteel een aantal VB6 programma's waarmee de stations Enschede, Hengelo en Oldenzaal bestuurd worden. De besturing beperkt zich (voorlopig?) vooral to wissels, seinen, en bezetmelders (RS-Bus).

Ondanks deze focus op wissels, seinen en terugmelders, zijn de meeste XpressNet commando's gewoon geïmplementeerd. Niet alles kon echter even rigoureus getest worden, dus fouten zijn niet uitgesloten. Vanwege de grote overeenkomst tussen het XpressNet en Z21 protocol, is ook een subset van het Z21 protocol geïmplementeerd.

Om precies te weten welke commando's geïmplementeerd zijn, download de [PDF files in Documentation/Commands](Documentation/Commands) voor details. De eerste kolom bevat de VB6 stub / event / logfile bericht dat bij het specifieke commando behoort; indien deze leeg is, is het betreffende commando niet geïmplementeerd. De kleuren geven de graad van testen aan.

[Voorbeelden](Examples/Voorbeelden.md) staan onder de "Examples" tak van deze repository.

Een PDF versie van de [Gebruikers handleiding](Documentation/HowToUse/XpressNet-DLL_NL.md) kan gedownload worden vanuit de "Documentation/Howto" tak van deze repository.
