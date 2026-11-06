# Versiegeschiedenis

## Versie 2.0 (november 2026)

Tot en met versie 1.3 was de library alleen beschikbaar als VB6 DLL. Vanaf versie 2.0 is er ook een .NET versie, een macOS versie en een Linux versie (Qt), met dezelfde functionaliteit. De repository is daarom ingedeeld in de directories `VB6`, `DotNet`, `macOS` en `Linux`.

**Nieuw**
- .NET versie (XpressNet.NET, C#, .NET Framework 4.7.2) voor VB.NET en C# applicaties. De namen en parameters zijn dezelfde als bij de VB6 DLL; de DLL hoeft niet geregistreerd te worden. De voorbeelden zijn er ook in VB.NET (behalve XpressNet Regfree Test en XpressNet Project Group Test).
- macOS versie (Objective-C), met vier Xcode voorbeelden.
- Linux versie (XpressNetQt, C++17 met Qt 6, CMake), met vijf voorbeelden. De library gebruikt alleen Qt, en zou daarom ook op macOS en Windows bruikbaar moeten zijn.
- Na een foutmelding van het interface verstuurt de library een bericht opnieuw, en zolang de centrale het interface niet adresseert, wordt het versturen gepauzeerd.
- Nieuwe events `SendError`, `InterfaceNichtAdressiert` en `InterfaceWiederAdressiert`. Een bericht dat niet verstuurd kan worden, verdwijnt niet meer zonder melding.
- `CmdConnectUsb` heeft de optionele parameter `Handshake`.
- De logfile staat standaard in `%PUBLIC%\TMC-Logfiles` (macOS: `~/Library/Logs/TMC-Logfiles`, Linux: `~/TMC-Logfiles`). In de logfile begint ook de beschrijving van een bericht met `=>` (verstuurd) of `<=` (ontvangen).
- In de logfile begint iedere regel met de letter van het XpressNet object (A, B, …). Bij het opbouwen van een verbinding zegt een regel als `A = 192.168.0.200, port 5550, local port 54321` bij welke verbinding de letter hoort (1.3: het IP adres of de COM poort aan het eind van de regel).
- TCP: een verbindingspoging stopt na 5 seconden (Windows: ca. 21 seconden).
- Z21: na `CmdInitUdp` zet de library de broadcast flags (LAN_SET_BROADCASTFLAGS), zodat de Z21 o.a. Alles An / Alles Aus en wisselmeldingen stuurt. Nieuw commando `CmdZ21SetBroadcastFlags`. Bij een nieuwe verbinding en bij afsluiten meldt de library zich af (LAN_LOGOFF).
- Nieuw voorbeeld RS-Bus Monitor, in alle vier de versies: toont de terugmeld informatie van alle 128 RS-Bus adressen. De VB6 en macOS versies bestonden al, met een eigen kopie van de library code; ze gebruiken nu de library.
- De voorbeeld programma's hebben in alle vier de versies dezelfde namen: XpressNet Basic Test, XpressNet Test, Z21 Basic Test, XpressNet Multiple Instances, XpressNet Regfree Test en XpressNet Project Group Test (alleen VB6), en RS-Bus Monitor. De directories van de voorbeelden hebben ook deze namen (1.3: 1-Basic_test, 2-Extended_test, …).
- Documentatie: een [Reference manual](Documentation/Reference/Reference.md) (Engels) en de [handleiding](Documentation/HowToUse/Handleiding.md) in Markdown.

**Wijzigingen**
- `CmdAccessoryAuto`: als de wachtrij geleegd wordt, krijgt het wissel dat op dat moment bekrachtigd is eerst nog een UIT commando.
- Ontvangen: na een bericht met een XOR fout zoekt de library naar het begin van het volgende bericht.
- Een nieuwe verbinding sluit eerst een bestaande verbinding van hetzelfde object.

**Let op bij het overstappen**: applicaties moeten opnieuw gebouwd worden (door de extra parameter van `CmdConnectUsb`). Bij .NET heet het attribuut `ZentraleVersion` `ZentraleSoftwareVersion`, en sluit `Dispose()` de verbinding. De VB6 DLL krijgt bij het compileren nieuwe GUID's: VB6 applicaties moeten de referentie opnieuw zetten.

## Versie 1.3 (oktober 2026)

Alleen de VB6 DLL.

- Correcties in de PoM commando's (CV nummer en bit schrijven), de vergelijking van de versie van de centrale (als getal), het decoderen van de snelheid (14 en 128 rijstappen) en van Z21 berichten, en de invoerbuffer bij veel ontvangen data.
- De response time-out (5 seconden) geldt ook voor USB en serieel; voor de Z21 is er geen flow control meer.
- `CmdAccessoryAuto`: na een deactivering volgt een korte pauze voordat het volgende wissel wordt geactiveerd.
- De DLL heeft nieuwe GUID's: applicaties moeten de referentie opnieuw zetten en opnieuw gecompileerd worden.

## Versie 1.2 (augustus 2025)

- Het applicatie programma kan de directory voor de logfiles instellen (`LogDirectory`). Zie `XpressNet Basic Test` voor een voorbeeld.
- De FIFO voor XpressNet berichten is vergroot van 32 naar 128 plaatsen.
- De tijd tussen twee wissel commando's (`AccessoryInterval`) is veranderd van 250 naar 125 ms.

## Versie 1.1 (december 2024)

- `CmdSwitchInfoRequest` is toegevoegd, om de wisselinformatie op te vragen. `CmdFeedbackRequest` is nu bedoeld voor (RS-Bus) terugmelddecoders.
- De voorbeelden vragen de juiste terugmeldgegevens op, gebruiken de standaard IP adressen van Lenz en Roco, en stand 1 is nu rechtdoor.

## Versie 1.0 (november 2024)

- Eerste versie.
