# Voorbeelden (Linux, Qt)
Deze directory bevat voorbeeld applicaties in C++ (Qt 6 Widgets) die de Linux library XpressNetQt gebruiken. Het zijn dezelfde voorbeelden als bij [.NET](../../DotNet/Examples/Voorbeelden.md): de indeling van de forms (de .ui files) is omgezet van de .NET forms, en de code volgt de .NET code. De .ui files kunnen met Qt Designer (of Qt Creator) aangepast worden.

Er zijn geen gecompileerde programma's: de voorbeelden worden vanuit de source code gebouwd, samen met de library. Zie de [installatie handleiding](../Code/Install/Installeren-Library.md) voor de pakketten die nodig zijn. In de directory Linux:
```
cmake -S . -B build
cmake --build build
"./build/Examples/XpressNet Basic Test/XpressNet Basic Test"
```
De namen van de programma's zijn in alle versies van de library (VB6, .NET, macOS en Linux) gelijk; vanwege de spaties moet een naam in de terminal tussen aanhalingstekens staan. Ieder voorbeeld kan ook los gebouwd worden, met de CMakeLists.txt in zijn directory (bijvoorbeeld `cmake -S "Examples/RS-Bus Monitor" -B build-rs`); de library wordt dan ook gebouwd.

Verschillen met de .NET voorbeelden:
- De USB poort (en de seriële poort) is de naam van het seriële device, bijvoorbeeld /dev/ttyUSB0. De lijst bevat de seriële poorten die Qt vindt; een andere naam kan ingetypt worden.
- De logfiles worden geschreven in ~/TMC-Logfiles.
- De lampjes zijn een kleine widget (Led.h / Led.cpp, in ieder voorbeeld). Een Led kan ook in de code gemaakt worden, bijvoorbeeld voor een raster met veel lampjes.

## XpressNet Basic Test ##
Het basis gebruik van de library: verbinden via TCP/IP of USB, versie informatie opvragen, wissels schakelen (met automatische deactivering), wissel- en RS-Bus terugmeldingen opvragen, en locomotieven en hun functies (F0 t/m F68) besturen. Daarnaast laat het zien hoe de applicatie de directory voor de logfiles kan instellen. Meldingen van de library (SendError, interface nicht / wieder adressiert) staan in de statusregel.

## XpressNet Test ##
Met dit test programma kunnen een groot aantal XpressNet commando's worden getest (ook seriële verbindingen, modeltijd, programmeren en de erweiterte Versionsinformation), en bekeken worden welke responses door de centrale worden verstuurd.

## Z21 Basic Test ##
Met dit test programma kunnen een beperkt aantal Z21 commando's worden getest, en bekeken worden welke responses door de centrale worden verstuurd. Na "Set" zet de library de broadcast flags van de Z21; daardoor werken de status lampjes (NotAus, NotHalt, Programmier Mode) en de wisselmeldingen.

## XpressNet Multiple Instances ##
Het is mogelijk om meerdere XpressNet objecten te maken binnen één applicatie. Iedere klik op "Add XpressNet" of "Add Z21" opent een nieuw form (XpressNet Test of Z21 Basic Test) met een eigen XpressNet object. In de logfile heeft ieder XpressNet object een eigen letter (A, B, …). Het sluiten van het start form sluit ook de andere forms.

## RS-Bus Monitor ##
Toont de terugmeld informatie van alle 128 RS-Bus adressen: 4 blokken van 32 rijen, iedere rij met 8 lampjes (twee nibbles van 4 bits) en een klein vierkantje voor het type van de laatste terugmelding (blauw: terugmeld decoder, groen: wissel decoder met terugmelding, rood: informatie van de centrale). "Get RS-Bus data" vraagt beide nibbles van alle 128 adressen op, in twee delen van 128 verzoeken (de FIFO van de library bevat 128 berichten); het tweede deel wordt opgevraagd zodra het antwoord op het laatste verzoek van het eerste deel ontvangen is. Daarnaast rekent het programma een RS-Bus adres om naar wissel adressen, en omgekeerd. De matrix wordt getekend door PanelManager (PanelManager.cpp, zoals bij VB6 en .NET). Het IP adres wordt bewaard voor de volgende start, in ~/.config/RS-Bus Monitor/Settings.conf (QSettings).
