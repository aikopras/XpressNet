# Voorbeelden (Qt: Linux)
Deze directory bevat voorbeeld applicaties in C++ (Qt 6 Widgets) die de library XpressNetQt gebruiken. Het zijn dezelfde voorbeelden als bij [.NET](../DotNet/Voorbeelden.md): de indeling van de forms (de .ui files) is omgezet van de .NET forms, en de code volgt de .NET code. De .ui files kunnen met Qt Designer (of Qt Creator) aangepast worden. Dezelfde source code wordt voor Windows gebruikt, zie [Voorbeelden-Windows.md](Voorbeelden-Windows.md).

De gecompileerde programma's voor Linux (x86-64, Qt 6) staan in [Examples/compiled/Linux-Qt](../../compiled/Linux-Qt), één zip file per voorbeeld; Qt 6 moet geïnstalleerd zijn (zie de [installatie handleiding](../../../Library/src/Qt/Install/Installeren-Library.md)). De namen van de programma's zijn in alle versies van de library (VB6, .NET, macOS, Windows en Linux) gelijk; vanwege de spaties moet een naam in de terminal tussen aanhalingstekens staan.

## Bouwen ##
Nodig: Qt 6 (Widgets, Network, SerialPort), CMake en een C++17 compiler; zie de [installatie handleiding](../../../Library/src/Qt/Install/Installeren-Library.md) voor de packages. Voor het alles-in-één script ook `zip` (`sudo apt install zip`).

Pak de complete repository uit (of clone hem) en laat de directory structuur intact. Alleen de directory met de voorbeelden kopiëren werkt niet: de voorbeelden halen de library source uit `Library/src/Qt/XpressNetQt`, en de RS-Bus Decoder Scanner haalt zijn CV overzicht uit `Examples/src/macOS`. Beide worden relatief ten opzichte van de directory met de voorbeelden gevonden.

Er zijn twee manieren om te bouwen:

| | Alles in één | Per voorbeeld |
|---|---|---|
| Hoe | [build-linux.sh](build-linux.sh) in een terminal | commando's in een terminal |
| Wat | de library en alle zes voorbeelden, altijd helemaal opnieuw | één voorbeeld (en de library) |
| Resultaat | een zip file per voorbeeld in `Examples/compiled/Linux-Qt` | het programma in de directory van het voorbeeld |
| Gebruik voor | alles (opnieuw) bouwen, bijvoorbeeld voor een nieuwe release | werken aan één voorbeeld, of uitzoeken welke stap mislukt |

### Alles in één: build-linux.sh ###
In de directory `Examples/src/Qt`:

    ./build-linux.sh

(of `sh build-linux.sh`, als de file na het uitpakken niet uitvoerbaar is). Het script controleert of de library source, `cmake` en `zip` gevonden worden, verwijdert de directory `build`, bouwt de library en alle voorbeelden (Release), en maakt `Examples/compiled/Linux-Qt/<voorbeeld>.zip`. De programma's kunnen ook vanuit `build/<voorbeeld>/` gestart worden, bijvoorbeeld `"./build/RS-Bus Decoder Scanner/RS-Bus Decoder Scanner"`.

### Per voorbeeld ###
In de directory van het voorbeeld, bijvoorbeeld:

    cd "Examples/src/Qt/RS-Bus Decoder Scanner"
    cmake -S . -B build
    cmake --build build
    "./build/RS-Bus Decoder Scanner"

De library wordt ook gebouwd. Alle voorbeelden kunnen ook samen gebouwd worden, zonder zip files, in de directory `Examples/src/Qt`, met dezelfde `cmake` commando's; de programma's staan dan in `build/<voorbeeld>/`. Meldt CMake dat `../../../../Library/src/Qt/XpressNetQt` geen bestaande directory is, dan is de directory structuur niet compleet.

## Verschillen met de .NET voorbeelden ##
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

## RS-Bus Decoder Scanner ##
Qt versie van het macOS voorbeeld: zoekt de decoders met RS-Bus terugmelding op maximaal 5 centrales (Lenz LAN/USB interfaces), en leest via PoM hun type, software versie en andere CVs, die in de preferences gekozen kunnen worden (File → Preferences…). De resultaten staan in een tabel, die als CSV file bewaard kan worden. De lift decoder, waarvan het PoM adres niet uit het RS-Bus adres volgt, kan via de preferences meegenomen worden. Dit voorbeeld laat zien hoe meerdere XpressNet objecten (één per centrale) gebruikt worden: de signals van ieder XpressNet object zijn verbonden met lambda functies die de centrale toevoegen (DecoderScanner.cpp). Het CV overzicht (Help → CV overview…) is dezelfde HTML file als in de macOS versie. De instellingen worden met QSettings bewaard, in `~/.config/RS-Bus Decoder Scanner/Settings.conf`.
