# Voorbeelden (Qt: Windows)
Deze directory bevat voorbeeld applicaties in C++ (Qt 6 Widgets) die de library XpressNetQt gebruiken. Het zijn dezelfde voorbeelden als bij [.NET](../DotNet/Voorbeelden.md): de indeling van de forms (de .ui files) is omgezet van de .NET forms, en de code volgt de .NET code. De .ui files kunnen met Qt Designer (of Qt Creator) aangepast worden. Dezelfde source code wordt voor Linux gebruikt, zie [Voorbeelden-Linux.md](Voorbeelden-Linux.md).

De gecompileerde programma's voor Windows (x64) staan in [Examples/compiled/Windows-Qt](../../compiled/Windows-Qt), één zip file per voorbeeld, met de Qt DLLs; uitpakken en de .exe starten. De programma's hebben de Microsoft Visual C++ Redistributable (x64) nodig, die op de meeste computers geïnstalleerd is.

## Bouwen ##
Nodig: Visual Studio 2019 (workload "Desktop development with C++", met de component "C++ CMake tools for Windows") en Qt 6.7 voor MSVC 2019 (Qt Online Installer: Qt 6.7.3 → MSVC 2019 64-bit en Additional Libraries → Qt Serial Port, in `C:\Qt`). Met Visual Studio 2022 kan op dezelfde manier Qt 6.8 of nieuwer gebruikt worden (`msvc2022_64` in plaats van `msvc2019_64`).

Pak de complete repository uit, bijvoorbeeld naar `C:\Users\<naam>\Desktop\XpressNet`, en laat de directory structuur intact. Alleen de directory met de voorbeelden kopiëren werkt niet: de voorbeelden halen de library source uit `Library\src\Qt\XpressNetQt`, en de RS-Bus Decoder Scanner haalt zijn CV overzicht uit `Examples\src\macOS`. Beide worden relatief ten opzichte van de directory met de voorbeelden gevonden.

Er zijn twee manieren om te bouwen:

| | Alles in één | Per voorbeeld |
|---|---|---|
| Hoe | dubbelklik op [build-windows.bat](build-windows.bat) | commando's in een command prompt |
| Wat | de library en alle zes voorbeelden, altijd helemaal opnieuw | één voorbeeld (en de library) |
| Resultaat | een zip file per voorbeeld in `Examples\compiled\Windows-Qt` | het programma in de directory van het voorbeeld |
| Gebruik voor | alles (opnieuw) bouwen, bijvoorbeeld voor een nieuwe release | werken aan één voorbeeld, of uitzoeken welke stap mislukt |

### Alles in één: build-windows.bat ###
Dubbelklik op `build-windows.bat` in `Examples\src\Qt`. De batch file:
1. controleert of de library source, Qt, Visual Studio en de CMake daarvan gevonden worden (zo niet, dan meldt hij wat ontbreekt);
2. verwijdert de directories `build` en `deploy`, en bouwt de library en alle voorbeelden (Release);
3. kopieert ieder programma naar `deploy\<voorbeeld>`, voegt met `windeployqt` de Qt DLLs toe, en maakt `Examples\compiled\Windows-Qt\<voorbeeld>.zip`.

Het venster blijft aan het eind open, zodat de meldingen gelezen kunnen worden. De programma's kunnen ook vanuit `deploy\<voorbeeld>` gestart worden. De paden van Qt en Visual Studio staan in drie regels bovenaan de batch file (`QTDIR`, `GENERATOR` en `VSVERSION`); pas die aan voor een andere versie van Qt of Visual Studio.

### Per voorbeeld ###
Open "x64 Native Tools Command Prompt for VS 2019" (Start menu, Visual Studio 2019), en ga naar de directory van het voorbeeld, bijvoorbeeld:

    cd "C:\Users\<naam>\Desktop\XpressNet\Examples\src\Qt\RS-Bus Decoder Scanner"
    dir ..\..\..\..\Library\src\Qt\XpressNetQt

Het `dir` commando moet `CMakeLists.txt` laten zien; meldt het "The system cannot find the path specified", dan is de directory structuur niet compleet. Bouw het voorbeeld (de library wordt ook gebouwd):

    cmake -S . -B build -G "Visual Studio 16 2019" -A x64 -DCMAKE_PREFIX_PATH=C:\Qt\6.7.3\msvc2019_64
    cmake --build build --config Release

De melding "Could NOT find WrapVulkanHeaders" kan genegeerd worden (Vulkan is optioneel in Qt en wordt niet gebruikt). Als het eerste `cmake` commando mislukt is, verwijder dan eerst de `build` directory (`rmdir /s /q build`) voordat je het opnieuw probeert.

Het programma staat nu in `build\Release`, bijvoorbeeld `build\Release\RS-Bus Decoder Scanner.exe`. Het heeft de Qt DLLs nodig; `windeployqt` kopieert die naast een kopie van het programma, in `deploy`:

    mkdir "deploy\RS-Bus Decoder Scanner"
    copy /y "build\Release\RS-Bus Decoder Scanner.exe" "deploy\RS-Bus Decoder Scanner"
    C:\Qt\6.7.3\msvc2019_64\bin\windeployqt.exe --release --no-translations --no-opengl-sw --no-system-d3d-compiler --no-compiler-runtime "deploy\RS-Bus Decoder Scanner\RS-Bus Decoder Scanner.exe"

Het programma kan daarna vanuit `deploy\RS-Bus Decoder Scanner` gestart worden. Om de zip file te maken (zoals de batch file doet; `del` en `mkdir` kunnen melden dat een file niet gevonden is of al bestaat, dat kan genegeerd worden):

    del "deploy\RS-Bus Decoder Scanner\dxcompiler.dll" "deploy\RS-Bus Decoder Scanner\dxil.dll"
    mkdir ..\..\..\compiled\Windows-Qt
    tar -a -c -f "..\..\..\compiled\Windows-Qt\RS-Bus Decoder Scanner.zip" -C deploy "RS-Bus Decoder Scanner"

De DirectX shader compilers (dxcompiler.dll, dxil.dll) die windeployqt kan toevoegen worden door de programma's niet gebruikt. Voor de andere voorbeelden zijn de commando's hetzelfde, met hun eigen naam.

## Verschillen met de .NET voorbeelden ##
- De USB poort (en de seriële poort) is de naam van de COM poort, bijvoorbeeld COM3. De lijst bevat de seriële poorten die Qt vindt; een andere naam kan ingetypt worden.
- De logfiles worden geschreven in `C:\Users\<naam>\TMC-Logfiles` (de voorbeelden stellen die directory in; de standaard directory van de library is `C:\Users\Public\TMC-Logfiles`, zoals bij .NET).
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
Toont de terugmeld informatie van alle 128 RS-Bus adressen: 4 blokken van 32 rijen, iedere rij met 8 lampjes (twee nibbles van 4 bits) en een klein vierkantje voor het type van de laatste terugmelding (blauw: terugmeld decoder, groen: wissel decoder met terugmelding, rood: informatie van de centrale). "Get RS-Bus data" vraagt beide nibbles van alle 128 adressen op, in twee delen van 128 verzoeken (de FIFO van de library bevat 128 berichten); het tweede deel wordt opgevraagd zodra het antwoord op het laatste verzoek van het eerste deel ontvangen is. Daarnaast rekent het programma een RS-Bus adres om naar wissel adressen, en omgekeerd. De matrix wordt getekend door PanelManager (PanelManager.cpp, zoals bij VB6 en .NET). Het IP adres wordt bewaard voor de volgende start, in de registry, `HKEY_CURRENT_USER\Software\RS-Bus Monitor\Settings` (QSettings).

## RS-Bus Decoder Scanner ##
Qt versie van het macOS voorbeeld: zoekt de decoders met RS-Bus terugmelding op maximaal 5 centrales (Lenz LAN/USB interfaces), en leest via PoM hun type, software versie en andere CVs, die in de preferences gekozen kunnen worden (File → Preferences…). De resultaten staan in een tabel, die als CSV file bewaard kan worden. De lift decoder, waarvan het PoM adres niet uit het RS-Bus adres volgt, kan via de preferences meegenomen worden. Dit voorbeeld laat zien hoe meerdere XpressNet objecten (één per centrale) gebruikt worden: de signals van ieder XpressNet object zijn verbonden met lambda functies die de centrale toevoegen (DecoderScanner.cpp). Het CV overzicht (Help → CV overview…) is dezelfde HTML file als in de macOS versie. De instellingen worden met QSettings bewaard, in de registry: `HKEY_CURRENT_USER\Software\RS-Bus Decoder Scanner\Settings`.
