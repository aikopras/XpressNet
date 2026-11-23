# Voorbeelden (macOS)
Deze directory bevat vijf Xcode projecten (Objective-C) die de macOS versie van de XpressNet library gebruiken. Het zijn macOS versies van de VB6 voorbeelden. Ieder project bevat een kopie van de library (de directories XpressNet en Communication).

De gecompileerde programma's staan in [Examples/compiled/macOS](../../compiled/macOS), één zip file per voorbeeld (universal: Apple silicon en Intel, macOS 12.4 of nieuwer). De namen van de programma's zijn in alle versies van de library (VB6, .NET, macOS en Linux) gelijk. De programma's zijn niet door Apple genotariseerd. De eerste keer weigert macOS een programma te openen: kies dan in Systeeminstellingen → Privacy en beveiliging de knop om het programma toch te openen ("Open Anyway"; tot en met macOS 14: rechter klik op het programma → Open). Daarna start het gewoon.

## XpressNet Test ##
macOS versie van het VB6 voorbeeld XpressNet Test: een groot aantal XpressNet commando's, en de antwoorden van de centrale. Meldingen van de library (SendError, interface nicht / wieder adressiert) staan in de statusregel.

## XpressNet Basic Test ##
macOS versie van het VB6 voorbeeld XpressNet Basic Test: verbinding maken via TCP/IP of USB (met de optie Handshake), versie informatie, wissels schakelen, wissel en RS-Bus terugmeldingen, en locomotieven en hun functies. De logfiles worden geschreven in ~/Library/Logs/TMC-Logfiles. Meldingen van de library (SendError, interface nicht / wieder adressiert) staan in de statusregel.

## Z21 Basic Test ##
macOS versie van het VB6 voorbeeld Z21 Basic Test: een beperkt aantal Z21 commando's, en de antwoorden van de centrale. Het event SendError staat in de statusregel.

## RS-Bus Monitor ##
Toont de terugmeld informatie van alle 128 RS-Bus adressen: 4 blokken van 32 rijen, iedere rij met 8 lampjes (twee nibbles van 4 bits) en een klein vierkantje voor het type van de laatste terugmelding (blauw: terugmeld decoder, groen: wissel decoder met terugmelding, rood: informatie van de centrale). "Get RS-Bus data" vraagt beide nibbles van alle 128 adressen op, in twee delen van 128 verzoeken (de FIFO van de library bevat 128 berichten); het tweede deel wordt opgevraagd zodra het antwoord op het laatste verzoek van het eerste deel ontvangen is. Daarnaast rekent het programma een RS-Bus adres om naar wissel adressen, en omgekeerd. Het IP adres wordt ingesteld in de voorkeuren (RS-Bus Monitor → Preferences…), en wordt bewaard voor de volgende start. Vanaf versie 2.0 gebruikt het programma de library, in plaats van eigen TCP code.

## RS-Bus Decoder Scanner ##
Zoekt de decoders met RS-Bus terugmelding op maximaal 5 centrales (Lenz LAN/USB interfaces), en leest via PoM hun type, software versie en andere CVs, die in de preferences gekozen kunnen worden. De resultaten staan in een tabel, die als CSV file bewaard kan worden. Dit voorbeeld laat zien hoe meerdere XpressNet objecten (één per centrale) gebruikt worden: omdat de events van de library niet het object bevatten dat ze veroorzaakte, heeft ieder XpressNet object een eigen delegate (StationClass in DecoderScanner.m). Het gecompileerde programma staat in [Examples/compiled/macOS/RS-Bus Decoder Scanner.zip](../../compiled/macOS/RS-Bus%20Decoder%20Scanner.zip).
