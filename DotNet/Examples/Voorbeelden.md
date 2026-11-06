# Voorbeelden (.NET)
Deze directory bevat voorbeeld applicaties in VB.NET (Visual Studio 2019, .NET Framework 4.7.2) die de .NET library XpressNet.NET.dll gebruiken. Het zijn dezelfde voorbeelden als bij [VB6](../../VB6/Examples/Voorbeelden.md): de indeling van de forms is overgenomen van de VB6 forms (waar nodig iets ruimer, omdat de teksten onder .NET meer ruimte nodig hebben), en de code volgt de VB6 code.

De projecten verwijzen naar de DLL in de directory DotNet/Code (Project → References → XpressNet.NET). Na het downloaden van de complete repository kunnen de voorbeelden direct gebouwd en gestart worden; het registreren van de DLL is niet nodig.

De gecompileerde programma's (.exe en .exe.config) staan, samen met de DLL, in [Examples-Compiled/DotNet_Examples.zip](../Examples-Compiled/DotNet_Examples.zip). De namen van de programma's zijn in alle versies van de library (VB6, .NET, macOS en Linux) gelijk. Het programma, zijn .exe.config en de DLL moeten op de productie machine in dezelfde directory worden gezet, waarna de applicatie gestart kan worden. Nodig is .NET Framework 4.7.2 of nieuwer (standaard aanwezig in actuele versies van Windows 10 en 11).

## XpressNet Basic Test ##
Het basis gebruik van de library: verbinden via TCP/IP of USB, versie informatie opvragen, wissels schakelen (met automatische deactivering), wissel- en RS-Bus terugmeldingen opvragen, en locomotieven en hun functies (F0 t/m F68) besturen. Daarnaast laat het zien hoe de applicatie de directory voor de logfiles kan instellen. Meldingen van de library (SendError, interface nicht / wieder adressiert) staan in de statusregel.

## XpressNet Test ##
Met dit test programma kunnen een groot aantal XpressNet commando's worden getest (ook seriële verbindingen, modeltijd, programmeren en de erweiterte Versionsinformation), en bekeken worden welke responses door de centrale worden verstuurd.

## Z21 Basic Test ##
Met dit test programma kunnen een beperkt aantal Z21 commando's worden getest, en bekeken worden welke responses door de centrale worden verstuurd. Na "Set" zet de library de broadcast flags van de Z21; daardoor werken de status lampjes (NotAus, NotHalt, Programmier Mode) en de wisselmeldingen.

## XpressNet Multiple Instances ##
Het is mogelijk om meerdere XpressNet objecten te maken binnen één applicatie. Iedere klik op "Add XpressNet" of "Add Z21" opent een nieuw form (XpressNet Test of Z21 Basic Test) met een eigen XpressNet object. In de logfile heeft ieder XpressNet object een eigen letter (A, B, …).

## RS-Bus Monitor ##
Toont de terugmeld informatie van alle 128 RS-Bus adressen: 4 blokken van 32 rijen, iedere rij met 8 lampjes (twee nibbles van 4 bits) en een klein vierkantje voor het type van de laatste terugmelding (blauw: terugmeld decoder, groen: wissel decoder met terugmelding, rood: informatie van de centrale). "Get RS-Bus data" vraagt beide nibbles van alle 128 adressen op, in twee delen van 128 verzoeken (de FIFO van de library bevat 128 berichten); het tweede deel wordt opgevraagd zodra het antwoord op het laatste verzoek van het eerste deel ontvangen is. Daarnaast rekent het programma een RS-Bus adres om naar wissel adressen, en omgekeerd. Het IP adres wordt bewaard voor de volgende start, met SaveSetting / GetSetting (dezelfde instelling als bij het VB6 programma).

## XpressNet Regfree Test en XpressNet Project Group Test ##
Deze VB6 voorbeelden hebben geen .NET versie: de .NET library hoeft niet geregistreerd te worden, en het library project (DotNet/Code/XpressNet.NET/XpressNet.NET.csproj) kan aan de solution van de applicatie worden toegevoegd (zie de [installatie handleiding](../Code/Install/Installeren-Library.md)).
