# Voorbeelden
In deze directory staan een aantal voorbeeld applicaties die bepaalde elementen van het gebruik van de XpressNet library tonen.

Behalve bij de laatste test, gaan alle overige test programma's er van uit dat de XpressNet DLL in de windows registry is geïnstalleerd. Zie de installatie handleiding hoe deze registratie kan plaatsvinden. Vergeet niet een Reference naar de DLL te leggen!

De voorbeelden zijn voor versie 2.0 van de DLL. De meldingen van de library (`SendError`, en bij de XpressNet voorbeelden ook interface nicht / wieder adressiert) staan in de statusregel; de Z21 voorbeelden hebben daarvoor een nieuwe statusregel. Bij USB is er een vakje Handshake. De logfiles staan in `C:\Users\Public\TMC-Logfiles`.

De gecompileerde programma's staan in [Examples-Compiled/VB6_Examples.zip](../Examples-Compiled/VB6_Examples.zip), samen met de DLL. De namen van de programma's zijn in alle versies van de library (VB6, .NET, macOS en Linux) gelijk. De voorbeelden gebruiken de geregistreerde DLL, behalve XpressNet Regfree Test (de DLL in dezelfde directory) en XpressNet Project Group Test (een eigen DLL); zie hieronder.

## XpressNet Basic Test ##
Dit test programma toont het basis gebruik van de XpressNet DLL: verbinden via TCP/IP of USB, versie informatie opvragen, wissels schakelen (met automatische deactivering), wissel- en RS-Bus terugmeldingen opvragen, en locomotieven en hun functies (F0 t/m F68) besturen. Daarnaast laat het zien hoe de applicatie de directory voor de logfiles kan instellen.

## XpressNet Test ##
Met dit test programma kunnen een groot aantal XpressNet commando's worden getest, en bekeken worden welke responses door de centrale worden verstuurd.
<br/><img src="XpressNet%20Test/A_Screenshot.png" alt="XpressNet Test" width="700"/>
<br/><br/>

## Z21 Basic Test ##
Met dit test programma kunnen een beperkt aantal Z21 commando's worden getest, en bekeken worden welke responses door de centrale worden verstuurd. Na "Set" zet de library de broadcast flags van de Z21; daardoor werken de status lampjes (NotAus, NotHalt, Programmier Mode) en de wisselmeldingen.
<br/><img src="Z21%20Basic%20Test/A_Screenshot.png" alt="Z21 Basic Test" width="600"/>
<br/><br/>

## XpressNet Multiple Instances ##
Het is mogelijk om meerdere XpressNet objecten te instantiëren binnen één applicatie. Dit test programma demonstreert dit. In de logfile heeft ieder XpressNet object een eigen letter (A, B, …).
<br/><img src="XpressNet%20Multiple%20Instances/A_Screenshot.png" alt="XpressNet Multiple Instances" width="700"/>
<br/><br/>

## XpressNet Regfree Test ##
Bij de meeste voorbeeld programma's is verondersteld dat de XpressNet DLL niet alleen op de ontwikkel, maar ook op de productie machines (in de Windows Registry) is geregistreerd. Om deze registratie te kunnen doen, zijn administrator rechten nodig op iedere machine. Op productie machines is het soms onmogelijk om administrator rechten te verkrijgen. In dergelijke gevallen kan de applicatie ook "Reg-Free" geschreven worden.

Hiertoe dient op de ontwikkel omgeving de XpressNet DLL geregistreerd te zijn (dus op deze machine zijn wel administrator rechten nodig). Daarnaast moet de DLL file ook worden toegevoegd aan de directory met de source files, alsmede een speciale class file (XPNRegFree.cls). Deze class file dient aan het eigen programma te worden toegevoegd. Zie onderstaand plaatje.
<br/><img src="XpressNet%20Regfree%20Test/A_Screenshot-1.png" alt="Kopieer files" width="700"/>
<br/>

Daarnaast moet één regel in de source file worden aangepast. Zie onderstaand plaatje voor details. Deze techniek is beschreven in [deze bijdrage in het VB6 forum](https://www.vbforums.com/showthread.php?908980-VB6-and-DLLs-Runtime-Error-429-ActiveX-Component-Can-t-Create-Object); met dank aan VanGoghGaming.
<br/><img src="XpressNet%20Regfree%20Test/A_Screenshot-2.png" alt="Kopieer files" width="700"/>
<br/>

In VB6_Examples.zip staan het gecompileerde programma (XpressNet Regfree Test.exe) en de DLL in dezelfde directory. Deze twee files moeten op de productie machine (in dezelfde directory) worden gezet, waarna de applicatie gestart kan worden (zonder dat de DLL geregistreerd is).

## XpressNet Project Group Test ##
Dit voorbeeld laat zien hoe de applicatie en de XpressNet code samen kunnen genomen in één project groep. Het voordeel van deze aanpak is dat de XpressNet code geïntegreerd wordt binnen de applicatie. Dit kan handig zijn tijdens software ontwikkeling, omdat de DLL op dat moment nog niet geïnstalleerd hoeft te worden.

De groep compileert de XpressNet code tot een eigen DLL, in de subdirectory DLL (Project Compatibility met `DLL\XpressNet.dll`, zodat de GUID's gelijk blijven). Het gecompileerde programma (in VB6_Examples.zip de directory XpressNet Project Group Test) gebruikt deze DLL, en niet de DLL in C:\XpressNetDll. Op een andere machine moet deze DLL daarom eerst geregistreerd worden (regsvr32).
<br/><img src="XpressNet%20Project%20Group%20Test/A_Screenshot.png" alt="XpressNet Project Group Test" width="300"/>

## RS-Bus Monitor ##
Toont de terugmeld informatie van alle 128 RS-Bus adressen: 4 blokken van 32 rijen, iedere rij met 8 lampjes (twee nibbles van 4 bits) en een klein vierkantje voor het type van de laatste terugmelding (blauw: terugmeld decoder, groen: wissel decoder met terugmelding, rood: informatie van de centrale). "Get RS-Bus data" vraagt beide nibbles van alle 128 adressen op, in twee delen van 128 verzoeken (de FIFO van de library bevat 128 berichten); het tweede deel wordt opgevraagd zodra het antwoord op het laatste verzoek van het eerste deel ontvangen is. Daarnaast rekent het programma een RS-Bus adres om naar wissel adressen, en omgekeerd. Het IP adres wordt bewaard voor de volgende start (SaveSetting / GetSetting). Voor versie 2.0 bevatte dit programma een kopie van de classes van de library; nu gebruikt het de geregistreerde DLL.

## VB.NET (Visual Studio 2019) ##
De voorbeelden (behalve XpressNet Regfree Test en XpressNet Project Group Test) zijn er ook in VB.NET, met de .NET library XpressNet.NET.dll: zie de [.NET voorbeelden](../../DotNet/Examples/Voorbeelden.md). Een VB.NET applicatie kan de VB6 DLL nog steeds via COM gebruiken (Project → Add Reference…, COM), maar de .NET library is eenvoudiger te installeren: die hoeft niet geregistreerd te worden.
