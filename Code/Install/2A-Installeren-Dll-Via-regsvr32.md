# Download en Installatie van de XpressNet DLL (regsvr32)

## Step 1: Make een folder ##

Als eerste wordt de File Explorer geopend, en een nieuwe folder gemaakt met de naam XpressNetDll. De locatie en naam van deze folder mag later niet meer veranderen. Een goede plaats is bijvoorbeeld onmiddellijk onder C:.
<br/><img src="Download/01-Create_Directory.png" alt="Maak een folder" width="500"/>
<br/><br/>

## Step 2: Ga naar de GitHub repository ##
Open de browser en ga naar de [XpressNet repository op Gihub](https://github.com/aikopras/XpressNet).
<br/><img src="Download/02-Goto_GitHub.png" alt="Ga naar GitHub" width="500"/>
<br/><br/>

## Step 3: Dubbel klik op XpressNet.dll ##
Navigeer naar de folder met de naam Code en dubbel klik op de file XpressNet.dll
<br/><img src="Download/03-2A-Double_Click.png" alt="klik XpressNet.dll" width="500"/>
<br/><br/>

## Step 4: Download als RAW file ##
Klik op de download RAW file icon aan de rechterkant om XpressNet.dll naar je lokale machine te downloaden.
<br/><img src="Download/04-2A-Download_raw_file.png" alt="Download RAW" width="500"/>
<br/><br/>

## Step 5: Save XpressNet.dll ##
Nadat XpressNet.dll is gedownload (naar je lokale Downloads folder), verschuif de dll file naar de folder die je hebt gemaakt under stap 1 hierboven ((b.v. C:/XpressNetDll).
<br/><img src="Regsvr32/01-Move_DLL-File.png" alt="Copy Files" width="600"/>
<br/><br/>

## Step 6: Open de Command prompt als Admin ##
Nadat de DLL file is opgeslagen, wordt het tijd de DLL te registreren op de locale machine. Hiertoe moet de Windows Command Prompt met Administrator rechten worden opgestart.
Klik Windows Start -> Microsoft System, gevolgd door een rechter muis klik op Command Prompt. Een nieuw venster verschijnt en onder More kan worden gekozen voor Run as Administrator.
<br/><img src="Regsvr32/02-Command_Prompt.png" alt="Open Cmd as Admin" width="600"/>
<br/><br/>

## Step 7: Bevestig dat wijzigingen zijn toegestaan ##
Afhankelijk van de versie van windows, wordt er een nieuw venster geopend, met daarin de vraag of de Windows Command Processor veranderingen mag aanbrengen op je locale computer. Klik YES.
<br/><img src="Regsvr32/03-Allow_Cmd_As_Admin.png" alt="Allow Cmd as Admin" width="400"/>
<br/><br/>

## Step 8: Ga naar de DLL folder ##
Ga naar de directory waar de XpressNet.dll is opgeslagen (cd c:\XpressnetDll). Deze folder mag later niet verplaats worden of een andere naam krijgen..
<br/><img src="Regsvr32/04-Cmd_regsvr32.png" alt="move dir" width="500"/>
<br/><br/>

## Step 9: Start regsvr32 ##
Start regsvr32, met XpressNet.dll als argument.
<br/><img src="Regsvr32/04-Detail_regsvr32.png" alt="Register" width="400"/>
<br/>
Als alles goed gaat, verschijnt er een window die zegt dat registratie van XpressNet.dll is gelukt.
<br/><img src="Regsvr32/05-Succeeded.png" alt="Registration Succeeded" width="400"/>
<br/>
Als je later de DLL weer uit de windows registry wilt verwijderen, omdat er bijvoorbeeld een nieuwe versie is verschenen, moet regsvr32 opnieuw worden aangeroepen, maar met /u als switch.
<br/><img src="Regsvr32/06-unregister.png" alt="Unregister" width="400"/>
<br/><br/>

## Step 10: Reference de XpressNet DLL ##
Voordat de XpressNet DLL library in je eigen programma gebruikt kan worden, moet je vanuit je programma een referentie leggen naar de DLL. Nadat je eigen programma is geopend, klik op Project -> References, en selecteer het XpressNet Interface. Zie beide onderstaande plaatjes.
<br/><img src="Reference/01-OpenPreferences.png" alt="Open Project -> Preferences" width="350"/>
<br/><br/>
<br/><img src="Reference/02-SelectXpressNet.png" alt="Select XpressNet" width="500"/>
<br/><br/>
