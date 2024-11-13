# Download en Installatie van de XpressNet DLL

## Step 1: Make een folder ##

Als eerste wordt de File Explorer geopend, en een nieuwe folder gemaakt met de naam XpressNetDll. De locatie en naam van deze folder mag later niet meer veranderen. Een goede plaats is bijvoorbeeld onmiddellijk onder C:.
<br/><img src="Download/01-Create_Directory.png" alt="Maak een flder" width="500"/>
<br/><br/>

## Step 2: Ga naar de GitHub repository ##
Open de browser en ga naar de [XpressNet repository op Gihub](https://github.com/aikopras/XpressNet).
<br/><img src="Download/02-Goto_GitHub.png" alt="Ga naar GitHub" width="500"/>
<br/><br/>

## Step 3: Dubbel klik op VB6_Source.zip ##
Navigeer naar de folder met de naam Code en dubbel klik op de file VB6_Source.zip
<br/><img src="Download/03-Double_Click.png" alt="klik VB6_Source.zip" width="500"/>
<br/><br/>

## Step 4: Download als RAW file ##
Klik op de download RAW file icon aan de rechterkant om VB6_Source.zip naar je lokale machine te downloaden.
<br/><img src="Download/04-Download_raw_file.png" alt="Download RAW" width="500"/>
<br/><br/>

## Step 5: Open VB6_Source.zip ##
Nadat VB6_Source.zip is gedownload (naar je lokale Downloads folder), klik Open file om VB6_Source.zip te openen.
<br/><img src="Download/05-Open_VB6_Source.png" alt="Open VB6 Source" width="500"/>
<br/><br/>

## Step 6: Kopieer alle files ##
Een nieuw venster opent, en toont alle files van VB6_Source.zip. Selecteer alle files, en kopieer ze in de folder die je onder Stap 1 hebt gemaakt (bijvoorbeeld C:/XpressNetDll).
<br/><img src="Download/06-Copy_Files.png" alt="Kopieer files" width="600"/>
<br/><br/>

## Step 7: Open VB6 als Administrator ##
Nadat alle files zijn gedownload en in de juiste folder opgeslagen, is het tijd de DLL code te vertalen en te installeren op je eigen computer. Om dit te kunnen doen, MOET VB6 als Administrator geopend worden, zodat het mag schrijven in de Windows Registry.
Klik Windows Start -> Microsoft Visual Studio, gevolgd door een rechter muis klik op Microsft Visual Basic 6. Een nieuw menu opent zich, en onder More wordt gekozen voor Run as Administrator.
<br/><img src="Make/01-Run_VB6_AsAdmin.png" alt="Open VB6 als Admin" width="600"/>
<br/><br/>

## Step 8: VB6 als Administrator toestaan ##
Afhankelijk van de versie van windows die je hebt, kan een nieuw venster geopend worden waarin gevraagd wordt of VB6 toegestaan wordt om veranderingen (in de registry) van je computer uit te voeren. Selecteer YES.
<br/><img src="Make/02-Allow_VB6_AsAdmin.png" alt="VB6 als Admin toestaan" width="400"/>
<br/><br/>

## Step 9: VB6 wordt geopend als Nieuw Project ##
Afhankelijk van je VB6 preferences, kan een window geopend worden waarin gevraagd wordt of je een nieuw VB6 project wilt openen.. Kies Cancel.
<br/><img src="Make/03-Cancel_New_Project.png" alt="Cancel nieuw project" width="500"/>
<br/><br/>

## Step 10: Open het XpressNet project ##
Open het VB6 DLL project (XpressNet_vbp), welke je in Stap 6 hierboven hebt opgeslagen (in C:/XpressNetDll).
<br/><img src="Make/04-OpenXpressNet_vbp.png" alt="Open XpressNet.vbp" width="500"/>
<br/><br/>

## Step 11: Make XpressNet.dll ##
Klik File -> Make XpressNet.dll, om de DLL code te vertalen.
<br/><img src="Make/05-MakeDll.png" alt="Make XpressNet.dll" width="500"/>
<br/><br/>

## Step 12: Save en registreer de DLL ##
Na de vorige stap wordt gevraagd waar de DLL file moet worden opgeslagen. Kies de folder waarin ook de source files zijn opgeslagen, en die later niet meer wordt verplaatst. Onzichtbaar voor de gebruiker zal VB6 ook de DLL in de Windows Registry registreren. Hierdoor kunnen alle VB6 programma's van de DLL gebruik maken. Je bent nu klaar met het installeren van de DLL.
<br/><img src="Make/06-SaveDll.png" alt="Save XpressNet.dll" width="500"/>
<br/><br/>

## Step 13: Reference de XpressNet DLL ##
Voordat de XpressNet DLL library in je eigen programma gebruikt kan worden, moet je vanuit je programma een referentie leggen naar de DLL. Nadat je eigen programma is geopend, klik op Project -> References, en selecteer het XpressNet Interface. Zie beide onderstaande plaatjes.
<br/><img src="Reference/01-OpenPreferences.png" alt="Open Project -> Preferences" width="350"/>
<br/><br/>
<br/><img src="Reference/02-SelectXpressNet.png" alt="Select XpressNet" width="500"/>
<br/><br/>
