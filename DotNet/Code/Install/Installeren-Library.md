# Installeren van de .NET library (XpressNet.NET)

XpressNet.NET is de .NET versie van de XpressNet library. Versie 2.0 heeft dezelfde functionaliteit als versie 2.0 van de VB6 DLL, de macOS versie en de Linux versie. De library kan gebruikt worden in VB.NET en C# applicaties (Windows Forms), met Visual Studio 2019 of nieuwer en .NET Framework 4.7.2 of nieuwer.

Vergeleken met de VB6 DLL is de .NET library eenvoudiger te installeren: de library hoeft niet geregistreerd te worden (geen regsvr32, geen administrator rechten), en er zijn geen ActiveX controls nodig (Winsock, MSComm). De library bestaat uit één file, XpressNet.NET.dll, die naast de applicatie (exe) wordt gezet.

## Optie 1: de DLL gebruiken
1. Download [XpressNet.NET.dll](../XpressNet.NET.dll) (en, om met de debugger in de library te kunnen kijken, [XpressNet.NET.pdb](../XpressNet.NET.pdb)). Klik op de file, en daarna op de download RAW file icon aan de rechterkant.
2. Kopieer de file(s) naar een directory van het project, bijvoorbeeld een directory met de naam Lib.
3. In Visual Studio: Project → Add Reference… → Browse… en selecteer XpressNet.NET.dll.

Bij het bouwen van het project kopieert Visual Studio XpressNet.NET.dll naar de directory van de exe (bin\Debug of bin\Release). Om de applicatie op een andere PC te installeren, kopieer de exe samen met XpressNet.NET.dll.

## Optie 2: de library bouwen vanuit de source code
De source code staat in de directory [XpressNet.NET](../XpressNet.NET). Open XpressNet.NET.sln in Visual Studio, kies Release en daarna Build → Build Solution. De DLL komt in bin\Release te staan. Het is ook mogelijk het project XpressNet.NET.csproj aan de solution van de applicatie toe te voegen (File → Add → Existing Project…), en een referentie naar dat project te maken (Project → Add Reference… → Projects).

## Gebruik van de library
Het gebruik van de library is beschreven in de handleiding van de VB6 DLL (Documentation/HowToUse). In VB.NET wordt het object als volgt gedeclareerd en worden de events gekoppeld:

```vb
Imports XpressNet

Public Class Form1
    Private XpressNet As XpressNetClass

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        XpressNet = New XpressNetClass()
        AddHandler XpressNet.InterfaceConnected, AddressOf InterfaceConnected
        AddHandler XpressNet.Feedback, AddressOf Feedback
        Call XpressNet.CmdConnectTcp("192.168.0.200", "5550")
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        XpressNet.Dispose()                     ' Sluit de verbinding
    End Sub

    Private Sub InterfaceConnected()
        Call XpressNet.CmdAccessoryAuto(12, 1)
    End Sub

    Private Sub Feedback(ByVal address As Integer, ByVal nibble As Byte, ByVal data As Byte)
        Label1.Text = address & ": " & data
    End Sub
End Class
```

In plaats van AddHandler kan het object ook met WithEvents gedeclareerd worden (`Private WithEvents XpressNet As XpressNetClass`), en kan Handles gebruikt worden (`Private Sub XpressNet_Feedback(...) Handles XpressNet.Feedback`).

De events worden aangeroepen op de thread die het XpressNetClass object heeft gemaakt. Als het object in Form_Load (of in de declaratie van het form) wordt gemaakt, mogen de event handlers dus direct de controls van het form aanpassen; Invoke is niet nodig.

## Nieuw in versie 2.0
- Na een foutmelding van het interface verstuurt de library een bericht opnieuw; zolang de centrale het interface niet adresseert, wordt het versturen gepauzeerd.
- Nieuwe events: `SendError(Reason, Description)` bij elk bericht dat niet verstuurd kan worden, en `InterfaceNichtAdressiert()` / `InterfaceWiederAdressiert()`. De waarden van Reason staan in `SendErrorReasonEnum`.
- `CmdConnectUsb` heeft de optionele parameter `Handshake` (standaard uit).
- De logfile staat standaard in `C:\Users\Public\TMC-Logfiles`.
- Z21: de library zet na `CmdInitUdp` de broadcast flags (nieuw commando `CmdZ21SetBroadcastFlags`).

De details staan in de [Reference manual](../../../Documentation/Reference/Reference.md) (Engels), hoofdstuk 10 en 11. Door de extra parameter van CmdConnectUsb moet een applicatie die versie 1.3 gebruikte opnieuw gebouwd worden; de aanroepen hoeven niet te veranderen.

## Van de VB6 DLL (COM) naar de .NET library
Een VB.NET applicatie die de VB6 DLL via COM gebruikt, kan als volgt overstappen op de .NET library:
1. Verwijder de referentie naar XpressNet (Solution Explorer → References → XpressNet → Remove).
2. Voeg een referentie naar XpressNet.NET.dll toe (zie optie 1).
3. Bouw de applicatie opnieuw.

De namen van de methods, events en attributen, en de typen van de parameters, zijn hetzelfde als in de VB6 DLL. De aanroepen en event handlers in de applicatie kunnen daarom blijven zoals ze zijn. Er zijn twee uitzonderingen:
- Het attribuut ZentraleVersion heet ZentraleSoftwareVersion. In .NET kunnen een event en een attribuut niet dezelfde naam hebben; het event ZentraleVersion houdt zijn naam.
- `XpressNet = Nothing` sluit de verbinding niet. Gebruik in plaats daarvan `XpressNet.Dispose()`.

## Verschillen met de VB6 DLL
- Het openen van een nieuwe verbinding sluit eerst een bestaande verbinding.
- Zolang de verbinding (nog) niet is opgebouwd, worden berichten niet verstuurd (logfile en event SendError). Bij de VB6 DLL sloot een bericht tijdens het opbouwen van de TCP verbinding de verbinding.
- Een parameter waarde die niet in een byte past, wordt niet verstuurd (logfile en event SendError). VB6: run time error "Overflow".
- Een fout (exception) in een event handler van de applicatie verstoort de library niet: de fout wordt in de logfile gezet, en daarna op de gebruikelijke manier aan de applicatie doorgegeven.
- Aanroepen vanuit een andere thread worden doorgegeven aan de thread die het object heeft gemaakt (er komt een regel in de logfile).
- Als de logfile niet geschreven kan worden, gaat de tekst verloren zonder foutmelding (VB6: run time error).

De logfile is dezelfde als bij de VB6 DLL: TMC-LogFile.txt in de directory C:\Users\Public\TMC-Logfiles, of in de directory die met LogDirectory is ingesteld.
