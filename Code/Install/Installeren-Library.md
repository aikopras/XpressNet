# Installatie van de XpressNet library


Er zijn meerdere manieren om de Library software te installeren en te gebruiken. Iedere manier heeft zijn eigen voor- en nadelen. Er dient hierbij onderscheid gemaakt te worden tussen de ontwikkel machine, en de productie machines.

## 1: Broncode toevoegen aan de applicatie
Het is mogelijk gedurende de applicatie ontwikkel fase de broncode van de library toe te voegen aan de broncode van de applicatie. Hiertoe moet, in de IDE, een VB groep worden gemaakt, welke de twee projecten samenvoegt: 1) de applicatie code en 2) de XpressNet library code. Zolang er binnen de IDE wordt gewerkt, kan het programma uitevoerd worden zonder administrator privileges. Als er echter een executable (.exe) gemaakt wordt, dan kan dat niet zonder administrator privileges. Mocht er toch een executable gewenst zijn, start dan VB6 met administrator rechten (zie onder *2: In windows registry registreren*), of registreer de DLL meteen in de windows registry.

**Voordelen:**</br>
    - Tijdens de ontwikkeling zijn geen administrator rechten nodig.</br>
    - Indien nodig, kan de XpressNet library code relatief eenvoudig worden aangepast.

**Nadelen:**</br>
    - Zonder administrator rechten kan geen executable (.exe) worden gemaakt.</br>
    - Applicatie programma ***moet*** in VB6 zijn geschreven (en niet in .NET).</br>
    - Details van de library (de broncode) zijn zichtbaar voor de applicatie ontwikkelaar. Hierdoor zou (per ongeluk) de library kunnen worden aangepast, waardoor het geheel niet meer werkt.</br>

**Installatie:** zie deze [handleiding](1-VB6_Groep.md) voor details.

## 2: In windows registry registreren
De DLL file (XpressNet.dll) wordt in de windows registry geïnstalleerd. Dit moet gebeuren op de ontwikkel- maar ook de productie machines.

**Voordelen:**</br>
    - Eenvoudig in gebruik.</br>
    - Details van de library (zoals de broncode) blijven verborgen voor de applicatie ontwikkelaar.</br>
    - De library (dll file) hoeft op iedere machine slechts eenmaal te worden geïnstalleerd (=geregistreerd). Daarna kan ieder applicatie programma (.exe file) het gebruiken.</br>
    - Werkt ook met applicaties die niet geschreven zijn in VB6, maar in bijvoorbeeld VB.NET (VB2019/2022), C#, of andere .NET talen. Merk op dat in dat geval de applicatie gebouwd moet zijn als 32-bit (x86), of dat er gebruik gemaakt moet worden van een zogeheten COM+ bridge (zie de  [speciale instructies hiervoor](2A-Installeren-Dll-Via-Com-Plus.md)).

**Nadelen:**</br>
    - Voor installatie van de DLL zijn administrator rechten nodig, niet alleen op de ontwikkel machine, maar ook alle productie machines.</br>
    - Verschillende versies van de DLL kunnen voor problemen zorgen</br>
    - De DLL moet op een "stabiele" plaats worden opgeslagen, en mag later niet verplaatst worden.

**Installatie:** </br>
Installatie is, naast de ontwikkel machine, ook nodig voor alle productie machines. Er zijn twee methoden om de DLL in de windows registry te installeren.</br>
    1. Door gebruik te maken van ***regsvr32***. Deze methode is relatief eenvoudig, maar wordt op Internet soms afgeraden. Zie deze [installatie handleiding](2A-Installeren-Dll-Via-regsvr32.md) hoe dit kan worden gedaan.</br>
    2. De library code wordt in VB6 vertaald en vervolgens (voor de gebruiker onzichtbaar) door VB6 geregistreerd. Deze methode is wat omslachtiger, maar vermijd het gebruik van regsvr32. Hoe dit moet staat in deze [installatie handleiding](2B-Installeren-Dll-Via-VB6.md).


## 3: De DLL samen met de applicatie te verspreiden.
Een derde mogelijkheid is de XpressNet DLL file naast het applicatie programma (.exe file) op de productie machines te zetten. De DLL file hoeft hierbij *niet* op de productie machines in de registry gezet te worden, zodat er geen administrator rechten nodig zijn. Dit heet *RegFree* installatie.

Op de ontwikkel machine moet de DLL wel in de registry worden geïnstalleerd. Hiertoe kan één van de methoden genoemd onder 2. hierboven worden gebruikt.

**Voordelen:**</br>
    - Geen administrator rechten nodig op de productie machines.</br>
    - Details van de library (zoals de broncode) blijven verborgen voor de applicatie ontwikkelaar.</br>

**Nadelen:**</br>
    - Op de ontwikkel machine zijn administrator rechten nodig.</br>
    - Aan de broncode van de applicatie moet een extra class (.cls) file worden toegevoegd, en één regel in het applicatie programma moet worden aangepast.

**Installatie:** </br>
De source code moet worden gedownload zoals beschreven onder optie 1. Daarnaast moet er een extra class file (XPNRegFree.cls) worden toegevoegd. Details zijn beschreven in de voorbeeld applicatie [Regfree_test](../../Examples/Voorbeelden.md).
