# Installatie van de XpressNet library


Er zijn meerdere manieren om de Library software (dll file) te installeren en te gebruiken. Iedere manier heeft zijn eigen voor- en nadelen. Er dient hierbij onderscheid gemaakt te worden tussen de ontwikkel machine, en de productie machines.

## 1: Broncode toevoegen aan de applicatie
Het is mogelijk de broncode van de library toe te voegen aan de broncode van de applicatie. Hiertoe wordt een VB groep gemaakt, welke twee projecten samenvoegt: 1) de applicatie code en 2) de XpressNet library code.

**Voordelen:**</br>
    - Nergens administrator rechten nodig.</br>
    - Er is geen DLL file nodig (we gebruiken immers de broncode). Dus geen problemen met het registreren van DLLs.</br>
    - Indien nodig, kan de XpressNet library code relatief eenvoudig worden aangepast.</br>
    - Nadat het applicatie programma is ontwikkeld en een .exe file is gemaakt, hoeft alleen deze .exe file op de productie machines te worden gezet.

**Nadelen:**</br>
    - Applicatie programma moet in VB6 zijn geschreven.</br>
    - Details van de library (de broncode) zijn zichtbaar voor de applicatie ontwikkelaar. Hierdoor zou (per ongeluk) de library kunnen worden aangepast, waardoor het geheel niet meer werkt.</br>

**Installatie:** zie deze [handleiding](abc.md) voor details.

## 2: In windows registry registreren
De dll file (XpressNet.dll) wordt in de windows registry geïnstalleerd. Dit moet gebeuren op de ontwikkel- maar ook de productie machines.

**Voordelen:**</br>
    - Eenvoudig in gebruik.</br>
    - Details van de library (zoals de broncode) blijven verborgen voor de applicatie ontwikkelaar.</br>
    - De library (dll file) hoeft op iedere machine slechts eenmaal te worden geïnstalleerd (=geregistreerd). Daarna kan ieder applicatie programma (.exe file) het gebruiken.</br>
    - Zou ook moeten werken met applicaties die niet geschreven zijn in VB6, maar in bijvoorbeeld C en .Net.

**Nadelen:**</br>
    - Voor installatie van de DLL zijn administrator rechten nodig, niet alleen op de ontwikkel machine, maar ook alle productie machines.</br>
    - Verschillende versies van de dll kunnen voor problemen zorgen

**Installatie:** </br>
Installatie is, naast de ontwikkel machine, ook nodig voor alle productie machines. Er zijn twee methoden om de DLL in de windows registry te installeren.</br>
    1. Door gebruik te maken van ***regsvr32***. Deze methode is relatief eenvoudig, maar wordt op Internet soms afgeraden. Zie deze [installatie handleiding](abc.md) hoe dit kan worden gedaan.</br>
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
