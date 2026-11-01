# Versiegeschiedenis

## Versie 1.3 (oktober 2026)

### XpressNet DLL (Code/VB6_Source)

**Correcties**

- `XpressNet-Class-Main.cls`
  - PoM commando's: de lage byte van het CV nummer werd berekend als `CV Mod 255` in plaats van `CV Mod 256`. Hierdoor werd CV 256 gelijk aan CV 1, en waren CV 257 t/m 1024 één te hoog.
  - PoM bit schrijven (`CmdPomBitWrite`): `0xE8`, zoals in de V3.6 specificatie (3.2.28.3). Versie 1.2 gebruikte `0x7C` uit de V4 specificatie (3.43.3); dat lijkt een fout in die specificatie.
  - De versie van de centrale (`ZentraleVersion`) wordt als getal vergeleken (`Val`). Voorheen werd deze als string vergeleken, of impliciet omgezet. Een string vergelijking gaat fout bij versies vanaf 10.0, en de impliciete omzetting hangt af van de Windows landinstellingen (decimaalteken).
  - `CmdZ21GetExtAccessory` stuurde één databyte te veel (Z21 specificatie 5.5).
  - Lok informatie: de snelheidsbyte wordt gedecodeerd volgens het aantal rijstappen. Voorheen werd deze altijd als 27/28 rijstappen gedecodeerd, wat bij 14 en 128 rijstappen een verkeerde snelheid gaf.
  - De firmware versie van de Z21 is BCD gecodeerd (Z21 specificatie 2.15). `0x01 0x23` wordt nu getoond als 1.23 (was 1.35).
- `XpressNet-Class-Receive-Z21.cls`
  - Het XpressNet bericht en de XOR byte worden genomen uit het Z21 bericht dat gedecodeerd wordt. Voorheen werden ze uit de invoerbuffer genomen; dat ging alleen goed als een UDP pakket één enkel Z21 bericht bevatte.
  - Een bericht met een XOR fout wordt niet langer gedecodeerd (wel gelogd).
  - DataLen bestaat uit twee bytes (little endian). Een DataLen kleiner dan 4 geeft geen eindeloze lus meer, en een Z21 bericht langer dan 23 bytes geeft geen "subscript out of range" meer.
- `XpressNet-Class-Receive.cls`: als de invoerbuffer vol is, worden eerst de complete XpressNet berichten verwerkt. Voorheen werden de resterende bytes weggegooid en bleef de buffer vol, waardoor niets meer ontvangen kon worden.
- `XpressNet-Class-AccAuto.cls`: `emptyFifo` (aangeroepen na een response time-out) stopt ook de deactiveringstimer. Voorheen bleef de timer lopen met een lege FIFO, wat een overflow fout gaf en een deactivering van een verkeerd adres.

**Wijzigingen**

- `XpressNet-Class-Send.cls`: de response time-out (5 seconden, 90 seconden in de programmeer mode) wordt gebruikt voor TCP, USB en seriële verbindingen. Voorheen alleen voor TCP, waardoor bij USB en seriële verbindingen een ontbrekend antwoord de FIFO blokkeerde. Voor het USB interface geldt dezelfde regel als voor het LAN interface (LI-USB 1.1, Interface 23151 1.3). Het keep alive bericht na een time-out wordt alleen bij TCP verstuurd.
- `XpressNet-Class-Send.cls`: geen flow control bij de Z21 (UDP). De Z21 beantwoordt niet ieder commando (Z21 specificatie 1.1), daarom worden berichten voor de Z21 nu direct verstuurd. Voorheen kon de FIFO, inclusief het keep alive bericht, blokkeren tot er toevallig een ander XpressNet bericht van de Z21 binnenkwam.
- `XpressNet-Class-AccAuto.cls`: na een deactivering volgt een pauze van `AccessoryInterval` / 2 (62 ms) voordat het volgende wissel wordt geactiveerd, zoals in het voorbeeld van de Z21 specificatie (5.2.1). Dit beperkt de stroom als veel wissels geschakeld worden.
- `XpressNet.vbp`: versie 1.3 ("Version October 2026").

**Let op bij het overstappen**

Het DLL project gebruikt "No Compatibility". Versie 1.3 heeft daarom nieuwe GUID's. Applicaties die met een eerdere versie van de DLL zijn gecompileerd, moeten de referentie naar de XpressNet DLL opnieuw zetten (Project → References) en opnieuw gecompileerd worden. Bij Visual Studio (.NET) applicaties moet de COM referentie opnieuw worden toegevoegd.

### Voorbeelden

- `1-Basic_test`, `2-Extended_test`, `4_Multiple_Instance_test`, `5_Regfree_test`, `6-Project_Group`: `FunctionsZustandF29F68` sloeg de functiegroepen 6 t/m 10 op in `FB(4)` en `FB(5)`; deze worden nu opgeslagen in `FB(6)` t/m `FB(10)`.
- Z21 forms (`3_Z21_Basic test`, `4_Multiple_Instance_test`, `5_Regfree_test`, `6-Project_Group`): de Steps lijst bevat 14, 27, 28 en 128, en "Number of Z21Messages" wordt bijgewerkt (nieuwe Timer `TimerMessageCounter`).
- `5_Regfree_test`, Z21 form: `=>` en `<=` schakelden de verkeerde stand, en Feedback vroeg RS-Bus terugmeldingen op in plaats van de wisselinformatie.
- `6-Project_Group`: de DLL source files zijn bijgewerkt naar versie 1.3 (ze waren van vóór augustus 2025).
- Alle executables zijn opnieuw gecompileerd met DLL versie 1.3. De project files bevatten nu de referentie naar de DLL.

### Documentatie

- `Examples.md` / `Voorbeelden.md`: beschrijving van de Basic test toegevoegd; de executables van de Regfree test staan in de subdirectory `Exe`; opmerking bij het VS2019 voorbeeld over de COM referentie.
- `2A-Install-Dll-Via-regsvr32.md` / `2A-Installeren-Dll-Via-regsvr32.md`: op 64-bit Windows (zoals Windows 11) geldt stap 9B, ook voor VB6 applicaties.
- `Version_History.md` / `Versiegeschiedenis.md` toegevoegd.

## Versie 1.2 (augustus 2025)

- Het applicatie programma kan de directory voor de logfiles instellen (`LogDirectory`). Zie `1-Basic_test` voor een voorbeeld.
- De FIFO voor XpressNet berichten is vergroot van 32 naar 128 plaatsen.
- De tijd tussen twee wissel commando's (`AccessoryInterval`) is veranderd van 250 naar 125 ms.

## Versie 1.1 (december 2024)

- `CmdSwitchInfoRequest` is toegevoegd, om de wisselinformatie op te vragen. `CmdFeedbackRequest` is nu bedoeld voor (RS-Bus) terugmelddecoders.
- De voorbeelden vragen de juiste terugmeldgegevens op, gebruiken de standaard IP adressen van Lenz en Roco, en stand 1 is nu rechtdoor.

## Versie 1.0 (november 2024)

- Eerste versie.
