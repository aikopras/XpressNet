# XpressNet DLL installeren met de COM+ bridge

De XpressNet DLL is een 32-bit COM-component.  
Als je je eigen applicatie als **64-bit** bouwt, kan de DLL niet direct (in-process) geladen worden.  
Windows kan dit oplossen door de DLL te draaien in een **COM+ server applicatie** (een apart 32-bit proces).  
Je 64-bit applicatie kan dan toch gewoon via COM met de DLL praten.

---

## 1. Voorwaarden
- Je hebt administrator-rechten nodig op je pc.
- Kopieer de DLL eerst naar een vaste locatie, bijvoorbeeld:  
  `C:\Program Files\XpressNet\XpressNet.dll`
- Registreer de DLL in het **32-bit COM-register**:  
  ```bat
  C:\Windows\SysWOW64\regsvr32.exe "C:\Program Files\XpressNet\XpressNet.dll"
  ```

---

## 2. Maak een COM+ applicatie aan

1. Open het startmenu van Windows, typ **Component Services** en start dit programma.  
   (Je kunt ook `dcomcnfg` uitvoeren.)
2. In de boomstructuur links open je:  
   `Component Services → Computers → My Computer → COM+ Applications`
3. Rechtsklik op **COM+ Applications** → **New → Application…**
4. In de wizard:  
   - Kies **Server application**  
   - Geef een naam, bijvoorbeeld **XpressNetHost**  
   - Bij Identity kies je **Interactive user** (de eenvoudigste optie)

---

## 3. Voeg de DLL toe aan de COM+ applicatie

1. Rechtsklik je nieuwe applicatie (bijv. *XpressNetHost*) → **New → Component…**  
2. Kies **Install new component** en blader naar je DLL  
   (`C:\Program Files\XpressNet\XpressNet.dll`).  
3. Rond de wizard af. De klassen uit de DLL verschijnen nu onder *Components*.

---

## 4. Gebruik in je eigen applicatie

- Je hoeft in je code **niets** aan te passen.  
- Maak gewoon het object zoals je gewend bent:  
  ```vbnet
  Dim xp As New XpressNet.XpressNetClass
  ```
- Als je applicatie 64-bit is, start Windows automatisch een 32-bit **dllhost.exe** proces voor de COM+ applicatie.  
- Alle aanroepen worden automatisch doorgestuurd tussen jouw proces en de COM+ host.

---

## 5. De COM+ applicatie verwijderen

Wil je de COM+ bridge later verwijderen?

1. Open opnieuw **Component Services**.  
2. Zoek je applicatie (bijv. *XpressNetHost*) onder **COM+ Applications**.  
3. Rechtsklik en kies **Delete**.  
4. Wil je ook de DLL uit het register halen? Gebruik:  
   ```bat
   C:\Windows\SysWOW64\regsvr32.exe /u "C:\Program Files\XpressNet\XpressNet.dll"
   ```

---

## 6. Problemen oplossen

- Krijg je *Class not registered* fouten?  
  → Voer het `regsvr32` commando (stap 1) opnieuw uit met admin-rechten.  
- Werkt het niet of loopt het vast?  
  → Kijk in **Windows Event Viewer → Applications and Services Logs → Microsoft → Windows → COM+**.  
- Meer controle nodig?  
  → In de eigenschappen van je COM+ applicatie kun je identity, security of idle timeout aanpassen.

---

### Samenvatting
- 32-bit applicaties (VB6, VB.NET x86) kunnen de DLL direct gebruiken.  
- 64-bit applicaties kunnen de DLL niet direct laden.  
- Met de COM+ bridge kan een 64-bit applicatie de DLL tóch transparant via COM gebruiken.
