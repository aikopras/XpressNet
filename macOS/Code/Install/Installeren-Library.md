# Gebruik van de macOS library

De macOS versie van de XpressNet library is geschreven in Objective-C. Versie 2.0 heeft dezelfde functionaliteit als versie 2.0 van de VB6 DLL, de .NET library en de Linux versie. De source code bestaat uit twee directories:
- [XpressNet](../XpressNet): de XpressNet classes (XpressNetClass, Send, Receive, ReceiveZ21, AccAuto en LogFile)
- [Communication](../Communication): de TCP, UDP en seriële (USB) verbindingen

Om de library in een Xcode project te gebruiken, kopieer beide directories naar de directory van het project, en voeg ze toe aan het project (File → Add Files to "…"). De library gebruikt manual reference counting (geen ARC), net als de voorbeeld projecten.

De methods, events en attributen hebben dezelfde namen als in de VB6 DLL. De events zijn de methods van het XpressNetDelegate protocol; de applicatie implementeert alleen de events die nodig zijn. De events worden aangeroepen op de main thread. Voorbeeld:

```objc
_XpressNet = [[XpressNetClass alloc] init];
[_XpressNet setDelegate:self];
[[_XpressNet Logfile] setLogDirectory:@"~/Library/Logs/TMC-Logfiles"];
[_XpressNet CmdConnectTcp:@"192.168.0.200" Port:@"5550"];

- (void)Feedback:(int)address Nibble:(int)Nibble data:(int)data {
  ...
}
```

Zie XpressNet.h voor alle methods en events, en voor de verschillen met de VB6 DLL. De voorbeeld projecten (macOS/Examples) bevatten een kopie van beide directories.

## Nieuw in versie 2.0
- Na een foutmelding van het interface verstuurt de library een bericht opnieuw; zolang de centrale het interface niet adresseert, wordt het versturen gepauzeerd.
- Nieuwe events (XpressNetDelegate): `SendError:Description:` bij elk bericht dat niet verstuurd kan worden, en `InterfaceNichtAdressiert` / `InterfaceWiederAdressiert`. De waarden van Reason staan in `SendErrorReasonEnum`.
- `CmdConnectUsb:Settings:Handshake:` (hardware handshake RTS/CTS); zonder Handshake staat die uit.
- De logfile staat standaard in `~/Library/Logs/TMC-Logfiles`. Iedere regel begint met de letter van het XpressNet object (A, B, …).
- Z21: na `CmdInitUdp:Port:` zet de library de broadcast flags (nieuw commando `CmdZ21SetBroadcastFlags:`).
- Berichten die aangeboden worden terwijl de TCP of UDP verbinding wordt opgebouwd, worden verstuurd zodra de verbinding klaar is (VB6, .NET en Linux: event SendError).

De details staan in de [Reference manual](../../../Documentation/Reference/Reference.md) (Engels), hoofdstuk 10, 11 en 13.
