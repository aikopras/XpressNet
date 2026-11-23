[🇬🇧 English](README.md) | [🇩🇪 Deutsch](README.de.md) | [🇳🇱 Nederlands](README.nl.md)

# XpressNet library
De library heeft in alle code bases dezelfde methoden, events en attributen. Per code base:

| Code base | Source | Gecompileerd | Installatie |
|---|---|---|---|
| macOS (Objective-C, Xcode) | [src/macOS](src/macOS) | - (de source wordt aan de applicatie toegevoegd) | [handleiding](src/macOS/Install/Installeren-Library.md) |
| Qt (C++, Qt 6, CMake; Linux en Windows) | [src/Qt](src/Qt) | - (wordt samen met de applicatie gebouwd) | [handleiding](src/Qt/Install/Installeren-Library.md) |
| .NET (C#, .NET Framework 4.7.2) | [src/DotNet](src/DotNet) | [compiled/DotNet](compiled/DotNet) | [handleiding](src/DotNet/Install/Installeren-Library.md) |
| VB6 (COM DLL) | [src/VB6](src/VB6) | [compiled/VB6](compiled/VB6) | [handleiding](src/VB6/Install/Installeren-Library.md) |

De [documentatie](Documentation) geldt voor alle code bases:
- [Reference manual](Documentation/Reference/Reference.md) (Engels): alle methoden, events en attributen, met de XpressNet berichten en de hoofdstukken in de specificaties.
- [Commands](Documentation/Commands): PDF files met de geïmplementeerde commando's, en hoe goed ze getest zijn.
- [HowToUse](Documentation/HowToUse/Handleiding.md): gebruikers handleiding voor VB.NET en VB6.

De [voorbeelden](../Examples/README.nl.md) laten zien hoe de library gebruikt wordt.
