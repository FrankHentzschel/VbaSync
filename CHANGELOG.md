# Changelog

Schema: MAJOR.MINOR.PATCH. PATCH = Fehlerbehebung, MINOR = neue Funktion.

## 1.5.3 – 05.11.2026 11:20
- Fix Add-in: Klick-Ereignis über `FindConnectionPoint` mit der IID aus der VBIDE-Typbibliothek abonnieren (Aufzählen schlug mit E_OUTOFMEMORY fehl)

## 1.5.2 – 05.11.2026 11:15
- Fix Add-in: Bereichsprüfungsfehler beim Formatieren von COM-Fehlercodes; Menüleiste über `CommandBars.ActiveMenuBar` (Fallback „Menu Bar“)

## 1.5.1 – 05.11.2026 11:10
- Fix Add-in: „Mitglied nicht gefunden“ beim Anlegen des Menüs; COM-Aufrufe jetzt mit Argumenten by value und expliziten Flags
- Fehlermeldung nennt den fehlgeschlagenen Schritt

## 1.5.0 – 05.11.2026 11:05
- Neu: VBE-Add-in `VbaSyncAddin.dll` (Win32) mit Menüeintrag „VbaSync...“ in der VBA-IDE; öffnet das VbaSync-Fenster mit der aktiven Mappe vorausgewählt
- Registrierung pro Benutzer per `regsvr32`, ohne Administratorrechte
- Delphi-Projektdateien (`.dproj`) im Repository

## 1.4.3 – 05.11.2026 10:45
- Fix: Hauptfenster erscheint in Alt+Tab und in der Taskleiste

## 1.4.2 – 05.11.2026 10:40
- Fix: Programmsymbol wird jetzt als eigene Ressource (`src/VbaSyncIcon.res`) eingebunden und im Fenster, in der Taskleiste und im Info-Dialog angezeigt, unabhängig von den Projektoptionen

## 1.4.1 – 05.11.2026 10:35
- Programmsymbol überarbeitet: modernes Whiteboard mit farbigem Code und Wischer

## 1.4.0 – 05.11.2026 10:30
- Neues Programmsymbol (Tafel mit Schwamm), `res/VbaSync.ico` mit 16–256 px
- Info-Dialog zeigt das Programmsymbol statt des Emojis

## 1.3.1 – 05.11.2026 09:00
- Fix: Compilerfehler E2003 `fsBold` (fehlende Unit Vcl.Graphics)
- Fix: Warnung W1010 (Methode DoShow umbenannt)

## 1.3.0 – 05.11.2026 08:55
- Info-Dialog (Button unten rechts) mit Programmname, Version, Versionsdatum und Erstellungszeit der EXE
- Version im Fenstertitel

## 1.2.0 – 05.11.2026 08:50
- „Dateien importieren...“: Datei-öffnen-Dialog mit Mehrfachauswahl. Vorhandene Komponenten werden ersetzt, fehlende hinzugefügt, Dokumentmodule erhalten nur neuen Code (auch aus VBE-exportierter `.cls`).

## 1.1.0 – 05.11.2026 08:48
- Mehrfachauswahl (Shift/Strg) in der Komponentenliste; ohne Häkchen wirken die „Markierte“-Aktionen auf die Auswahl
- Rückfragen nennen die betroffenen Komponenten

## 1.0.2 – 05.11.2026 08:40
- Fix: Compilerfehler E2029 (Zeiger auf generisches Array)

## 1.0.1 – 05.11.2026 08:35
- Fix: Arbeitsmappen werden in allen Excel-Instanzen gefunden (statt nur der ersten registrierten)
- Checkbox „Tabellen-Code leeren“ einzeilig

## 1.0.0 – 05.11.2026 00:15
- Erste Version: GUI und Kommandozeile mit Export, Import, Löschen und Clean
