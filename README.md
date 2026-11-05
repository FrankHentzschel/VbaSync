# VbaSync (Prototyp)

Export, Import, Löschen und „Clean“ von Excel-VBA-Code für die Versionsverwaltung mit Git.
Stand: Prototyp 1, Kommandozeile. Ein VBE-Add-in mit Menü folgt und nutzt dieselbe Unit `VbaSync.Core.pas`.

## Bauen (Delphi 12)
1. `src/vbasync.dpr` in Delphi öffnen (die IDE legt die `.dproj` an).
2. Zielplattform Win32 (Win64 geht auch). Da die EXE Excel von außen steuert, funktioniert sie mit 32- und 64-Bit-Office.
3. Kompilieren. Es werden keine Typbibliotheken importiert, alles läuft per Late Binding.

## Voraussetzung in Excel
Trust Center → Makroeinstellungen → **„Zugriff auf das VBA-Projektobjektmodell vertrauen“**.

## Aufruf (Excel muss laufen)
```
vbasync list   Projekt.xlsm
vbasync export Projekt.xlsm C:\Repos\Projekt\src
vbasync import Projekt.xlsm C:\Repos\Projekt\src            (alle)
vbasync import Projekt.xlsm C:\Repos\Projekt\src modTools   (nur einzelne)
vbasync delete Projekt.xlsm --all [--clear-docs]
vbasync delete Projekt.xlsm modTools clsLogger
vbasync clean  Projekt.xlsm
```
Gespeichert wird **nicht** automatisch.

## Dateiformate
| Typ | Datei | Import-Verhalten |
|---|---|---|
| Standardmodul | `.bas` | vorhandenes Modul wird ersetzt (vorher umbenannt und entfernt) |
| Klasse | `.cls` | wie oben; Attribute wie `VB_PredeclaredId` bleiben erhalten |
| UserForm | `.frm` + `.frx` | wie oben; Warnung, wenn die `.frx` fehlt |
| Dokumentmodul (DieseArbeitsmappe, Tabelle1 …) | `.doccls` | nur der Code wird ersetzt; fehlt das Blatt, wird es übersprungen und eine Warnung ausgegeben |

Der Modulname wird beim Import aus `Attribute VB_Name` gelesen, nicht aus dem Dateinamen.
`export` löscht vorher die verwalteten Dateitypen im Zielordner, damit in Git entfernte Module als gelöscht erscheinen.
`clean` exportiert zuerst in einen Temp-Ordner, der als Backup erhalten bleibt.

## Bekannte Grenzen / ungetestet
- Der Code ist **ungetestet** und wurde nicht kompiliert (er entstand ohne Delphi und ohne Excel).
- `GetActiveOleObject` findet nur die erste registrierte Excel-Instanz.
- Kennwortgeschützte Projekte werden abgelehnt.
- Encoding: ANSI (System-Codepage), passend zum VBE-Export.

## Test-Checkliste
1. `list` auf einer Test-Mappe mit Modul, Klasse, UserForm und Code in Tabelle1/DieseArbeitsmappe.
2. `export` → Dateien prüfen (Umlaute korrekt?).
3. In `.bas` und `.doccls` etwas ändern → `import` → in der IDE prüfen.
4. `delete modX`, dann `delete --all --clear-docs` → `import` → alles wieder da?
5. `clean` → Mappe speichern, erneut öffnen, kompilieren (Debuggen → Kompilieren).
