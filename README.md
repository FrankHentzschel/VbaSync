# VbaSync (Prototyp)

Export, Import, Löschen und „Clean“ von Excel-VBA-Code für die Versionsverwaltung mit Git.
Zwei Programme mit derselben Kern-Unit `VbaSync.Core.pas`:
- `vbasyncgui.exe`: Oberfläche (VCL) mit Komponentenliste, Checkboxen und Buttons
- `vbasync.exe`: Kommandozeile, z. B. für Batch-Export

- `VbaSyncAddin.dll`: VBE-Add-in mit Menü „VbaSync...“ in der VBA-IDE (siehe unten)

## Bauen (Delphi 12)
1. `src/vbasyncgui.dpr` (GUI) bzw. `src/vbasync.dpr` (Kommandozeile) in Delphi öffnen (die IDE legt die `.dproj` an).
2. Zielplattform Win32 (Win64 geht auch). Da die EXE Excel von außen steuert, funktioniert sie mit 32- und 64-Bit-Office.
3. Einmalig: Projekt → Optionen → Anwendung → Symbol → „Symbol laden…“ → `res\VbaSync.ico` (für beide Projekte).
4. Kompilieren. Es werden keine Typbibliotheken importiert, alles läuft per Late Binding.

## Voraussetzung in Excel
Trust Center → Makroeinstellungen → **„Zugriff auf das VBA-Projektobjektmodell vertrauen“**.

## VBE-Add-in (Menü in der VBA-IDE)
1. `src/VbaSyncAddin.dpr` in Delphi öffnen und als **Win32** bauen (passend zu 32-Bit-Office; für 64-Bit-Office Win64).
2. Excel schließen. In einer Eingabeaufforderung (ohne Administrator):
   ```
   regsvr32 C:\Work\VbaSync\src\Win32\Debug\VbaSyncAddin.dll
   ```
   Der Pfad hängt von deiner Ausgabe-Einstellung ab. Die DLL nach dem Registrieren nicht verschieben.
3. Excel starten, VBA-Editor öffnen (Alt+F11). In der Menüleiste erscheint **VbaSync...**. Unter Add-Ins → Add-In-Manager ist „VbaSync“ gelistet.
4. Entfernen: `regsvr32 /u <Pfad>\VbaSyncAddin.dll`.

Hinweis: Während Excel läuft, ist die DLL gesperrt. Zum Neubauen Excel schließen.

## GUI
Excel starten, Mappe öffnen, `vbasyncgui.exe` starten. Oben Mappe wählen (Vorschlag für den Ordner: `src` neben der Mappe), rechts die Aktionen. „Markierte …“ bezieht sich auf die Häkchen in der Liste. Löschen, Import und Clean fragen vorher nach.

## Kommandozeile (Excel muss laufen)
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
