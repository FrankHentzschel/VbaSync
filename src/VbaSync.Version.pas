unit VbaSync.Version;

{
  Versionsinformationen. Schema MAJOR.MINOR.PATCH:
    PATCH (rechts)  +1 bei Fehlerbehebungen
    MINOR (Mitte)   +1 bei neuen Funktionen, PATCH wird 0
    MAJOR (links)   +1 bei grundlegenden/inkompatiblen Aenderungen
  Bei jeder Aenderung AppVersion und AppVersionDate anpassen und CHANGELOG.md ergaenzen.
}

interface

const
  AppName        = 'VbaSync';
  AppDescription = 'Export, Import und Clean von Excel-VBA-Code fuer die Versionsverwaltung';
  AppVersion     = '1.7.1';
  AppVersionDate = '06.11.2026 00:20';
  AppRepoUrl     = 'https://github.com/FrankHentzschel/VbaSync';

implementation

end.
