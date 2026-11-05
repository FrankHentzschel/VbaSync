library VbaSyncAddin;

{ VBE-Add-in fuer Excel. Registrieren (ohne Administrator):
    regsvr32 VbaSyncAddin.dll
  Entfernen:
    regsvr32 /u VbaSyncAddin.dll }

uses
  System.Win.ComServ,
  VbaSync.Core in 'VbaSync.Core.pas',
  VbaSync.Version in 'VbaSync.Version.pas',
  VbaSync.MainForm in 'VbaSync.MainForm.pas',
  VbaSync.Addin in 'VbaSync.Addin.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  VbaSyncRegisterServer name 'DllRegisterServer',
  VbaSyncUnregisterServer name 'DllUnregisterServer';

begin
end.
