program vbasyncgui;

uses
  Vcl.Forms,
  VbaSync.Core in 'VbaSync.Core.pas',
  VbaSync.Version in 'VbaSync.Version.pas',
  VbaSync.MainForm in 'VbaSync.MainForm.pas';

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'VbaSync';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
