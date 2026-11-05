program vbasync;

{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  System.Variants,
  Winapi.ActiveX,
  System.Win.ComObj,
  VbaSync.Core in 'VbaSync.Core.pas';

procedure Usage;
begin
  Writeln('vbasync - Export/Import/Loeschen von Excel-VBA-Code');
  Writeln;
  Writeln('  vbasync list   <Mappe>');
  Writeln('  vbasync export <Mappe> <Ordner>');
  Writeln('  vbasync import <Mappe> <Ordner> [Modul ...]');
  Writeln('  vbasync delete <Mappe> (--all | Modul ...) [--clear-docs]');
  Writeln('  vbasync clean  <Mappe>');
  Writeln;
  Writeln('<Mappe>: Name wie in Excel (z. B. Projekt.xlsm) oder Pfad.');
  Writeln('Excel muss laufen. Speichern erfolgt NICHT automatisch.');
end;

var
  Cmd: string;
  Excel, WB: OleVariant;
  Sync: TVbaSync;
  Names: TArray<string>;
  I: Integer;
  All, ClearDocs: Boolean;
  S: string;
begin
  CoInitialize(nil);
  try
    try
      if ParamCount < 2 then
      begin
        Usage;
        ExitCode := 1;
        Exit;
      end;
      Cmd := LowerCase(ParamStr(1));

      try
        Excel := GetActiveOleObject('Excel.Application');
      except
        raise Exception.Create('Keine laufende Excel-Instanz gefunden.');
      end;
      WB := OpenWorkbook(Excel, ParamStr(2));

      Sync := TVbaSync.Create(WB,
        procedure(const Msg: string)
        begin
          Writeln(Msg);
        end);
      try
        if Cmd = 'list' then
          Sync.ListComponents
        else if Cmd = 'export' then
        begin
          if ParamCount < 3 then raise Exception.Create('Zielordner fehlt.');
          Sync.ExportAll(ExpandFileName(ParamStr(3)));
        end
        else if Cmd = 'import' then
        begin
          if ParamCount < 3 then raise Exception.Create('Quellordner fehlt.');
          for I := 4 to ParamCount do
            Names := Names + [ParamStr(I)];
          Sync.ImportFolder(ExpandFileName(ParamStr(3)), Names);
        end
        else if Cmd = 'delete' then
        begin
          All := False;
          ClearDocs := False;
          for I := 3 to ParamCount do
          begin
            S := ParamStr(I);
            if SameText(S, '--all') then All := True
            else if SameText(S, '--clear-docs') then ClearDocs := True
            else Names := Names + [S];
          end;
          if not All and (Length(Names) = 0) then
            raise Exception.Create('Module angeben oder --all verwenden.');
          Sync.DeleteComponents(Names, All, ClearDocs);
        end
        else if Cmd = 'clean' then
          Sync.Clean
        else
        begin
          Usage;
          ExitCode := 1;
        end;
      finally
        Sync.Free;
      end;
    except
      on E: Exception do
      begin
        Writeln(ErrOutput, 'FEHLER: ', E.Message);
        ExitCode := 2;
      end;
    end;
  finally
    WB := Unassigned;
    Excel := Unassigned;
    CoUninitialize;
  end;
end.
