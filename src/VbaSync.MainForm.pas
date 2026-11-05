unit VbaSync.MainForm;

{
  VCL-Oberflaeche fuer VbaSync. Die Steuerelemente werden im Code erzeugt
  (keine .dfm), damit die Unit ohne Formular-Designer versionierbar bleibt.
  Dasselbe Formular soll spaeter auch aus dem VBE-Add-in heraus aufgerufen werden.
}

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Variants,
  Vcl.Forms, Vcl.Controls, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.Dialogs,
  VbaSync.Core, VbaSync.Version;

type
  TMainForm = class(TForm)
  private
    FWorkbooks: TArray<OleVariant>;
    FFullNames: TStringList;
    cbWorkbook: TComboBox;
    btnRefresh: TButton;
    edFolder: TEdit;
    btnBrowse: TButton;
    lvComps: TListView;
    chkClearDocs: TCheckBox;
    memLog: TMemo;
    procedure BuildUI;
    function AddButton(AParent: TWinControl; const ACaption: string; ATop: Integer;
      AOnClick: TNotifyEvent): TButton;
    procedure Log(const Msg: string);
    function CurrentWorkbook: OleVariant;
    function CheckedNames: TArray<string>;
    function RequireFolder: string;
    procedure RefreshComponents;
    procedure RunAction(const Action: TProc<TVbaSync>);
    procedure DoRefreshWorkbooks(Sender: TObject);
    procedure DoWorkbookChanged(Sender: TObject);
    procedure DoBrowse(Sender: TObject);
    procedure DoSelectAll(Sender: TObject);
    procedure DoSelectNone(Sender: TObject);
    procedure DoExport(Sender: TObject);
    procedure DoImportAll(Sender: TObject);
    procedure DoImportSelected(Sender: TObject);
    procedure DoImportFiles(Sender: TObject);
    procedure DoDeleteSelected(Sender: TObject);
    procedure DoDeleteAll(Sender: TObject);
    procedure DoClean(Sender: TObject);
    procedure HandleFirstShow(Sender: TObject);
    procedure DoAbout(Sender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  MainForm: TMainForm;

implementation

uses
  System.IOUtils, System.UITypes, System.StrUtils, System.Win.ComObj, Vcl.Graphics;

const
  FormW = 920;
  BtnW  = 170;

{ TMainForm }

constructor TMainForm.Create(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  FFullNames := TStringList.Create;
  BuildUI;
  OnShow := HandleFirstShow;
end;

destructor TMainForm.Destroy;
begin
  FFullNames.Free;
  FWorkbooks := nil;
  inherited;
end;

function TMainForm.AddButton(AParent: TWinControl; const ACaption: string; ATop: Integer;
  AOnClick: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := AParent;
  Result.SetBounds(8, ATop, BtnW - 16, 28);
  Result.Caption := ACaption;
  Result.OnClick := AOnClick;
end;

procedure TMainForm.BuildUI;
var
  pnlTop, pnlRight, pnlInfo: TPanel;
  lbl: TLabel;
  Col: TListColumn;
  Y: Integer;
begin
  Caption := AppName + ' ' + AppVersion + ' - Excel-VBA Export / Import';
  Position := poScreenCenter;
  ClientWidth := FormW;
  ClientHeight := 640;
  Constraints.MinWidth := 700;
  Constraints.MinHeight := 450;

  { Kopfbereich: Arbeitsmappe und Ordner }
  pnlTop := TPanel.Create(Self);
  pnlTop.Parent := Self;
  pnlTop.Align := alTop;
  pnlTop.Height := 76;
  pnlTop.BevelOuter := bvNone;

  lbl := TLabel.Create(Self);
  lbl.Parent := pnlTop;
  lbl.SetBounds(8, 12, 90, 17);
  lbl.Caption := 'Arbeitsmappe:';

  cbWorkbook := TComboBox.Create(Self);
  cbWorkbook.Parent := pnlTop;
  cbWorkbook.Style := csDropDownList;
  cbWorkbook.SetBounds(100, 8, FormW - 100 - 120, 23);
  cbWorkbook.Anchors := [akLeft, akTop, akRight];
  cbWorkbook.OnChange := DoWorkbookChanged;

  btnRefresh := TButton.Create(Self);
  btnRefresh.Parent := pnlTop;
  btnRefresh.SetBounds(FormW - 112, 7, 104, 25);
  btnRefresh.Anchors := [akTop, akRight];
  btnRefresh.Caption := 'Aktualisieren';
  btnRefresh.OnClick := DoRefreshWorkbooks;

  lbl := TLabel.Create(Self);
  lbl.Parent := pnlTop;
  lbl.SetBounds(8, 46, 90, 17);
  lbl.Caption := 'Ordner:';

  edFolder := TEdit.Create(Self);
  edFolder.Parent := pnlTop;
  edFolder.SetBounds(100, 42, FormW - 100 - 120, 23);
  edFolder.Anchors := [akLeft, akTop, akRight];

  btnBrowse := TButton.Create(Self);
  btnBrowse.Parent := pnlTop;
  btnBrowse.SetBounds(FormW - 112, 41, 104, 25);
  btnBrowse.Anchors := [akTop, akRight];
  btnBrowse.Caption := 'Durchsuchen...';
  btnBrowse.OnClick := DoBrowse;

  { Protokoll unten }
  memLog := TMemo.Create(Self);
  memLog.Parent := Self;
  memLog.Align := alBottom;
  memLog.Height := 160;
  memLog.ReadOnly := True;
  memLog.ScrollBars := ssVertical;
  memLog.Font.Name := 'Consolas';

  { Buttons rechts }
  pnlRight := TPanel.Create(Self);
  pnlRight.Parent := Self;
  pnlRight.Align := alRight;
  pnlRight.Width := BtnW;
  pnlRight.BevelOuter := bvNone;

  pnlInfo := TPanel.Create(Self);
  pnlInfo.Parent := pnlRight;
  pnlInfo.Align := alBottom;
  pnlInfo.Height := 40;
  pnlInfo.BevelOuter := bvNone;
  AddButton(pnlInfo, 'Info...', 4, DoAbout);

  Y := 4;
  AddButton(pnlRight, 'Alle markieren', Y, DoSelectAll);        Inc(Y, 32);
  AddButton(pnlRight, 'Keine markieren', Y, DoSelectNone);      Inc(Y, 48);
  AddButton(pnlRight, 'Alle exportieren', Y, DoExport);         Inc(Y, 48);
  AddButton(pnlRight, 'Alle importieren', Y, DoImportAll);      Inc(Y, 32);
  AddButton(pnlRight, 'Markierte importieren', Y, DoImportSelected); Inc(Y, 32);
  AddButton(pnlRight, 'Dateien importieren...', Y, DoImportFiles); Inc(Y, 48);
  AddButton(pnlRight, 'Markierte loeschen', Y, DoDeleteSelected); Inc(Y, 32);
  AddButton(pnlRight, 'Alle loeschen', Y, DoDeleteAll);         Inc(Y, 32);

  chkClearDocs := TCheckBox.Create(Self);
  chkClearDocs.Parent := pnlRight;
  chkClearDocs.SetBounds(8, Y, BtnW - 16, 22);
  chkClearDocs.Caption := 'Tabellen-Code leeren';
  chkClearDocs.Hint := 'Beim Loeschen auch den Code in Tabellen und DieseArbeitsmappe leeren';
  chkClearDocs.ShowHint := True;
  Inc(Y, 40);

  AddButton(pnlRight, 'Clean (Code Cleaner)', Y, DoClean);

  { Komponentenliste }
  lvComps := TListView.Create(Self);
  lvComps.Parent := Self;
  lvComps.Align := alClient;
  lvComps.ViewStyle := vsReport;
  lvComps.Checkboxes := True;
  lvComps.RowSelect := True;
  lvComps.ReadOnly := True;
  lvComps.HideSelection := False;
  lvComps.MultiSelect := True; // Shift/Strg-Mehrfachauswahl
  Col := lvComps.Columns.Add; Col.Caption := 'Name';   Col.Width := 300;
  Col := lvComps.Columns.Add; Col.Caption := 'Typ';    Col.Width := 120;
  Col := lvComps.Columns.Add; Col.Caption := 'Zeilen'; Col.Width := 80;
  Col.Alignment := taRightJustify;
end;

procedure TMainForm.Log(const Msg: string);
begin
  memLog.Lines.Add(Msg);
  Application.ProcessMessages;
end;

procedure TMainForm.DoAbout(Sender: TObject);
var
  Dlg: TForm;
  Icon, Lbl: TLabel;
  Btn: TButton;
  Y: Integer;
  ExeTime: TDateTime;

  procedure AddLine(const Text: string; Size: Integer; Bold: Boolean);
  begin
    Lbl := TLabel.Create(Dlg);
    Lbl.Parent := Dlg;
    Lbl.Left := 110;
    Lbl.Top := Y;
    Lbl.Font.Size := Size;
    if Bold then Lbl.Font.Style := [fsBold];
    Lbl.Caption := Text;
    Inc(Y, Lbl.Height + 6);
  end;

begin
  Dlg := TForm.CreateNew(Self);
  try
    Dlg.Caption := 'Info ueber ' + AppName;
    Dlg.BorderStyle := bsDialog;
    Dlg.Position := poOwnerFormCenter;
    Dlg.ClientWidth := 520;

    // Putzeimer-Symbol (Unicode U+1FAA3), dargestellt mit der Windows-Emoji-Schrift
    Icon := TLabel.Create(Dlg);
    Icon.Parent := Dlg;
    Icon.SetBounds(20, 16, 80, 80);
    Icon.Font.Name := 'Segoe UI Emoji';
    Icon.Font.Size := 44;
    Icon.Caption := #$D83E#$DEA3;

    Y := 18;
    AddLine(AppName, 18, True);
    AddLine('Version ' + AppVersion, 11, False);
    AddLine('Versionsdatum: ' + AppVersionDate, 9, False);
    if FileAge(ParamStr(0), ExeTime) then
      AddLine('EXE erstellt: ' + FormatDateTime('dd.mm.yyyy hh:nn', ExeTime), 9, False);
    Inc(Y, 6);
    AddLine(AppDescription, 9, False);
    AddLine(AppRepoUrl, 9, False);
    {$IFDEF WIN64}
    AddLine('Plattform: Win64', 9, False);
    {$ELSE}
    AddLine('Plattform: Win32', 9, False);
    {$ENDIF}

    Btn := TButton.Create(Dlg);
    Btn.Parent := Dlg;
    Btn.Caption := 'OK';
    Btn.ModalResult := mrOk;
    Btn.Default := True;
    Btn.Cancel := True;
    Btn.SetBounds(Dlg.ClientWidth - 100, Y + 10, 88, 28);
    Dlg.ClientHeight := Btn.Top + Btn.Height + 12;
    Dlg.ShowModal;
  finally
    Dlg.Free;
  end;
end;

procedure TMainForm.HandleFirstShow(Sender: TObject);
begin
  OnShow := nil;
  DoRefreshWorkbooks(nil);
end;

procedure TMainForm.DoRefreshWorkbooks(Sender: TObject);
var
  I: Integer;
  WB: OleVariant;
  Prev: string;
begin
  if cbWorkbook.ItemIndex >= 0 then
    Prev := FFullNames[cbWorkbook.ItemIndex];
  cbWorkbook.Items.Clear;
  FFullNames.Clear;
  lvComps.Items.Clear;
  try
    FWorkbooks := GetOpenWorkbooks;
  except
    on E: Exception do
    begin
      FWorkbooks := nil;
      Log('FEHLER beim Suchen der Arbeitsmappen: ' + E.Message);
    end;
  end;
  for WB in FWorkbooks do
  begin
    cbWorkbook.Items.Add(VarToStr(WB.Name));
    FFullNames.Add(VarToStr(WB.FullName));
  end;
  if cbWorkbook.Items.Count = 0 then
  begin
    Log('Keine geoeffnete Arbeitsmappe gefunden. Hinweis: VbaSync nicht "als Administrator" ' +
      'starten, wenn Excel normal laeuft (und umgekehrt).');
    Exit;
  end;
  Log(Format('%d Arbeitsmappe(n) gefunden.', [cbWorkbook.Items.Count]));
  I := FFullNames.IndexOf(Prev);
  if I < 0 then I := 0;
  cbWorkbook.ItemIndex := I;
  DoWorkbookChanged(nil);
end;

procedure TMainForm.DoWorkbookChanged(Sender: TObject);
var
  Full: string;
begin
  if cbWorkbook.ItemIndex < 0 then Exit;
  Full := FFullNames[cbWorkbook.ItemIndex];
  // Vorschlag: Unterordner "src" neben der Mappe (ungespeicherte Mappen haben keinen Pfad)
  if ExtractFilePath(Full) <> '' then
    edFolder.Text := TPath.Combine(ExtractFilePath(Full), 'src')
  else
    edFolder.Text := '';
  RefreshComponents;
end;

function TMainForm.CurrentWorkbook: OleVariant;
begin
  if (cbWorkbook.ItemIndex < 0) or (cbWorkbook.ItemIndex >= Length(FWorkbooks)) then
    raise Exception.Create('Keine Arbeitsmappe ausgewaehlt.');
  Result := FWorkbooks[cbWorkbook.ItemIndex];
end;

procedure TMainForm.RefreshComponents;
var
  Sync: TVbaSync;
  Info: TVbaComponentInfo;
  Item: TListItem;
begin
  lvComps.Items.BeginUpdate;
  try
    lvComps.Items.Clear;
    try
      Sync := TVbaSync.Create(CurrentWorkbook, Log);
      try
        for Info in Sync.GetComponents do
        begin
          Item := lvComps.Items.Add;
          Item.Caption := Info.Name;
          Item.SubItems.Add(TypeName(Info.CompType));
          Item.SubItems.Add(IntToStr(Info.LineCount));
        end;
      finally
        Sync.Free;
      end;
    except
      on E: Exception do
        Log('FEHLER: ' + E.Message);
    end;
  finally
    lvComps.Items.EndUpdate;
  end;
end;

function TMainForm.CheckedNames: TArray<string>;
var
  I: Integer;
begin
  // Haekchen haben Vorrang; ohne Haekchen zaehlt die Auswahl (Shift/Strg + Klick).
  Result := nil;
  for I := 0 to lvComps.Items.Count - 1 do
    if lvComps.Items[I].Checked then
      Result := Result + [lvComps.Items[I].Caption];
  if Length(Result) = 0 then
    for I := 0 to lvComps.Items.Count - 1 do
      if lvComps.Items[I].Selected then
        Result := Result + [lvComps.Items[I].Caption];
end;

function TMainForm.RequireFolder: string;
begin
  Result := Trim(edFolder.Text);
  if Result = '' then
    raise Exception.Create('Bitte einen Ordner angeben.');
end;

procedure TMainForm.RunAction(const Action: TProc<TVbaSync>);
var
  Sync: TVbaSync;
begin
  Screen.Cursor := crHourGlass;
  try
    try
      Sync := TVbaSync.Create(CurrentWorkbook, Log);
      try
        Action(Sync);
      finally
        Sync.Free;
      end;
    except
      on E: Exception do
      begin
        Log('FEHLER: ' + E.Message);
        MessageDlg(E.Message, mtError, [mbOK], 0);
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  RefreshComponents;
end;

function Confirm(const Msg: string): Boolean;
begin
  Result := MessageDlg(Msg, mtConfirmation, [mbYes, mbNo], 0) = mrYes;
end;

procedure TMainForm.DoBrowse(Sender: TObject);
var
  Dlg: TFileOpenDialog;
begin
  Dlg := TFileOpenDialog.Create(nil);
  try
    Dlg.Options := [fdoPickFolders];
    Dlg.Title := 'Ordner fuer die VBA-Dateien';
    if DirectoryExists(edFolder.Text) then
      Dlg.DefaultFolder := edFolder.Text;
    if Dlg.Execute then
      edFolder.Text := Dlg.FileName;
  finally
    Dlg.Free;
  end;
end;

procedure TMainForm.DoSelectAll(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to lvComps.Items.Count - 1 do
    lvComps.Items[I].Checked := True;
end;

procedure TMainForm.DoSelectNone(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to lvComps.Items.Count - 1 do
    lvComps.Items[I].Checked := False;
end;

procedure TMainForm.DoExport(Sender: TObject);
var
  Folder: string;
begin
  Folder := RequireFolder;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.ExportAll(Folder);
    end);
end;

procedure TMainForm.DoImportAll(Sender: TObject);
var
  Folder: string;
begin
  Folder := RequireFolder;
  if not Confirm('Alle Dateien aus' + sLineBreak + Folder + sLineBreak +
    'importieren? Gleichnamige Module werden ersetzt.') then Exit;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.ImportFolder(Folder, nil);
    end);
end;

procedure TMainForm.DoImportSelected(Sender: TObject);
var
  Folder: string;
  Names: TArray<string>;
begin
  Folder := RequireFolder;
  Names := CheckedNames;
  if Length(Names) = 0 then
  begin
    MessageDlg('Keine Komponente markiert oder ausgewaehlt.', mtInformation, [mbOK], 0);
    Exit;
  end;
  if not Confirm(Format('%d Komponente(n) aus dem Ordner neu importieren?', [Length(Names)]) + sLineBreak + sLineBreak + String.Join(', ', Names)) then Exit;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.ImportFolder(Folder, Names);
    end);
end;

procedure TMainForm.DoImportFiles(Sender: TObject);
var
  Dlg: TFileOpenDialog;
  FT: TFileTypeItem;
  Files: TArray<string>;
  I: Integer;
begin
  Dlg := TFileOpenDialog.Create(nil);
  try
    Dlg.Title := 'VBA-Dateien zum Importieren auswaehlen';
    Dlg.Options := [fdoAllowMultiSelect, fdoFileMustExist, fdoPathMustExist];
    FT := Dlg.FileTypes.Add;
    FT.DisplayName := 'VBA-Komponenten (*.bas;*.cls;*.frm;*.doccls)';
    FT.FileMask := '*.bas;*.cls;*.frm;*.doccls';
    FT := Dlg.FileTypes.Add;
    FT.DisplayName := 'Alle Dateien (*.*)';
    FT.FileMask := '*.*';
    if DirectoryExists(Trim(edFolder.Text)) then
      Dlg.DefaultFolder := Trim(edFolder.Text);
    if not Dlg.Execute then Exit;
    Files := nil;
    for I := 0 to Dlg.Files.Count - 1 do
      Files := Files + [Dlg.Files[I]];
  finally
    Dlg.Free;
  end;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.ImportFiles(Files);
    end);
end;

procedure TMainForm.DoDeleteSelected(Sender: TObject);
var
  Names: TArray<string>;
  ClearDocs: Boolean;
begin
  Names := CheckedNames;
  if Length(Names) = 0 then
  begin
    MessageDlg('Keine Komponente markiert oder ausgewaehlt.', mtInformation, [mbOK], 0);
    Exit;
  end;
  ClearDocs := chkClearDocs.Checked;
  if not Confirm(Format('%d Komponente(n) loeschen?', [Length(Names)]) + sLineBreak + sLineBreak + String.Join(', ', Names)) then Exit;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.DeleteComponents(Names, False, ClearDocs);
    end);
end;

procedure TMainForm.DoDeleteAll(Sender: TObject);
var
  ClearDocs: Boolean;
begin
  ClearDocs := chkClearDocs.Checked;
  if not Confirm('ALLE Module, Klassen und Formulare loeschen?' +
    IfThen(ClearDocs, sLineBreak + 'Der Code in Tabellen/DieseArbeitsmappe wird ebenfalls geleert.', '') +
    sLineBreak + sLineBreak + 'Tipp: vorher exportieren.') then Exit;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.DeleteComponents(nil, True, ClearDocs);
    end);
end;

procedure TMainForm.DoClean(Sender: TObject);
begin
  if not Confirm('Clean ausfuehren? Alles wird in einen Backup-Ordner exportiert, ' +
    'entfernt und neu importiert.') then Exit;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.Clean;
    end);
end;

end.
