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
    FOwnerWnd: HWND;       // Add-in: Fenster der VBA-IDE als Besitzer
    FPreferred: string;    // Add-in: Pfad der aktiven Mappe vorauswaehlen
    FFullNames: TStringList;
    cbWorkbook: TComboBox;
    btnRefresh: TButton;
    edFolder: TEdit;
    btnBrowse: TButton;
    lvComps: TListView;
    chkClearDocs: TCheckBox;
    memLog: TMemo;
    pnlTop, pnlRight: TPanel;
    lblWorkbook, lblFolder: TLabel;
    FSideCtrls: TArray<TControl>;  // Schaltflaechen rechts, von oben nach unten
    FSideGaps: TArray<Integer>;    // Abstand davor: 0 normal, 1 Gruppe, 2 Info
    procedure BuildUI;
    procedure LayoutControls;
    procedure DoDpiChanged(Sender: TObject; OldDPI, NewDPI: Integer);
    procedure AddSide(AControl: TControl; AGap: Integer);
    function AddButton(const ACaption: string; AGap: Integer;
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
    procedure DoCompare(Sender: TObject);
    procedure DoImportAll(Sender: TObject);
    procedure DoImportSelected(Sender: TObject);
    procedure DoImportFiles(Sender: TObject);
    procedure DoDeleteSelected(Sender: TObject);
    procedure DoDeleteAll(Sender: TObject);
    procedure DoClean(Sender: TObject);
    procedure HandleFirstShow(Sender: TObject);
    procedure ApplyAppIcon;
    procedure DoAbout(Sender: TObject);
  protected
    procedure CreateParams(var Params: TCreateParams); override;
  public
    constructor Create(AOwner: TComponent); override;
    /// Aufruf aus dem VBE-Add-in: Besitzerfenster und vorzuwaehlende Mappe
    constructor CreateForAddin(AOwnerWnd: HWND; const APreferredFullName: string);
    destructor Destroy; override;
  end;

var
  MainForm: TMainForm;

implementation

{$R 'VbaSyncIcon.res'}  // Programmsymbol als Ressource VBASYNCICON (16-256 px)

uses
  System.IOUtils, System.UITypes, System.StrUtils, System.Win.ComObj, Vcl.Graphics;

const
  FormW = 920;

{ TMainForm }

procedure TMainForm.CreateParams(var Params: TCreateParams);
begin
  inherited;
  // Eigenstaendiges Top-Level-Fenster: erscheint in Alt+Tab und in der Taskleiste.
  // (Das Fenster wird schon im Konstruktor erzeugt, bevor Application.MainForm
  //  gesetzt ist; dadurch griff MainFormOnTaskbar nicht.)
  if FOwnerWnd <> HWND(0) then
    Params.WndParent := FOwnerWnd  // im Add-in: ueber der VBA-IDE
  else
  begin
    Params.ExStyle := Params.ExStyle or WS_EX_APPWINDOW;
    Params.WndParent := 0;
  end;
end;

procedure TMainForm.ApplyAppIcon;

  function Load(Size: Integer): HICON;
  begin
    Result := LoadImage(HInstance, 'VBASYNCICON', IMAGE_ICON, Size, Size, LR_DEFAULTCOLOR);
  end;

var
  H: HICON;
begin
  // Fenster- und Taskleistensymbol unabhaengig von den Projektoptionen setzen.
  // TIcon uebernimmt das Handle und gibt es frei, daher je ein eigenes Handle.
  H := Load(GetSystemMetrics(SM_CXICON));
  if H = HICON(0) then Exit;
  Icon.Handle := H;
  Application.Icon.Handle := Load(GetSystemMetrics(SM_CXICON));
end;

constructor TMainForm.CreateForAddin(AOwnerWnd: HWND; const APreferredFullName: string);
begin
  FOwnerWnd := AOwnerWnd;
  FPreferred := APreferredFullName;
  Create(nil);
end;

constructor TMainForm.Create(AOwner: TComponent);
begin
  inherited CreateNew(AOwner);
  FFullNames := TStringList.Create;
  ApplyAppIcon;
  BuildUI;
  LayoutControls;
  OnShow := HandleFirstShow;
  OnAfterMonitorDpiChanged := DoDpiChanged;
end;

destructor TMainForm.Destroy;
begin
  FFullNames.Free;
  FWorkbooks := nil;
  inherited;
end;

procedure TMainForm.AddSide(AControl: TControl; AGap: Integer);
begin
  AControl.Parent := pnlRight;
  FSideCtrls := FSideCtrls + [AControl];
  FSideGaps := FSideGaps + [AGap];
end;

function TMainForm.AddButton(const ACaption: string; AGap: Integer;
  AOnClick: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Caption := ACaption;
  Result.OnClick := AOnClick;
  AddSide(Result, AGap);
end;

{ Ordnet Kopfbereich und Schaltflaechen an. Alle Schaltflaechen (auch
  Aktualisieren/Durchsuchen) erhalten dieselbe Breite, die sich nach der
  laengsten Beschriftung in der aktuellen Schrift richtet. Wird nach dem
  Anzeigen und bei DPI-Wechsel erneut aufgerufen. }
procedure TMainForm.LayoutControls;

  function S(V: Integer): Integer;
  begin
    Result := MulDiv(V, CurrentPPI, 96);
  end;

  procedure Widen(var W: Integer; TW: Integer);
  begin
    if TW > W then W := TW;
  end;

var
  I, W, H, Y, M, LblW, Row1, Row2, X: Integer;
  C: TControl;
const
  Gaps: array[0..2] of Integer = (4, 16, 28);
begin
  if pnlRight = nil then Exit;
  Canvas.Font := Font;
  M := S(8);
  H := Canvas.TextHeight('Ag') + S(12);
  if H < S(26) then H := S(26);

  // Einheitliche Breite: laengste Beschriftung plus Rand
  W := S(150);
  for I := 0 to High(FSideCtrls) do
  begin
    C := FSideCtrls[I];
    if C is TButton then
      Widen(W, Canvas.TextWidth(TButton(C).Caption) + S(28))
    else if C is TCheckBox then
      Widen(W, Canvas.TextWidth(TCheckBox(C).Caption) + GetSystemMetrics(SM_CXMENUCHECK) + S(16));
  end;
  Widen(W, Canvas.TextWidth(btnRefresh.Caption) + S(28));
  Widen(W, Canvas.TextWidth(btnBrowse.Caption) + S(28));

  // Kopfbereich
  LblW := Canvas.TextWidth(lblWorkbook.Caption);
  Widen(LblW, Canvas.TextWidth(lblFolder.Caption));
  Inc(LblW, S(12));
  Row1 := M;
  Row2 := Row1 + H + S(6);
  pnlTop.Height := Row2 + H + M;
  X := pnlTop.ClientWidth - M - W;

  btnRefresh.SetBounds(X, Row1, W, H);
  btnBrowse.SetBounds(X, Row2, W, H);
  cbWorkbook.SetBounds(M + LblW, Row1 + (H - cbWorkbook.Height) div 2,
    X - M - (M + LblW), cbWorkbook.Height);
  edFolder.SetBounds(M + LblW, Row2 + (H - edFolder.Height) div 2,
    X - M - (M + LblW), edFolder.Height);
  lblWorkbook.SetBounds(M, Row1 + (H - lblWorkbook.Height) div 2, LblW - S(4), lblWorkbook.Height);
  lblFolder.SetBounds(M, Row2 + (H - lblFolder.Height) div 2, LblW - S(4), lblFolder.Height);

  // Schaltflaechen rechts, Info-Button direkt unter der letzten
  pnlRight.Width := W + 2 * M;
  Y := 0;
  for I := 0 to High(FSideCtrls) do
  begin
    Inc(Y, S(Gaps[FSideGaps[I]]));
    FSideCtrls[I].SetBounds(M, Y, W, H);
    Inc(Y, H);
  end;
end;

procedure TMainForm.DoDpiChanged(Sender: TObject; OldDPI, NewDPI: Integer);
begin
  LayoutControls;
end;

procedure TMainForm.BuildUI;
var
  Col: TListColumn;
begin
  Caption := AppName + ' ' + AppVersion + ' - Excel-VBA Export / Import';
  Position := poScreenCenter;
  ClientWidth := FormW;
  ClientHeight := 640;
  Constraints.MinWidth := 700;
  Constraints.MinHeight := 450;

  { Kopfbereich: Arbeitsmappe und Ordner (Positionen in LayoutControls) }
  pnlTop := TPanel.Create(Self);
  pnlTop.Parent := Self;
  pnlTop.Align := alTop;
  pnlTop.Height := 76;
  pnlTop.BevelOuter := bvNone;

  lblWorkbook := TLabel.Create(Self);
  lblWorkbook.Parent := pnlTop;
  lblWorkbook.Caption := 'Arbeitsmappe:';

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

  lblFolder := TLabel.Create(Self);
  lblFolder.Parent := pnlTop;
  lblFolder.Caption := 'Ordner:';

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

  { Schaltflaechen rechts (Positionen in LayoutControls) }
  pnlRight := TPanel.Create(Self);
  pnlRight.Parent := Self;
  pnlRight.Align := alRight;
  pnlRight.Width := 180;
  pnlRight.BevelOuter := bvNone;

  AddButton('Alle markieren', 0, DoSelectAll);
  AddButton('Keine markieren', 0, DoSelectNone);
  AddButton('Alle exportieren', 1, DoExport);
  AddButton('Mit Ordner vergleichen', 0, DoCompare);
  AddButton('Alle importieren', 1, DoImportAll);
  AddButton('Markierte importieren', 0, DoImportSelected);
  AddButton('Dateien importieren...', 0, DoImportFiles);
  AddButton('Markierte loeschen', 1, DoDeleteSelected);
  AddButton('Alle loeschen', 0, DoDeleteAll);

  chkClearDocs := TCheckBox.Create(Self);
  chkClearDocs.Caption := 'Tabellen-Code leeren';
  chkClearDocs.Hint := 'Beim Loeschen auch den Code in Tabellen und DieseArbeitsmappe leeren';
  chkClearDocs.ShowHint := True;
  AddSide(chkClearDocs, 0);

  AddButton('Clean (Code Cleaner)', 1, DoClean);
  AddButton('Info...', 2, DoAbout);

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
  Lbl: TLabel;
  Img: TImage;
  hIco: HICON;
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

    // Programmsymbol (Ressource VBASYNCICON) in 64x64
    hIco := LoadImage(HInstance, 'VBASYNCICON', IMAGE_ICON, 64, 64, LR_DEFAULTCOLOR);
    if hIco <> HICON(0) then
    begin
      Img := TImage.Create(Dlg);
      Img.Parent := Dlg;
      Img.SetBounds(24, 20, 64, 64);
      Img.Picture.Icon.Handle := hIco; // TIcon gibt das Handle frei
    end;

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
  LayoutControls;
  DoRefreshWorkbooks(nil);
end;

procedure TMainForm.DoRefreshWorkbooks(Sender: TObject);
var
  I: Integer;
  WB: OleVariant;
  Prev: string;
begin
  if cbWorkbook.ItemIndex >= 0 then
    Prev := FFullNames[cbWorkbook.ItemIndex]
  else
    Prev := FPreferred;
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

procedure TMainForm.DoCompare(Sender: TObject);
var
  Folder: string;
begin
  Folder := RequireFolder;
  RunAction(
    procedure(S: TVbaSync)
    begin
      S.CompareWithFolder(Folder);
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
