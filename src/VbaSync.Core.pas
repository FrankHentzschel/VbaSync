unit VbaSync.Core;

{
  VbaSync - Kernlogik fuer Export / Import / Loeschen / Clean von Excel-VBA-Komponenten.

  Arbeitet bewusst per Late Binding (OleVariant), damit
    - keine Typbibliotheken (VBE6EXT.OLB, MSO.DLL) importiert werden muessen,
    - die Unit unveraendert von der EXE und spaeter vom VBE-Add-in genutzt werden kann,
    - 32- und 64-Bit-Office gleichermassen funktionieren.

  Voraussetzung in Excel:
    Datei > Optionen > Trust Center > Einstellungen fuer das Trust Center >
    Makroeinstellungen > "Zugriff auf das VBA-Projektobjektmodell vertrauen".

  Dateiformate im Zielordner:
    .bas     Standardmodul         (VBComponent.Export / VBComponents.Import)
    .cls     Klassenmodul          (VBComponent.Export / VBComponents.Import)
    .frm/.frx UserForm             (VBComponent.Export / VBComponents.Import)
    .doccls  Dokumentmodul (DieseArbeitsmappe, Tabelle1, ...): nur der reine Code,
             beim Import wird der Inhalt des vorhandenen Moduls ersetzt.
}

interface

uses
  System.SysUtils, System.Classes, System.Variants;

const
  ct_StdModule       = 1;
  ct_ClassModule     = 2;
  ct_MSForm          = 3;
  ct_ActiveXDesigner = 11;
  ct_Document        = 100;

  pp_Locked = 1; // vbext_pp_locked

type
  TLogProc = reference to procedure(const Msg: string);

  TVbaComponentInfo = record
    Name: string;
    CompType: Integer;
    LineCount: Integer;
  end;

  TVbaSync = class
  private
    FWorkbook: OleVariant;
    FProject: OleVariant;
    FLog: TLogProc;
    FDelCounter: Integer;
    procedure Log(const Msg: string);
    function FindComponent(const AName: string; out Comp: OleVariant): Boolean;
    procedure ReplaceDocumentCode(const Comp: OleVariant; const Code: string);
    procedure RemoveComponentSafe(const Comp: OleVariant);
  public
    constructor Create(const AWorkbook: OleVariant; ALog: TLogProc);
    procedure ListComponents;
    function GetComponents: TArray<TVbaComponentInfo>;
    /// Exportiert alle Komponenten. Vorhandene .bas/.cls/.frm/.frx/.doccls im Ordner
    /// werden vorher geloescht, damit entfernte Module auch in Git als geloescht erscheinen.
    procedure ExportAll(const Folder: string);
    /// Importiert alle (oder nur die genannten) Komponenten aus dem Ordner.
    procedure ImportFolder(const Folder: string; const Names: TArray<string>);
    /// Loescht genannte oder alle Nicht-Dokument-Komponenten.
    /// ClearDocs: leert zusaetzlich den Code der Dokumentmodule.
    procedure DeleteComponents(const Names: TArray<string>; All, ClearDocs: Boolean);
    /// Code-Cleaner: Export in Temp-Ordner, alles entfernen, neu importieren.
    /// Gibt den Backup-Ordner zurueck.
    function Clean: string;
  end;

function ComponentType(const Comp: OleVariant): Integer;
function ExtensionForType(T: Integer): string;
function TypeName(T: Integer): string;
/// Sucht eine geoeffnete Mappe per Name oder Pfad; oeffnet sie, falls die Datei existiert.
function OpenWorkbook(const ExcelApp: OleVariant; const Spec: string): OleVariant;
/// Liefert alle geoeffneten Arbeitsmappen ueber ALLE laufenden Excel-Instanzen.
/// Sucht ueber die Excel-Fenster (XLMAIN/XLDESK/EXCEL7) statt GetActiveOleObject,
/// weil Letzteres nur die erste registrierte (evtl. leere/unsichtbare) Instanz findet.
function GetOpenWorkbooks: TArray<OleVariant>;

implementation

uses
  System.IOUtils, System.StrUtils, System.Generics.Collections, Winapi.Windows, Winapi.ActiveX, System.Win.ComObj;

const
  OBJID_NATIVEOM = DWORD($FFFFFFF0);

function AccessibleObjectFromWindow(hwnd: HWND; dwId: DWORD; const riid: TGUID;
  out ppvObject): HResult; stdcall; external 'oleacc.dll';

function EnumXlMain(Wnd: HWND; Param: LPARAM): BOOL; stdcall;
var
  Buf: array[0..63] of Char;
begin
  if (GetClassName(Wnd, Buf, Length(Buf)) > 0) and (StrComp(Buf, 'XLMAIN') = 0) then
    TList<HWND>(Param).Add(Wnd);
  Result := True;
end;

function GetOpenWorkbooks: TArray<OleVariant>;
var
  Mains: TList<HWND>;
  Main, Desk, Child: HWND;
  Disp: IDispatch;
  WB: OleVariant;
  Seen: TStringList;
  Full: string;
begin
  Result := nil;
  Mains := TList<HWND>.Create;
  Seen := TStringList.Create;
  try
    EnumWindows(@EnumXlMain, LPARAM(Mains));
    Seen.CaseSensitive := False;
    for Main in Mains do
    begin
      Desk := FindWindowEx(Main, 0, 'XLDESK', nil);
      if Desk = HWND(0) then Continue;
      Child := FindWindowEx(Desk, 0, 'EXCEL7', nil);
      while Child <> HWND(0) do
      begin
        Disp := nil;
        if Succeeded(AccessibleObjectFromWindow(Child, OBJID_NATIVEOM, IDispatch, Disp))
          and Assigned(Disp) then
        try
          WB := OleVariant(Disp).Parent; // Excel.Window.Parent = Workbook
          Full := VarToStr(WB.FullName);
          if Seen.IndexOf(Full) < 0 then
          begin
            Seen.Add(Full);
            Result := Result + [WB];
          end;
        except
          // Fenster ohne Zugriff (z. B. anderes Rechte-Level) ignorieren
        end;
        Child := FindWindowEx(Desk, Child, 'EXCEL7', nil);
      end;
    end;
  finally
    Seen.Free;
    Mains.Free;
  end;
end;

const
  ManagedExts: array[0..4] of string = ('.bas', '.cls', '.frm', '.frx', '.doccls');
  ImportExts:  array[0..3] of string = ('.bas', '.cls', '.frm', '.doccls');

{ "Type" ist in Pascal ein reserviertes Wort, daher Property-Get direkt ueber IDispatch. }
function OleGetProp(const Obj: OleVariant; const Name: WideString): OleVariant;
var
  Disp: IDispatch;
  DispId: Integer;
  Params: TDispParams;
  PName: PWideChar;
  ExcepInfo: TExcepInfo;
  ArgErr: Cardinal;
begin
  Disp := IDispatch(Obj);
  PName := PWideChar(Name);
  OleCheck(Disp.GetIDsOfNames(GUID_NULL, @PName, 1, LOCALE_USER_DEFAULT, @DispId));
  FillChar(Params, SizeOf(Params), 0);
  FillChar(ExcepInfo, SizeOf(ExcepInfo), 0);
  OleCheck(Disp.Invoke(DispId, GUID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
    Params, @Result, @ExcepInfo, @ArgErr));
end;

function ComponentType(const Comp: OleVariant): Integer;
begin
  Result := OleGetProp(Comp, 'Type');
end;

function ExtensionForType(T: Integer): string;
begin
  case T of
    ct_StdModule:   Result := '.bas';
    ct_ClassModule: Result := '.cls';
    ct_MSForm:      Result := '.frm';
    ct_Document:    Result := '.doccls';
  else
    Result := '';
  end;
end;

function TypeName(T: Integer): string;
begin
  case T of
    ct_StdModule:       Result := 'Modul';
    ct_ClassModule:     Result := 'Klasse';
    ct_MSForm:          Result := 'UserForm';
    ct_ActiveXDesigner: Result := 'Designer';
    ct_Document:        Result := 'Dokument';
  else
    Result := 'Typ ' + IntToStr(T);
  end;
end;

function OpenWorkbook(const ExcelApp: OleVariant; const Spec: string): OleVariant;
var
  I: Integer;
  WB: OleVariant;
  FullSpec: string;
begin
  FullSpec := ExpandFileName(Spec);
  for I := 1 to ExcelApp.Workbooks.Count do
  begin
    WB := ExcelApp.Workbooks.Item(I);
    if SameText(VarToStr(WB.Name), Spec) or SameText(VarToStr(WB.FullName), FullSpec) then
      Exit(WB);
  end;
  // Add-ins (.xlam) tauchen bei der Aufzaehlung nicht auf, sind aber per Name erreichbar.
  try
    Result := ExcelApp.Workbooks.Item(ExtractFileName(Spec));
    Exit;
  except
    // nicht gefunden -> weiter
  end;
  if FileExists(FullSpec) then
    Result := ExcelApp.Workbooks.Open(FullSpec)
  else
    raise Exception.CreateFmt('Arbeitsmappe "%s" ist nicht geoeffnet und als Datei nicht vorhanden.', [Spec]);
end;

function ReadVbName(const FileName: string): string;
var
  Line: string;
  P: Integer;
begin
  Result := TPath.GetFileNameWithoutExtension(FileName);
  for Line in TFile.ReadAllLines(FileName, TEncoding.ANSI) do
    if StartsText('Attribute VB_Name', Line) then
    begin
      P := Pos('"', Line);
      if P > 0 then
        Result := Copy(Line, P + 1, LastDelimiter('"', Line) - P - 1);
      Exit;
    end;
end;

{ Liest eine .doccls-Datei. Toleriert auch eine per VBE exportierte .cls eines
  Dokumentmoduls (VERSION/BEGIN/END-Kopf und Attribute-Zeilen werden entfernt). }
function ReadDocCode(const FileName: string): string;
var
  Lines: TArray<string>;
  SB: TStringBuilder;
  I: Integer;
begin
  Lines := TFile.ReadAllLines(FileName, TEncoding.ANSI);
  I := 0;
  if (Length(Lines) > 0) and StartsText('VERSION ', Lines[0]) then
  begin
    while (I < Length(Lines)) and not SameText(Trim(Lines[I]), 'END') do
      Inc(I);
    Inc(I); // END selbst ueberspringen
  end;
  SB := TStringBuilder.Create;
  try
    while I < Length(Lines) do
    begin
      if not StartsText('Attribute VB_', Lines[I]) then
        SB.Append(Lines[I]).Append(sLineBreak);
      Inc(I);
    end;
    Result := SB.ToString.TrimRight([#13, #10]);
  finally
    SB.Free;
  end;
end;

{ TVbaSync }

constructor TVbaSync.Create(const AWorkbook: OleVariant; ALog: TLogProc);
begin
  inherited Create;
  FWorkbook := AWorkbook;
  FLog := ALog;
  try
    FProject := FWorkbook.VBProject;
  except
    on E: Exception do
      raise Exception.Create('Kein Zugriff auf das VBA-Projekt. Bitte in Excel ' +
        '"Zugriff auf das VBA-Projektobjektmodell vertrauen" aktivieren. (' + E.Message + ')');
  end;
  if Integer(FProject.Protection) = pp_Locked then
    raise Exception.Create('Das VBA-Projekt ist kennwortgeschuetzt und gesperrt.');
end;

procedure TVbaSync.Log(const Msg: string);
begin
  if Assigned(FLog) then
    FLog(Msg);
end;

function TVbaSync.FindComponent(const AName: string; out Comp: OleVariant): Boolean;
var
  I: Integer;
  C: OleVariant;
begin
  for I := 1 to FProject.VBComponents.Count do
  begin
    C := FProject.VBComponents.Item(I);
    if SameText(VarToStr(C.Name), AName) then
    begin
      Comp := C;
      Exit(True);
    end;
  end;
  Comp := Unassigned;
  Result := False;
end;

procedure TVbaSync.ReplaceDocumentCode(const Comp: OleVariant; const Code: string);
var
  CM: OleVariant;
  N: Integer;
begin
  CM := Comp.CodeModule;
  N := CM.CountOfLines;
  if N > 0 then
    CM.DeleteLines(1, N);
  if Code <> '' then
    CM.AddFromString(Code);
end;

{ Vor dem Entfernen umbenennen: verhindert Namenskonflikte ("Modul11"),
  falls VBE das Entfernen verzoegert ausfuehrt. }
procedure TVbaSync.RemoveComponentSafe(const Comp: OleVariant);
begin
  Inc(FDelCounter);
  Comp.Name := Format('zzVbaSyncDel%d', [FDelCounter]);
  FProject.VBComponents.Remove(Comp);
end;

function TVbaSync.GetComponents: TArray<TVbaComponentInfo>;
var
  I: Integer;
  C: OleVariant;
  Info: TVbaComponentInfo;
begin
  Result := nil;
  for I := 1 to FProject.VBComponents.Count do
  begin
    C := FProject.VBComponents.Item(I);
    Info.Name := VarToStr(C.Name);
    Info.CompType := ComponentType(C);
    Info.LineCount := C.CodeModule.CountOfLines;
    Result := Result + [Info];
  end;
end;

procedure TVbaSync.ListComponents;
var
  I: Integer;
  C: OleVariant;
begin
  Log(Format('Projekt "%s" in %s', [VarToStr(FProject.Name), VarToStr(FWorkbook.FullName)]));
  for I := 1 to FProject.VBComponents.Count do
  begin
    C := FProject.VBComponents.Item(I);
    Log(Format('  %-10s %-32s %6d Zeilen',
      [TypeName(ComponentType(C)), VarToStr(C.Name), Integer(C.CodeModule.CountOfLines)]));
  end;
end;

procedure TVbaSync.ExportAll(const Folder: string);
var
  F, Ext, Name, Path, Code: string;
  I, T, N, Count: Integer;
  C, CM: OleVariant;
begin
  ForceDirectories(Folder);
  for F in TDirectory.GetFiles(Folder) do
    if MatchText(ExtractFileExt(F), ManagedExts) then
      TFile.Delete(F);

  Count := 0;
  for I := 1 to FProject.VBComponents.Count do
  begin
    C := FProject.VBComponents.Item(I);
    T := ComponentType(C);
    Name := VarToStr(C.Name);
    Ext := ExtensionForType(T);
    if Ext = '' then
    begin
      Log(Format('  uebersprungen: %s (%s)', [Name, TypeName(T)]));
      Continue;
    end;
    Path := TPath.Combine(Folder, Name + Ext);
    if T = ct_Document then
    begin
      CM := C.CodeModule;
      N := CM.CountOfLines;
      if N = 0 then
        Code := ''
      else
        Code := VarToStr(CM.Lines[1, N]).TrimRight([#13, #10]);
      if Trim(Code) = '' then
        Continue; // leere Dokumentmodule nicht exportieren
      TFile.WriteAllText(Path, Code + sLineBreak, TEncoding.ANSI);
    end
    else
      C.Export(Path);
    Inc(Count);
    Log('  exportiert: ' + ExtractFileName(Path));
  end;
  Log(Format('%d Komponenten nach %s exportiert.', [Count, Folder]));
end;

procedure TVbaSync.ImportFolder(const Folder: string; const Names: TArray<string>);
var
  F, Ext, Name: string;
  C: OleVariant;
  Count: Integer;
begin
  if not TDirectory.Exists(Folder) then
    raise Exception.CreateFmt('Ordner "%s" existiert nicht.', [Folder]);
  Count := 0;
  for F in TDirectory.GetFiles(Folder) do
  begin
    Ext := LowerCase(ExtractFileExt(F));
    if not MatchText(Ext, ImportExts) then
      Continue;
    if Ext = '.doccls' then
      Name := TPath.GetFileNameWithoutExtension(F)
    else
      Name := ReadVbName(F);
    if (Length(Names) > 0) and not MatchText(Name, Names) then
      Continue;

    if Ext = '.doccls' then
    begin
      if FindComponent(Name, C) and (ComponentType(C) = ct_Document) then
      begin
        ReplaceDocumentCode(C, ReadDocCode(F));
        Log('  ersetzt (Dokument): ' + Name);
        Inc(Count);
      end
      else
        Log('  WARNUNG: Dokumentmodul "' + Name + '" gibt es in der Mappe nicht - uebersprungen.');
      Continue;
    end;

    if (Ext = '.frm') and not FileExists(ChangeFileExt(F, '.frx')) then
      Log('  WARNUNG: ' + ExtractFileName(F) + ' ohne .frx - Formular evtl. unvollstaendig.');

    if FindComponent(Name, C) then
    begin
      if ComponentType(C) = ct_Document then
      begin
        Log('  WARNUNG: "' + Name + '" ist ein Dokumentmodul, Datei ' + ExtractFileName(F) + ' uebersprungen.');
        Continue;
      end;
      RemoveComponentSafe(C);
    end;
    FProject.VBComponents.Import(F);
    Log('  importiert: ' + ExtractFileName(F));
    Inc(Count);
  end;
  Log(Format('%d Komponenten aus %s importiert.', [Count, Folder]));
end;

procedure TVbaSync.DeleteComponents(const Names: TArray<string>; All, ClearDocs: Boolean);
var
  I, T: Integer;
  C: OleVariant;
  ToRemove: TArray<OleVariant>;
  Name: string;
begin
  // erst sammeln, dann entfernen (Indizes verschieben sich beim Entfernen)
  for I := 1 to FProject.VBComponents.Count do
  begin
    C := FProject.VBComponents.Item(I);
    Name := VarToStr(C.Name);
    if not (All or MatchText(Name, Names)) then
      Continue;
    T := ComponentType(C);
    if T = ct_Document then
    begin
      if ClearDocs then
      begin
        ReplaceDocumentCode(C, '');
        Log('  geleert (Dokument): ' + Name);
      end
      else if not All then
        Log('  Hinweis: "' + Name + '" ist ein Dokumentmodul; nur mit --clear-docs leerbar.');
    end
    else
      ToRemove := ToRemove + [C];
  end;
  for C in ToRemove do
  begin
    Name := VarToStr(C.Name);
    RemoveComponentSafe(C);
    Log('  geloescht: ' + Name);
  end;
  Log(Format('%d Komponenten geloescht.', [Length(ToRemove)]));
end;

function TVbaSync.Clean: string;
begin
  Result := TPath.Combine(TPath.GetTempPath,
    'VbaSync_' + ChangeFileExt(VarToStr(FWorkbook.Name), '') + '_' + FormatDateTime('yyyymmdd_hhnnss', Now));
  Log('Backup/Export nach ' + Result);
  ExportAll(Result);
  DeleteComponents(nil, True, False);
  ImportFolder(Result, nil);
  Log('Clean abgeschlossen. Mappe jetzt speichern. Backup bleibt in ' + Result);
end;

end.
