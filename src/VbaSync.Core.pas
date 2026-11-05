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
    FQuiet: Boolean;  // Log unterdruecken (Vergleichs-Export in Temp-Ordner)
    procedure Log(const Msg: string);
    function FindComponent(const AName: string; out Comp: OleVariant): Boolean;
    procedure ReplaceDocumentCode(const Comp: OleVariant; const Code: string);
    procedure RemoveComponentSafe(const Comp: OleVariant);
    procedure EnsureGitAttributes(const Folder: string);
  public
    constructor Create(const AWorkbook: OleVariant; ALog: TLogProc);
    procedure ListComponents;
    function GetComponents: TArray<TVbaComponentInfo>;
    /// Exportiert alle Komponenten. Vorhandene .bas/.cls/.frm/.frx/.doccls im Ordner
    /// werden vorher geloescht, damit entfernte Module auch in Git als geloescht erscheinen.
    procedure ExportAll(const Folder: string);
    /// Importiert alle (oder nur die genannten) Komponenten aus dem Ordner.
    procedure ImportFolder(const Folder: string; const Names: TArray<string>);
    /// Importiert einzelne Dateien: vorhandene Komponenten werden ersetzt,
    /// fehlende hinzugefuegt, Dokumentmodule bekommen nur neuen Code.
    procedure ImportFiles(const Files: TArray<string>);
    function ImportFile(const F: string): Boolean;
    /// Loescht genannte oder alle Nicht-Dokument-Komponenten.
    /// ClearDocs: leert zusaetzlich den Code der Dokumentmodule.
    procedure DeleteComponents(const Names: TArray<string>; All, ClearDocs: Boolean);
    /// Code-Cleaner: Export in Temp-Ordner, alles entfernen, neu importieren.
    /// Gibt den Backup-Ordner zurueck.
    function Clean: string;
    /// Vergleicht den Code der Mappe mit den Dateien im Ordner (Export in einen
    /// Temp-Ordner, dann byteweiser Vergleich). Liefert die Anzahl der Abweichungen.
    function CompareWithFolder(const Folder: string): Integer;
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
  if Assigned(FLog) and not FQuiet then
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
  EnsureGitAttributes(Folder);
end;

{ Legt im Exportordner eine .gitattributes an, falls dort noch keine liegt:
  - eol=crlf: die VBA-IDE importiert Dateien mit reinen LF-Zeilenenden fehlerhaft
  - working-tree-encoding: Git speichert UTF-8 im Repository (Umlaute in
    GitHub-Diffs lesbar) und schreibt die Dateien lokal in der ANSI-Codepage,
    die der VBE-Export und -Import verwenden
  - .frx ist binaer
  Eine vorhandene Datei wird nie veraendert. }
procedure TVbaSync.EnsureGitAttributes(const Folder: string);
const
  Exts: array[0..3] of string = ('bas', 'cls', 'frm', 'doccls');
var
  Path, Enc, Txt: string;
  E: string;
  CP: Cardinal;
begin
  Path := TPath.Combine(Folder, '.gitattributes');
  if TFile.Exists(Path) then
    Exit;
  CP := GetACP;
  if CP = CP_UTF8 then
    Enc := ''  // System arbeitet bereits mit UTF-8: keine Umkodierung
  else
    Enc := ' working-tree-encoding=CP' + IntToStr(CP);
  Txt := '# Angelegt von ' + 'VbaSync: Einstellungen fuer exportierten VBA-Code' + sLineBreak;
  for E in Exts do
    Txt := Txt + Format('*.%-7s text eol=crlf%s', [E, Enc]) + sLineBreak;
  Txt := Txt + '*.frx    binary' + sLineBreak;
  TFile.WriteAllText(Path, Txt, TEncoding.ASCII);
  Log('  angelegt: .gitattributes (CRLF' + IfThen(Enc <> '', ', CP' + IntToStr(CP), '') + ')');
end;

function TVbaSync.ImportFile(const F: string): Boolean;
var
  Ext, Name: string;
  C: OleVariant;
begin
  Result := False;
  Ext := LowerCase(ExtractFileExt(F));
  if not MatchText(Ext, ImportExts) then
  begin
    if Ext <> '.frx' then // .frx gehoert zur .frm und wird mit ihr importiert
      Log('  uebersprungen (Dateityp): ' + ExtractFileName(F));
    Exit;
  end;
  if Ext = '.doccls' then
    Name := TPath.GetFileNameWithoutExtension(F)
  else
    Name := ReadVbName(F);

  // Dokumentmodul (DieseArbeitsmappe, Tabelle1 ...): nur den Code ersetzen.
  // Gilt auch fuer eine per VBE exportierte .cls eines Dokumentmoduls.
  if FindComponent(Name, C) and (ComponentType(C) = ct_Document) then
  begin
    if Ext = '.frm' then
    begin
      Log('  WARNUNG: "' + Name + '" ist ein Dokumentmodul, ' + ExtractFileName(F) + ' uebersprungen.');
      Exit;
    end;
    ReplaceDocumentCode(C, ReadDocCode(F));
    Log('  ersetzt (Dokument): ' + Name);
    Exit(True);
  end;
  if Ext = '.doccls' then
  begin
    Log('  WARNUNG: Dokumentmodul "' + Name + '" gibt es in der Mappe nicht - uebersprungen ' +
      '(Tabellenblaetter legt VbaSync nicht an).');
    Exit;
  end;

  if (Ext = '.frm') and not FileExists(ChangeFileExt(F, '.frx')) then
    Log('  WARNUNG: ' + ExtractFileName(F) + ' ohne .frx - Formular evtl. unvollstaendig.');

  if not VarIsEmpty(C) then
  begin
    RemoveComponentSafe(C);
    FProject.VBComponents.Import(F);
    Log('  ersetzt: ' + Name + '  (' + ExtractFileName(F) + ')');
  end
  else
  begin
    FProject.VBComponents.Import(F);
    Log('  neu hinzugefuegt: ' + Name + '  (' + ExtractFileName(F) + ')');
  end;
  Result := True;
end;

procedure TVbaSync.ImportFiles(const Files: TArray<string>);
var
  F: string;
  Count: Integer;
begin
  Count := 0;
  for F in Files do
    if ImportFile(F) then
      Inc(Count);
  Log(Format('%d von %d Datei(en) importiert.', [Count, Length(Files)]));
end;

procedure TVbaSync.ImportFolder(const Folder: string; const Names: TArray<string>);
var
  F, Ext, Name: string;
  Files: TArray<string>;
begin
  if not TDirectory.Exists(Folder) then
    raise Exception.CreateFmt('Ordner "%s" existiert nicht.', [Folder]);
  Files := nil;
  for F in TDirectory.GetFiles(Folder) do
  begin
    Ext := LowerCase(ExtractFileExt(F));
    if not MatchText(Ext, ImportExts) then
      Continue;
    if Length(Names) > 0 then
    begin
      if Ext = '.doccls' then
        Name := TPath.GetFileNameWithoutExtension(F)
      else
        Name := ReadVbName(F);
      if not MatchText(Name, Names) then
        Continue;
    end;
    Files := Files + [F];
  end;
  Log('Import aus ' + Folder);
  ImportFiles(Files);
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

function SameFileBytes(const A, B: string): Boolean;
var
  BA, BB: TBytes;
begin
  BA := TFile.ReadAllBytes(A);
  BB := TFile.ReadAllBytes(B);
  Result := (Length(BA) = Length(BB)) and
    ((Length(BA) = 0) or CompareMem(@BA[0], @BB[0], Length(BA)));
end;

function TVbaSync.CompareWithFolder(const Folder: string): Integer;
var
  Tmp, F, N, Other: string;
begin
  if not TDirectory.Exists(Folder) then
    raise Exception.Create('Ordner nicht gefunden: ' + Folder);
  Tmp := TPath.Combine(TPath.GetTempPath,
    'VbaSync_cmp_' + FormatDateTime('yyyymmdd_hhnnsszzz', Now));
  Log('Vergleiche Mappe mit ' + Folder + ' ...');
  FQuiet := True;
  try
    ExportAll(Tmp);
  finally
    FQuiet := False;
  end;
  Result := 0;
  try
    for F in TDirectory.GetFiles(Tmp) do
    begin
      N := ExtractFileName(F);
      if not MatchText(ExtractFileExt(N), ManagedExts) then Continue;
      Other := TPath.Combine(Folder, N);
      if not TFile.Exists(Other) then
      begin
        Log('  nur in der Mappe: ' + N);
        Inc(Result);
      end
      else if not SameFileBytes(F, Other) then
      begin
        Log('  abweichend:       ' + N);
        Inc(Result);
      end;
    end;
    for F in TDirectory.GetFiles(Folder) do
    begin
      N := ExtractFileName(F);
      if MatchText(ExtractFileExt(N), ManagedExts) and
         not TFile.Exists(TPath.Combine(Tmp, N)) then
      begin
        Log('  nur im Ordner:    ' + N);
        Inc(Result);
      end;
    end;
  finally
    TDirectory.Delete(Tmp, True);
  end;
  if Result = 0 then
    Log('Vergleich: Mappe und Ordner sind identisch.')
  else
    Log(Format('Vergleich: %d Abweichung(en).', [Result]));
end;

end.
