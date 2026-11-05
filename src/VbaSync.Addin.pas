unit VbaSync.Addin;

{
  VBE-Add-in (COM, IDTExtensibility2) fuer den VBA-Editor von Excel.
  - Fuegt der Menueleiste der VBA-IDE einen Eintrag "VbaSync..." hinzu.
  - Ein Klick oeffnet das VbaSync-Fenster (modal) mit dem aktiven VBA-Projekt.
  - Registrierung pro Benutzer (HKCU), kein Administrator noetig:
      regsvr32 VbaSyncAddin.dll        (registrieren)
      regsvr32 /u VbaSyncAddin.dll     (entfernen)
    32-Bit-Office: DLL als Win32 bauen, 64-Bit-Office: als Win64.
}

interface

uses
  Winapi.Windows, Winapi.ActiveX, System.SysUtils, System.Classes, System.Variants,
  System.Win.ComObj, System.Win.ComServ, System.Win.Registry;

const
  CLASS_VbaSyncConnect: TGUID = '{1BD53B9E-6B23-48F4-A70E-23E53A38AC3F}';
  AddinProgID = 'VbaSync.Connect';

type
  ext_ConnectMode = TOleEnum;
  ext_DisconnectMode = TOleEnum;

  // aus der Typbibliothek "Microsoft Add-In Designer" (MSADDNDR.DLL)
  IDTExtensibility2 = interface(IDispatch)
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode;
      const AddInInst: IDispatch; var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  end;

  TClickProc = procedure of object;

  { Ereignisempfaenger fuer den Menue-Button (CommandBarEvents.Click, DISPID 1).
    Die Ereignis-Schnittstelle wird zur Laufzeit beim Connection Point erfragt,
    daher ist keine feste GUID noetig. }
  TCommandBarSink = class(TObject, IUnknown, IDispatch)
  private
    FRefCount: Integer;
    FEventIID: TGUID;
    FOnClick: TClickProc;
  protected
    function QueryInterface(const IID: TGUID; out Obj): HResult; stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
  public
    constructor Create(const AEventIID: TGUID; AOnClick: TClickProc);
  end;

  TVbaSyncConnect = class(TComObject, IDispatch, IDTExtensibility2)
  private
    FVBE: OleVariant;
    FButton: OleVariant;
    FEvents: IUnknown;
    FConnPoint: IConnectionPoint;
    FCookie: Integer;
    FSink: IDispatch;
    procedure DoConnect(const App: IDispatch);
    procedure DoDisconnect;
    procedure CreateMenu;
    procedure RemoveMenu;
    procedure ShowMainForm;
  protected
    { IDispatch (falls der Host spaet bindet) }
    function GetTypeInfoCount(out Count: Integer): HResult; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; stdcall;
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; stdcall;
    { IDTExtensibility2 }
    procedure OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode;
      const AddInInst: IDispatch; var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  end;

function VbaSyncRegisterServer: HResult; stdcall;
function VbaSyncUnregisterServer: HResult; stdcall;

implementation

uses
  VbaSync.MainForm, VbaSync.Version;

const
  msoControlButton = 1;
  msoButtonCaption = 2;

{ TCommandBarSink }

constructor TCommandBarSink.Create(const AEventIID: TGUID; AOnClick: TClickProc);
begin
  inherited Create;
  FEventIID := AEventIID;
  FOnClick := AOnClick;
end;

function TCommandBarSink.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if IsEqualGUID(IID, IUnknown) or IsEqualGUID(IID, IDispatch) or IsEqualGUID(IID, FEventIID) then
  begin
    // Obj kommt vom Aufrufer uninitialisiert: erst nil setzen, sonst ruft
    // die Zuweisung _Release auf einem Zufallszeiger auf (Zugriffsverletzung).
    Pointer(Obj) := nil;
    IDispatch(Obj) := Self;
    Result := S_OK;
  end
  else
  begin
    Pointer(Obj) := nil;
    Result := E_NOINTERFACE;
  end;
end;

function TCommandBarSink._AddRef: Integer;
begin
  Result := InterlockedIncrement(FRefCount);
end;

function TCommandBarSink._Release: Integer;
begin
  Result := InterlockedDecrement(FRefCount);
  if Result = 0 then
    Destroy;
end;

function TCommandBarSink.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count := 0;
  Result := S_OK;
end;

function TCommandBarSink.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Pointer(TypeInfo) := nil;
  Result := E_NOTIMPL;
end;

function TCommandBarSink.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

function TCommandBarSink.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
begin
  Result := S_OK;
  if (DispID = 1) and Assigned(FOnClick) then // Click
  try
    FOnClick;
  except
    on E: Exception do
      MessageBox(0, PChar(E.Message), 'VbaSync', MB_ICONERROR or MB_OK);
  end;
end;

{ TVbaSyncConnect }

function TVbaSyncConnect.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Count := 0;
  Result := S_OK;
end;

function TVbaSyncConnect.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Pointer(TypeInfo) := nil;
  Result := E_NOTIMPL;
end;

function TVbaSyncConnect.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

{ Spaete Bindung: DISPID 1 = OnConnection, 2 = OnDisconnection, 3..5 ohne Funktion.
  Argumente liegen in umgekehrter Reihenfolge vor. }
function TVbaSyncConnect.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
var
  DP: TDispParams;
  V: OleVariant;
begin
  Result := S_OK;
  DP := TDispParams(Params);
  try
    case DispID of
      1:
        if DP.cArgs >= 4 then
        begin
          V := POleVariant(@DP.rgvarg^[DP.cArgs - 1])^;
          DoConnect(IUnknown(V) as IDispatch);
        end;
      2:
        DoDisconnect;
    end;
  except
    Result := E_FAIL;
  end;
end;

procedure TVbaSyncConnect.OnConnection(const Application: IDispatch;
  ConnectMode: ext_ConnectMode; const AddInInst: IDispatch; var custom: PSafeArray);
begin
  DoConnect(Application);
end;

procedure TVbaSyncConnect.OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray);
begin
  DoDisconnect;
end;

procedure TVbaSyncConnect.OnAddInsUpdate(var custom: PSafeArray);
begin
end;

procedure TVbaSyncConnect.OnStartupComplete(var custom: PSafeArray);
begin
end;

procedure TVbaSyncConnect.OnBeginShutdown(var custom: PSafeArray);
begin
end;

procedure TVbaSyncConnect.DoConnect(const App: IDispatch);
begin
  if not VarIsEmpty(FVBE) then
    Exit; // schon verbunden
  // Die VBE-Add-in-Registrierung gilt fuer alle Office-Programme (Excel, Access,
  // Word ...). VbaSync arbeitet nur mit Excel-Mappen -> in anderen Hosts nichts tun.
  if not SameText(ExtractFileName(GetModuleName(0)), 'EXCEL.EXE') then
    Exit;
  FVBE := App; // Application ist hier das VBE-Objekt
  try
    CreateMenu;
  except
    on E: Exception do
      MessageBox(0, PChar('VbaSync-Menue konnte nicht angelegt werden: ' + E.Message),
        'VbaSync', MB_ICONWARNING or MB_OK);
  end;
end;

procedure TVbaSyncConnect.DoDisconnect;
begin
  try
    RemoveMenu;
  except
    // beim Beenden ignorieren
  end;
  FVBE := Unassigned;
end;

{ Late-Binding-Aufruf mit Argumenten BY VALUE und expliziten Flags.
  (Delphis eingebaute Variant-Aufrufe uebergeben Variablen by reference,
   was die VBE-Objekte u. a. mit "Mitglied nicht gefunden" quittieren.) }
function DispCall(const Obj: OleVariant; const Name: string;
  const Args: array of OleVariant; Flags: Word): OleVariant;
var
  Disp: IDispatch;
  DispId, PutId, I, N: Integer;
  WName: WideString;
  PName: PWideChar;
  VArgs: array of TVariantArg;
  DP: TDispParams;
  Excep: TExcepInfo;
  ArgErr: Cardinal;
  HR: HResult;
  Msg: string;
begin
  Disp := IDispatch(Obj);
  if Disp = nil then
    raise Exception.Create(Name + ': Objekt ist leer');
  WName := Name;
  PName := PWideChar(WName);
  HR := Disp.GetIDsOfNames(GUID_NULL, @PName, 1, LOCALE_USER_DEFAULT, @DispId);
  if Failed(HR) then
    raise Exception.CreateFmt('%s: unbekannter Name (0x%.8x)', [Name, HR]);
  N := Length(Args);
  SetLength(VArgs, N);
  for I := 0 to N - 1 do
    VArgs[I] := PVariantArg(@Args[N - 1 - I])^; // umgekehrte Reihenfolge
  FillChar(DP, SizeOf(DP), 0);
  if N > 0 then
    DP.rgvarg := PVariantArgList(@VArgs[0]);
  DP.cArgs := N;
  if Flags = DISPATCH_PROPERTYPUT then
  begin
    PutId := DISPID_PROPERTYPUT;
    DP.rgdispidNamedArgs := PDispIDList(@PutId);
    DP.cNamedArgs := 1;
  end;
  FillChar(Excep, SizeOf(Excep), 0);
  ArgErr := 0;
  if Flags = DISPATCH_PROPERTYPUT then
    HR := Disp.Invoke(DispId, GUID_NULL, LOCALE_USER_DEFAULT, Flags, DP, nil, @Excep, @ArgErr)
  else
    HR := Disp.Invoke(DispId, GUID_NULL, LOCALE_USER_DEFAULT, Flags, DP, @Result, @Excep, @ArgErr);
  if Failed(HR) then
  begin
    Msg := Excep.bstrDescription;
    if Msg = '' then
      Msg := SysErrorMessage(Cardinal(HR)); // HRESULT ist negativ -> Cast noetig
    raise Exception.CreateFmt('%s: %s (0x%.8x)', [Name, Msg, HR]);
  end;
end;

function DispGet(const Obj: OleVariant; const Name: string;
  const Args: array of OleVariant): OleVariant;
begin
  Result := DispCall(Obj, Name, Args, DISPATCH_PROPERTYGET or DISPATCH_METHOD);
end;

procedure DispPut(const Obj: OleVariant; const Name: string; const Value: OleVariant);
begin
  DispCall(Obj, Name, [Value], DISPATCH_PROPERTYPUT);
end;

{ Ermittelt die IID der Ereignis-Schnittstelle _dispCommandBarControlEvents
  aus der Typbibliothek, zu der das CommandBarEvents-Objekt gehoert (VBIDE).
  So muss keine GUID fest im Code stehen. }
function FindCommandBarEventIID(const Unk: IUnknown; out IID: TGUID): Boolean;
const
  LIBID_VBIDE: TGUID = '{0002E157-0000-0000-C000-000000000046}'; // VBA Extensibility 5.3
var
  Disp: IDispatch;
  PCI: IProvideClassInfo;
  TI, TI2: ITypeInfo;
  TL: ITypeLib;
  Idx, I: Integer;
  Name: WideString;
  PA: PTypeAttr;
begin
  Result := False;
  TL := nil;
  // 1. ueber IDispatch.GetTypeInfo
  if Supports(Unk, IDispatch, Disp) and Succeeded(Disp.GetTypeInfo(0, 0, TI)) and (TI <> nil) then
    TI.GetContainingTypeLib(TL, Idx);
  // 2. ueber IProvideClassInfo
  if (TL = nil) and Supports(Unk, IProvideClassInfo, PCI) and
     Succeeded(PCI.GetClassInfo(TI)) and (TI <> nil) then
    TI.GetContainingTypeLib(TL, Idx);
  // 3. registrierte VBIDE-Typbibliothek direkt laden
  if TL = nil then
    LoadRegTypeLib(LIBID_VBIDE, 5, 3, 0, TL);
  if TL = nil then
    Exit;
  for I := 0 to TL.GetTypeInfoCount - 1 do
  begin
    Name := '';
    if Succeeded(TL.GetDocumentation(I, @Name, nil, nil, nil)) and
       SameText(Name, '_dispCommandBarControlEvents') then
    begin
      if Succeeded(TL.GetTypeInfo(I, TI2)) and Succeeded(TI2.GetTypeAttr(PA)) then
      begin
        IID := PA^.guid;
        TI2.ReleaseTypeAttr(PA);
        Exit(True);
      end;
    end;
  end;
end;

procedure TVbaSyncConnect.CreateMenu;
var
  Bars, Bar, Controls, Events, EvObj: OleVariant;
  CPC: IConnectionPointContainer;
  Enum: IEnumConnectionPoints;
  CP: IConnectionPoint;
  EventIID: TGUID;
  Step: string;
begin
  try
    // "Menu Bar" ist der sprachunabhaengige Name der Menueleiste der VBA-IDE
    Step := 'VBE.CommandBars';
    Bars := DispGet(FVBE, 'CommandBars', []);
    Step := 'CommandBars.ActiveMenuBar';
    try
      Bar := DispGet(Bars, 'ActiveMenuBar', []);
    except
      Step := 'CommandBars("Menu Bar")';
      Bar := DispGet(Bars, 'Item', ['Menu Bar']);
    end;
    Step := 'Controls.Add';
    Controls := DispGet(Bar, 'Controls', []);
    FButton := DispCall(Controls, 'Add',
      [msoControlButton, EmptyParam, EmptyParam, EmptyParam, True], DISPATCH_METHOD);
    Step := 'Button-Eigenschaften';
    DispPut(FButton, 'Caption', 'VbaSync...');
    DispPut(FButton, 'Style', msoButtonCaption);
    DispPut(FButton, 'Tag', AddinProgID);
    DispPut(FButton, 'TooltipText', 'Module exportieren, importieren, loeschen, Clean');

    // Click-Ereignis ueber VBE.Events.CommandBarEvents(Button) abonnieren
    Step := 'Events.CommandBarEvents';
    Events := DispGet(FVBE, 'Events', []);
    EvObj := DispGet(Events, 'CommandBarEvents', [FButton]);
    FEvents := IUnknown(EvObj);
    Step := 'Ereignis-Schnittstelle (Typbibliothek)';
    if not FindCommandBarEventIID(FEvents, EventIID) then
      EventIID := StringToGUID('{0002E131-0000-0000-C000-000000000046}'); // Annahme VBIDE
    Step := 'FindConnectionPoint ' + GUIDToString(EventIID);
    CPC := FEvents as IConnectionPointContainer;
    if Failed(CPC.FindConnectionPoint(EventIID, CP)) or (CP = nil) then
    begin
      // letzter Versuch: aufzaehlen
      Step := 'EnumConnectionPoints';
      OleCheck(CPC.EnumConnectionPoints(Enum));
      if Enum.Next(1, CP, nil) <> S_OK then
        raise Exception.Create('Kein Connection Point fuer CommandBarEvents gefunden.');
      OleCheck(CP.GetConnectionInterface(EventIID));
    end;
    Step := 'Advise ' + GUIDToString(EventIID);
    FSink := TCommandBarSink.Create(EventIID, ShowMainForm);
    OleCheck(CP.Advise(FSink, FCookie));
    FConnPoint := CP;
  except
    on E: Exception do
      raise Exception.Create('Schritt ' + Step + ': ' + E.Message);
  end;
end;

procedure TVbaSyncConnect.RemoveMenu;
begin
  if Assigned(FConnPoint) then
  begin
    FConnPoint.Unadvise(FCookie);
    FConnPoint := nil;
  end;
  FSink := nil;
  FEvents := nil;
  if not VarIsEmpty(FButton) then
  begin
    try
      DispCall(FButton, 'Delete', [], DISPATCH_METHOD);
    except
    end;
    FButton := Unassigned;
  end;
end;

procedure TVbaSyncConnect.ShowMainForm;
var
  Form: TMainForm;
  OwnerWnd: HWND;
  Preferred: string;
begin
  OwnerWnd := 0;
  Preferred := '';
  try
    OwnerWnd := HWND(Integer(DispGet(DispGet(FVBE, 'MainWindow', []), 'HWnd', [])));
  except
  end;
  try
    // ungespeicherte Projekte haben keinen Dateinamen -> Fehler abfangen
    Preferred := VarToStr(DispGet(DispGet(FVBE, 'ActiveVBProject', []), 'FileName', []));
  except
  end;
  Form := TMainForm.CreateForAddin(OwnerWnd, Preferred);
  try
    Form.ShowModal;
  finally
    Form.Free;
  end;
end;

{ Registrierung pro Benutzer }

function AddinsKey: string;
begin
  {$IFDEF WIN64}
  Result := 'Software\Microsoft\VBA\VBE\6.0\Addins64\' + AddinProgID;
  {$ELSE}
  Result := 'Software\Microsoft\VBA\VBE\6.0\Addins\' + AddinProgID;
  {$ENDIF}
end;

function VbaSyncRegisterServer: HResult;
var
  R: TRegistry;
  Clsid, DllPath: string;
begin
  Result := S_OK;
  Clsid := GUIDToString(CLASS_VbaSyncConnect);
  DllPath := GetModuleName(HInstance);
  R := TRegistry.Create(KEY_ALL_ACCESS);
  try
    try
      R.RootKey := HKEY_CURRENT_USER;
      if R.OpenKey('Software\Classes\CLSID\' + Clsid, True) then
      begin
        R.WriteString('', 'VbaSync VBE Add-in');
        R.CloseKey;
      end;
      if R.OpenKey('Software\Classes\CLSID\' + Clsid + '\InprocServer32', True) then
      begin
        R.WriteString('', DllPath);
        R.WriteString('ThreadingModel', 'Apartment');
        R.CloseKey;
      end;
      if R.OpenKey('Software\Classes\CLSID\' + Clsid + '\ProgID', True) then
      begin
        R.WriteString('', AddinProgID);
        R.CloseKey;
      end;
      if R.OpenKey('Software\Classes\' + AddinProgID + '\CLSID', True) then
      begin
        R.WriteString('', Clsid);
        R.CloseKey;
      end;
      if R.OpenKey(AddinsKey, True) then
      begin
        R.WriteString('FriendlyName', AppName + ' ' + AppVersion);
        R.WriteString('Description', AppDescription);
        R.WriteInteger('LoadBehavior', 3); // beim Start laden + geladen
        R.WriteInteger('CommandLineSafe', 0);
        R.CloseKey;
      end;
    except
      Result := SELFREG_E_CLASS;
    end;
  finally
    R.Free;
  end;
end;

function VbaSyncUnregisterServer: HResult;
var
  R: TRegistry;
  Clsid: string;
begin
  Result := S_OK;
  Clsid := GUIDToString(CLASS_VbaSyncConnect);
  R := TRegistry.Create(KEY_ALL_ACCESS);
  try
    R.RootKey := HKEY_CURRENT_USER;
    R.DeleteKey(AddinsKey);
    R.DeleteKey('Software\Classes\CLSID\' + Clsid + '\InprocServer32');
    R.DeleteKey('Software\Classes\CLSID\' + Clsid + '\ProgID');
    R.DeleteKey('Software\Classes\CLSID\' + Clsid);
    R.DeleteKey('Software\Classes\' + AddinProgID + '\CLSID');
    R.DeleteKey('Software\Classes\' + AddinProgID);
  finally
    R.Free;
  end;
end;

initialization
  TComObjectFactory.Create(ComServer, TVbaSyncConnect, CLASS_VbaSyncConnect,
    'Connect', 'VbaSync VBE Add-in', ciMultiInstance, tmApartment);
end.
