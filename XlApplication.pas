unit XlApplication;

{$WARN SYMBOL_PLATFORM OFF}
{$ALIGN ON}

interface

uses
  ComObj, ActiveX, XlComAddin, OleServer, Office2010, Excel2010, Contnrs, StdVcl, Dialogs, Vcl.Grids;

type
  TXLComAddinFactory = class(TAutoObjectFactory)
    procedure UpdateRegistry(Register: Boolean); override;
  end;

// Constants for enum ext_ConnectMode
type
  ext_ConnectMode = TOleEnum;

const
  ext_cm_AfterStartup = $00000000;
  ext_cm_Startup = $00000001;
  ext_cm_External = $00000002;
  ext_cm_CommandLine = $00000003;

// Constants for enum ext_DisconnectMode
type
  ext_DisconnectMode = TOleEnum;

const
  ext_dm_HostShutdown = $00000000;
  ext_dm_UserClosed = $00000001;

type
  IDTExtensibility2 = interface(IDispatch)
    ['{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}']
    procedure OnConnection(const Application: IDispatch;
                           ConnectMode: ext_ConnectMode;
                           const AddInInst: IDispatch;
                           var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode;
                              var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
  end;

type
  IRibbonExtensibility = interface(IDispatch)
    ['{000C0396-0000-0000-C000-000000000046}']
    function GetCustomUI(const RibbonID: WideString): WideString; safecall;
  end;

type
  IRibbonControl = interface(IDispatch)
    ['{000C0395-0000-0000-C000-000000000046}']
    function Get_Id: WideString; safecall;
    function Get_Context: IDispatch; safecall;
    function Get_Tag: WideString; safecall;
    property Id: WideString read Get_Id;
    property Context: IDispatch read Get_Context;
    property Tag: WideString read Get_Tag;
  end;

type
  CommandBarButton_Click = procedure(const Ctrl: CommandBarButton;
                                     var CancelDefault: WordBool) of object;

type
  TCommandBarButton = class(TOleServer)
  private
    FIntf: _CommandBarButton;
    FOnClick: CommandBarButton_Click;
    function GetDefaultInterface: _CommandBarButton;
  protected
    procedure InitServerData; override;
    procedure InvokeEvent(DispID: TDispID; var Params: TVariantArray); override;
  public
    procedure Connect; override;
    procedure ConnectTo(svrIntf: _CommandBarButton);
    procedure Disconnect; override;
    property DefaultInterface: _CommandBarButton read GetDefaultInterface;
  published
    property OnClick: CommandBarButton_Click read FOnClick write FOnClick;
  end;

type
  DAddin = class(TAutoObject, IXlAddin,
                        IDTExtensibility2, IRibbonExtensibility)
  private
    FApp: IDispatch;
    FBList: TObjectList;
    procedure InitButtons;
    procedure DestroyButtons;
    function GetResPic(const ImgName: String;
      const Res: integer = 0): IPictureDisp;
    { IDTExtensibility2 }
    procedure OnConnection(const Application: IDispatch;
                           ConnectMode: ext_ConnectMode;
                           const AddInInst: IDispatch;
                           var custom: PSafeArray); safecall;
    procedure OnDisconnection(RemoveMode: ext_DisconnectMode;
                              var custom: PSafeArray); safecall;
    procedure OnAddInsUpdate(var custom: PSafeArray); safecall;
    procedure OnStartupComplete(var custom: PSafeArray); safecall;
    procedure OnBeginShutdown(var custom: PSafeArray); safecall;
    { IRibbonExtensibility }
    function GetCustomUI(const RibbonID: WideString): WideString; safecall;
  protected
    function GetImage(const ImageID: WideString): IPictureDisp; safecall;
    procedure RibbonClick(const Control: IDispatch); safecall;
  public
    procedure Initialize; override;
    Destructor Destroy; override;
  end;

var
  XlApp: ExcelApplication;
  XlWorkbook: ExcelWorkbook;
  XlSheet: ExcelWorksheet;
  XlTApp: TExcelApplication;

implementation

uses
  ComServ, Windows, Registry, Variants, Classes, SysUtils, StrUtils, UITypes, XlActivation, StringCtrl, RegistryCtrl,
  XlAbout, XlUtilites, XlCurrencyRate, XlYahooStock, XlFxInspector, XlStatusMarker, XlFxConversion, XlLinkManager, XlPriceIndex,
  XlSheetHeader;

{ TXLComAddinFactory }

procedure TXLComAddinFactory.UpdateRegistry(Register: Boolean);
var
  RootKey: HKEY;
  AddInKey: String;
  r: TRegistry;
begin
  Rootkey:=HKEY_CURRENT_USER;
  AddInKey:='Software\Microsoft\Office\Excel\Addins\' + ProgID;
  r:=TRegistry.Create;
  r.RootKey:=RootKey;
  try
    if Register then
      if r.OpenKey(AddInKey, True) then begin
        r.WriteInteger('LoadBehavior', 3);
        r.WriteInteger('CommandLineSafe', 0);
        r.WriteString('FriendlyName', 'ZTools Add-In');
        r.WriteString('Description', 'ZTools by Ziv Tal');
        r.CloseKey;
      end else
        raise EOleError.Create('Can''t register Add-In ' + ProgID)
    else
      if r.KeyExists(AddInKey) then
        r.DeleteKey(AddInKey);
  finally
    r.Free;
  end;
  inherited;
end;

{ TCommandBarButton }

procedure TCommandBarButton.InitServerData;
const
  CServerData: TServerData = (
    ClassID:    '{55F88891-7708-11D1-ACEB-006008961DA5}';
    // CLASS_CommandBarButton;
    IntfIID:    '{000C030E-0000-0000-C000-000000000046}';
    // IID__CommandBarButton;
    EventIID:   '{000C0351-0000-0000-C000-000000000046}';
    // DIID__CommandBarButtonEvents;
    LicenseKey: nil;
    Version:    500);
begin
  ServerData:= @CServerData;
end;

procedure TCommandBarButton.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk:= GetServer;
    ConnectEvents(punk);
    FIntf:= punk as _CommandBarButton;
  end;
end;

procedure TCommandBarButton.ConnectTo(svrIntf: _CommandBarButton);
begin
  Disconnect;
  FIntf:= svrIntf;
  ConnectEvents(FIntf);
end;

procedure TCommandBarButton.DisConnect;
begin
  if Fintf <> nil then begin
    DisconnectEvents(FIntf);
    FIntf:= nil;
  end;
end;

function TCommandBarButton.GetDefaultInterface: _CommandBarButton;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL.');
  Result:= FIntf;
end;

procedure TCommandBarButton.InvokeEvent(DispID: TDispID; var Params: TVariantArray);
begin
  case DispID of
   1 : if Assigned(FOnClick) then
         FOnClick(IUnknown(TVarData(Params[0]).VPointer)
                    as _CommandBarButton {const CommandBarButton},
                  WordBool((TVarData(Params[1]).VPointer)^)
                    {var WordBool});
  end;
end;

{ GDI+ }

const
  WINGDIPDLL = 'GdiPlus.dll';

type
  UINT32 = type Cardinal;
  ARGB   = DWORD;
  GpBitmap = Pointer;
  GpImage = Pointer;
  GpStatus = Cardinal;

  GdiplusStartupInput = packed record
    GdiplusVersion           : UINT32;  // always 1
    DebugEventCallback       : Pointer; // DebugEventProc
    SuppressBackgroundThread : BOOL;
    SuppressExternalCodecs   : BOOL;
  end;
  PGdiplusStartupInput = ^GdiplusStartupInput;

function GdiplusStartup(out token: PULONG; const input: PGdiplusStartupInput;
  {out} output: Pointer {PGdiplusStartupOutput}):
  GpStatus; stdcall; external WINGDIPDLL;

procedure GdiplusShutdown(token: PULONG); stdcall; external WINGDIPDLL;

function GdipCreateBitmapFromStream(stream: ISTREAM; out bitmap: GPBITMAP): GpStatus; stdcall; external WINGDIPDLL;

function GdipCreateHBITMAPFromBitmap(bitmap: GpBitmap; out hbmReturn: HBITMAP; background: ARGB): GpStatus; stdcall; external WINGDIPDLL;

function GdipCreateHICONFromBitmap(bitmap: GpBitmap; out hbmReturn: HICON): GpStatus; stdcall; external WINGDIPDLL;

function GdipDisposeImage(image: GpImage): GpStatus; stdcall; external WINGDIPDLL;

{ DAddin }

function DAddin.GetResPic(const ImgName: string; const Res: integer = 0): ActiveX.IPictureDisp;
// Res: 0 GDI+ Bitmap RCDATA
//      1 LoadResource ICON
//      2 LoadResource BITMAP
var
  PictureDesc: TPictDesc;
  GPInput: GdiplusStartupInput;
  Status: GpStatus;
  Token: PULONG;
  ResStream: TResourceStream;
  ResStreamI: IStream;
  GPBM: GpBitmap;
  HBM: HBITMAP;
begin
  try
    case Res of
    0: begin
         FillChar(GPInput, SizeOf(GPInput), 0);
         GPInput.GdiplusVersion:= 1;
         Status:= GdiplusStartup(Token, @GPInput, nil);
         if Status = 0 then begin
           try
             ResStream:= TResourceStream.Create(HInstance, ImgName, RT_RCDATA);
             try
               ResStreamI:= TStreamAdapter.Create(ResStream);
               Status:= GdipCreateBitmapFromStream(ResStreamI, GPBM);
               if Status = 0 then begin
                 try
                   Status:= GdipCreateHBITMAPFromBitmap(GPBM, HBM, $00FFFFFF);
                   if Status = 0 then begin
                     FillChar(PictureDesc,SizeOf(PictureDesc),0);
                     PictureDesc.cbSizeOfStruct:= SizeOf(PictureDesc);
                     PictureDesc.picType:= PICTYPE_BITMAP;
                     PictureDesc.hbitmap:= HBM;
                     OleCheck(OleCreatePictureIndirect(PictureDesc,
                        ActiveX.IPicture, true, Result));
                   end;
                 finally
                   GdipDisposeImage(GPBM);
                 end;
               end;
             finally
               ResStream.Free;
               ResStreamI:= nil;
             end;
           finally
             GdiplusShutdown(Token);
           end;
         end;
       end;
    1: begin
         FillChar(PictureDesc, SizeOf(PictureDesc), 0);
         PictureDesc.cbSizeOfStruct:= SizeOf(PictureDesc);
         PictureDesc.picType := PICTYPE_ICON;
         PictureDesc.hIcon := LoadIcon(HInstance, PChar(ImgName));
         OleCheck(OleCreatePictureIndirect(PictureDesc,
                    ActiveX.IPicture, true, Result));
       end;
    2: begin
         FillChar(PictureDesc, SizeOf(PictureDesc), 0);
         PictureDesc.cbSizeOfStruct:= SizeOf(PictureDesc);
         PictureDesc.picType := PICTYPE_BITMAP;
         PictureDesc.hbitmap := LoadBitmap(HInstance, PChar(ImgName));
         OleCheck(OleCreatePictureIndirect(PictureDesc,
                    ActiveX.IPicture, true, Result));
       end;
    else
      Result:= nil;
    end;
  except
    Result:= nil;
  end;
end;

procedure DAddin.Initialize;
begin
  inherited;
  // container for button handlers
  FBList:= TObjectList.Create;
end;

destructor DAddin.Destroy;
begin
  FBList.Free;
  FBList := nil;
  inherited;
end;

procedure DAddin.InitButtons;
begin
end;

procedure DAddin.DestroyButtons;
begin
  // disconnect button handlers
  FBList.Clear;
  XlApp:=nil;
end;

function DAddin.GetImage(const ImageID: WideString): IPictureDisp;
var
  s: string;
  i: integer;
begin
  s:= imageID;
  i:= length(s);
  if i > 3 then
    s[i-3]:= '_';

  if imageID = 'ztool.ico' then begin
    Result:= GetResPic(s, 1);
    exit;
  end;

  Result:= GetResPic(s);
end;

{ DAddin - IDTExtensibility2}

procedure DAddin.OnConnection(const Application: IDispatch; ConnectMode: ext_ConnectMode; const AddInInst: IDispatch; var custom: PSafeArray);
begin
  FApp:= Application;
  if ConnectMode = ext_cm_AfterStartup then
    InitButtons;
  XlTApp := TExcelApplication.Create(nil);
end;

procedure DAddin.OnDisconnection(RemoveMode: ext_DisconnectMode; var custom: PSafeArray);
begin
  DestroyButtons;
  // release internal reference
  XlTApp:=nil;
  FApp:=nil;
end;

procedure DAddin.OnAddInsUpdate(var custom: PSafeArray);
begin
end;

procedure DAddin.OnStartupComplete(var custom: PSafeArray);
begin
  InitButtons;
end;

procedure DAddin.OnBeginShutdown(var custom: PSafeArray);
begin
  RegDelete('\SOFTWARE\Ziv Tal\ZTools\Cache\');
end;

{ DAddin - IRibbonExtensibility}

function DAddin.GetCustomUI(const RibbonID: WideString): WideString;
begin
  result :=
    '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" loadImage="GetImage">'#13#10 +
    '  <ribbon>'#13#10 +
    '    <tabs>'#13#10 +
    '      <tab id="TZTool" label="ZTools">'#13#10 +
    '        <group id="gOnline" autoScale="true" label="Online">'#13#10 +
		'			     <splitButton id="sbCurrencyRate" size="large">'#13#10 +
		'				     <menu id="mCurrencyRate">'#13#10 +
    '			        <button id="currate1" label="Currency rate" image="CurrencyRate.png" onAction="RibbonClick" supertip="A currencies converter is designed to convert one currency into another in order to check its corresponding value."/>'#13#10 +
		'					    <button id="currate2" label="Settings" image="Settings.png" onAction="RibbonClick" supertip="Currency Rate Settings" />'#13#10 +
		'				     </menu>'#13#10 +
		'			      </splitButton>'#13#10 +
    '          <button id="pindex1" label="Price Index" size="large" image="PriceIndex.png" onAction="RibbonClick" supertip="Price Index add-in is designed to import a price index by type, date and basis."/>'#13#10 +
    '          <button id="yahfin1" label="Stock Market" size="large" image="Stocks.png" onAction="RibbonClick"/>'#13#10 +
    '        </group>'#13#10 +
    '        <group id="gFormula" autoScale="true" label="Formula">'#13#10 +
    '			     <button id="fxinspector" label="Formula Inspector" image="FxInspector.png" onAction="RibbonClick" supertip="Inspect all cells mentioned in the formula of the selected cells." size="large"/>'#13#10 +
		'			     <splitButton id="sbFxConv" size="large">'#13#10 +
		'				     <menu id="mFxConv">'#13#10 +
    '             <button id="fxconv" label="Formula conversion" image="FxConversion.png" onAction="RibbonClick" supertip="Convert a complex dynamic formula to a static variables." />'#13#10 +
		'					    <button id="fxconvundo" label="Undo fx conversion" image="Undo.png" onAction="RibbonClick" supertip="Undo all formula conversions." />'#13#10 +
		'				     </menu>'#13#10 +
		'			      </splitButton>'#13#10 +
    '        </group>'#13#10 +
    '        <group id="gDesign" autoScale="true" label="General">'#13#10 +
    '           <button id="statusmrk" label="Status marker" image="StatusMarker.png" onAction="RibbonClick" visible="true" size="large"/>'#13#10 +
    '           <button id="lnkmgr" label="Link manager" image="LinkManager.png" onAction="RibbonClick" visible="true" size="large"/>'#13#10 +
    '        </group>'#13#10 +
    '        <group id="gSettings" autoScale="true" label="Settings">'#13#10 +
    '           <button id="ztxl" label="About" size="large" image="ZTools.png" onAction="RibbonClick"/>'#13#10 +
    '        </group>'#13#10 +
    '      </tab>'#13#10 +
    '   </tabs>'#13#10 +
    '  </ribbon>'#13#10 +
    '</customUI>';
end;

procedure DAddin.RibbonClick(const Control: IDispatch);
begin
  if Assigned(FApp) and Supports(FApp, ExcelApplication, XlApp) then
    begin
      // Free tools
      case AnsiIndexStr((Control as IRibbonControl).Id, ['ztxl','statusmrk','headersht','lnkmgr']) of
        0: // Show about dialog
          AboutGUI;
        1: // Status marker (statusmrk)
          StatusMarkerGUI;
        2: // Add sheet header (headersht)
          SheetHeaderGUI;
        3: // Link manager (lnkmgr)
          LinkManagerGUI;
        else // If not activate exit procedure
          if not Activated then
            exit;
      end;
      // Licensed tools
      case AnsiIndexStr((Control as IRibbonControl).Id, ['currate1','currate2','yahfin1','pindex1','fxinspector','fxconv','fxconvundo']) of
        0: // Currency rate (currate1)
          CurrencyGUI;
        1: // Currency rate settings (currate2)
          CurrencyGUI(true);
        2: // Yahoo stock (yahfin1)
          YahooStockGUI;
        3: // Price Index (pindex1)
          PriceIndexGUI;
        4: // Mark references cells (fxinspector)
          FxInspectorGUI;
        5: // Convert complex dynamic fx to static variants (fxconv)
          FxConversionGUI;
        6: // Undo fx conversion (fxconvundo)
          FxConversionUndo;
      end;
    end;
end;

initialization
  TXLComAddinFactory.Create(ComServer, DAddin, Class_DelphiAddin4,
    ciMultiInstance, tmApartment);

end.
