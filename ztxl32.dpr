library ztxl32;

{$R 'XlImages.res' 'XlImages.rc'}

uses
  Windows,
  ComServ,
  Forms,
  StringCtrl in 'StringCtrl.pas',
  RegistryCtrl in 'RegistryCtrl.pas',
  WebCtrl in 'WebCtrl.pas',
  XlActivation in 'XlActivation.pas' {Activation},
  XlApplication in 'XlApplication.pas' {XlComAddin} {DelphiAddin4: CoClass},
  XlComAddin in 'XlComAddin.pas',
  XlProgress in 'XlProgress.pas' {ProgressBar},
  XlAbout in 'XlAbout.pas' {About},
  XlUtilites in 'XlUtilites.pas',
  XlCurrencyRate in 'XlCurrencyRate.pas' {CurrencyRate},
  XlCurrencyRateSet in 'XlCurrencyRateSet.pas' {CurrencyRateSet},
  XlYahooStock in 'XlYahooStock.pas' {YahooStock},
  XlStatusMarker in 'XlStatusMarker.pas' {StatusMarker},
  XlFxConversion in 'XlFxConversion.pas' {FxConversion},
  XlFxInspector in 'XlFxInspector.pas' {FxInspector},
  XlLinkManager in 'XlLinkManager.pas' {LinkManager},
  XlPriceIndex in 'XlPriceIndex.pas' {PriceIndex},
  XlSheetHeader in 'XlSheetHeader.pas' {SheetHeader};

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$IF CompilerVersion >= 22.0}
exports
  DllInstall;
{$IFEND}

{$R *.TLB}

{$R *.RES}

begin
  SystemParametersInfo(SPI_SETBEEP, 0, nil, SPIF_SENDWININICHANGE);
end.
