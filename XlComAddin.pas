unit XlComAddin;

{$TYPEDADDRESS OFF}
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
{$ALIGN 4}
interface

uses Windows, ActiveX, Classes, Variants;

const
  XLCOMMajorVersion = 1;
  XLCOMMinorVersion = 0;

  LIBID_XLCOM: TGUID = '{6C29D36D-1B89-4936-A8D5-6066CA044FFB}';

  IID_IXlAddin: TGUID = '{EF97A441-613E-413F-BC31-CD14F3B92FCE}';
  CLASS_DelphiAddin4: TGUID = '{D026CA98-1DC5-4531-8005-325D80646D0E}';

type
  IXlAddin = interface;
  IXlAddinDisp = dispinterface;

  DelphiAddin4 = IXlAddin;

  IXlAddin = interface(IDispatch)
    ['{EF97A441-613E-413F-BC31-CD14F3B92FCE}']
    procedure RibbonClick(const Control: IDispatch); safecall;
    function GetImage(const ImageID: WideString): IPictureDisp; safecall;
  end;

  IXlAddinDisp = dispinterface
    ['{EF97A441-613E-413F-BC31-CD14F3B92FCE}']
    procedure RibbonClick(const Control: IDispatch); dispid 1;
    function GetImage(const ImageID: WideString): IPictureDisp; dispid 2;
  end;

  CoDelphiAddin4 = class
    class function Create: IXlAddin;
    class function CreateRemote(const MachineName: string): IXlAddin;
  end;

implementation

uses ComObj;

class function CoDelphiAddin4.Create: IXlAddin;
begin
  Result := CreateComObject(CLASS_DelphiAddin4) as IXlAddin;
end;

class function CoDelphiAddin4.CreateRemote(const MachineName: string): IXlAddin;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DelphiAddin4) as IXlAddin;
end;

end.
