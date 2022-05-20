unit StringCtrl;


interface

uses
  Classes;

type
  TStringArray = array of string;

function StrSplit(const Separator, Input: string; Max: integer = 0): TStringArray; stdcall;
function StrCut(Input: string; sText: string; eText: string = ';'; Default: string = ''; Count: smallint = 1; Headers: boolean = false): string; stdcall;
function StrDynamic(Input: string; Date: TDateTime; Source: string = ''; Target: string = ''): string;
function GetBlock(Input, Find, StartSgmt, CloseSgmt: string): string; stdcall;
function UnixDate(DateIn: double): double; stdcall;
function StrEncrypt(const Input :WideString; Key: Word): string; stdcall;
function StrDecrypt(const Input: String; Key: Word): string; stdcall;
function GetMotherBoardSerial:string;

implementation

uses
  SysUtils,
  StrUtils,
  ActiveX,
  ComObj,
  Variants,
  Dialogs;

const CKEY1 = 53761;
      CKEY2 = 32618;

function StrEncrypt(const Input :WideString; Key: Word): string; stdcall;
var
  i: integer;
  rstr :RawByteString;
  rstrb :TBytes absolute rstr;
begin
  result:= '';
  rstr:= UTF8Encode(Input);
  for i := 0 to length(rstr)-1 do begin
    rstrB[i] := rstrB[i] xor (key shr 8);
    key := (rstrb[i] + Key) * CKEY1 + CKEY2;
  end;
  for i := 0 to length(rstr)-1 do begin
    result:= result + inttohex(rstrb[i], 2);
  end;
end;

function StrDecrypt(const Input: String; Key: Word): string; stdcall;
var
  i, tmpkey :integer;
  rstr :RawByteString;
  rstrb :TBytes absolute rstr;
  tmpstr :string;
begin
  tmpstr:= UpperCase(Input);
  setlength(rstr, length(tmpstr) div 2);
  i:= 1;
  try
    while (i < length(tmpstr)) do begin
      rstrb[i div 2]:= strtoint('$' + tmpstr[i] + tmpstr[i+1]);
      inc(i, 2);
    end;
  except
    result:= '';
    exit;
  end;
  for i := 0 to length(rstr)-1 do begin
    tmpkey:= rstrb[i];
    rstrb[i] := rstrb[i] xor (key shr 8);
    key := (tmpkey + Key) * CKEY1 + CKEY2;
  end;
  result:= UTF8Decode(rstr);
end;

function GetMotherBoardSerial:string;
var
  objWMIService, colItems, colItem: OLEVariant;
  oEnum: IEnumvariant;
  iValue: LongWord;

  function GetWMIObject(const objectName: String): IDispatch;
  var
    chEaten: Integer;
    BindCtx: IBindCtx;
    Moniker: IMoniker;
  begin
    OleCheck(CreateBindCtx(0, bindCtx));
    OleCheck(MkParseDisplayName(BindCtx, StringToOleStr(objectName), chEaten, Moniker));
    OleCheck(Moniker.BindToObject(BindCtx, nil, IDispatch, Result));
  end;

begin
  Result:='';
  objWMIService := GetWMIObject('winmgmts:\\localhost\root\cimv2');
  colItems      := objWMIService.ExecQuery('SELECT SerialNumber FROM Win32_BaseBoard','WQL',0);
  oEnum         := IUnknown(colItems._NewEnum) as IEnumVariant;
  if oEnum.Next(1, colItem, iValue) = 0 then
  result:=VarToStr(colItem.SerialNumber);
end;

function StrSplit(const Separator, Input: string; Max: integer = 0): TStringArray;
var
  i, strt, cnt: Integer;
  sepLen: Integer;
  procedure AddString(aEnd: Integer = -1);
  var
    endPos: Integer;
  begin
    if (aEnd = -1) then
      endPos := i
    else
      endPos := aEnd + 1;
    if (strt < endPos) then
      result[cnt] := Copy(Input, strt, endPos - strt)
    else
      result[cnt] := '';

    Inc(cnt);
  end;
begin
  if (Input = '') or (Max < 0) then
  begin
    SetLength(result, 0);
    exit;
  end;

  if (Separator = '') then
  begin
    SetLength(result, 1);
    result[0] := Input;
    exit;
  end;

  sepLen := Length(Separator);
  SetLength(result, (Length(Input) div sepLen) + 1);

  i     := 1;
  strt  := i;
  cnt   := 0;
  while (i <= (Length(Input)- sepLen + 1)) do
  begin
    if (Input[i] = Separator[1]) then
      if (Copy(Input, i, sepLen) = Separator) then
      begin
        AddString;

        if (cnt = Max) then
        begin
          SetLength(result, cnt);
          EXIT;
        end;
        Inc(i, sepLen - 1);
        strt := i + 1;
      end;
    Inc(i);
  end;
  AddString(Length(Input));
  SetLength(result, cnt);
end;

function StrCut(Input: string; sText: string; eText: string = ';'; Default: string = ''; Count: smallint = 1; Headers: boolean = false): string; stdcall;
var
  Index: Integer;
begin
  result := Input;
  if ContainsText(result, sText) then
    begin
      for Index := 1 to Count do
        if not (sText = '') then
        begin
          if AnsiContainsStr(result, sText) then
            result := Copy(result, AnsiPos(sText, result) + Length(sText), Length(result))
          else
            result := '';
        end;
      if not (eText = '') and ContainsText(result, eText) then
        result := Copy(result, 0, AnsiPos(eText, result) - 1);
    end
  else
    begin
      if Default = '' then
        result := ''
      else
        result := Default;
    end;
end;

function StrDynamic(Input: string; Date: TDateTime; Source: string = ''; Target: string = ''): string;
begin
  result := Input;
  result := StringReplace(result, '[DD]', FormatDateTime('dd', Date), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[MM]', FormatDateTime('mm', Date), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[YYYY]', FormatDateTime('yyyy', Date), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[SOURCE]', UpperCase(Source), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[TARGET]', UpperCase(Target), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[source]', LowerCase(Source), [rfReplaceAll, rfIgnoreCase]);
  result := StringReplace(result, '[target]', LowerCase(Target), [rfReplaceAll, rfIgnoreCase]);
end;

function GetBlock(Input, Find, StartSgmt, CloseSgmt: string): string; stdcall;
var
  Index: integer;
begin
  if Input = '' then
    exit(Input)
  else
    begin
      Index := 0;
      while (StrCut(Input, StartSgmt, CloseSgmt, '', Index) <> '') and (result = '') do
      begin
        Index := Index +1;
        if AnsiContainsStr(StrCut(Input, StartSgmt, CloseSgmt, '', Index), Find) then
          result := StrCut(Input, StartSgmt, CloseSgmt, '', Index);
      end;
    end;
end;

function UnixDate(DateIn: double): double; stdcall;
begin
  result := (DateIn - 25569) * 86400
end;

end.
