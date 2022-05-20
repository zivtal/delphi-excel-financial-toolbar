unit XlPriceIndex;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,  Vcl.Controls, Vcl.Forms,
  Vcl.Dialogs, Vcl.StdCtrls, StrUtils, Wininet, XMLDoc, XMLIntf, MSXML, Vcl.ComCtrls, DateUtils, Vcl.ExtCtrls, Vcl.Mask,
  RegularExpressions, Excel2010, Registry, Vcl.Imaging.pngimage, Vcl.Grids, Math, XlApplication, XlProgress;

  procedure PriceIndexGUI(); stdcall;

type
  TStringArray = array of string;
  TPriceIndex = class(TForm)
    IGet: TButton;
    IBase: TComboBox;
    LBase: TLabel;
    LIndex: TLabel;
    IName: TComboBox;
    LProvider: TLabel;
    IProvider: TComboBox;
    IDate: TDateTimePicker;
    IKnown: TCheckBox;
    Status: TStatusBar;
    SDate: TRadioButton;
    SYears: TRadioButton;
    IYears: TMaskEdit;
    INewSheet: TCheckBox;
    LExport: TLabel;
    function CheckFields: boolean;
    procedure IGetClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure INameChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure IKnownClick(Sender: TObject);
    procedure IDateChange(Sender: TObject);
    procedure SClick(Sender: TObject);
    procedure IYearsExit(Sender: TObject);
    procedure IDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure IYearsChange(Sender: TObject);
    procedure IBaseChange(Sender: TObject);
    procedure IBaseExit(Sender: TObject);
  private
    MODE: integer;
    CACHE_INDEXNAME: array of string;
    CACHE_INDEXCODE: array of string;
    CACHE_BASENAME: array of string;
    CACHE_BASEVALUE: array of double;
    RELEASE: integer;
    BASEDATA: widestring;
  end;

var
  PriceIndex: TPriceIndex;

implementation

{$R *.dfm}

const
  Registry: string = '\SOFTWARE\Ziv Tal\PriceIndex';

var
  CACHEURL: string;
  CACHEDATA: widestring;
  TEMPLATE: widestring;

function Quarter(Input: TDate; Offset: integer; MinDate, MaxDate: TDate): TDate;
var
  Year, Month, Day: Word;
  Step: integer;
begin
  Step := MonthOf(Input) - (Round(MonthOf(Input) div 3) * 3);
  if Step > 0 then
    if Offset < 0 then
      Offset := Offset + 1
    else if Offset > 0 then
      Offset := Offset - 1;
  result := IncMonth(Input, - (-Offset * 3 + Step));
  DecodeDate(result, Year, Month, Day);
  result := EndOfAMonth(Year, Month);

  if (result < MinDate) or (result > MaxDate) then
    result := Input;
end;

function RegRead(Path, Key: string; Default: string = ''): string;
var
  RegKey: TRegistry;
begin
  RegKey := TRegistry.Create;
  try
    RegKey.OpenKeyReadOnly(Path);
    try
      result := RegKey.ReadString(Key);
    except
        result := Default;
    end;
  finally
    RegKey.Free;
  end;
end;

procedure RegWrite(Path, Key, Value: string; Overwrite: boolean = true; Count: boolean = false); stdcall;
var
  Index: integer;
  RegKey: TRegistry;
label
  process, drop;
begin
  if not Overwrite and (RegRead(Path, Key) <> '') then
    goto drop;
  if Count then
    begin
      Index := 0;
      while (RegRead(Path, Key + IntToStr(Index)) <> '') do
        Index := Index + 1;
      Key := Key + IntToStr(Index);
    end;
process:
  RegKey := TRegistry.Create;
  try
    RegKey.OpenKey(Path, true);
    try
      RegKey.WriteString(Key, Value);
    except
    end;
  finally
    RegKey.Free;
  end;
drop:
  exit;
end;

function ActiveConnection(): boolean;
var
  connection: cardinal;
begin
  if not InternetGetConnectedState(@connection,0) then
    begin
      MessageDlg('Internet connection not available.',mtError, [mbOK],0);
      result := false;
    end
  else
    result := true;
end;

function HttpGet(const URL: string): widestring; stdcall;
var
  hInet: HINTERNET;
  hURL: HINTERNET;
  Buffer: array[0..1023] of AnsiChar;
  BufferLen: cardinal;
begin
  if CacheUrl = URL then
    exit(CacheData);
  result := '';
  hInet := InternetOpen('Delphi 5.x', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  if hInet = nil then RaiseLastOSError;
  try
    hURL := InternetOpenUrl(hInet, PChar(URL), nil, 0, 0, 0);
    if hURL = nil then RaiseLastOSError;
    try
      repeat
        if not InternetReadFile(hURL, @Buffer, SizeOf(Buffer), BufferLen) then
          RaiseLastOSError;
        result := result + UTF8ToWideString(Copy(Buffer, 1, BufferLen))
      until BufferLen = 0;
    finally
      InternetCloseHandle(hURL);
    end;
  finally
    InternetCloseHandle(hInet);
  end;
  CACHEURL := URL;
  CACHEDATA := result;
end;

function IndexOfArray(Input: array of string; Value: string): integer;
var
  Index: integer;
begin
  for Index := 0 to High(Input) do
    if Input[Index] = Value then
      exit(Index);
end;

function GetKey(Input: string; StartText: string; Default: string = ''; EndString: string = ';'): string; stdcall;
begin
  result := Input;
  if ContainsText(result, StartText) then
    begin
      if not (StartText = '') then
        begin
          if AnsiContainsStr(result, StartText) then
            result := Copy(result, AnsiPos(StartText, result) + Length(StartText), Length(result))
          else
            result := '';
        end;
      if not (EndString = '') and ContainsText(result, EndString) then
        result := Copy(result, 0, AnsiPos(EndString, result) - 1);
    end
  else
    begin
      if Default = '' then
        result := ''
      else
        result := Default;
    end;
end;

function GetValue(Path: string; NODES: IXMLDOMNodeList): string;
  function SubPath(Value: string; NODES: IXMLDOMNodeList): IXMLDOMNodeList;
  var
    Index: integer;
  begin
    if Value[1] = '#' then
      result := NODES
    else
      for Index := 0 to NODES.length-1 do
        if NODES[Index].nodeName = Value then
          exit(NODES[Index].childNodes);
  end;
  function GetAttr(Name: string; NODES: IXMLDOMNodeList): string;
  var
    Index: integer;
  begin
    for Index := 0 to NODES.item[0].attributes.length-1 do
      if NODES.item[0].attributes.item[Index].nodeName = Name then
        exit(NODES.item[0].attributes.item[Index].text);
  end;
var
  Value: string;
begin
  for Value in SplitString(Path, '/') do
    if not (Value = '') then
      try
          NODES := SubPath(Value, NODES);
        if Value[1] = '#' then
          result := GetAttr(Value, NODES)
        else
          result := NODES.item[0].text;
      except
      end;
end;

function FindNode(Name, Value: TStringArray; Path: WideString; XML: IXMLDOMDocument): IXMLDOMNodeList;
var
  Match: boolean;
  Index1, Index2, Index3: integer;
begin
  result := XML.selectNodes('//' + Path);
  for Index1 := 0 to result.length-1 do
  begin
    Match := True;
    for Index2 := 0 to result[Index1].childNodes.length-1 do
      for Index3 := 0 to High(Name) do
        begin
          if Name[Index3][1] = '#' then
            begin
              if (result[Index1].childNodes[Index2].attributes.getNamedItem(Copy(Name[Index3],2,Length(Name[Index3]))).text <> Value[Index3]) then
                Match := False;
            end
          else
            begin
              if (result[Index1].childNodes[Index2].nodeName = Name[Index3]) and (result[Index1].childNodes[Index2].text <> Value[Index3]) then
                Match := False;
            end;
        end;
    if Match then
      exit(result[Index1].childNodes);
  end;
end;

function GetIndex(Month, Year: integer; IndexCode: string; out IndexOut: double; out BaseOut: string): boolean;
var
  XML: IXMLDOMDocument;
  NODES: IXMLDOMNodeList;
  Data: WideString;
  URL, TREE, VALUE, BASE: string;
begin
  URL := GetKey(Template,'URL=');
  URL := StringReplace(URL, '[M]', IntToStr(Month), [rfReplaceAll]);
  URL := StringReplace(URL, '[YYYY]', IntToStr(Year), [rfReplaceAll]);
  URL := StringReplace(URL, '[CODE]', IndexCode, [rfReplaceAll]);
  VALUE := GetKey(Template,'VALUE=');
  BASE := GetKey(Template,'BASE=');
  TREE := GetKey(Template,'TREE=');
  try
    Data := HttpGet(URL);
    XML := CoDOMDocument.Create;
    XML.loadXML(Data);
    NODES := FindNode(['year','month'],[IntToStr(Year),IntToStr(Month)], TREE, XML);
    if TryStrToFloat(GetValue(VALUE, NODES), IndexOut) then
      begin
        BaseOut := GetValue(BASE, NODES);
        result := true;
      end;
  except
    result := false;
  end;
  XML := nil;
end;

procedure PriceIndexGUI(); stdcall;
begin
  if ActiveConnection then
    begin
      Application.CreateForm(TPriceIndex, PriceIndex);
      try
        PriceIndex.ShowModal;
      finally
        PriceIndex.Destroy;
      end;
    end
end;

procedure TPriceIndex.IBaseChange(Sender: TObject);
begin
  if IBase.ItemIndex = -1 then
    LBase.Font.Color := clRed
  else
    LBase.Font.Color := clBlack;
  IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.IBaseExit(Sender: TObject);
begin
  if IBase.ItemIndex = -1 then
    begin
      LBase.Font.Color := clRed;
      IBase.SetFocus;
    end
  else
    LBase.Font.Color := clBlack;
end;

procedure TPriceIndex.IDateChange(Sender: TObject);
begin
  if not IKnown.Checked then
      Status.Panels[1].Text := IntToStr(MonthOf(IDate.Date)) + '/' + IntToStr(YearOf(IDate.Date))
  else
    if DayOf(IDate.Date) >= RELEASE then
      Status.Panels[1].Text := IntToStr(MonthOf(IncMonth(IDate.Date, -1))) + '/' + IntToStr(YearOf(IncMonth(IDate.Date, -1)))
    else
      Status.Panels[1].Text := IntToStr(MonthOf(IncMonth(IDate.Date, -2))) + '/' + IntToStr(YearOf(IncMonth(IDate.Date, -2)));
  IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.IDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    32:
      IDate.Date := IDate.MaxDate;
    33:
      IDate.Date := Quarter(IDate.Date, 1, IDate.MinDate, IDate.MaxDate);
    34:
      IDate.Date := Quarter(IDate.Date, -1, IDate.MinDate, IDate.MaxDate);
    35:
      IDate.Date := Quarter(Now, 1, IDate.MinDate, IDate.MaxDate);
  end;
  if (Key >= 32) and (Key <= 35) then
    IKnownClick(Self.IDate);
  FormKeyDown(Self.IDate, Key, Shift);
end;

procedure TPriceIndex.IGetClick(Sender: TObject);
var
  Index: integer;
  IndexBase: string;
  IndexDate: TDate;
  IdxValue, IdxBase: double;
  Row, Col: integer;
  Month, Year, SYear, EYear, EMonth: integer;
label
  drop;
begin
  if not ActiveConnection then
    exit;
  case MODE of
    0: // export signle index
      try
        if not IKnown.Checked then
          IndexDate := IDate.Date
        else
          if DayOf(IDate.Date) >= RELEASE then
            IndexDate := IncMonth(IDate.Date, -1)
          else
            IndexDate := IncMonth(IDate.Date, -2);
        if GetIndex(MonthOf(IndexDate), YearOf(IndexDate), CACHE_INDEXCODE[IName.ItemIndex], IdxValue, IndexBase) then
          begin
            IdxBase := CACHE_BASEVALUE[IndexOfArray(CACHE_BASENAME, IndexBase)];
            if IBase.ItemIndex = -1 then
              begin
                IBase.ItemIndex := IBase.Items.IndexOf(IndexBase);
                MessageDlg('The index base was not selected, the result used the default base "' + IndexBase + '".',mtInformation, [mbOK],0);
              end;
            XlApp.ActiveCell.Value2 := IdxValue * (CACHE_BASEVALUE[IBase.ItemIndex]/IdxBase);
            RegWrite(Registry + '\LastSettings', 'Date', DateToStr(IDate.Date));
            RegWrite(Registry + '\LastSettings', 'Base', IntToStr(IBase.ItemIndex));
            RegWrite(Registry + '\LastSettings', 'Known', BoolToStr(IKnown.Checked));
            PriceIndex.Close;
          end;
      except
        MessageDlg('Error while retrieving price index.',mtError, [mbOK],0);
        exit;
      end;
    1: // export years range into new sheet
      try
        if TryStrToInt(Copy(IYears.Text,1,4), EYear) and TryStrToInt(Copy(IYears.Text,8,4), SYear) and (SYear >= EYear) then
          begin
            if INewSheet.Checked then
              begin
                XlWorkbook := XlApp.ActiveWorkbook;
                XlWorkbook.Worksheets.Add(EmptyParam, EmptyParam, EmptyParam, EmptyParam, 0);
                XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
                XlSheet.Name := 'Price Index ' + FormatDateTime('yyyymmddhhmmss', Now);
              end;
            for Col := Low(CACHE_BASENAME) to High(CACHE_BASENAME) do
              begin
                XlApp.ActiveCell.Offset[0, Col + 1].Value2 := CACHE_BASENAME[Col];
                XlApp.ActiveCell.Offset[0, Col + 1].Font.Bold := True;
              end;
            for Year := SYear downto EYear do
              begin
                if ProgressBar.Process(SYear-Year, SYear-EYear, 'Getting price index of ' + IntToStr(Year) + ' ...', 'Get price index from "' + IProvider.Text + '" ...') = false then
                  exit;
                EMonth := 12;
                if Year = YearOf(Now) then
                  if DayOf(Now) >= RELEASE then
                    EMonth := MonthOf(Now) -1
                  else
                    EMonth := MonthOf(Now) -2;
                for Month := EMonth downto 1 do
                  try
                    Row := Row + 1;
                    if ProgressBar.Process(SYear-Year, SYear-EYear, IName.Text + ' ' + IntToStr(Month) + '/' + IntToStr(Year) + ' ...', 'Get price index from "' + IProvider.Text + '" ...') = false then
                      exit;
                    XlApp.ActiveCell.Offset[Row, 0].Value2 := IntToStr(Month) + '/' + IntToStr(Year);
                    XlApp.ActiveCell.Offset[Row, 0].Font.Bold := True;
                    if GetIndex(Month, Year, CACHE_INDEXCODE[IName.ItemIndex], IdxValue, IndexBase) then
                      begin
                        IdxBase := CACHE_BASEVALUE[IndexOfArray(CACHE_BASENAME, IndexBase)];
                        for Index := IndexOfArray(CACHE_BASENAME, IndexBase) to High(CACHE_BASEVALUE) do
                          try
                            XlApp.ActiveCell.Offset[Row, Index + 1].Value2 := IdxValue * CACHE_BASEVALUE[Index]/IdxBase;
                          finally
                          end;
                      end;
                  finally
                  end;
              end;
            if INewSheet.Checked then
              XlApp.Range[XlApp.ActiveCell.Address[False, False, xlA1, False, False],XlApp.ActiveCell.Offset[Row,Col]].Columns.AutoFit;
            RegWrite(Registry + '\LastSettings', 'Range', IYears.Text);
            PriceIndex.Close;
          end;
      except
        MessageDlg('Error while retrieving price index.',mtError, [mbOK],0);
        exit;
      end;
  end;
  RegWrite(Registry + '\LastSettings', 'Index', IntToStr(IName.ItemIndex));
end;

procedure TPriceIndex.IKnownClick(Sender: TObject);
begin
  if IKnown.Checked then
    IDate.MaxDate := Now
  else
    IDate.MaxDate := EndOfAMonth(YearOf(IncMonth(Now, -1)), MonthOf(IncMonth(Now, -1)));
  IDateChange(Self.IKnown);
end;

procedure TPriceIndex.INameChange(Sender: TObject);
var
  Index, AttIndex: integer;
  XML: IXMLDOMDocument;
  NODES: IXMLDOMNodeList;
begin
  // Read bases
  try
    CACHE_BASENAME := nil;
    CACHE_BASEVALUE := nil;
    IBase.Items.Clear;
    XML := CoDOMDocument.Create;
    XML.loadXML(BASEDATA);
    NODES := FindNode(['name'],[IName.Text], 'date/code', XML);
    for Index := 0 to NODES.length -1 do
      for AttIndex := 0 to NODES[Index].attributes.length - 1 do
        if NODES[Index].attributes.item[AttIndex].baseName = 'base' then
          begin
            IBase.Items.Add(NODES[Index].attributes.item[AttIndex].text);
            SetLength(CACHE_BASENAME, Length(CACHE_BASENAME)+1);
            CACHE_BASENAME[High(CACHE_BASENAME)] := NODES[Index].attributes.item[AttIndex].text;
            SetLength(CACHE_BASEVALUE, Length(CACHE_BASEVALUE)+1);
            TryStrToFloat(NODES[Index].text, CACHE_BASEVALUE[High(CACHE_BASEVALUE)]);
          end;
  finally
    XML := nil;
  end;
end;

function TPriceIndex.CheckFields: boolean;
var
  SYear, EYear: integer;
begin
  case MODE of
    0:
      begin
        if IBase.ItemIndex = -1 then
          begin
            LBase.Font.Color := clRed;
            SYears.Font.Color := clBlack;
            IYears.Font.Color := clBlack;
            exit(false);
          end;
      end;
    1:
      begin
        if not TryStrToInt(Copy(IYears.Text,1,4), EYear) or not TryStrToInt(Copy(IYears.Text,8,4), SYear) or (SYear > YearOf(Now)) or (EYear < StrToInt(GetKey(Template,'MINYEAR='))) then
          begin
            LBase.Font.Color := clBlack;
            SYears.Font.Color := clRed;
            IYears.Font.Color := clRed;
            exit(false);
          end;
      end;
  end;
  result := true;
end;

procedure TPriceIndex.IYearsChange(Sender: TObject);
var
  SYear, EYear: integer;
begin
  if not TryStrToInt(Copy(IYears.Text,1,4), EYear) or not TryStrToInt(Copy(IYears.Text,8,4), SYear) or (SYear > YearOf(Now)) or (EYear < StrToInt(GetKey(Template,'MINYEAR='))) then
    begin
      SYears.Font.Color := clRed;
      IYears.Font.Color := clRed;
    end
  else
    begin
      SYears.Font.Color := clBlack;
      IYears.Font.Color := clBlack;
    end;
  IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.IYearsExit(Sender: TObject);
var
  EYear, SYear: integer;
begin
  if not TryStrToInt(Copy(IYears.Text,1,4), EYear) or not TryStrToInt(Copy(IYears.Text,8,4), SYear) or (SYear > YearOf(Now)) or (EYear < StrToInt(GetKey(Template,'MINYEAR='))) then
    begin
      MessageDlg('Invalid range of years, Use the range from ' + GetKey(Template,'MINYEAR=') + ' to ' + IntToStr(YearOf(Now)) + '.',mtError, [mbOK],0);
      IYears.SetFocus;
    end
  else
    IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.SClick(Sender: TObject);
begin
  IDate.Enabled := SDate.Checked;
  IKnown.Enabled := SDate.Checked;
  IBase.Enabled := SDate.Checked;
  IYears.Enabled := SYears.Checked;
  INewSheet.Enabled := SYears.Checked;
  if SDate.Checked then
    MODE := 0;
  if SYears.Checked then
    MODE := 1;
  IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.FormCreate(Sender: TObject);
var
  Index: integer;
  URL: string;
  XML: IXMLDOMDocument;
  NODES: IXMLDOMNodeList;
begin
  TEMPLATE :=
    'URL=https://api.cbs.gov.il/index/data/price?lang=he&format=xml&id=[CODE]&download=false&startPeriod=1-[YYYY]&endPeriod=12-[YYYY];' +
    'BASEURL=https://api.cbs.gov.il/index/data/price_selected?lang=he&format=xml&download=false;' +
    'TREE=month/CodeMonth/date/DateMonth;' +
    'VALUE=currBase/value;' +
    'BASE=currBase/baseDesc;' +
    'RELEASE=15;' +
    'MINYEAR=1951;';
  URL := GetKey(Template,'BASEURL=');
  RELEASE := StrToInt(GetKey(Template,'RELEASE='));
  BASEDATA := HttpGet(URL);
  try
    XML := CoDOMDocument.Create;
    XML.loadXML(BASEDATA);
    NODES := XML.selectNodes('//date/code');
    for Index := 0 to NODES.length-1 do
      begin
        SetLength(CACHE_INDEXCODE, Length(CACHE_INDEXCODE)+1);
        CACHE_INDEXCODE[High(CACHE_INDEXCODE)] := NODES[Index].attributes.getNamedItem('code').text;
      end;
    NODES := XML.selectNodes('//date/code/name');
    for Index := 0 to NODES.length-1 do
      begin
        SetLength(CACHE_INDEXNAME, Length(CACHE_INDEXNAME)+1);
        CACHE_INDEXNAME[High(CACHE_INDEXNAME)] := NODES[Index].text;
        IName.Items.Add(CACHE_INDEXNAME[High(CACHE_INDEXNAME)]);
      end;
  finally
    IName.ItemIndex := 0;
    INameChange(Self.IName);
    XML := nil;
  end;
  try
    IDate.Date := StrToDate(RegRead(Registry + '\LastSettings', 'Date', DateToStr(Now)));
    try
      IName.ItemIndex := StrToInt(RegRead(Registry + '\LastSettings', 'Index', '0'));
    finally
      INameChange(Self.IName);
      IBase.ItemIndex := StrToInt(RegRead(Registry + '\LastSettings', 'Base', '0'));
    end;
    try
      IKnown.Checked := StrToBool(RegRead(Registry + '\LastSettings', 'Known', 'true'));
    finally
      IKnownClick(Self.IDate);
    end;
    try
      IYears.Text := RegRead(Registry + '\LastSettings', 'Range', '2020 - ' + IntToStr(YearOf(Now)));
    finally
      IDateChange(Self.IDate);
    end;
  except
    IName.ItemIndex := 0;
  end;
  IGet.Enabled := CheckFields;
end;

procedure TPriceIndex.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    13:
      if IGet.Enabled then
        IGetClick(Self.IGet);
    27:
      PriceIndex.Close;
  end;
end;

end.
