unit XlCurrencyRate;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.UITypes, System.DateUtils,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Imaging.pngimage, Vcl.ExtCtrls,
  Vcl.Imaging.jpeg, StrUtils, Wininet, Excel2010, Registry, RegularExpressions, StringCtrl, XlApplication, XlProgress,
  XlCurrencyRateSet;

  procedure CurrencyGUI(Settings: boolean = false); stdcall;

type
  TCurrencyRate = class(TForm)
    Source: TComboBox;
    Target: TComboBox;
    Exchange: TButton;
    StartDate: TDateTimePicker;
    EndDate: TDateTimePicker;
    ManualProv: TCheckBox;
    LSource: TLabel;
    LTarget: TLabel;
    Provaiders: TComboBox;
    ThridPart: TCheckBox;
    LStartDate: TLabel;
    LEndDate: TLabel;
    Average: TCheckBox;
    LastestRate: TCheckBox;
    function IndexOfBank(Name: string): integer;
    procedure Enabled(input: boolean);
    procedure CurrencyChange(Sender: TObject);
    procedure ExchangeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure StartDateChange(Sender: TObject);
    procedure AverageClick(Sender: TObject);
    procedure ManualProvClick(Sender: TObject);
    procedure SourceChange(Sender: TObject);
    procedure TargetChange(Sender: TObject);
    procedure ProvaidersChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure StartDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EndDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ProvaidersKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  public
    Formula, Replace: string;
    InSource, InTarget: string;
  end;

implementation

var
  CurrencyRate: TCurrencyRate;

{$R *.dfm}

const
  Registry: string = '\SOFTWARE\Ziv Tal\CurrencyRate';

var
  CACHEURL: string;
  CACHEDATA: widestring;

function IsQuarter(Input: TDate): boolean;
var
  Year, Month, Day: Word;
begin
  DecodeDate(Input, Year, Month, Day);
  if (DayOf(Input) = DayOf(EndOfAMonth(Year, Month))) and ((MonthOf(Input) - (Round(MonthOf(Input) div 3) * 3)) = 0) then
    result := true
  else
    result := false;
end;

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

function Extract(Input: string; out Output, Source, Target: string): boolean;
var
  RegEx: TRegEx;
  Day, Month, Year: Word;
  DateType: TFormatSettings;
begin
  RegEx := TRegEx.Create('([A-Z]{6})_+(\d{6}_\d{6}|\d{8})');
  result := RegEx.Match(Input).Success;
  if result then
    begin
      Output := RegEx.Match(Input).Value;
      Source := Copy(Output, 1, 3);
      Target := Copy(Output, 4, 3);
    end;
end;

function DateToString(Date1, Date2: TDate): string;
begin
  if Date1 = Date2 then
    result := FormatDateTime('ddmmyyyy', Date1)
  else
    result := FormatDateTime('ddmmyy', Date1) + '_' + FormatDateTime('ddmmyy', Date2);
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

procedure RegDelete(Path: string; Key: string = '');
var
  RegKey: TRegistry;
begin
  RegKey := TRegistry.Create;
  try
    if Key = '' then
      RegKey.DeleteKey(Path)
    else
      begin
        RegKey.OpenKey(Path, False);
        RegKey.DeleteValue(Key);
      end;
  finally
    RegKey.Free;
  end;
end;

function RegCache(Path, Key: string; out Data: double): boolean;
var
  RegKey: TRegistry;
begin
  RegKey := TRegistry.Create;
  try
    RegKey.OpenKeyReadOnly(Path);
    try
      Data := StrToFloat(RegKey.ReadString(Key));
      result := true;
    except
        result := false;
    end;
  finally
    RegKey.Free;
  end;
end;

procedure RegClearCache(Month: integer = 12);
var
  BankIndex, DateIndex: integer;
  RegKey: TRegistry;
  Banks, Dates: TStringList;
  KeyDate, MinDate: TDate;
begin
  MinDate := IncMonth(Now, -Month);
  MinDate := EncodeDate(YearOf(MinDate), MonthOf(MinDate), 1);
  RegKey := TRegistry.Create;
  try
    if RegKey.OpenKey(Registry + '\Cache\', False) then
      begin
        Banks := TStringList.Create;
        try
          RegKey.GetKeyNames(Banks);
          for BankIndex := 0 to Banks.Count - 1 do
            begin
              if RegKey.OpenKey(Registry + '\Cache\' + Banks[BankIndex] + '\', False) then
                begin
                  Dates := TStringList.Create;
                  RegKey.GetKeyNames(Dates);
                  for DateIndex := 0 to Dates.Count - 1 do
                    begin
                      KeyDate := EncodeDate(StrToInt(Copy(Dates[DateIndex],1,4)),StrToInt(Copy(Dates[DateIndex],5,2)),StrToInt(Copy(Dates[DateIndex],7,2)));
                      if (KeyDate < MinDate) then
                        RegDelete(Registry + '\Cache\' + Banks[BankIndex] + '\' + Dates[DateIndex]);
                    end;
                end;
            end;
        finally
          Banks.Free;
        end;
        RegKey.CloseKey;
      end;
  finally
    RegKey.Free;
  end;
end;

function RegDecrypt(Path, Key: string; Name: string = ''): string; stdcall;
begin
  result := StrDecrypt(RegRead(Path, Key), 553);
  if Name <> '' then
    result := StrCut(result, Name + '=', '|');
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

function GetRate(Source, Target: PWideChar; Date: TDateTime; BankIndex: integer = 0; LastKnown: boolean = true; Recall: boolean = false): Double; stdcall;
var
  Switch: PWideChar;
  TryDate: TDateTime;
  DATA, BANK, URL, NAT, SOM, EOM, SOS, EOS, SON, EON, SOR, EOR: string;
  INVERSE: boolean;
  Name: string;
begin
  Name := string(Source) + string(Target);
  TryDate := Date;
  if RegCache(Registry + '\Cache\' + IntToStr(BankIndex) + '\' + FormatDateTime('yyyymmdd', Date), Name, result) then
    exit;
  try
    BANK := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex));
    INVERSE := StrToBool(StrCut(BANK, 'INVERSE=','|','false'));
    URL := StrCut(BANK, 'MASKURL=','|');
    NAT := StrCut(BANK, 'DEFCODE=','|');
    if string(Source) = NAT then
      begin
        Switch := Source;
        Source := Target;
        Target := Switch;
      end;
    while (StrCut(DATA,SOR,EOR) = '') and ((Date-TryDate) < 14) do
      begin
        if (Date-TryDate > 0) and RegCache(Registry + '\Cache\' + IntToStr(BankIndex) + '\' + FormatDateTime('yyyymmdd', TryDate), Name, result) then
          exit
        else
          begin
            SOM := StrDynamic(StrCut(BANK, 'SOM=','|'), TryDate, Source, Target);
            EOM := StrDynamic(StrCut(BANK, 'EOM=','|'), TryDate, Source, Target);
            SOS := StrDynamic(StrCut(BANK, 'SOS=','|'), TryDate, Source, Target);
            EOS := StrDynamic(StrCut(BANK, 'EOS=','|'), TryDate, Source, Target);
            SON := StrDynamic(StrCut(BANK, 'SON=','|'), TryDate, Source, Target);
            EON := StrDynamic(StrCut(BANK, 'EON=','|'), TryDate, Source, Target);
            SOR := StrDynamic(StrCut(BANK, 'SOR=','|'), TryDate, Source, Target);
            EOR := StrDynamic(StrCut(BANK, 'EOR=','|'), TryDate, Source, Target);
            DATA := HttpGet(StrDynamic(URL, TryDate, Target, Source));
            DATA := GetBlock(DATA, Source, SOM, EOM);
            if (SOS <> '') and (EOS <> '') then
              DATA := StrCut(DATA, SOS, EOS);
          end;
        TryDate := TryDate - 1;
        if (StrCut(DATA,SOR,EOR) = '') and not LastKnown then
          exit(0);
      end;
    try
      result := (StrToFloat(StrCut(DATA,SOR,EOR)) / StrToFloat(StrCut(DATA,SON,EON,'1')));
      if (string(Source) <> NAT) and (string(Target) <> NAT) then
        result := result / GetRate(PWideChar(NAT), Target, TryDate, BankIndex, LastKnown, True);
      if (not Recall and (not INVERSE and (string(Switch) = NAT)) or (INVERSE and (string(Switch) <> NAT))) then
        result := 1 / result;
    except
      exit(-1);
    end;
  finally
    if (TryDate + 1 = Date) and (result > 0) then
      RegWrite(Registry + '\Cache\' + IntToStr(BankIndex) + '\' + FormatDateTime('yyyymmdd', TryDate), Name, FloatToStr(result));
  end;
end;

function Currency(Source, Target: PWideChar; StartDate, EndDate: TDateTime; BankIndex: integer = 0; LastKnown: boolean = true): Double; stdcall;
var
  Bank, Url: string;
  Index, Count, Days: integer;
  Sum, Currency: double;
begin
    begin
      Sum := 0;
      Days := 0;
      Count := Round(EndDate - StartDate);
      Bank := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME');
      Url := StrCut(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'MASKURL'),'//','/');
      for Index := 0 to Count do
        begin
          if ProgressBar.Process(Index, Count, 'Convert from "' + Source + '" to "' + Target + '", Date: ' + DateToStr(StartDate + Index) + ', Rate: ' + FloatToStr(Currency), 'Get exchange rate from ' + Bank + ' (' + Url + ')') = false then
            exit(-2);
          Currency := GetRate(Source, Target, StartDate + Index, BankIndex, LastKnown);
          case Round(Currency) of
            -1, -2:
              exit(Currency);
            0:
              Days := Days - 1;
            else
              Sum := (Sum + Currency);
          end;
          Days := Days + 1;
        end;
      if (Sum > 0) then
        result := Sum / Days
      else
        result := -1;
      OutputDebugString(PChar(FloatToStr(Sum) + ' / ' + IntToStr(Days) + ' -> ' + FloatToStr(result)));
    end;
end;

procedure CurrencyGUI(Settings: boolean = false);
var
  BankIndex: integer;
begin
  BankIndex := 0;
  while RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME') <> '' do
    BankIndex := BankIndex + 1;

  if (BankIndex > 0) and (Settings = false) then
    begin
      Application.CreateForm(TCurrencyRate, CurrencyRate);
      try
        CurrencyRate.ShowModal;
      finally
        CurrencyRate.Destroy;
      end;
    end
  else
    begin
      if not Settings then
        MessageDlg('Exchange rates'''' providers are missing. Please set exchange rates'''' providers before using "Currency Rate" tool.',mtError, [mbOK],0);
      Application.CreateForm(TCurrencyRateSet, CurrencyRateSet);
      try
        CurrencyRateSet.ShowModal;
      finally
        CurrencyRateSet.Destroy;
      end;
    end;
end;

procedure TCurrencyRate.Enabled(input: boolean);
begin
  Source.Enabled := input;
  Target.Enabled := input;
  StartDate.Enabled := input;
  EndDate.Enabled := input;
  Average.Enabled := input;
  ManualProv.Enabled := input;
  Provaiders.Enabled := input;
  ThridPart.Enabled := input;
  Exchange.Enabled := input;
end;

function TCurrencyRate.IndexOfBank(Name: string): integer;
begin
  result := 0;
  while ((RegDecrypt(Registry + '\Providers','BankData' + IntToStr(result), 'BANKNAME') <> Name) and (RegDecrypt(Registry + '\Providers','BankData' + IntToStr(result), 'BANKNAME') <> '')) do
    result := result + 1;
end;

procedure TCurrencyRate.CurrencyChange(Sender: TObject);
var
  BankIndex: integer;
  BankName, BankCode, BankCodes, TrdPart: string;
begin
  Provaiders.Items.Clear;
  BankIndex := 0;
  while RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME') <> '' do
    begin
      BankName := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME');
      BankCode := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'DEFCODE');
      BankCodes := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'SUPPORT');
      if ((Source.Text = BankCode) and ContainsText(BankCodes, Target.Text)) or ((Target.Text = BankCode) and ContainsText(BankCodes, Source.Text)) then
        if Provaiders.Items.IndexOf(BankName) = -1 then
          Provaiders.Items.Add(BankName);
      BankIndex := BankIndex + 1;
    end;
  if Provaiders.Items.Count = 0 then
    ThridPart.Checked := True;
  if ThridPart.Checked then
    begin
      BankIndex := 0;
      while RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME') <> '' do
        begin
          BankName := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME');
          BankCode := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'DEFCODE');
          BankCodes := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'SUPPORT');
          TrdPart := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), '3RDPART');
          if StrToBool(TrdPart) and ContainsText(BankCodes, Source.Text) and ContainsText(BankCodes, Target.Text) then
            if Provaiders.Items.IndexOf(BankName) = -1 then
              Provaiders.Items.Add(BankName);
          BankIndex := BankIndex + 1;
        end;
    end;
  if Provaiders.Items.Count = 0 then
    begin
      StartDate.Enabled := False;
      EndDate.Enabled := False;
      Average.Enabled := False;
      ManualProv.Enabled := False;
      Provaiders.Enabled := False;
      ThridPart.Enabled := False;
      Exchange.Enabled := False;
    end
  else
    begin
      StartDate.Enabled := True;
      EndDate.Enabled := True;
      Average.Enabled := True;
      ManualProv.Enabled := True;
      Provaiders.Enabled := True;
      ThridPart.Enabled := True;
      Exchange.Enabled := True;
      if CurrencyRate.Provaiders.ItemIndex = -1 then
        CurrencyRate.Provaiders.ItemIndex := 0;
    end;
end;

procedure TCurrencyRate.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  RegClearCache(12+MonthOf(Now));
end;

procedure TCurrencyRate.FormCreate(Sender: TObject);
var
  RegEx: TRegEx;
  CurCode: string;
  BankIndex: integer;
begin
  StartDate.MaxDate := Now;
  EndDate.MaxDate := Now;
  try
    StartDate.Date := StrToDate(RegRead(Registry + '\LastSettings','Start'));
  except
    StartDate.Date := Now;
  end;
  EndDate.MinDate := StartDate.Date;
  try
    EndDate.Date := StrToDate(RegRead(Registry + '\LastSettings','End'));
  except
    EndDate.Date := StartDate.Date;
  end;
  BankIndex := 0;
  while RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME') <> '' do
  begin
    Provaiders.Items.Add(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME'));
    for CurCode in RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'SUPPORT').Split([',']) do
    begin
      if Source.Items.IndexOf(CurCode) = -1 then Source.Items.Add(CurCode);
      if Target.Items.IndexOf(CurCode) = -1 then Target.Items.Add(CurCode);
    end;
    BankIndex := BankIndex + 1;
  end;
  if XlApp.ActiveCell.HasFormula and Extract(XlApp.ActiveCell.Formula, Replace, InSource, InTarget) then
    begin
      Formula := XlApp.ActiveCell.Formula;
      Source.ItemIndex := Source.Items.IndexOf(InSource);
      Target.ItemIndex := Target.Items.IndexOf(InTarget);
    end
  else if (Source.Items.Count > 0) and (Target.Items.Count > 0) then
    try
      Source.ItemIndex := Source.Items.IndexOf(RegRead(Registry + '\LastSettings','From'));
      Target.ItemIndex := Target.Items.IndexOf(RegRead(Registry + '\LastSettings','To'));
    except
      Source.ItemIndex := 0;
      Target.ItemIndex := 1;
    end;
  try
    ManualProv.Checked := StrToBool(RegRead(Registry + '\LastSettings','Manual'));
  except
    ManualProv.Checked := false;
  end;
  Provaiders.ItemIndex := 0;
  CurrencyChange(Self);
  ManualProvClick(Self.ManualProv);
  AverageClick(Self.Average);
end;

procedure TCurrencyRate.TargetChange(Sender: TObject);
begin
  Exchange.Enabled := not (CurrencyRate.Source.Text = CurrencyRate.Target.Text);
  ThridPart.Checked := False;
  CurrencyChange(Self.Target);
end;

procedure TCurrencyRate.ManualProvClick(Sender: TObject);
begin
  if not ManualProv.Checked then
    ThridPart.Checked := false;
  Provaiders.Enabled := ManualProv.Checked;
  ThridPart.Enabled := ManualProv.Checked;
end;

procedure TCurrencyRate.ProvaidersChange(Sender: TObject);
begin
  Exchange.Caption := 'Get rate from ' + Provaiders.Text;
end;

procedure TCurrencyRate.SourceChange(Sender: TObject);
begin
  Exchange.Enabled := not (CurrencyRate.Source.Text = CurrencyRate.Target.Text);
  ThridPart.Checked := False;
  CurrencyChange(Self.Source);
end;

procedure TCurrencyRate.StartDateChange(Sender: TObject);
begin
  EndDate.MinDate := StartDate.Date;
  if not Average.Checked then EndDate.Date := StartDate.Date;
end;

procedure TCurrencyRate.AverageClick(Sender: TObject);
begin
  if Average.Checked and not EndDate.Visible then
    begin
      if IsQuarter(EndDate.Date) then
        StartDate.Date := Quarter(EndDate.Date, -1, StartDate.MinDate, StartDate.MaxDate);
      LStartDate.Caption := 'From date:';
    end
  else
    begin
      LStartDate.Caption := 'Date:';
      StartDate.Date := EndDate.Date;
    end;
  if not Average.Checked then
    EndDate.Date := StartDate.Date;
  EndDate.Enabled := Average.Checked;
  LEndDate.Enabled := Average.Checked;
  EndDate.Visible := Average.Checked;
  LEndDate.Visible := Average.Checked;
end;

procedure TCurrencyRate.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Index: integer;
begin
  case Key of
    13:
      if Exchange.Enabled then
        ExchangeClick(Self.Exchange);
    27:
      CurrencyRate.Close;
    192:
      begin
        Index := Source.ItemIndex;
        Source.ItemIndex := Target.ItemIndex;
        Target.ItemIndex := Index;
      end;
  end;
end;

procedure TCurrencyRate.ProvaidersKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    13:
      if Exchange.Enabled then
        ExchangeClick(Self.Exchange);
    27:
      CurrencyRate.Close;
    192:
      begin
        ThridPart.Checked := not ThridPart.Checked;
        CurrencyChange(Self.Provaiders);
      end;
  end;
end;

procedure TCurrencyRate.StartDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    32:
      StartDate.Date := StartDate.MaxDate;
    33:
      StartDate.Date := Quarter(StartDate.Date, 1, StartDate.MinDate, StartDate.MaxDate);
    34:
      StartDate.Date := Quarter(StartDate.Date, -1, StartDate.MinDate, StartDate.MaxDate);
    35:
      StartDate.Date := Quarter(Now, 1, StartDate.MinDate, StartDate.MaxDate);
    192:
      begin
        Average.Checked := not Average.Checked;
        EndDate.Visible := not EndDate.Visible;
        AverageClick(Self.Average);
      end;
    else
      FormKeyDown(Sender, Key, Shift);
  end;
  StartDateChange(Self.StartDate);
end;

procedure TCurrencyRate.EndDateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    32:
      EndDate.Date := EndDate.MaxDate;
    33:
      EndDate.Date := Quarter(EndDate.Date, 1, EndDate.MinDate, StartDate.MaxDate);
    34:
      EndDate.Date := Quarter(EndDate.Date, -1, EndDate.MinDate, StartDate.MaxDate);
    35:
      EndDate.Date := Quarter(Now, 1, EndDate.MinDate, StartDate.MaxDate);
    192:
      begin
        Average.Checked := not Average.Checked;
        EndDate.Visible := not EndDate.Visible;
        AverageClick(Self.Average);
      end;
    else
      FormKeyDown(Sender, Key, Shift);
  end;
end;

procedure TCurrencyRate.ExchangeClick(Sender: TObject);
var
  Name: string;
  Value: double;
begin
  Enabled(false);
  if ActiveConnection then
    begin
      Value := Currency(PWideChar(Source.Text),PWideChar(Target.Text),StartDate.Date,EndDate.Date,IndexOfBank(Provaiders.Text), LastestRate.Checked);
      if Value > 0 then
        try
          Name := Source.Text + Target.Text + '_' + DateToString(StartDate.Date, EndDate.Date);
          XlApp.Names.Add(Name, Value, True, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
//          XlApp.ActiveCell.Value2 := Value;
          if not (ContainsText(Formula, Replace) and (Formula <> '=' + Replace) and (Copy(Replace, 1, 6) = Source.Text + Target.Text)) then
            XlApp.ActiveCell.Formula := '=' + Name
          else
            if MessageDlg('Do you want update convertion from "' + InSource + '" to "' + InTarget + '" in formula ?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then
              XlApp.ActiveCell.Formula := StringReplace(Formula, Replace, Name, []);

        finally
          if StartDate.Date = EndDate.Date then Average.Checked := false;
          RegWrite(Registry + '\LastSettings','Average',BoolToStr(Average.Checked));
          RegWrite(Registry + '\LastSettings','Manual',BoolToStr(ManualProv.Checked));
          RegWrite(Registry + '\LastSettings','From',Source.Text);
          RegWrite(Registry + '\LastSettings','To',Target.Text);
          if StartDate.Date = Date() then
            begin
              RegDelete(Registry + '\LastSettings','Start');
              RegDelete(Registry + '\LastSettings','End');
            end
          else
            begin
              RegWrite(Registry + '\LastSettings','Start',DateToStr(StartDate.Date));
              RegWrite(Registry + '\LastSettings','End',DateToStr(EndDate.Date));
            end;
        end
    else
      case Round(Value) of
        -1: MessageDlg('Currency exchange failed.',mtError, [mbOK],0);
        -2: MessageDlg('Currency exchange canceled.',mtInformation, [mbOK],0);
      end;
    end;
  CurrencyRate.Close;
end;

end.
