unit XlCurrencyRateSet;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.StdCtrls,
  Vcl.ComCtrls,
  Registry,
  StrUtils,
  Wininet,
  StringCtrl,
  RegistryCtrl,
  WebCtrl;

type
  TCurrencyRateSet = class(TForm)
    Label1: TLabel;
    BANKNAME: TComboBox;
    Save: TButton;
    Remove: TButton;
    Export: TButton;
    Import: TButton;
    Label2: TLabel;
    MASKURL: TEdit;
    Label3: TLabel;
    LISTURL: TEdit;
    GroupBox1: TGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    SOM: TEdit;
    EOM: TEdit;
    GroupBox2: TGroupBox;
    Label6: TLabel;
    Label7: TLabel;
    SOS: TEdit;
    EOS: TEdit;
    GroupBox3: TGroupBox;
    Label8: TLabel;
    Label9: TLabel;
    SOC: TEdit;
    EOC: TEdit;
    GroupBox4: TGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    SON: TEdit;
    EON: TEdit;
    GroupBox5: TGroupBox;
    Label12: TLabel;
    Label13: TLabel;
    SOR: TEdit;
    EOR: TEdit;
    GroupBox6: TGroupBox;
    Label14: TLabel;
    Label15: TLabel;
    DEFCODE: TEdit;
    SUPPORT: TEdit;
    Load: TButton;
    GroupBox7: TGroupBox;
    Label17: TLabel;
    Label16: TLabel;
    TestDate: TDateTimePicker;
    TestTarget: TEdit;
    Label18: TLabel;
    PTarget: TEdit;
    Label19: TLabel;
    PSource: TEdit;
    Label20: TLabel;
    PNominal: TEdit;
    PRate: TEdit;
    Label21: TLabel;
    GroupBox8: TGroupBox;
    Test: TButton;
    INVERSE: TCheckBox;
    DATA: TMemo;
    THIRDPART: TCheckBox;
    ClearCache: TButton;
    function CRP(): string;
    procedure CheckChanges(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BANKNAMEChange(Sender: TObject);
    procedure TestClick(Sender: TObject);
    procedure SaveClick(Sender: TObject);
    procedure RemoveClick(Sender: TObject);
    procedure ImportClick(Sender: TObject);
    procedure ExportClick(Sender: TObject);
    procedure LoadClick(Sender: TObject);
    procedure LISTURLChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ClearCacheClick(Sender: TObject);
  end;

var
  CurrencyRateSet: TCurrencyRateSet;

implementation

{$R *.dfm}

const
  Registry: string = '\SOFTWARE\Ziv Tal\CurrencyRate';

function ReadFile(Filename: string): string;
var
   fileData : TStringList;
   saveLine : String;
   lines, i : Integer;
 begin
   fileData := TStringList.Create;        // Create the TSTringList object
   fileData.LoadFromFile(Filename);     // Load from Testing.txt file

   // Reverse the sequence of lines in the file
   lines := fileData.Count;

   for i := lines-1 downto (lines div 2) do
   begin
     saveLine := fileData[lines-i-1];
     fileData[lines-i-1] := fileData[i];
     fileData[i] := saveLine;
   end;

   // Now display the file
   for i := 0 to lines-1 do
     result := result + fileData[i];

   fileData.SaveToFile(Filename);       // Save the reverse sequence file
 end;

procedure WriteFile(Filename, Data: string);
var
  myFile : TextFile;
begin
  try
    AssignFile(myFile, Filename);
    ReWrite(myFile);
    WriteLn(myFile, Data);
  finally
    CloseFile(myFile);
  end;
end;

function RegDecrypt(Path, Key: string; Name: string = ''): string; stdcall;
begin
  result := StrDecrypt(RegRead(Path, Key), 553);
  if Name <> '' then
    result := StrCut(result, Name + '=', '|');
end;

function GetBlock(Input, Find, StartSgmt, CloseSgmt: string): string; stdcall;
var
  Index: integer;
begin
  if Input = '' then
    result := ''
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

function TCurrencyRateSet.CRP(): string;
begin
  result := result + 'BANKNAME=' + BANKNAME.Text + '|';
  result := result + 'MASKURL=' + MASKURL.Text + '|';
  result := result + 'LISTURL=' + LISTURL.Text + '|';
  result := result + 'DEFCODE=' + DEFCODE.Text + '|';
  result := result + 'SUPPORT=' + SUPPORT.Text + '|';
  result := result + 'INVERSE=' + BoolToStr(INVERSE.Checked) + '|';
  result := result + '3RDPART=' + BoolToStr(THIRDPART.Checked) + '|';
  if SOM.Text <> '' then result := result + 'SOM=' + SOM.Text + '|';
  if EOM.Text <> '' then result := result + 'EOM=' + EOM.Text + '|';
  if SOS.Text <> '' then result := result + 'SOS=' + SOS.Text + '|';
  if EOS.Text <> '' then result := result + 'EOS=' + EOS.Text + '|';
  if SOC.Text <> '' then result := result + 'SOC=' + SOC.Text + '|';
  if EOC.Text <> '' then result := result + 'EOC=' + EOC.Text + '|';
  if SON.Text <> '' then result := result + 'SON=' + SON.Text + '|';
  if EON.Text <> '' then result := result + 'EON=' + EON.Text + '|';
  if SOR.Text <> '' then result := result + 'SOR=' + SOR.Text + '|';
  if EOR.Text <> '' then result := result + 'EOR=' + EOR.Text + '|';
end;

procedure TCurrencyRateSet.CheckChanges(Sender: TObject);
begin
  Save.Enabled := false;
  // Check for changes
  if BANKNAME.Text <> '' then
    if not (BANKNAME.Items.IndexOf(BANKNAME.Text) = -1) then
      begin
        Save.Caption := 'Save';
        Remove.Enabled := true;
        if MASKURL.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'MASKURL') then Save.Enabled := true;
        if LISTURL.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'LISTURL') then Save.Enabled := true;
        if DEFCODE.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'DEFCODE') then Save.Enabled := true;
        if SUPPORT.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SUPPORT') then Save.Enabled := true;
        if INVERSE.Checked <> StrToBool(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'INVERSE')) then Save.Enabled := true;
        if THIRDPART.Checked <> StrToBool(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), '3RDPART')) then Save.Enabled := true;
        if SOM.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOM') then Save.Enabled := true;
        if EOM.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOM') then Save.Enabled := true;
        if SOS.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOS') then Save.Enabled := true;
        if EOS.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOS') then Save.Enabled := true;
        if SOC.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOC') then Save.Enabled := true;
        if EOC.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOC') then Save.Enabled := true;
        if SON.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SON') then Save.Enabled := true;
        if EON.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EON') then Save.Enabled := true;
        if SOR.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOR') then Save.Enabled := true;
        if EOR.Text <> RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOR') then Save.Enabled := true;
      end
    else
      begin
        Save.Caption := 'Add';
        Remove.Enabled := false;
        if (MASKURL.Text <> '')
        and (DEFCODE.Text <> '')
        and (SUPPORT.Text <> '')
        then Save.Enabled := true;
      end;
end;

procedure TCurrencyRateSet.ClearCacheClick(Sender: TObject);
begin
  if BANKNAME.ItemIndex > -1 then
    begin
      RegDelete(Registry + '\Cache\' + IntToStr(BANKNAME.ItemIndex));
      MessageDlg('Exchange cache of "' + BANKNAME.Text + '" has been cleared.', mtInformation, [mbOk], 0);
    end
  else
    begin
      RegDelete(Registry + '\Cache');
      MessageDlg('All exchange cache has been cleared.', mtInformation, [mbOk], 0);
    end;
end;

procedure TCurrencyRateSet.BANKNAMEChange(Sender: TObject);
begin
  if not (BANKNAME.Items.IndexOf(BANKNAME.Text) = -1) then
    begin
      Save.Caption := 'Save';
      Remove.Enabled := true;
      MASKURL.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'MASKURL');
      LISTURL.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'LISTURL');
      DEFCODE.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'DEFCODE');
      SUPPORT.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SUPPORT');
      INVERSE.Checked := StrToBool(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'INVERSE'));
      THIRDPART.Checked := StrToBool(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), '3RDPART'));
      SOM.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOM');
      EOM.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOM');
      SOS.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOS');
      EOS.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOS');
      SOC.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOC');
      EOC.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOC');
      SON.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SON');
      EON.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EON');
      SOR.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'SOR');
      EOR.Text := RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BANKNAME.Items.IndexOf(BANKNAME.Text)), 'EOR');
      LISTURLChange(Self.LISTURL);
    end
  else
    begin
      Save.Caption := 'Add';
      Remove.Enabled := false;
    end;
end;

procedure TCurrencyRateSet.ExportClick(Sender: TObject);
var
  SaveDialog : TSaveDialog;
begin
  SaveDialog := TSaveDialog.Create(self);
  SaveDialog.Title := 'Save currency rate preset';
  SaveDialog.Filter := 'Currency Rate Preset|*.crp';
  SaveDialog.DefaultExt := 'crp';
  SaveDialog.FilterIndex := 1;
  SaveDialog.FileName := BANKNAME.Text + '.crp';
  SaveDialog.Options := SaveDialog.Options + [ofOverwritePrompt];
  if SaveDialog.Execute then
    WriteFile(SaveDialog.FileName, StrEncrypt('CRP:'+CRP(),553))
  else
    MessageDlg('File was not saved.', mtInformation, [mbOk], 0);
end;

procedure TCurrencyRateSet.FormCreate(Sender: TObject);
var
  BankIndex: integer;
begin
  BankIndex := 0;
  while RegRead(Registry + '\Providers','BankData' + IntToStr(BankIndex)) <> '' do
  begin
    BANKNAME.Items.Add(RegDecrypt(Registry + '\Providers','BankData' + IntToStr(BankIndex), 'BANKNAME'));
    BankIndex := BankIndex + 1;
  end;
end;

procedure TCurrencyRateSet.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 84) and (Shift = [ssCtrl]) then
    if (GroupBox7.Visible = False)  then
      begin
        CurrencyRateSet.Height := CurrencyRateSet.Height + GroupBox7.Height + 12;
        GroupBox7.Visible := True;
      end
    else
      begin
        CurrencyRateSet.Height := CurrencyRateSet.Height - GroupBox7.Height - 12;
        GroupBox7.Visible := False;
      end;
end;

procedure TCurrencyRateSet.ImportClick(Sender: TObject);
var
  OpenDialog: TOpenDialog;
  CRP: string;
begin
  OpenDialog := TOpenDialog.Create(self);
  OpenDialog.Options := [ofFileMustExist];
  OpenDialog.Filter := 'Currency Rate Preset|*.crp';
  OpenDialog.FilterIndex := 1;
  if OpenDialog.Execute then
  begin
    CRP := StrDecrypt(ReadFile(OpenDialog.FileName),553);
    if Copy(CRP,1,3) = 'CRP' then
      begin
        BANKNAME.Text := StrCut(CRP, 'BANKNAME=','|');
        MASKURL.Text := StrCut(CRP, 'MASKURL=','|');
        LISTURL.Text := StrCut(CRP, 'LISTURL=','|');
        DEFCODE.Text := StrCut(CRP, 'DEFCODE=','|');
        SUPPORT.Text := StrCut(CRP, 'SUPPORT=','|');
        INVERSE.Checked := StrToBool(StrCut(CRP, 'INVERSE=','|'));
        THIRDPART.Checked := StrToBool(StrCut(CRP, '3RDPART=','|'));
        SOM.Text := StrCut(CRP, 'SOM=','|');
        EOM.Text := StrCut(CRP, 'EOM=','|');
        SOS.Text := StrCut(CRP, 'SOS=','|');
        EOS.Text := StrCut(CRP, 'EOS=','|');
        SOC.Text := StrCut(CRP, 'SOC=','|');
        EOC.Text := StrCut(CRP, 'EOC=','|');
        SON.Text := StrCut(CRP, 'SON=','|');
        EON.Text := StrCut(CRP, 'EON=','|');
        SOR.Text := StrCut(CRP, 'SOR=','|');
        EOR.Text := StrCut(CRP, 'EOR=','|');
        CheckChanges(Self.Import);
      end
    else
      MessageDlg('File not compatible or corrupted.', mtError, [mbOk], 0);
  end;
end;

procedure TCurrencyRateSet.LISTURLChange(Sender: TObject);
begin
//  if LISTURL.Text = '' then
//    Load.Enabled := false
//  else
//    Load.Enabled := true;
  CheckChanges(Self.LISTURL);
end;

procedure TCurrencyRateSet.LoadClick(Sender: TObject);
var
  HttpData: string;
begin
  if (LISTURL.Text = '') then
    HttpData := HttpGet(StrDynamic(MASKURL.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text))
  else
    HttpData := HttpGet(StrDynamic(LISTURL.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text));
  SUPPORT.Text := DEFCODE.Text;
  while StrCut(HttpData,SOC.Text,EOC.Text) <> '' do
    begin
      if not ContainsText(SUPPORT.Text, StrCut(HttpData,SOC.Text,EOC.Text)) then
        SUPPORT.Text := SUPPORT.Text + ',' + StrCut(HttpData,SOC.Text,EOC.Text);
      HttpData := StrCut(HttpData,SOC.Text,'');
    end;
end;

procedure TCurrencyRateSet.RemoveClick(Sender: TObject);
var
  BankReIndex: integer;
begin
  RegDelete(Registry + '\Cache');
//  RegDelete(Registry + '\Providers','BankData' + IntToStr(BANKNAME.ItemIndex));
  BankReIndex := BANKNAME.ItemIndex +1;
  while RegRead(Registry + '\Providers','BankData' + IntToStr(BankReIndex)) <> '' do
    begin
      RegWrite(Registry + '\Providers','BankData' + IntToStr(BankReIndex-1), RegRead(Registry + '\Providers','BankData' + IntToStr(BankReIndex)));
      BankReIndex := BankReIndex +1;
    end;
  RegDelete(Registry + '\Providers','BankData' + IntToStr(BankReIndex-1));
  BANKNAME.Items.Delete(BANKNAME.ItemIndex);
end;

procedure TCurrencyRateSet.SaveClick(Sender: TObject);
begin
  if BANKNAME.Items.IndexOf(BANKNAME.Text) = -1 then
    begin
      RegWrite(Registry + '\Providers','BankData', StrEncrypt('CRP:'+CRP,553), false, true);
      BANKNAME.Items.Add(BANKNAME.Text);
    end
  else
    RegWrite(Registry + '\Providers','BankData' + IntToStr(BANKNAME.ItemIndex), StrEncrypt('CRP:'+CRP,553), true, false);
  CheckChanges(Self.Save);
end;

procedure TCurrencyRateSet.TestClick(Sender: TObject);
var
  connection: cardinal;
  HttpData: string;
begin
  if not InternetGetConnectedState(@connection,0) then
    MessageDlg('Internet connection not available.',mtError, [mbOK],0)
  else
    begin
      HttpData := HttpGet(StrDynamic(MASKURL.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text));
      if (SOM.Text <> '') and (EOM.Text <> '') then
        DATA.Text := GetBlock(HttpData, TestTarget.Text, StrDynamic(SOM.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text), StrDynamic(EOM.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text))
      else
        DATA.Text := HttpData;
      if (SOS.Text <> '') and (EOS.Text <> '') then
        DATA.Text := StrCut(DATA.Text, StrDynamic(SOS.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text), StrDynamic(EOS.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text));
      PTarget.Text := TestTarget.Text;
      PSource.Text := DEFCODE.Text;
      PNominal.Text := StrCut(DATA.Text, StrDynamic(SON.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text), StrDynamic(EON.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text));
      PRate.Text := StrCut(DATA.Text, StrDynamic(SOR.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text), StrDynamic(EOR.Text, TestDate.Date, DEFCODE.Text, TestTarget.Text));
    end;
end;

end.
