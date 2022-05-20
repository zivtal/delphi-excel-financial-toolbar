unit XlActivation;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Variants,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  StrUtils,
  Clipbrd,
  WebCtrl,
  RegistryCtrl,
  StringCtrl,
  AES128;

  function ActiveNow(): boolean;
  function Activated(Warning: boolean = true): boolean;
  function ActDays(): integer;
  function ActName(): string;
  function ActEmail(): string;

type
  TActivationGUI = class(TForm)
    SerialBox: TEdit;
    LSerial: TLabel;
    LKey: TLabel;
    KeyBox: TMemo;
    OkButton: TButton;
    LoadButton: TButton;
    Title: TLabel;
    Copyrights: TLabel;
    LName: TLabel;
    LEmail: TLabel;
    Name: TEdit;
    EMail: TEdit;
    CopyClipboard: TButton;
    procedure LoadButtonClick(Sender: TObject);
    procedure OkButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure NameExit(Sender: TObject);
    procedure EMailExit(Sender: TObject);
    procedure CopyClipboardClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

var
  AppKey: string = 'Key';
  AppPath: string = '\SOFTWARE\Ziv Tal\ZTools';
  AppName: string = 'ztxl';
  ActivationGUI: TActivationGUI;
  Ok: boolean = false;
  masterkey: string = '&ST+t}wcqkN23:3s"4?U';

implementation

{$R *.dfm}

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

// create activation form
function ActiveNow(): boolean;
begin
  Ok := false;
  Application.CreateForm(TActivationGUI, ActivationGUI);
  try
    ActivationGUI.ShowModal;
    if (ActivationGUI.KeyBox.Text <> '') and Ok then
      RegWrite(AppPath, AppKey, ActivationGUI.KeyBox.Text);
  finally
    result := Activated(false);
    ActivationGUI.Destroy;
  end;
end;

// activate or renew
function Activate(Mb, Nm, Em: string; De: integer): string;
var
  ReGen: string;
begin
  ReGen := 'APP=' + AppName + ';DE=' + IntToStr(De) + ';AD=' + datetostr(now) + ';ED=' + DateToStr(Now + De) + ';PA=FALSE' + ';LU=' + DateToStr(Now) + ';MB=' + Mb + ';DU=0;NM=' + Nm +';EM=' + Em + ';';
  result := ReGen;
  ReGen := StrEncrypt( AES128_Encrypt(ReGen, masterkey) , 322);
  RegWrite(AppPath, AppKey, ReGen);
  MessageDlg('Activation successful, application will expire in ' + inttostr(De) + ' days.',mtInformation, [mbOK],0);
end;

// generate new serial number
function Serial(Name, Email: string): string;
  function Username: String;
  var
    nSize: DWord;
  begin
    nSize := 1024;
    SetLength(Result, nSize);
    if GetUserName(PChar(Result), nSize) then
      SetLength(Result, nSize-1)
    else
      RaiseLastOSError;
  end;
var
  SN, Nm, Em, Id, Un: string;
begin
  Nm := 'NM=' + Name + ';';
  Em := 'EM=' + EMail + ';';
  Id := 'ID=' + GetMotherBoardSerial + ';';
  Un := 'UN=' + Username + ';';
  SN := Id + Un + Nm + Em;
  result := StrEncrypt(SN,398);
end;

// update usge status
function DailyCheck(Reg: string): boolean;
var
  De, Ad, Ed, Lu, Du: string;
  ReDe: integer;
begin
  result := true;
  De := StrCut(Reg, 'DE=', ';'); // days to end
  Ad := StrCut(Reg, 'AD=', ';'); // activation date
  Ed := StrCut(Reg, 'ED=', ';'); // expire date
  Lu := StrCut(Reg, 'LU=', ';'); // expire date
  Du := StrCut(Reg, 'DU=', ';'); // expire date
  ReDe := Round(((StrToDate(Ed) - StrToDate(Ad)) + (StrToDate(Ad) - Date)));
  if ((StrToDate(Ad) > Date) or (ReDe > StrToInt(De)) or (StrToDate(Lu) > Date)) then
    begin
      MessageDlg('Invalid activation key.',mtError, [mbOK],0);
      RegDelete(AppPath, AppKey);
      result := false;
    end
  else
    begin
      Reg := StringReplace(reg, 'DE=' + De + ';', 'DE=' + IntToStr(ReDe) + ';', [rfReplaceAll, rfIgnoreCase]);
      Reg := StringReplace(reg, 'DU=' + Du + ';', 'DU=' + IntToStr(Round(StrToDate(Ad) - Date)) + ';', [rfReplaceAll, rfIgnoreCase]);
      Reg := StringReplace(reg, 'LU=' + Lu + ';', 'LU=' + DateToStr(Now) + ';', [rfReplaceAll, rfIgnoreCase]);
      Reg := StrEncrypt( AES128_Encrypt(Reg, masterkey) , 322);
      RegWrite(AppPath, AppKey, Reg);
    end;
end;

// Activarion check
function Activated(Warning: boolean = true): boolean;
var
  Reg, Ap, Nm, Em, Mb, Ad, Ed, De, Md, Pa, Lu, Du: string;
  ReDe: integer;
begin
  Reg := RegRead(AppPath, AppKey);
  try
    Reg := StrDecrypt(Reg, 322);
    Reg := AES128_Decrypt(Reg , masterkey);
    Ap := StrCut(Reg, 'APP=', ';'); // application name
    Nm := StrCut(Reg, 'NM=', ';'); // name
    Em := StrCut(Reg, 'EM=', ';'); // email
    Mb := StrCut(Reg, 'MB=', ';'); // motherboard serial number
    De := StrCut(Reg, 'DE=', ';'); // days to end
    Md := StrCut(Reg, 'MD=', ';'); // maximum days for activation
    Pa := StrCut(Reg, 'PA=', ';', 'FALSE'); // pre-activate
    if ( StrToBool(Pa) and (Mb = GetMotherBoardSerial()) and (StrToDate(Md) >= Now) ) then
      Reg := Activate(Mb, Nm, Em, StrToInt(De));
    Ed := StrCut(Reg, 'ED=', ';'); // expire date
    Ad := StrCut(Reg, 'AD=', ';'); // activation date
  finally
    result := false;
    if (Ap = AppName) and (Mb = GetMotherBoardSerial()) and (StrToDate(Ed) >= Now) and DailyCheck(Reg) then
      result := true
    else
      begin
        if (Ap = AppName) then
          begin
            if Mb <> GetMotherBoardSerial() then MessageDlg('Invalid activation key.',mtError, [mbOK],0);
            if StrToDate(Ed) < Now then MessageDlg('Application has been expired.',mtInformation, [mbOK],0);
          end
        else
          if Warning then
            begin
              MessageDlg('Application not activated.',mtInformation, [mbOk],0);
              result := ActiveNow;
            end;
      end;
  end;
end;

function ActDays(): integer;
var
  Reg: string;
begin
  Reg := RegRead(AppPath, AppKey);
  try
    Reg := StrDecrypt(Reg, 322);
    Reg := AES128_Decrypt(Reg , masterkey);
  finally
    result := round(StrToDate(StrCut(Reg, 'ED=', ';')) - Date);
  end;
end;

function ActName(): string;
var
  Reg: string;
begin
  Reg := RegRead(AppPath, AppKey);
  try
    Reg := StrDecrypt(Reg, 322);
    Reg := AES128_Decrypt(Reg , masterkey);
  finally
    result := StrCut(Reg, 'NM=', ';');
  end;
end;

function ActEmail(): string;
var
  Reg: string;
begin
  Reg := RegRead(AppPath, AppKey);
  try
    Reg := StrDecrypt(Reg, 322);
    Reg := AES128_Decrypt(Reg , masterkey);
  finally
    result := StrCut(Reg, 'EM=', ';');
  end;
end;

procedure TActivationGUI.LoadButtonClick(Sender: TObject);
var
  OpenDialog: TOpenDialog;
  Data, Decrypted: string;
begin
  OpenDialog := TOpenDialog.Create(self);
  OpenDialog.Options := [ofFileMustExist];
  OpenDialog.Filter := 'ZTool key file|*.key';
  OpenDialog.FilterIndex := 1;
  if OpenDialog.Execute then
  begin
    Data := ReadFile(OpenDialog.FileName);
    Decrypted := StrDecrypt(Data, 322);
    Decrypted := AES128_Decrypt(Decrypted, masterkey);
    if StrCut(Decrypted, 'APP=', ';') = AppName then
      KeyBox.Text := Data
    else
      MessageDlg('File not compatible or corrupted.', mtError, [mbOk], 0);
  end;
end;

procedure TActivationGUI.NameExit(Sender: TObject);
begin
  if Name.Text <> '' then
    begin
      RegWrite('\SOFTWARE\Ziv Tal\','Username',Name.Text);
      LName.Font.Color := clWindowText;
    end
  else
    begin
      Name.SetFocus;
      LName.Font.Color := clRed;
      MessageDlg('You must enter full name to continue.', mtError, [mbOk], 0);
    end;

  if (Name.Text <> '') and IsValidEmail(EMail.Text) then
    SerialBox.Text := Serial(Name.Text, EMail.Text)
  else
    SerialBox.Text := '';
end;

procedure TActivationGUI.CopyClipboardClick(Sender: TObject);
begin
  ClipBoard.AsText := SerialBox.Text;
end;

procedure TActivationGUI.EMailExit(Sender: TObject);
begin
  if IsValidEmail(EMail.Text) then
    begin
      RegWrite('\SOFTWARE\Ziv Tal\','Email',Email.Text);
      LEmail.Font.Color := clWindowText;
    end
  else
    begin
      Email.SetFocus;
      LEmail.Font.Color := clRed;
      MessageDlg('You must enter valid email address to continue.', mtError, [mbOk], 0);
    end;

  if (Name.Text <> '') and IsValidEmail(EMail.Text) then
    SerialBox.Text := Serial(Name.Text, EMail.Text)
  else
    SerialBox.Text := '';
end;

procedure TActivationGUI.OkButtonClick(Sender: TObject);
var
  Decrypted: string;
begin
  Decrypted := StrDecrypt(KeyBox.Text, 322);
  Decrypted := AES128_Decrypt(Decrypted, masterkey);
  if (StrCut(Decrypted, 'APP=', ';') = AppName) and (StrCut(Decrypted, 'NM=', ';') = Name.Text) and (StrCut(Decrypted, 'EM=', ';') = Email.Text) then
    Ok := true
  else
    MessageDlg('Invalid activation key.', mtError, [mbOk], 0);
  ActivationGUI.Close;
end;

procedure TActivationGUI.FormCreate(Sender: TObject);
begin
  Title.Width := ActivationGUI.Width;
  Copyrights.Width := ActivationGUI.Width;
  try
    Name.Text := RegRead('\SOFTWARE\Ziv Tal\','Username');
    Email.Text := RegRead('\SOFTWARE\Ziv Tal\','Email');
  finally
    if (Name.Text <> '') then
      NameExit(Self.Name);
    if (EMail.Text <> '') then
      EMailExit(Self.EMail);
  end;
end;

procedure TActivationGUI.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    27:
      ActivationGUI.Close;
  end;
end;

end.
