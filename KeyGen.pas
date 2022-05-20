unit KeyGen;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, ActiveX, ComObj, IdCoderMIME, IdGlobal,
  Vcl.StdCtrls, Vcl.ComCtrls, StrUtils, Vcl.Clipbrd;

type
  TKeyGenerator = class(TForm)
    GroupBox1: TGroupBox;
    Serial: TEdit;
    Email: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Motherboard: TLabel;
    Name: TLabel;
    ChkDe: TRadioButton;
    De: TEdit;
    ChkEd: TRadioButton;
    Ed: TDateTimePicker;
    ExportButton: TButton;
    Output: TMemo;
    Decrypted: TMemo;
    CopyClipboard: TButton;
    procedure SerialChange(Sender: TObject);
    procedure DeChange(Sender: TObject);
    procedure ChkEdClick(Sender: TObject);
    procedure ChkDeClick(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure OutputChange(Sender: TObject);
    procedure CopyClipboardClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  KeyGenerator: TKeyGenerator;

implementation

{$R *.dfm}

var
  masterkey: string = '&ST+t}wcqkN23:3s"4?U';
  AppName: string = 'ztxl';

const CKEY1 = 53761;
      CKEY2 = 32618;

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

// Base64 Encode/Decode

function Base64_Encode(Value: TBytes): string;
var
  Encoder: TIdEncoderMIME;
begin
  Encoder := TIdEncoderMIME.Create(nil);
  try
    Result := Encoder.EncodeBytes(TIdBytes(Value));
  finally
    Encoder.Free;
  end;
end;

function Base64_Decode(Value: string): TBytes;
var
  Encoder: TIdDecoderMIME;
begin
  Encoder := TIdDecoderMIME.Create(nil);
  try
    Result := TBytes(Encoder.DecodeBytes(Value));
  finally
    Encoder.Free;
  end;
end;

// WinCrypt.h

type
  HCRYPTPROV  = Cardinal;
  HCRYPTKEY   = Cardinal;
  ALG_ID      = Cardinal;
  HCRYPTHASH  = Cardinal;

const
  _lib_ADVAPI32    = 'ADVAPI32.dll';
  CALG_SHA_256     = 32780;
  CALG_AES_128     = 26126;
  CRYPT_NEWKEYSET  = $00000008;
  PROV_RSA_AES     = 24;
  KP_MODE          = 4;
  CRYPT_MODE_CBC   = 1;

function CryptAcquireContext(var Prov: HCRYPTPROV; Container: PChar; Provider: PChar; ProvType: LongWord; Flags: LongWord): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptAcquireContextW';
function CryptDeriveKey(Prov: HCRYPTPROV; Algid: ALG_ID; BaseData: HCRYPTHASH; Flags: LongWord; var Key: HCRYPTKEY): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptDeriveKey';
function CryptSetKeyParam(hKey: HCRYPTKEY; dwParam: LongInt; pbData: PBYTE; dwFlags: LongInt): LongBool stdcall; stdcall; external _lib_ADVAPI32 name 'CryptSetKeyParam';
function CryptEncrypt(Key: HCRYPTKEY; Hash: HCRYPTHASH; Final: LongBool; Flags: LongWord; pbData: PBYTE; var Len: LongInt; BufLen: LongInt): LongBool;stdcall;external _lib_ADVAPI32 name 'CryptEncrypt';
function CryptDecrypt(Key: HCRYPTKEY; Hash: HCRYPTHASH; Final: LongBool; Flags: LongWord; pbData: PBYTE; var Len: LongInt): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptDecrypt';
function CryptCreateHash(Prov: HCRYPTPROV; Algid: ALG_ID; Key: HCRYPTKEY; Flags: LongWord; var Hash: HCRYPTHASH): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptCreateHash';
function CryptHashData(Hash: HCRYPTHASH; Data: PChar; DataLen: LongWord; Flags: LongWord): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptHashData';
function CryptReleaseContext(hProv: HCRYPTPROV; dwFlags: LongWord): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptReleaseContext';
function CryptDestroyHash(hHash: HCRYPTHASH): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptDestroyHash';
function CryptDestroyKey(hKey: HCRYPTKEY): LongBool; stdcall; external _lib_ADVAPI32 name 'CryptDestroyKey';

//-------------------------------------------------------------------------------------------------------------------------

{$WARN SYMBOL_PLATFORM OFF}

function __CryptAcquireContext(ProviderType: Integer): HCRYPTPROV;
begin
  if (not CryptAcquireContext(Result, nil, nil, ProviderType, 0)) then
  begin
    if HRESULT(GetLastError) = NTE_BAD_KEYSET then
      Win32Check(CryptAcquireContext(Result, nil, nil, ProviderType, CRYPT_NEWKEYSET))
    else
      RaiseLastOSError;
  end;
end;

function __AES128_DeriveKeyFromPassword(m_hProv: HCRYPTPROV; Password: string): HCRYPTKEY;
var
  hHash: HCRYPTHASH;
  Mode: DWORD;
begin
  Win32Check(CryptCreateHash(m_hProv, CALG_SHA_256, 0, 0, hHash));
  try
    Win32Check(CryptHashData(hHash, PChar(Password), Length(Password) * SizeOf(Char), 0));
    Win32Check(CryptDeriveKey(m_hProv, CALG_AES_128, hHash, 0, Result));
    // Wine uses a different default mode of CRYPT_MODE_EBC
    Mode := CRYPT_MODE_CBC;
    Win32Check(CryptSetKeyParam(Result, KP_MODE, Pointer(@Mode), 0));
  finally
    CryptDestroyHash(hHash);
  end;
end;

function AES128_Encrypt(Value, Password: string): string;
var
  hCProv: HCRYPTPROV;
  hKey: HCRYPTKEY;
  lul_datalen: Integer;
  lul_buflen: Integer;
  Buffer: TBytes;
begin
  Assert(Password <> '');
  if (Value = '') then
    Result := ''
  else begin
    hCProv := __CryptAcquireContext(PROV_RSA_AES);
    try
      hKey := __AES128_DeriveKeyFromPassword(hCProv, Password);
      try
        // allocate buffer space
        lul_datalen := Length(Value) * SizeOf(Char);
        Buffer := TEncoding.Unicode.GetBytes(Value + '        ');
        lul_buflen := Length(Buffer);
        // encrypt to buffer
        Win32Check(CryptEncrypt(hKey, 0, True, 0, @Buffer[0], lul_datalen, lul_buflen));
        SetLength(Buffer, lul_datalen);
        // base 64 result
        Result := Base64_Encode(Buffer);
      finally
        CryptDestroyKey(hKey);
      end;
    finally
      CryptReleaseContext(hCProv, 0);
    end;
  end;
end;

function AES128_Decrypt(Value, Password: string): string;
var
  hCProv: HCRYPTPROV;
  hKey: HCRYPTKEY;
  lul_datalen: Integer;
  Buffer: TBytes;
begin
  Assert(Password <> '');
  if Value = '' then
    Result := ''
  else begin
    hCProv := __CryptAcquireContext(PROV_RSA_AES);
    try
      hKey := __AES128_DeriveKeyFromPassword(hCProv, Password);
      try
        // decode base64
        Buffer := Base64_Decode(Value);
        // allocate buffer space
        lul_datalen := Length(Buffer);
        // decrypt buffer to to string
        Win32Check(CryptDecrypt(hKey, 0, True, 0, @Buffer[0], lul_datalen));
        Result := TEncoding.Unicode.GetString(Buffer, 0, lul_datalen);
      finally
        CryptDestroyKey(hKey);
      end;
    finally
      CryptReleaseContext(hCProv, 0);
    end;
  end;
end;

function GenerateActivation(Mb, Nm, Em: string; De: integer): string;
var
  Reg: string;
begin
  Reg := 'APP=' + AppName + ';DE=' + IntToStr(De) + ';MB=' + Mb + ';MD=' + DateToStr(now + 14) + ';PA=TRUE;NM=' + Nm + ';EM=' + Em + ';';
  Reg := StrEncrypt(AES128_Encrypt(Reg, masterkey), 322);
  result := Reg;
end;

procedure TKeyGenerator.ChkDeClick(Sender: TObject);
begin
  SerialChange(Self.Serial);
end;

procedure TKeyGenerator.ChkEdClick(Sender: TObject);
begin
  SerialChange(Self.Serial);
end;

procedure TKeyGenerator.CopyClipboardClick(Sender: TObject);
begin
  ClipBoard.AsText := Output.Text;
end;

procedure TKeyGenerator.DeChange(Sender: TObject);
begin
  SerialChange(Self.Serial);
end;

procedure TKeyGenerator.ExportButtonClick(Sender: TObject);
var
  SaveDialog : TSaveDialog;
begin
  SaveDialog := TSaveDialog.Create(self);
  SaveDialog.Title := 'Save currency rate preset';
  SaveDialog.Filter := 'ZTool key file|*.key';
  SaveDialog.DefaultExt := 'key';
  SaveDialog.FilterIndex := 1;
  SaveDialog.Options := SaveDialog.Options + [ofOverwritePrompt];
  if SaveDialog.Execute then
    WriteFile(SaveDialog.FileName, Output.Text)
  else
    MessageDlg('File was not saved.', mtInformation, [mbOk], 0);
end;

procedure TKeyGenerator.OutputChange(Sender: TObject);
begin
  Decrypted.Text := Output.Text;
  Decrypted.Text := StrDecrypt(Decrypted.Text,322);
  Decrypted.Text := AES128_Decrypt(Decrypted.Text, masterkey);
end;

procedure TKeyGenerator.SerialChange(Sender: TObject);
begin
  Motherboard.Caption := StrCut(StrDecrypt(Serial.Text,398), 'ID=',';');
  Name.Caption := StrCut(StrDecrypt(Serial.Text,398), 'NM=',';');
  Email.Caption := StrCut(StrDecrypt(Serial.Text,398), 'EM=',';');
  Output.Text := GenerateActivation(Motherboard.Caption, Name.Caption, Email.Caption, StrToInt(De.Text));
  ExportButton.Enabled := Output.Text <> '';
end;

end.
