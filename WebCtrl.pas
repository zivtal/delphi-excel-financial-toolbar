unit WebCtrl;

interface

uses
  Wininet, XMLDoc, XMLIntf;

function ActiveConnection(): boolean;
function HttpGet(const URL: string): widestring; stdcall;
function IsValidEmail(const EmailAddress: string): boolean; stdcall;
function UnixDate(DateIn: double): Double; stdcall;

implementation

uses
  SysUtils,
  Dialogs,
  RegularExpressions;

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
end;

function IsValidEmail(const EmailAddress: string): boolean;
const
  EMAIL_REGEX = '^((?>[a-zA-Z\d!#$%&''*+\-/=?^_`{|}~]+\x20*|"((?=[\x01-\x7f])'
             +'[^"\\]|\\[\x01-\x7f])*"\x20*)*(?<angle><))?((?!\.)'
             +'(?>\.?[a-zA-Z\d!#$%&''*+\-/=?^_`{|}~]+)+|"((?=[\x01-\x7f])'
             +'[^"\\]|\\[\x01-\x7f])*")@(((?!-)[a-zA-Z\d\-]+(?<!-)\.)+[a-zA-Z]'
             +'{2,}|\[(((?(?<!\[)\.)(25[0-5]|2[0-4]\d|[01]?\d?\d))'
             +'{4}|[a-zA-Z\d\-]*[a-zA-Z\d]:((?=[\x01-\x7f])[^\\\[\]]|\\'
             +'[\x01-\x7f])+)\])(?(angle)>)$';
begin
  result := TRegEx.IsMatch(EmailAddress, EMAIL_REGEX);
end;

function UnixDate(DateIn: double): double; stdcall;
begin
  result := (DateIn - 25569) * 86400
end;

end.
