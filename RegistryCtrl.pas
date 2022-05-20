unit RegistryCtrl;

interface
function RegRead(Path, Key: string; Default: string = ''): string;
procedure RegWrite(Path, Key, Value: string; Overwrite: boolean = true; Count: boolean = false); stdcall;
procedure RegDelete(Path: string; Key: string = '');

implementation

uses
  System.SysUtils,
  Registry;

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

end.
