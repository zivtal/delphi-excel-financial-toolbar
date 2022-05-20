function LoadCache(Address, Name: string; out Value: string): boolean;
var
  RegKey: TRegistry;
  Workbook, Sheet, Cell: string;
begin
  if not XlDecodeRange(Address, Workbook, Sheet, Cell) then
    exit(false);
  RegKey := TRegistry.Create;
  try
    RegKey.OpenKeyReadOnly(Cache + '\' + Workbook + '\' + Sheet + '\' + Cell + '\');
    try
      Value := RegKey.ReadString(Name);
      result := (Value <> '');
    except
      result := false;
    end;
  finally
    RegKey.Free;
  end;
end;

function SaveCache(Address, Name, Value: string; Overwrite: boolean = true): boolean;
var
  RegKey: TRegistry;
  Workbook, Sheet, Cell: string;
begin
  if not XlDecodeRange(Address, Workbook, Sheet, Cell) then
    exit(false);
  RegKey := TRegistry.Create;
  try
    RegKey.OpenKey(Cache + '\' + Workbook + '\' + Sheet + '\' + Cell + '\', True);
    if (RegKey.ReadString(Name) <> '') and not Overwrite then
      begin
        RegKey.Free;
        exit(false);
      end
    else
      try
        result := true;
        RegKey.WriteString(Name, Value);
      except
        result := false;
      end;
  finally
    RegKey.Free;
  end;
end;






if not LoadCache(XlApp.ActiveCell.Address[False, False, xlA1, True, False], Formula, result) then
	try
	finally
		SaveCache(XlApp.ActiveCell.Address[False, False, xlA1, True, False], Formula, result);
	end;

