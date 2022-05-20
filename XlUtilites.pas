unit XlUtilites;

interface

uses Windows, Classes, Dialogs, UITypes, Variants, StrUtils, SysUtils, VarUtils, DateUtils, Math, Registry, RegularExpressions, Excel2010;

type
  T1DStringArray = array of string;
  T2DStringArray = array of array of string;

// general functions
procedure ExcelForeground;
function ExtractFormulas(Formula: string): T1DStringArray; stdcall;
function ExtractCells(Formula: string): T1DStringArray; stdcall;
function ExtractVars(Formula: string): T1DStringArray; stdcall;
// excel features
function XlTypeOf(Input: OleVariant): integer; stdcall;
function XlIsEmpty(const Value: OleVariant): boolean; stdcall;
function XlFind(InRange, Value, LookIn: OleVariant; MatchCase, SearchFormat: boolean; out Cell: string): boolean; stdcall;
function XlFindAll(InRange, Value, LookIn: OleVariant; MatchCase, SearchFormat: boolean): string; stdcall;
function XlDecodeRange(Input: string; out Workbook, Sheet, Cells: string; IsRange: boolean = False; UsedRange: boolean = True): boolean; stdcall;
function XlEncodeRange(Workbook, Sheet, Cells: string; out Range: string): boolean; stdcall;
function XlGetName(Name: string; out Value: string): boolean; stdcall;
function XlIsFormula(Value: string; out Formula: string): boolean; stdcall;
function XlIsRange(Value: string; out Range: string): boolean; stdcall;
function XlSelection(Seperate: boolean = false): T1DStringArray; stdcall;
function XlUnion(Cells: string): string; stdcall;
function XlNonunion(Input: string; EmptyCell: boolean = true): string; stdcall;
function XlTryOleToStr(Input: OleVariant; out Output: string): boolean; stdcall;
function XlOleToArray(Input: string): T1DStringArray; stdcall;
function XlCriteria(Address, Value: string; Evaluate: boolean = false): boolean; stdcall;
function XlEvaluate(const Input: string): string;
function XlWorkbookOpen(Name: string): boolean;
procedure XlActivate(Range: string; Warning: boolean = True);
// excel formulas emulation
function FxIFERROR(Formula: string; Error: string = '#N/A'): string;
function FxIF(Formula: string): string;
function FxMIN(Formula: string; Evaluate: boolean): string;
function FxMAX(Formula: string; Evaluate: boolean): string;
function FxROW(Formula: string): string;
function FxCOLUMN(Formula: string): string;
function FxINDIRECT(Formula: string): string;
function FxINDEX(Formula: string; Default: string = '#N/A'): string; stdcall;
function FxMATCH(Formula: string): string; stdcall;
function FxSUMIF(Formula: string; Union: boolean = True): string;
function FxSUMIFS(Formula: string; Union: boolean = True): string;
function FxAVERAGEIF(Formula: string; Union: boolean = True): string;
function FxAVERAGEIFS(Formula: string; Union: boolean = True): string;
function FxVLOOKUP(Formula: string; Default: string = '#N/A'): string;
function FxHLOOKUP(Formula: string; Default: string = '#N/A'): string;

const
  Cache: string = '\SOFTWARE\Ziv Tal\ZTools\Cache';

var
  XlWorkbook: ExcelWorkbook;
  XlSheet: ExcelWorksheet;

implementation

uses
  XlApplication, XlProgress, RegistryCtrl, StringCtrl;

procedure ExcelForeground;
var
  Handle: HWND;
begin
  Handle := FindWindow('XLMAIN', nil);
  if Handle <> 0 then
    begin
      if IsIconic(Handle) then
        ShowWindow(Handle, SW_RESTORE); // in case Excel is minimized
      SetForegroundWindow(Handle);
    end;
end;

function IsDate(Value: string): boolean;
begin
  try
    exit(TRegEx.Create('^(((0[1-9]|[12][0-9]|3[01])[- /.](0[13578]|1[02])|(0[1-9]|[12][0-9]|30)[- /.](0[469]|11)|(0[1-9]|1\d|2[0-8])[- /.]02)[- /.]\d{4}|29[- /.]02[- /.](\d{2}(0[48]|[2468][048]|[13579][26])|([02468][048]|[1359][26])00))$').Match(Value).Success);
  except
    exit(false);
  end;
end;

function IsCell(Value: string; Range: boolean = False): boolean;
const
  Template: array of string = [
    '^(?:''[^'']+''!|(\w|)+!)?([A-Z\$]+\d{1,9}$)',
    '^(?:''[^'']+''!|(\w|)+!)?([A-Z\$]+(?:\d)+(?:\:[A-Z\$]+\d+)?|(?:\$?[A-Z\$]\:?(?:\$)?[A-Z\$]))'
  ];
var
  RegEx: TRegEx;
begin
  if Range then
    RegEx := TRegex.Create(Template[1])
  else
    RegEx := TRegex.Create(Template[0]);
  try
    exit(RegEx.Match(Value).Success);
  except
    exit(false);
  end;
end;

function ExtractMath(Input: string; out Symbol: string): boolean;
var
  RegEx: TRegEx;
begin
  RegEx := TRegEx.Create('(^<>|^<=|^>=|^=|^<|^>)');
  result := RegEx.Match(Input).Success;
  if result then
    Symbol := RegEx.Match(Input).Value
  else
    Symbol := '=';
end;

function ExtractFormulas(Formula: string): T1DStringArray; stdcall;
var
  RegExt: TRegEx;
  Match: TMatch;
  Matches: TMatchCollection;
begin
  RegExt := TRegex.Create('(?:[a-zA-Z]*\((?:[^)(]+|(?R))*+\))');
  Matches := RegExt.Matches(Formula);
  for Match in Matches do
    begin
      SetLength(result, Length(result)+1);
      result[High(result)] := Match.Value;
    end;
end;

function ExtractCells(Formula: string): T1DStringArray; stdcall;
const
  Template: string = '(?:''[^'']+''!|(\w|)+!)?([A-Z\$]+(?:\d)+(?:\:[A-Z\$]+\d+)?|(?:\$?[A-Z\$]\:?(?:\$)?[A-Z\$]))';
var
  RegExt: TRegEx;
  Match: TMatch;
  Matches: TMatchCollection;
begin
  RegExt := TRegex.Create(Template);
  Matches := RegExt.Matches(Formula);
  for Match in Matches do
    begin
      SetLength(result, Length(result)+1);
      result[High(result)] := Match.Value;
    end;
end;

function ExtractVars(Formula: string): T1DStringArray; stdcall;
var
  Index, LastComma, Brackets, ExcelArray: integer;
  Quotation: boolean;
begin
  Brackets := 0;
  ExcelArray := 0;
  Quotation := False;
  LastComma := 1;
  Formula := Copy(Formula, (Pos('(', Formula)+1), Length(Formula) - (Pos('(', Formula)+1));
  for Index := 1 to Length(Formula) do
    case AnsiIndexStr(Copy(Formula, Index, 1), [',','(',')','{','}','"']) of
      0:
        if (Brackets = 0) and (ExcelArray = 0) and not Quotation then
          begin
            SetLength(result, Length(result)+1);
            result[High(result)] := Copy(Formula, LastComma, Index - LastComma);
            LastComma := Index + 1;
          end;
      1: Brackets := Brackets + 1;
      2: Brackets := Brackets - 1;
      3: ExcelArray := ExcelArray + 1;
      4: ExcelArray := ExcelArray - 1;
      5: Quotation := not Quotation;
    end;
  SetLength(result, Length(result)+1);
  result[High(result)] := Copy(Formula, LastComma, Index - LastComma);
end;

function RemoveFromArray(Input: T1DStringArray; Index: integer): T1DStringArray;
var
  I: Integer;
begin
  if Length(Input) = 1 then
    exit(nil);
  for I := Index to (High(Input) -1) do
    Input[I] := Input[I+1];
  SetLength(Input, Length(Input)-1);
  result := Input;
end;

function IsContian(Contain, Text: string; out Position: integer): boolean;
var
  Index: integer;
  Quotation: boolean;
begin
  result := false;
  Quotation := False;
  for Index := 1 to (Length(Text)) do
  begin
    case IndexStr(Text[Index], [Contain,'"']) of
      0:
        if not Quotation then
          begin
            Position := Index;
            exit(true);
          end;
      1: Quotation := not Quotation;
    end;
  end;
end;

// excel features

function XlTypeOf(Input: OleVariant): integer;
var
  Code: string;
  Numeric: Integer;
begin
  Numeric := VarType(Input) and VarTypeMask;
  case Numeric of
    varEmpty     : Code := 'varEmpty';      // 0
    varNull      : Code := 'varNull';       // 1
    varSmallInt  : Code := 'varSmallInt';   // 2
    varInteger   : Code := 'varInteger';    // 3
    varSingle    : Code := 'varSingle';     // 4
    varDouble    : Code := 'varDouble';     // 5
    varCurrency  : Code := 'varCurrency';   // 6
    varDate      : Code := 'varDate';       // 7
    varOleStr    : Code := 'varOleStr';     // 8
    varDispatch  : Code := 'varDispatch';   // 9
    varError     : Code := 'varError';      // 10
    varBoolean   : Code := 'varBoolean';    // 11
    varVariant   : Code := 'varVariant';    // 12
    varUnknown   : Code := 'varUnknown';    // 13
    varByte      : Code := 'varByte';       // 14
    varWord      : Code := 'varWord';       // 15
    varLongWord  : Code := 'varLongWord';   // 16
    varInt64     : Code := 'varInt64';      // 17
    varStrArg    : Code := 'varStrArg';     // 18
    varString    : Code := 'varString';     // 19
    varAny       : Code := 'varAny';        // 20
    varTypeMask  : Code := 'varTypeMask';   // 21
  end;
  exit(Numeric);
end;

function XlIsEmpty(const Value: OleVariant): boolean;
begin
  result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual);
  if not result and VarIsStr(Value) then
    result := Value = '';
end;

function XlFind(InRange, Value, LookIn: OleVariant; MatchCase, SearchFormat: boolean; out Cell: string): boolean; stdcall;
var
  XlFound: OleVariant;
begin
  try
    XlFound := InRange.Find('', EmptyParam, xlValues, xlPart, xlByColumns, xlNext, EmptyParam, MatchCase, SearchFormat);
  finally
    try
      Cell := XlFound.Address[False, False, xlA1, True, False];
      result := (Cell <> '');
    except
      result := false;
    end;
    XlApp.FindFormat.Clear;
  end;
end;

function XlFindAll(InRange, Value, LookIn: OleVariant; MatchCase, SearchFormat: boolean): string; stdcall;
var
  Range, Output: string;
  LastCell: string;
  Row, Col, LastCol, LastRow: integer;
  XlRange, XlFound, XlColumn: OleVariant;
begin
  XlRange := InRange;
  LastCell := XlRange.Address[False, False, xlA1, False, False];
  LastCell := Copy(LastCell, AnsiPos(':', LastCell) + 1, Length(LastCell)-AnsiPos(':', LastCell));
  LastRow := XlApp.Range[LastCell, EmptyParam].Row;
  LastCol := XlApp.Range[LastCell, EmptyParam].Column;
  for Col := 1 to XlRange.Columns.Count do
    try
      XlColumn := XlRange.Columns[Col];
      LastCell := XlColumn.Address[False, False, xlA1, False, False];
      LastCell := Copy(LastCell, AnsiPos(':', LastCell) + 1, Length(LastCell)-AnsiPos(':', LastCell));
      XlFound := XlColumn.Find(Value, EmptyParam, xlValues, xlPart, xlByRows, xlNext, EmptyParam, MatchCase, SearchFormat);
      while not (VarIsClear(XlFound) or VarIsEmpty(XlFound)) do
        begin
          Output := Output + ',' + XlFound.Address[False, False, xlA1, False, False];
          if XlFound.Row < LastRow then
            try
              XlColumn := XlApp.Range[XlFound.Offset[1,0].Address[False, False, xlA1, False, False], LastCell];
              XlFound := XlColumn.Find(Value, EmptyParam, xlValues, xlPart, xlByRows, xlNext, EmptyParam, MatchCase, SearchFormat);
            except
              VarClear(XlFound);
            end;
        end;
    finally
    end;
  XlApp.FindFormat.Clear;
  result := XlUnion(Output);
  OutputDebugString(PChar(result));
end;

function XlGetName(Name: string; out Value: string): boolean;
begin
  Value := Name;
  try
    XlWorkbook := XlApp.ActiveWorkbook;
    Value := XlWorkbook.Names.Item(Value, EmptyParam, EmptyParam).Value;
    Value := Copy(Value, 2, Length(Value) - 1);
  except
    exit(false);
  end;
  result := Value<>Name;
  OutputDebugString(PChar(Name + ' -> ' + Value));
end;

function XlIsFormula(Value: string; out Formula: string): boolean;
var
  RegEx: TRegEx;
begin
  RegEx := TRegex.Create('(?:[a-zA-Z]*\((?:[^)(]+|(?R))*+\))');
  result := RegEx.Match(Value).Success;
  Formula := Value;
end;

function XlIsRange(Value: string; out Range: string): boolean;
var
  RegExt: TRegEx;
begin
  Range := Value;
  try
    RegExt := TRegex.Create('(((?:\$)?[A-Z]{1,9}(?:\$)?[0-9]{0,9}|(?:\$)?[A-Z]{0,9}(?:\$)?[0-9]{1,9})+\:+((?:\$)?[A-Z]{1,9}(?:\$)?[0-9]{0,9}|(?:\$)?[A-Z]{0,9}(?:\$)?[0-9]{1,9}))');
    result := RegExt.Match(Range).Success;
    if not result then
      try
        XlGetName(Value, Range);
        exit(RegExt.Match(Range).Success);
      except
        exit(false);
      end;
  except
    exit(false);
  end;
end;

function XlEndOfRange(Input: string; Limit: boolean = True): string;
  function ExtractCol(Cell: string): string;
  begin
    result := TRegEx.Create('[A-Z]{1,4}').Match(Cell).Value;
  end;
  function ExtractRow(Cell: string): string;
  begin
    result := TRegEx.Create('[0-9]{1,7}').Match(Cell).Value;
  end;
var
  RegEx: TRegEx;
  Workbook, Sheet, Range: string;
  Cell: array[0..2] of string;
  XlRange, XlLastCell: OleVariant;
begin
  try
    if not XlIsRange(Input, Range) then
      exit(Range);
    XlWorkbook := XlApp.ActiveWorkbook;
    if ContainsText(Range, '!') then
      begin
        result := SplitString(Range, '!')[0] + '!';
        Sheet := SplitString(Range, '!')[0];
        if ContainsText(Sheet, '[') and ContainsText(Sheet, ']') then
          begin
            Workbook := Copy(Sheet, AnsiPos('[', Sheet) + 1, AnsiPos(']', Sheet) - AnsiPos('[', Sheet) - 1);
            Sheet := StringReplace(Sheet, '[' + Workbook + ']','', [rfReplaceAll, rfIgnoreCase]);
            XlWorkbook := XlApp.Workbooks[Workbook];
          end;
        if (Sheet[1] = '''') and (Sheet[Length(Sheet)] = '''') then
          Sheet := Copy(Sheet,2,Length(Sheet)-2);
        XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
        Cell[0] := SplitString(Range, '!')[1];
      end
    else
      begin
        Cell[0] := Range;
        XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
      end;
    XlRange := XlSheet.UsedRange[0].Address[False, False, xlA1, False, False];
    XlLastCell := Copy(XlRange, AnsiPos(':', XlRange) + 1, Length(XlRange)-AnsiPos(':', XlRange));
    Cell[1] := SplitString(Cell[0], ':')[0];
    Cell[2] := SplitString(Cell[0], ':')[1];
    Cell[1] := StringReplace(Cell[1], '$', '', [rfReplaceAll, rfIgnoreCase]);
    Cell[2] := StringReplace(Cell[2], '$', '', [rfReplaceAll, rfIgnoreCase]);
    if not TRegEx.Create('(\$?[A-Z]{1,9})').Match(Cell[1]).Success and (StrToFloat(Cell[2]) > XlSheet.UsedRange[0].Rows.Count) then
      if XlEncodeRange(Workbook, Sheet, 'A1:' + XlLastCell, result) then
        exit;
    if not TRegEx.Create('(\$?[A-Z]{1,9})').Match(Cell[1]).Success then
      Cell[1] := 'A' + Cell[1];
    if not TRegEx.Create('(\$?[A-Z]{1,9})').Match(Cell[2]).Success then
      Cell[2] := TRegEx.Create('(\$?[A-Z]{1,9})').Match(XlLastCell).Value + Cell[2];
    if not TRegEx.Create('(\$?\d{1,9})').Match(Cell[1]).Success then
      Cell[1] := Cell[1] + '1';
    if not TRegEx.Create('(\$?\d{1,9})').Match(Cell[2]).Success then
      Cell[2] := Cell[2] + TRegEx.Create('(\$?\d{1,9})').Match(XlLastCell).Value;
    if Limit then
      begin
        if XlApp.Range[Cell[2], EmptyParam].Row > XlApp.Range[XlLastCell, EmptyParam].Row then
          Cell[2] := ExtractCol(Cell[2]) + ExtractRow(XlLastCell);
        if XlApp.Range[Cell[2], EmptyParam].Column > XlApp.Range[XlLastCell, EmptyParam].Column then
          Cell[2] := ExtractCol(XlLastCell) + ExtractRow(Cell[2]);
      end;
    if result = Cell[0] then
      result := Cell[1] + ':' + Cell[2]
    else
      result := result + Cell[1] + ':' + Cell[2];
  except
    exit(Range);
  end;
end;

function XlDecodeRange(Input: string; out Workbook, Sheet, Cells: string; IsRange: boolean = False; UsedRange: boolean = True): boolean;
var
  Range: string;
begin
  XlGetName(Input, Range);
  if IsRange and not TRegex.Create('[A-Z]{0,9}[0-9]{0,9}\:[A-Z]{0,9}[0-9]{0,9}').Match(Range).Success then
    exit(false);
  try
    Range := XlApp.Range[Range, EmptyParam].Address[False, False, xlA1, True, False];
    Range := XlEndOfRange(Range, UsedRange);
    if ContainsText(Range, '!') then
      begin
        Sheet := SplitString(Range, '!')[0];
        Sheet := StringReplace(Sheet, '''''','''', [rfReplaceAll]);
        if (Sheet[1] = '''') and (Sheet[Length(Sheet)] = '''') then
          Sheet := Copy(Sheet,2,Length(Sheet)-2);
        if ContainsText(Sheet, '[') and ContainsText(Sheet, ']') then
          begin
            Workbook := Copy(Sheet, Pos('[', Sheet) + 1, Pos(']', Sheet) - Pos('[', Sheet) - 1);
            Sheet := StringReplace(Sheet, '[' + Workbook + ']','', [rfReplaceAll]);
          end;
        Cells := SplitString(Range, '!')[1];
        exit(true);
      end
    else
      exit(false);
  except
    exit(false);
  end;
end;

function XlEncodeRange(Workbook, Sheet, Cells: string; out Range: string): boolean;
begin
  try
    XlWorkbook := XlApp.ActiveWorkbook;
    XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
    if (Workbook = '') then
      Workbook := XlWorkbook.Name;
    if (Sheet = '') then
      Sheet := XlSheet.Name;
    Range := ('''[' + Workbook + ']' + Sheet + '''!' + Cells);
    exit(true);
  except
    exit(false);
  end;
end;

function XlSelection(Seperate: boolean = false): T1DStringArray; stdcall;
var
  Selection: OleVariant;
  Address: string;
  Row, Col: integer;
begin
  try
    Selection := XlApp.Selection[0];
    for Address in ExtractCells(Selection.Address[False, False, xlA1, False, False]) do
      if ContainsText(Address, ':') and (Seperate = True) then
        begin
          for Row := 0 to XlApp.Range[Address, EmptyParam].Rows.Count-1 do
            for Col := 0 to XlApp.Range[Address, EmptyParam].Columns.Count-1 do
              begin
                SetLength(result, Length(result)+1);
                result[High(result)] := XlApp.Range[SplitString(Address, ':')[0], EmptyParam].Offset[Row, Col].Address[False, False, xlA1, False, False];
              end;
        end
      else
        begin
          SetLength(result, Length(result)+1);
          result[High(result)] := Address;
        end;
  except
  end;
end;

function XlSplitRange(Input: string; out Output: T1DStringArray; EmptyCell: boolean = True): boolean;
var
  Range, Workbook, Sheet, Cells: string;
  Cell: array[0..2] of string;
  Row, Col, ERow, ECol: integer;
begin
  result := true;
  for Cells in ExtractCells(Input) do
    try
      if XlIsRange(Cells, Range) then
        begin
          Range := XlEndOfRange(Range);
          XlDecodeRange(Range, Workbook, Sheet, Cell[0]);
          Cell[1] := SplitString(Cell[0],':')[0];
          Cell[2] := SplitString(Cell[0],':')[1];
          ECol := XlApp.Range[Cell[2],EmptyParam].Column - XlApp.Range[Cell[1],EmptyParam].Column;
          ERow := XlApp.Range[Cell[2],EmptyParam].Row - XlApp.Range[Cell[1],EmptyParam].Row;
          for Row := 0 to ERow do
            for Col := 0 to ECol do
              begin
                if XlEncodeRange(Workbook, Sheet, XlApp.Range[Cell[1],EmptyParam].Offset[Row,Col].Address[False, False, xlA1, False, False], Range) then
                  if EmptyCell or (string(XlApp.Range[Range, EmptyCell].Value2) <> '') then
                    begin
                      SetLength(Output, Length(Output)+1);
                      Output[High(Output)] := Range;
                    end;
              end;
        end
      else
        if XlDecodeRange(Cells, Workbook, Sheet, Cell[0]) and XlEncodeRange(Workbook, Sheet, Cell[0], Range) and (EmptyCell or (string(XlApp.Range[Range, EmptyParam].Value2) <> '')) then
          if EmptyCell or (string(XlApp.Range[Range, EmptyCell].Value2) <> '') then
            begin
              SetLength(Output, Length(Output)+1);
              Output[High(Output)] := Range;
            end;
    except
      exit(false);
    end;
  if Length(Output) = 0 then
    exit(false);;
end;

function XlRangeToArray(Input: string; out Output: T1DStringArray; Sort: integer = 1; EmptyCell: boolean = False): boolean;
var
  Numeric: array[1..2] of double;
  Index, EndOfArray: integer;
  Temp: T1DStringArray;
begin
  Output := ExtractCells(Xlnonunion(Input, EmptyCell));
  case Sort of
     1: exit(Length(Output) > 0);
    -1:
      try
        EndOfArray := High(Output);
        for Index := Low(Output) to EndOfArray do
          Temp[EndOfArray - Index] := Output[Index];
      finally
        Output := Temp;
      end;
     2:
      try
        EndOfArray := High(Output) - 1;
        for Index := Low(Output) to EndOfArray do
          if TryStrToFloat(XlEvaluate(Output[Index]), Numeric[1]) and TryStrToFloat(XlEvaluate(Output[Index+1]), Numeric[2]) and (Numeric[2] > Numeric[1]) then
            begin
              Output[Index] := FloatToStr(Numeric[2]);
              Output[Index+1] := FloatToStr(Numeric[1]);
            end;
      finally end;
    -2:
      try
        EndOfArray := High(Output) - 1;
        for Index := Low(Output) to EndOfArray do
          if TryStrToFloat(XlEvaluate(Output[Index]), Numeric[1]) and TryStrToFloat(XlEvaluate(Output[Index+1]), Numeric[2]) and (Numeric[1] > Numeric[2]) then
            begin
              Output[Index] := FloatToStr(Numeric[2]);
              Output[Index+1] := FloatToStr(Numeric[1]);
            end;
      finally end;
  end;
  result := (Length(Output) > 0);
end;

function XlRangeTo2DArray(Input: string; out Output: T2DStringArray; EmptyCell: boolean = True): boolean;
var
  Range, Workbook, Sheet, Cells: string;
  Cell: array[0..2] of string;
  Row, Col, ERow, ECol: integer;
begin
  result := true;
  for Cells in ExtractCells(Input) do
    try
      if XlIsRange(Cells, Range) then
        begin
          Range := XlEndOfRange(Range);
          XlDecodeRange(Range, Workbook, Sheet, Cell[0]);
          Cell[1] := SplitString(Cell[0],':')[0];
          Cell[2] := SplitString(Cell[0],':')[1];
          ECol := XlApp.Range[Cell[2],EmptyParam].Column - XlApp.Range[Cell[1],EmptyParam].Column;
          ERow := XlApp.Range[Cell[2],EmptyParam].Row - XlApp.Range[Cell[1],EmptyParam].Row;
          for Row := 0 to ERow do
            begin
              SetLength(Output, Row+1);
              for Col := 0 to ECol do
                begin
                  if XlEncodeRange(Workbook, Sheet, XlApp.Range[Cell[1],EmptyParam].Offset[Row,Col].Address[False, False, xlA1, False, False], Range) then
                    if EmptyCell or (string(XlApp.Range[Range, EmptyCell].Value2) <> '') then
                      begin
                        SetLength(Output[Row], Col+1);
                        Output[Row,Col] := Range;
                      end;
                end;
            end;
        end
      else
        exit(false);
    except
      exit(false);
    end;
  if Length(Output) = 0 then
    exit(false);;
end;

function XlArray(Input: string; out Output: T2DStringArray): boolean;
var
  Quotation: boolean;
  Index, Comma, Row, Col: integer;
  Add: array[1..2] of string;
  Position: array[1..2] of integer;
begin
  result := true;
  if XlRangeTo2DArray(Input, Output) then
    exit(true);
  Row := 0;
  Col := 0;
  Comma := 1;
  Quotation := False;
  SetLength(Output, Row+1);
  if IsContian('{', Input, Position[1]) and IsContian('}', Input, Position[2]) then
    begin
      Add[1] := Copy(Input, 0, Position[1] -1);
      Add[2] := Copy(Input, Position[2] +1, Length(Input));
      Input := Copy(Input, Position[1] +1, Position[2] - Position[1] -1);
      for Index := 1 to (Length(Input)) do
        case IndexStr(Input[Index], [',',';','"']) of
          0,1:
            if not Quotation then
              begin
                SetLength(Output[Row], Col+1);
                Output[Row, Col] := Add[1] + Copy(Input, Comma, Index - Comma) + Add[2];
                Comma := Index + 1;
                if (Input[Index] = ';') then
                  begin
                    Col := 0;
                    Row := Row +1;
                    SetLength(Output, Row+1);
                  end
                else
                  Col := Col +1;
              end;
          2: Quotation := not Quotation;
        end;
      begin
        SetLength(Output[Row], Col+1);
        Output[Row, Col] := Add[1] + Copy(Input, Comma, Length(Input)) + Add[2];
      end;
    end
  else
    begin
      SetLength(Output[Row], Col+1);
      Output[Row, Col] := Input;
    end;
//  for Row := Low(Output) to High(Output) do
//    for Col := Low(Output[Row]) to High(Output[Row]) do
//      showmessage(Output[Row, Col]);
end;

function XlUnion(Cells: string): string;
  function Fix(Cells: string; Sheet: string = ''): string;
  var
    Cell, Temp: string;
  begin
    if (Sheet <> '') and (Sheet[1] <> '''') and (Sheet[Length(Sheet)] <> '''') then
      Sheet := '''' + Sheet + '''';
    for Cell in ExtractCells(Cells) do
      begin
        Temp := Cell;
        if ContainsText(Temp, '!') then
          begin
            if Sheet = '' then
              Sheet := SplitString(Temp, '!')[0];
            Temp := SplitString(Temp, '!')[1];
          end;
        Temp := Sheet + '!' + Temp;
        result := result + ',' + Temp;
      end;
    result := Copy(result, 2, Length(result));
  end;
var
  Cell: string;
  Create: boolean;
  Range: ExcelRange;
begin
  try
    Create := true;
    for Cell in ExtractCells(Cells) do
      if Cell <> '' then
        if Create then
          begin
            Range := XlApp.Range[Cell, EmptyParam];
            Create := false;
          end
        else
          Range := XlTApp.Union(Range, XlApp.Range[Cell, EmptyParam]);
    result := Range.Address[False, False, xlA1, True, False];
    result := Fix(result);
  except
    exit(Cells);
  end;
end;

function XlNonunion(Input: string; EmptyCell: boolean = true): string;
var
  SPos, EPos: integer;
  Range: T1DStringArray;
  Workbook, Sheet, Cell, Cells, Split: string;
begin
  result := Copy(Input, 0, Pos('(', Input));
  if IsContian('(', Input, SPos) and IsContian('(', Input, EPos) then
    Input := Copy(Input, SPos + 1, Length(Input) - EPos -1);
  for Cells in ExtractCells(Input) do
    if XlSplitRange(Cells, Range) then
      for Split in Range do
        result := result + Split + ','
    else
      if XlDecodeRange(Cells, Workbook, Sheet, Cell) and XlEncodeRange(Workbook, Sheet, Cell, Split) and (EmptyCell or (string(XlApp.Range[Split, EmptyParam].Value2) <> '')) then
        result := result + Split + ',';
  result := Copy(result, 0, Length(result) -1) + ')';
end;

function XlTryOleToStr(Input: OleVariant; out Output: string): boolean;
var
  Index, Row, Col: integer;
  Value: string;
  Range: T2DStringArray;
  Numeric: double;
begin
  case XlTypeOf(Input) of
    2,3,4,5:
      begin
        Output := string(Input);
        exit(true);
      end;
    7:
      if XlTryOleToStr(XlApp.Evaluate('=NUMBERVALUE(' + Input + ')',0), Output) then
        exit(true);
    8,9:
      if XlRangeTo2DArray(string(Input), Range) then
        begin
          Output := '{';
          for Row := Low(Range) to High(Range) do
            begin
              for Col := Low(Range[Row]) to High(Range[Row]) do
                if XlTryOleToStr(XlApp.Evaluate('=' + Range[Row, Col], 0), Value) then
                  begin
                    Output := Output + Value;
                    if Col < High(Range[Row]) then
                      Output := Output + ',';
                  end;
              if Row < High(Range) then
                Output := Output + ';';
//                Output := Output + ',';
            end;
            Output := Output + '}';
          exit(true);
        end
      else
        begin
          Output := string(Input);
          if not TryStrToFloat(Output, Numeric) then
            Output := '"' + Output + '"';
          exit(true);
        end;
    12:
      begin
        Output := '{';
        for Index := VarArrayLowBound(Input,1) to VarArrayHighBound(Input,1) do
          if XlTryOleToStr(Input[Index], Value) then
            Output := Output + Value + ',';
        Output := Copy(Output, 0, Length(Output)-1) + '}';
        exit(true);
      end;
    else
      exit(false);
  end;
end;

function XlOleToArray(Input: string): T1DStringArray;
var
  Index, Row, Col: integer;
  OleInput: OleVariant;
  Range, Workbook, Sheet, Cell, Value: string;
  RangeArray: T1DStringArray;
  Range2DArray: T2DStringArray;
begin
  if XlIsRange(Input, Range) then
    begin
      if not XlDecodeRange(Range, Workbook, Sheet, Cell) then
        exit;
      if XlSplitRange(Range, RangeArray) then
        for Cell in RangeArray do
          if XlTryOleToStr(XlApp.Evaluate(Cell,0), Value) then
            begin
              SetLength(result, Length(result)+1);
              result[High(result)] := Value;
            end;
    end
  else if XlArray(Input,Range2DArray) then
    for Row := 0 to High(Range2DArray) do
      for Col := 0 to High(Range2DArray[Row]) do
        begin
          SetLength(result, Length(result)+1);
          result[High(result)] := Range2DArray[Row,Col];
        end
  else
    begin
      OleInput := XlApp.Evaluate(Input,0);
      case XlTypeOf(OleInput) of
        12:
          for Index := VarArrayLowBound(OleInput,1) to VarArrayHighBound(OleInput,1) do
            if XlTryOleToStr(OleInput[Index], Value) then
              begin
                SetLength(result, Length(result)+1);
                result[High(result)] := Value;
              end;
        else
          if XlTryOleToStr(OleInput, Value) then
            begin
              SetLength(result, Length(result)+1);
              result[High(result)] := Value;
            end;
      end;
    end;
end;

function XlCriteria(Address, Value: string; Evaluate: boolean = false): boolean;
var
  Math, Criteria: string;
  VString, Item: string;
  VNumeric, CNumeric: double;
begin
  try
    result := false;
    for Item in XlOleToArray(Value) do
    begin
      Criteria := Item;
      if Evaluate or ContainsText(Criteria, '"') or IsCell(Criteria) then
        Criteria := XlApp.Evaluate('=' + Criteria,0);
      if IsDate(Criteria) then
        Criteria := XlApp.Evaluate('=NUMBERVALUE(' + Item + ')',0);
      VString := string(XlApp.Range[Address, EmptyParam].Value2);
      if ExtractMath(Criteria, Math) then
        Criteria := Copy(Criteria, Length(Math) + 1, Length(Criteria));
      if (Criteria <> '') and TryStrToFloat(Criteria,CNumeric) and TryStrToFloat(VString,VNumeric) then
        case AnsiIndexStr(Math, ['=','>','<','<>','>=','<=']) of
          0: if (VNumeric=CNumeric) then exit(true);
          1: if (VNumeric>CNumeric) then exit(true);
          2: if (VNumeric<CNumeric) then exit(true);
          3: if (VNumeric<>CNumeric) then exit(true);
          4: if (VNumeric>=CNumeric) then exit(true);
          5: if (VNumeric<=CNumeric) then exit(true);
        end
      else
        case AnsiIndexStr(Math, ['=','<>']) of
          0:
            begin
              if Length(Criteria) = 0 then
                if VString = '' then exit(true);
              if (Criteria[1] = '*') and (Criteria[Length(Criteria)] = '*') then
                if ContainsText(VString, Copy(Criteria,2,Length(Criteria)-2)) then exit(true);
              if (Criteria[1] = '*') and (Criteria[Length(Criteria)] <> '*') then
                if Copy(Criteria, 2, Length(Criteria) -1) = Copy(VString, Length(VString) - Length(Criteria) + 2, Length(Criteria)) then exit(true);
              if (Criteria[1] <> '*') and (Criteria[Length(Criteria)] = '*') then
                if Copy(Criteria,1,Length(Criteria)-1) = Copy(VString,1,Length(Criteria)-1) then exit(true);
              if (VString=Criteria) then exit(true);
            end;
          1:
            begin
              if Length(Criteria) = 0 then
                if VString <> '' then exit(true);
              if (Criteria[1] = '*') and (Criteria[Length(Criteria)] = '*') then
                if not ContainsText(VString, Copy(Criteria,2,Length(Criteria)-2)) then exit(true);
              if (Criteria[1] = '*') and (Criteria[Length(Criteria)] <> '*') then
                if not (Copy(Criteria, 2, Length(Criteria) -1) = Copy(VString, Length(VString) - Length(Criteria) + 2, Length(Criteria)))  then exit(true);
              if (Criteria[1] <> '*') and (Criteria[Length(Criteria)] = '*') then
                if not (Copy(Criteria,1,Length(Criteria)-1) = Copy(VString,1,Length(Criteria)-1)) then exit(true);
              if (VString<>Criteria) then exit(true);
            end;
        end;
    end;
  except
    exit(false);
  end;
end;

function XlEvaluate(const Input: string): string;
begin
  if not XlTryOleToStr(XlApp.Evaluate(Input,0), result) then
    result := Input
end;

function XlWorkbookOpen(Name: string): boolean;
var
  Index: integer;
begin
  result := false;
  try
    for Index := 1 to XlApp.Workbooks.Count do
      if (XlApp.Workbooks[Index].Name = Name) then
        exit(true);
  except
    exit(false);
  end;
end;

procedure XlActivate(Range: string; Warning: boolean = True);
var
  Workbook, Sheet, Cell: string;
begin
  if XlDecodeRange(Range, Workbook, Sheet, Cell) then
    try
      if not (Workbook = '') then
        XlApp.Workbooks[Workbook].Activate(0);
      if not (Sheet = '') then
        begin
          XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
          if XlSheet.Visible[0] = xlSheetVisible then
            XlSheet.Activate(0)
          else
            begin
              if Warning and (XlSheet.Visible[0] = xlSheetHidden) and (MessageDlg('The sheet "' + Sheet + '" is hidden, Would you like to unhide sheet?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
                begin
                  XlSheet.Visible[0] := xlSheetVisible;
                  XlSheet.Activate(0);
                end
              else
                exit;
            end;
        end
      else
        XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
      if not (Cell = '') then
        XlSheet.Range[Cell, EmptyParam].Activate;
    except
    end;
end;

// basic formula

function FxIFERROR(Formula: string; Error: string = '#N/A'): string;
var
  Index: integer;
  Vars: T1DStringArray;
begin
  Vars := ExtractVars(Formula);
  for Index := Low(Vars) to High(Vars) - 1 do
    if Vars[Index] <> Error then
      result := result + ',' + Vars[Index]
    else
      result := result + ',' + Vars[High(Vars)];
  result := Copy(result, 2, Length(result)-1);
end;

function FxIF(Formula: string): string;
var
  Vars: T1DStringArray;
begin
  try
    Vars := ExtractVars(Formula);
    if StrToBool(XlApp.Evaluate(Vars[0],0)) then
      exit(Vars[1])
    else
      exit(Vars[2]);
  except
    exit(Formula);
  end;
end;

function FxMIN(Formula: string; Evaluate: boolean): string;
var
  Index: integer;
  Vars: T1DStringArray;
begin
  if Evaluate then
    try
      if not XlTryOleToStr(XlApp.Evaluate(Formula,0), result) then
        result := Formula;
    except
      result := Formula;
    end
  else
    try
      Vars := ExtractVars(Formula);
      result := Vars[0];
      for Index := 1 to High(Vars) do
        try
          if XlApp.Evaluate('=' + result,0) > XlApp.Evaluate('=' + Vars[Index],0) then
            result := Vars[Index];
        except end;
    except
      result := Formula;
    end;
end;

function FxMAX(Formula: string; Evaluate: boolean): string;
var
  Index: integer;
  Vars: T1DStringArray;
begin
  if Evaluate then
    try
      if not XlTryOleToStr(XlApp.Evaluate(Formula,0), result) then
        result := Formula;
    except
      result := Formula;
    end
  else
    try
      Vars := ExtractVars(Formula);
      result := Vars[0];
      for Index := 1 to High(Vars) do
        try
          if XlApp.Evaluate('=' + result,0) < XlApp.Evaluate('=' + Vars[Index],0) then
            result := Vars[Index];
        except end;
    except
      result := Formula;
    end;
//  showmessage(result);
end;

function FxROW(Formula: string): string;
var
  Vars: T1DStringArray;
begin
  try
    Vars := ExtractVars(Formula);
    result := IntToStr(XlApp.Range[Vars[0], EmptyParam].Row);
  except
    result := Formula;
  end;
end;

function FxCOLUMN(Formula: string): string;
var
  Vars: T1DStringArray;
begin
  try
    Vars := ExtractVars(Formula);
    result := IntToStr(XlApp.Range[Vars[0], EmptyParam].Column);
  except
    result := Formula;
  end;
end;

function FxINDIRECT(Formula: string): string;
begin
  try
    Formula := Copy(Formula, Pos('(',Formula) + 1, Length(Formula) - Pos('(',Formula) - 1);
    result := XlApp.Evaluate('=' + Formula,0);
  except
    result := Formula;
  end;
  OutputDebugString(PChar(Formula + ' -> ' + result));
end;

function FxMATCH(Formula: string): string;
var
  Index, Row, Col: integer;
  Vars: T1DStringArray;
  Data: T2DStringArray;
  Calculate, Value: string;
  Multi: boolean;
begin
  result := Formula;
  Vars := ExtractVars(Formula);
  if XlArray(Vars[0], Data) then
    begin
      Multi := False;
      result := '';
      for Row := Low(Data) to High(Data) do
      begin
        for Col := Low(Data[Row]) to High(Data[Row]) do
          begin
            Calculate := '=MATCH(' + Data[Row, Col];
            for Index := 1 to High(Vars) do
              Calculate := Calculate + ',' + Vars[Index];
            Calculate := Calculate + ')';
            if XlTryOleToStr(XlEvaluate(Calculate), Value) then
              begin
                if (result <> '') then
                  Multi := True;
                result := result + Value;
              end;
            if Col < High(Data[Row]) then
              result := result + ',';
          end;
        if Row < High(Data) then
          result := result + ';';
      end;
        if Multi then
          result := '{' + result + '}';
    end;
end;

function FxSUMIF(Formula: string; Union: boolean = True): string;
var
  Numeric: double;
  Row, Col: integer;
  Workbook, Sheet, Cells, Value: string;
  Cell: array[0..2] of string;
  Vars: T1DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      if not XlDecodeRange(Vars[0], Workbook, Sheet, Cell[0], True) then
        exit(Formula);
      Value := Vars[1];
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      if High(Vars) = 2 then
        begin
          if not XlIsRange(Vars[2], Cell[0]) then
            exit(Formula);
          Cell[0] := XlEndOfRange(Cell[0]);
          Cell[0] := SplitString(Cell[0], ':')[0];
        end
      else
        Cell[0] := Cell[1];
      try
        for Col := 0 to (XlSheet.Range[Cell[1], Cell[2]].Columns.Count - 1) do
          for Row := 0 to (XlSheet.Range[Cell[1], Cell[2]].Rows.Count - 1) do
            try
              if XlCriteria(XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False], Value) and (TryStrToFloat(XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2, Numeric) or (XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2 = '')) then
                Cells := Cells + ',' + XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False];
            except end;
      except end;
      if (Cells = '') then
        exit(Formula)
      else
        begin
          Cells := Copy(Cells, 2, Length(Cells));
          if Union then
            Cells := XlUnion(Cells);
          result := 'SUM(' + Cells + ')';
        end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxSUMIFS(Formula: string; Union: boolean = True): string;
var
  Add: boolean;
  Numeric: double;
  Index, Row, Col: integer;
  Workbook, Sheet, Cells: string;
  Cell: array[0..2] of string;
  Vars: T1DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      if not XlDecodeRange(Vars[0], Workbook, Sheet, Cell[0], True) then
        exit(Formula);
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      try
        for Col := 0 to (XlSheet.Range[Cell[1], Cell[2]].Columns.Count - 1) do
          for Row := 0 to (XlSheet.Range[Cell[1], Cell[2]].Rows.Count - 1) do
            begin
              Add := true;
              for Index := 1 to (Length(Vars) div 2) {step 2} do
                begin
                  if not XlIsRange(Vars[Index], Cell[0]) then
                    exit(Formula);
                  Cell[0] := XlEndOfRange(Cell[0]);
                  Cell[0] := SplitString(Cell[0], ':')[0];
                  if (Index < High(Vars)) and not XlCriteria(XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False], Vars[Index+1]) then
                    Add := false;
                  PInteger(@Index)^ := PInteger(@Index)^ + 2-1;
                end;
              if Add and (TryStrToFloat(XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Value2, Numeric) or (XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2 = '')) then
                Cells := Cells + ',' + XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False];
            end;
      except end;
      if (Cells = '') then
        exit(Formula)
      else
        begin
          Cells := Copy(Cells, 2, Length(Cells));
          if Union then
            Cells := XlUnion(Cells);
          result := 'SUM(' + Cells + ')';
        end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxAVERAGEIF(Formula: string; Union: boolean = True): string;
var
  Numeric: double;
  Row, Col: integer;
  Workbook, Sheet, Cells, Value: string;
  Cell: array[0..2] of string;
  Vars: T1DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      if not XlDecodeRange(Vars[0], Workbook, Sheet, Cell[0], True) then
        exit(Formula);
      Value := Vars[1];
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      if High(Vars) = 2 then
        begin

          if not XlIsRange(Vars[2], Cell[0]) then
            exit(Formula);
          Cell[0] := XlEndOfRange(Cell[0]);
          Cell[0] := SplitString(Cell[0], ':')[0];
        end
      else
        Cell[0] := Cell[1];
      try
        for Col := 0 to (XlSheet.Range[Cell[1], Cell[2]].Columns.Count - 1) do
          for Row := 0 to (XlSheet.Range[Cell[1], Cell[2]].Rows.Count - 1) do
            try
              if XlCriteria(XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False], Value) and (TryStrToFloat(XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2, Numeric) or (XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2 = '')) then
                Cells := Cells + ',' + XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False];
            except end;
      except end;
      if (Cells = '') then
        exit(Formula)
      else
        begin
          Cells := Copy(Cells, 2, Length(Cells));
          if Union then
            Cells := XlUnion(Cells);
          result := 'AVERAGE(' + Cells + ')';
        end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxAVERAGEIFS(Formula: string; Union: boolean = True): string;
var
  Add: boolean;
  Numeric: double;
  Index, Row, Col: integer;
  Workbook, Sheet, Cells: string;
  Cell: array[0..2] of string;
  Vars: T1DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      if not XlDecodeRange(Vars[0], Workbook, Sheet, Cell[0], True) then
        exit(Formula);
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      try
        for Col := 0 to (XlSheet.Range[Cell[1], Cell[2]].Columns.Count - 1) do
          for Row := 0 to (XlSheet.Range[Cell[1], Cell[2]].Rows.Count - 1) do
            begin
              Add := true;
              for Index := 1 to (Length(Vars) div 2) {step 2} do
                begin
                  if not XlIsRange(Vars[Index], Cell[0]) then
                    exit(Formula);
                  Cell[0] := XlEndOfRange(Cell[0]);
                  Cell[0] := SplitString(Cell[0], ':')[0];
                  if (Index < High(Vars)) and not XlCriteria(XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False], Vars[Index+1]) then
                    Add := false;
                  PInteger(@Index)^ := PInteger(@Index)^ + 2-1;
                end;
              if Add and (TryStrToFloat(XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Value2, Numeric) or (XlSheet.Range[Cell[0], EmptyParam].Offset[Row,Col].Value2 = '')) then
                Cells := Cells + ',' + XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False];
            end;
      except end;
      if (Cells = '') then
        exit(Formula)
      else
        begin
          Cells := Copy(Cells, 2, Length(Cells));
          if Union then
            Cells := XlUnion(Cells);
          result := 'AVERAGE(' + Cells + ')';
        end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxINDEX(Formula: string; Default: string = '#N/A'): string;
var
  Row, RowX, RowY, Col, ColX, ColY: integer;
  Workbook, Sheet, Range: string;
  Cell: array[0..2] of string;
  Vars: T1DStringArray;
  Rows, Cols: T2DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      if not XlIsRange(Vars[0], Range) or not XlDecodeRange(Range, Workbook, Sheet, Cell[0], True, True) then
        exit(Formula);
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      case High(Vars) of
        1:
          if XlArray(Vars[1], Rows) then
            for RowX := 0 to High(Rows) do
              for RowY := 0 to High(Rows[RowX]) do
                try
                  Row := StrToInt(XlEvaluate('=' + Rows[RowX,RowY])) - 1;
                  result := result + ',' + XlSheet.Range[Cell[1], EmptyParam].Offset[Row,0].Address[False, False, xlA1, True, False];
                except
                  result := result + ',' + Default;
                end;
        2:
          if XlArray(Vars[1], Rows) and XlArray(Vars[2], Cols) then
            begin
              for RowX := 0 to High(Rows) do
                for RowY := 0 to High(Rows[RowX]) do
                  for ColX := 0 to High(Cols) do
                    for ColY := 0 to High(Cols[ColX]) do
                      try
                        Row := StrToInt(XlEvaluate('=' + Rows[RowX,RowY])) - 1;
                        Col := StrToInt(XlEvaluate('=' + Cols[ColX,ColY])) - 1;
                        result := result + ',' + XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False];
                      except
                        result := result + ',' + Default;
                      end;
            end;
      end;
      result := Copy(result, 2, Length(result)-1);
    except
      try
        result := XlEvaluate('=CELL("address",' + Formula + ')');
      except
        exit(Formula);
      end;
    end;
  OutputDebugString(PChar(Formula + ' -> ' + result));
end;

function FxVLOOKUP(Formula: string; Default: string = '#N/A'): string;
var
  Row, Col, ERow, ECol, VCol: integer;
  Criteria, Value, Workbook, Sheet: string;
  Cell: array[0..2] of string;
  CArray, Vars: T1DStringArray;
  Found: ExcelRange;
begin
    try
      Vars := ExtractVars(Formula);
      CArray := XlOleToArray(Vars[0]);
      if not XlDecodeRange(Vars[1], Workbook, Sheet, Cell[0], True) or not TryStrToInt(XlApp.Evaluate(Vars[2],0), VCol) then
        exit(Formula);
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      ERow := XlSheet.Range[Cell[1], Cell[2]].Rows.Count;
      ECol := XlSheet.Range[Cell[1], Cell[2]].Columns.Count;
      case Length(CArray) of
        0:
          exit(Formula);
        else
          try
            for Criteria in CArray do
              try
                Found := nil;
                Value := XlApp.Evaluate('=' + Criteria,0);
                if (High(Vars) = 2) or ((High(Vars) = 3) and (StrToBool(Vars[3]) = True)) then
                  begin
                    MessageDlg('Could not convert "'+Formula+'"'+#13#10+'Formula conversion not support "Approximate match" yet.',mtError, [mbOk], 0);
                    exit(Formula);
                  end
                else
                  Found := XlSheet.Range[Cell[0], Cell[1]].Find(Value, XlSheet.Range[Cell[1],EmptyParam], xlValues, xlWhole, xlByRows, xlNext, EmptyParam, True, False);
              finally
                if not (Found = nil) then
                  begin
                    result := result + ',' + Found.Offset[0, VCol - 1].Address[False, False, xlA1, True, False];
                    CArray := RemoveFromArray(CArray, AnsiIndexStr(Criteria, CArray));
                  end;
              end;
            // old search procedure - slowly
            if (result = '') then
              for Row := 0 to (ERow - 1) do
                for Col := 0 to (ECol - 1) do
                  for Criteria in CArray do
                    if XlCriteria(XlSheet.Range[Cell[1], EmptyParam].Offset[Row,Col].Address[False, False, xlA1, True, False], Criteria) or ((High(Vars) = 3) and (Vars[3] = '-1') and ContainsText(string(XlSheet.Range[Cell[1],EmptyParam].Offset[Row, Col].Value2), Criteria)) then
                      begin
                        result := result + ',' + XlSheet.Range[Cell[1],EmptyParam].Offset[Row, Col].Offset[0, VCol -1].Address[False, False, xlA1, True, False];
                        CArray := RemoveFromArray(CArray, AnsiIndexStr(Criteria, CArray));
                        Break;
                      end;
            // end of old search procedure - slowly
          finally
            if not (result = '') then
              begin
                result := Copy(result, 2, Length(result)-1);
              end;
          end;
      end;
      if (result = '') then
        result := Default;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxHLOOKUP(Formula: string; Default: string = '#N/A'): string;
  function ExtractCol(Cell: string): string;
  begin
    result := TRegEx.Create('[A-Z]{1,4}').Match(Cell).Value;
  end;
  function ExtractRow(Cell: string): string;
  begin
    result := TRegEx.Create('[0-9]{1,7}').Match(Cell).Value;
  end;
var
  Col, ECol, VRow: integer;
  Criteria, Workbook, Sheet: string;
  Cell: array[0..2] of string;
  CArray, Vars: T1DStringArray;
begin
    try
      Vars := ExtractVars(Formula);
      CArray := XlOleToArray(Vars[0]);
      if not XlDecodeRange(Vars[1], Workbook, Sheet, Cell[0], True) or not TryStrToInt(string(XlApp.Evaluate(Vars[2],0)), VRow) then
        exit(Formula);
      XlWorkbook := XlApp.Workbooks[Workbook];
      XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
      Cell[1] := SplitString(Cell[0], ':')[0];
      Cell[2] := SplitString(Cell[0], ':')[1];
      Cell[2] := ExtractCol(Cell[2]) + ExtractRow(Cell[1]);
      ECol := XlSheet.Range[Cell[1], Cell[2]].Columns.Count;
      case Length(CArray) of
        0:
          exit(Formula);
        else
          try
            if (result = '') then
              for Col := 0 to (ECol - 1) do
                for Criteria in CArray do
                  if XlCriteria(XlSheet.Range[Cell[1], EmptyParam].Offset[0,Col].Address[False, False, xlA1, True, False], Criteria) or ((High(Vars) = 3) and (Vars[3] = '-1') and ContainsText(string(XlSheet.Range[Cell[1],EmptyParam].Offset[0, Col].Value2), Criteria)) then
                    begin
                      result := result + ',' + XlSheet.Range[Cell[1],EmptyParam].Offset[0, Col].Offset[VRow -1, 0].Address[False, False, xlA1, True, False];
                      CArray := RemoveFromArray(CArray, AnsiIndexStr(Criteria, CArray));
                      Break;
                    end;
          finally
            if not (result = '') then
              begin
                result := Copy(result, 2, Length(result)-1);
              end;
          end;
      end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

function FxXLOOKUP(Formula: string; Default: string = '#N/A'): string;
var
  Sort, Match: integer;
  Criteria, Address, Return: string;
  CArray, Addresses, Returns, Vars: T1DStringArray;
begin
    try
      Sort := 1;
      Vars := ExtractVars(Formula);
      CArray := XlOleToArray(Vars[0]);
      if not XlRangeToArray(Vars[1], Addresses) then
        exit(Formula);

      if High(Vars) >= 3 then
        Default := Vars[3];
      if High(Vars) >= 4 then
        Match := StrToInt(Vars[4]); // 0 = Exact match, -1 = Exact match or smaller item, 1 = Exact match or larger item, 2 = Wildmatch character match.
      if High(Vars) >= 5 then
        Sort := StrToInt(Vars[5]); // 1 = Search from first-to-last, -1 = Search from last-to-first, 2 = Binary search (sorted ascending order), -2 = Binary search (sorted descending order).

      if not XlRangeToArray(Vars[2], Returns) then
        exit(Formula);

      case Length(CArray) of
        0:
          exit(Formula);
        else
          try
            for Address in Addresses do
                for Criteria in CArray do
                begin
                  if XlCriteria(Address, Criteria) then
                    begin
                      result := result + ',' + Return[AnsiIndexStr(Address, Addresses)];
                      CArray := RemoveFromArray(CArray, AnsiIndexStr(Criteria, CArray));
                      Break;
                    end;
                end;
          finally
            if not (result = '') then
              begin
                result := Copy(result, 2, Length(result)-1);
              end;
          end;
      end;
      OutputDebugString(PChar(Formula + ' -> ' + result));
    except
      exit(Formula);
    end;
end;

begin
end.
