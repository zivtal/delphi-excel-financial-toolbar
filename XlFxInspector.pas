unit XlFxInspector;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, StrUtils, UITypes, Clipbrd, Math, Excel2010, Registry, Vcl.Grids,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Tabs, RegularExpressions, Vcl.CheckLst;

  procedure FxInspectorGUI(); stdcall;

type
  TArrayString = array of string;
  TFxInspector = class(TForm)
    Output: TStringGrid;
    FxFormula: TRichEdit;
    Status: TStatusBar;
    FxList: TCheckListBox;
    FxInput: TComboBox;
    function SplitFormula(Formula: string): TArrayString;
    function IndexOfArray(Input: array of string; Value: string): integer;
    function ConvertFx(Formula: string; Union: boolean = True; Recall: boolean = False; By: string = ''): string;
    procedure CheckFx(Formula: string; Recall: boolean = false);
    procedure Select(Row: integer);
    procedure Read(Input: string; Actual: boolean);
    procedure OutputSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FxFormulaEnter(Sender: TObject);
    procedure OutputEnter(Sender: TObject);
    procedure FxInputChange(Sender: TObject);
  private
    LastSheet, LastFormula: string;
    FxUncompiled, FxRange, FxSheet, FxWorkbook: array of string;
    CFormula, CValue, AFormula, AValue, ARange, ASheet, AWorkbook, ACell: string;
  public
  end;

var
  FxInspector: TFxInspector;
  XlWorkbook: ExcelWorkbook;
  XlSheet: ExcelWorksheet;

implementation

uses
  XlUtilites, XlApplication, RegistryCtrl, StringCtrl;

{$R *.dfm}

const
  Registry: string = 'SOFTWARE\Ziv Tal\ZTools\MarkReferences';
  Support: array of string =
    [
      'IF',           // 0
      'INDIRECT',     // 1
      'MATCH',        // 2
      'INDEX',        // 3
      'SUMIF',        // 4
      'SUMIFS',       // 5
      'AVERAGEIF',    // 6
      'AVERAGEIFS',   // 7
      'VLOOKUP',      // 8
      'HLOOKUP',      // 9
      'MAX',          // 10
      'MIN',          // 11
      'DATE',         // 12
      'DAY',          // 13
      'MONTH',        // 14
      'YEAR',         // 15
      'DAYS',         // 16
      'EOMONTH',      // 17
      'SUM',          // 18
      'AVERAGE',      // 19
      'COUNT',        // 20
      'IFNA',         // 21
      'IFERROR',      // 22
      'HYPERLINK',    // 23
      'ROW',          // 24
      'COLUMN'        // 25
    ];

function ColorRandom(): double;
begin
  result := RGB(RandomRange(150,250),RandomRange(150,250),RandomRange(150,250));
end;

procedure FxInspectorGUI();
var
  Formula: string;
begin
  Formula := XlApp.ActiveCell.Formula;
  if (Formula = '') or (Copy(Formula,1,1) <> '=') then
    exit
  else
    begin
      if (XlTypeOf(XlApp.ActiveCell.Value2) = 10) and (MessageDlg('The formula contain errors, Would you like to continue?', mtWarning, [mbYes, mbNo], 0, mbYes) = mrNo) then
        exit;
      Application.CreateForm(TFxInspector, FxInspector);
      try
        FxInspector.ShowModal;
      finally
        FxInspector.Close;
      end;
    end;
end;

procedure TFxInspector.CheckFx(Formula: string; Recall: boolean = false);
const
  Needed: array of string = ['INDIRECT'];
  Remove: array of string = ['HYPERLINK','IFNA','IFERROR'];
var
  Fx, SubFx: string;
begin
  if Copy(Formula, 1, 1) = '=' then
    Formula := Copy(Formula, 2, Length(Formula));
  for Fx in ExtractFormulas(Formula) do
    begin
      if (Copy(Fx, 0, Pos('(', Fx) -1) <> '') and not MatchStr(Copy(Fx, 0, Pos('(', Fx) -1), FxList.Items.ToStringArray) then
        begin
          FxList.Items.Add(Copy(Fx, 0, Pos('(', Fx) -1));
          if MatchStr(Copy(Fx, 0, Pos('(', Fx) -1), Support) then
            begin
              FxList.Checked[FxList.Items.Count-1] := True;
              if Recall and MatchStr(Copy(Fx, 0, Pos('(', Fx) -1), Needed) then
                FxList.ItemEnabled[FxList.Items.Count-1] := False;
            end
          else
            FxList.ItemEnabled[FxList.Items.Count-1] := False;
        end;
      for SubFx in ExtractVars(Fx) do
        CheckFx(SubFx, True);
      if not Recall then
        FxInput.Items.Add(Fx);
    end;
end;

function TFxInspector.SplitFormula(Formula: string): TArrayString;
var
  RegEx: TRegEx;
  Match: TMatch;
  Matches: TMatchCollection;
begin
// ((?:'[^']+'!|\w+!)?[A-Z\$]+\d+(?:\:[a-zA-Z\(\)\$]+\d+)?)|(?:[A-Z]{2,99}\((?:[^)(]+|(?R))*+\))
// ((?:'[^']+'!|\w+!)?[A-Z\$]+\d+(?:\:[a-zA-Z\$]+\d+)?)|(?:[A-Z]*\((?:[^)(]+|(?R))*+\))
  RegEx := TRegex.Create('((?:''[^'']+''!|\w+!)?[A-Z\$]+\d+(?:\:[a-zA-Z\$]+\d+)?)|(?:[A-Z]*\((?:[^)(]+|(?R))*+\))');
  Matches := RegEx.Matches(Formula);
  for Match in Matches do
    begin
      SetLength(result, Length(result)+1);
      result[High(result)] := Match.Value;
    end;
end;

function TFxInspector.IndexOfArray(Input: array of string; Value: string): integer;
var
  Index: integer;
begin
  for Index := 0 to High(Input) do
    if Input[Index] = Value then
      exit(Index);
end;

function TFxInspector.ConvertFx(Formula: string; Union: boolean = True; Recall: boolean = False; By: string = ''): string;
  function ReplaceFx(Input: string; Union: boolean; Recall: boolean = False; By: string = ''): string;
  var
    Fx: string;
  begin
    result := Input;
    Fx := Copy(Input, 0, Pos('(', Input) -1);
    if (Fx <> '') then
      case AnsiIndexStr(Fx, Support) of
        0: result := FxIF(result);                     // IF
        1: result := FxINDIRECT(result);               // INDIRECT
        2: result := FxMATCH(result);                  // MATCH
        3: result := FxINDEX(result);                  // INDEX
        4: result := FxSUMIF(result, Union);           // SUMIF
        5: result := FxSUMIFS(result, Union);          // SUMIFS
        6: result := FxAVERAGEIF(result, Union);       // AVERAGEIF
        7: result := FxAVERAGEIFS(result, Union);      // AVERAGEIFS
        8: result := FxVLOOKUP(result);                // VLOOKUP
        9: result := FxHLOOKUP(result);                // HLOOKUP
        10: result := FxMAX(result, Recall);           // MAX
        11: result := FxMIN(result, Recall);           // MIN
        12..17: result := XlEvaluate(result);          // DATE, DAY, MONTH, YEAR, DAYS, EOMONTH
        18..20: result := XlNonunion(result,false);    // NONUNION: SUM, COUNT, AVERAGE
        21: result := FxIFERROR(result,'#N/A');        // IFNA
        22: result := FxIFERROR(result,'#ERROR');      // IFERROR
        23: result := ExtractVars(result)[1];          // HYPERLINK
        24: result := FxROW(result);                   // ROW
        25: result := FxCOLUMN(result);                // COLUMN
      end;
  end;
var
  Variant, Tempory, Cell, Add: string;
begin
  result := Formula;
  if (By = '') then
    By := Copy(Formula, 1, Pos('(', Formula) -1);
  if Copy(result, 1, 1) = '=' then
    result := Copy(result, 2, Length(result));
  for Formula in ExtractFormulas(result) do
    begin
      Tempory := Formula;
      for Variant in ExtractVars(Tempory) do
        Tempory := StringReplace(Tempory, Variant, ConvertFx(Variant, Union, True, By), [rfReplaceAll, rfIgnoreCase]);
      Tempory := ReplaceFx(Tempory, Union, Recall, By);
      result := StringReplace(result, Formula, Tempory, [rfReplaceAll, rfIgnoreCase]);
      if not Recall then
        for Cell in ExtractCells(Tempory) do
          begin
            Add := '';
            if not ContainsText(Cell, '!') then
              begin
                Add := SplitString(ARange, '!')[0];
                if not ((Add[1] = '''') and (Add[Length(Add)] = '''')) then
                  Add := '''' + Add + '''';
                Add := Add + '!';
              end;
            Add := Add + Cell;
          end;
    end;
    result := StringReplace(result, sLineBreak, '', [rfReplaceAll]);
end;

procedure TFxInspector.Read(Input: string; Actual: boolean);
var
  Index, Count: integer;
  Numeric: double;
  Splitted, Formula, Cell, Range, Workbook, Sheet: string;
  Value: OleVariant;
begin
  OutputDebugString(PChar(Input));
  CFormula := Input;
  Output.RowCount := 2;
  SetLength(FxUncompiled, 0);
  for Splitted in SplitFormula(AFormula) do
    begin
      if XlIsFormula(Splitted, Formula) then
        Cell := ConvertFx(Splitted, False)
      else
        Cell := Splitted;
      CFormula := StringReplace(CFormula, Splitted, Cell, [rfReplaceAll]);
      for Cell in ExtractCells(Cell) do
        begin
          if XlDecodeRange(Cell, Workbook, Sheet, Range) then
            try
              XlWorkbook := XlApp.Workbooks[Workbook];
              XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
              Index := Index + 1;
              with Output do
                begin
                  // Uncompiled
                  SetLength(FxUncompiled, Length(FxUncompiled)+1);
                  FxUncompiled[High(FxUncompiled)] := Formula;
                  // Range
                  SetLength(FxRange, Length(FxRange)+1);
                  FxRange[High(FxRange)] := Range;
                  // Sheet
                  SetLength(FxSheet, Length(FxSheet)+1);
                  FxSheet[High(FxSheet)] := Sheet;
                  // Workbook
                  SetLength(FxWorkbook, Length(FxWorkbook)+1);
                  FxWorkbook[High(FxWorkbook)] := Workbook;
                  // Set output
                  if Output.RowCount < Index + 1 then
                    Output.RowCount := Index + 1;
                  Output.Cells[0, Index] := Range;
                  Output.Cells[1, Index] := Sheet;
                  Output.Cells[2, Index] := Workbook;
                end;
              XlWorkbook := XlApp.Workbooks[Workbook];
              XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
              Value := XlSheet.Range[Range,EmptyParam].Value2;
            finally
              if TryStrToFloat(Value, Numeric) then
                begin
                  Count := Count + 1;
                  Output.Cells[3, Index] := Format('%n', [Numeric]);
                  Status.Panels[1].Text := ' Count: ' + IntToStr(Count);
                end
              else
                Output.Cells[3, Index] := Value;
            end;
        end;
    end;
  if Length(FxRange) = 0 then
    begin
      MessageDlg('Could not detect refer cells in formula.', mtError, [mbOk], 0, mbOk);
      FxInspector.Destroy;
      exit;
    end
  else
    try
      CFormula := StringReplace(CFormula, '[' + AWorkbook + ']', '', [rfReplaceAll]);
      if (AValue = '') or not XlTryOleToStr(XlApp.Evaluate(CFormula,0), CValue) then
        Status.Panels[3].Text := ' Error'
      else if CValue = AValue then
        Status.Panels[3].Text := ' Success'
      else
        Status.Panels[3].Text := ' Warning';
      Select(1);
    except
    end;
end;

procedure TFxInspector.FormCreate(Sender: TObject);
var
  Value: double;
  StartTime: TDateTime;
begin
  StartTime := Now;
  XlTryOleToStr(XlApp.ActiveCell.Value2, AValue);
  AFormula := XlApp.ActiveCell.Formula;
  AFormula := StringReplace(StringReplace(AFormula, #10, '', [rfReplaceAll]), #13, '', [rfReplaceAll]);
  ARange := XlApp.ActiveCell.Address[False, False, xlA1, True, False];
  XlDecodeRange(XlApp.ActiveCell.Address[False, False, xlA1, True, False], AWorkbook, ASheet, ACell);
  FxFormula.Text := AFormula;
  CheckFx(AFormula);
  with Output do
    begin
      Cells[0, 0] := 'Cell';
      ColWidths[0] := Round((Output.Width - 5) * (10 / 100));
      Cells[1, 0] := 'Sheet';
      ColWidths[1] := Round((Output.Width - 5) * (32 / 100));
      Cells[2, 0] := 'Workbook';
      ColWidths[2] := Round((Output.Width - 5) * (42 / 100));
      Cells[3, 0] := 'Value';
      ColWidths[3] := Round((Output.Width - 5) * (16 / 100));
    end;
  Status.Font.Size := 8;
  Read(AFormula, True);
  if TryStrToFloat(AValue, Value) then
    Status.Panels[2].Text := Format(' Value: %n', [Value])
  else
    Status.Panels[2].Text := ' ' + AValue;
  Status.Font.Size := 8;
  Status.Panels[0].Text := TimeToStr(Now - StartTime);
end;

procedure TFxInspector.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
const
  UndoRegistry: string = '\SOFTWARE\Ziv Tal\ZTools\Undo\Formula';
begin
  XlWorkbook := XlApp.ActiveWorkbook;
  XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
  if (Shift = [ssCtrl]) then
    case Key of
      37:
        if XlApp.ActiveCell.Column > 1 then
          XlApp.ActiveCell.Offset[0, -1].Activate;
      38:
        if XlApp.ActiveCell.Row > 1 then
          XlApp.ActiveCell.Offset[-1, 0].Activate;
      39:
        if XlApp.ActiveCell.Column < 16384 then
          XlApp.ActiveCell.Offset[0, 1].Activate;
      40:
        if XlApp.ActiveCell.Row < 1048576 then
          XlApp.ActiveCell.Offset[1, 0].Activate;
      67:
        ClipBoard.AsText := CFormula;
      187:
        XlApp.ActiveWindow.Zoom := XlApp.ActiveWindow.Zoom + 10;
      189:
        XlApp.ActiveWindow.Zoom := XlApp.ActiveWindow.Zoom - 10;
    end
  else
    case Key of
      13:
        FxInspector.Close;
      27:
        begin
          FxFormulaEnter(Self.FxFormula);
          FxInspector.Close;
        end;
      107:
        XlApp.ActiveWindow.Zoom := XlApp.ActiveWindow.Zoom + 10;
      109:
        XlApp.ActiveWindow.Zoom := XlApp.ActiveWindow.Zoom - 10;
    end;
end;

procedure TFxInspector.FxFormulaEnter(Sender: TObject);
begin
  XlActivate(ARange);
end;

procedure TFxInspector.FxInputChange(Sender: TObject);
begin
  case FxInput.ItemIndex of
    0:
      Read(ConvertFx(Copy(AFormula,2,Length(AFormula))), True);
    else
      Read(ConvertFx(FxInput.Text), True);
  end;
end;

procedure TFxInspector.Select(Row: integer);
var
  Formula: array[1..2] of string;
  Range: string;
begin
  if not XlEncodeRange(Output.Cells[2,Row], Output.Cells[1,Row], Output.Cells[0,Row], Range) then
    exit;
  XlActivate(Range,LastSheet <> Output.Cells[1,Row]);
  if LastFormula <> FxUncompiled[Row -1] then
  try
    FxFormula.Clear;
    FxFormula.SelAttributes.Color := clBlack;
    FxFormula.SelText := AFormula;
    if not (Copy(FxFormula.Text, 1, 1) = '=') then
      FxFormula.Text := '=' + FxFormula.Text;
  finally
    try
      LastFormula := FxUncompiled[Row -1];
      Formula[1] := Copy(FxFormula.Text, 0, Pos(LastFormula, FxFormula.Text) - 1);
      Formula[2] := Copy(FxFormula.Text, Pos(LastFormula, FxFormula.Text) + Length(LastFormula), Length(FxFormula.Text));
      FxFormula.Clear;
      FxFormula.SelAttributes.Color := clBlack;
      FxFormula.SelText := Formula[1];
      FxFormula.SelAttributes.Color := clRed;
      FxFormula.SelAttributes.Style := [fsUnderline];
      FxFormula.SelText := LastFormula;
      FxFormula.SelAttributes.Color := clBlack;
      FxFormula.SelAttributes.Style := [];
      FxFormula.SelText := Formula[2];
      LastFormula := LastFormula;
    except end;
  end;
  LastSheet := Output.Cells[1,Row];
  ExcelForeground;
  if FxInspector.Visible and FxInspector.Enabled then
    FxInspector.SetFocus;
//  Status.Panels[0].Text := (' ' + FxUncompiled[Row-1]);
end;

procedure TFxInspector.OutputEnter(Sender: TObject);
begin
  Select(Output.Row);
end;

procedure TFxInspector.OutputSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  if (ARow = 0) or (Output.Cells[0,ARow] = '') then
    exit;
  Select(ARow);
end;

end.
