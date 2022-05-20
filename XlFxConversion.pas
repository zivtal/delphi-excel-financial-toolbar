unit XlFxConversion;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes, System.DateUtils, System.UITypes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, Vcl.CheckLst, Vcl.ComCtrls,
  Vcl.ExtCtrls, Variants, Excel2010, Registry, StrUtils, XlApplication, XlProgress, XlUtilites, RegistryCtrl;

  procedure FxConversionGUI; stdcall;
  procedure FxConversionUndo; stdcall;

type
  TFxConversion = class(TForm)
    ConvertButton: TButton;
    FxList: TCheckListBox;
    FxInput: TComboBox;
    FxOutput: TRichEdit;
    Label1: TLabel;
    procedure CheckFx(Formula: string; Recall: boolean = false);
    function ConvertFx(Formula: string; Union: boolean = True; Recall: boolean = False; By: string = ''): string;
    procedure ConvertButtonClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function FxListCheckedCount(): integer;
    procedure FxListClickCheck(Sender: TObject);
    procedure FxInputChange(Sender: TObject);
  private
    AFormula, ARange, AWorkbook, ASheet, ACell: string;
  public
  end;

var
  XlWorkbook: ExcelWorkbook;
  XlSheet: ExcelWorksheet;
  FxConversion: TFxConversion;

implementation

{$R *.dfm}

const
  Registry: string = '\SOFTWARE\Ziv Tal\ZTools\Undo\Formula';
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
      'COLUMN',       // 25
      'SUBSTITUTE',   // 26
      'ADDRESS'       // 27
    ];


procedure FxConversionGUI; stdcall;
var
  Formula, Address: string;
begin
  Formula := XlApp.ActiveCell.Formula;
  Address := XlApp.ActiveCell.Address[False, False, xlA1, True, False];
  if (RegRead(Registry, Address) <> '') and (MessageDlg('Do you want to cancel formula conversion?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes) then
    begin
      XlApp.ActiveCell.Formula := RegRead(Registry, Address);
      RegDelete(Registry, Address);
      MessageDlg('The formula has been restored.',mtInformation, [mbOK],0)
    end
  else
    if (Formula = '') or (Copy(Formula,1,1) <> '=') then
      begin
        MessageDlg('The cell does not contain formula',mtError, [mbOk], 0);
        exit;
      end
    else
      begin
        Application.CreateForm(TFxConversion, FxConversion);
        try
          FxConversion.ShowModal;
        finally
          FxConversion.Destroy;
        end;
      end;
end;

procedure FxConversionUndo;
var
  RegKey: TRegistry;
  Keys: TStringList;
  Index: Integer;
  Workbook, Sheet, Cell: string;
  Formula: OleVariant;
begin
  RegKey := TRegistry.Create;
  try
    if RegKey.OpenKey(Registry, False) then
    begin
      Keys := TStringList.Create;
      try
        RegKey.GetValueNames(Keys);
        for Index := 0 to Keys.Count - 1 do
          begin
            Formula := RegKey.ReadString(Keys.Strings[Index]);
            Cell := Keys.Strings[Index];
            Sheet := SplitString(Cell, '!')[0];
            Sheet := StringReplace(Sheet, '''','', [rfReplaceAll, rfIgnoreCase]);
            if ContainsText(Sheet, '[') and ContainsText(Sheet, ']') then
              begin
                Workbook := Copy(Cell, AnsiPos('[', Cell) + 1, AnsiPos(']', Cell) - AnsiPos('[', Cell) - 1);
                Sheet := StringReplace(Sheet, '[' + Workbook + ']','', [rfReplaceAll, rfIgnoreCase]);
                XlWorkbook := XlApp.Workbooks[Workbook];
              end;
            Cell := SplitString(Cell, '!')[1];
            XlSheet := XlWorkbook.Sheets[Sheet] as ExcelWorksheet;
            XlSheet.Range[Cell, EmptyParam].Formula := Formula;
            RegKey.DeleteValue(Keys.Strings[Index]);
          end;
      finally
        Keys.Free;
      end;
      RegKey.CloseKey;
    end;
  finally
    RegKey.Free;
  end;
end;

function TFxConversion.ConvertFx(Formula: string; Union: boolean = True; Recall: boolean = False; By: string = ''): string;
  function ReplaceFx(Input: string; Union: boolean; Recall: boolean = False; By: string = ''): string;
  var
    Fx: string;
  begin
    result := Input;
    Fx := Copy(Input, 0, Pos('(', Input) -1);
    if (Fx <> '') and MatchStr(Fx, FxList.Items.ToStringArray) and FxList.Checked[FxList.Items.IndexOf(Fx)] then
      case AnsiIndexStr(Fx, Support) of
        0: result := FxIF(result);                    // IF
        1: result := FxINDIRECT(result);              // INDIRECT
        2: result := FxMATCH(result);                 // MATCH
        3: result := FxINDEX(result);                 // INDEX
        4: result := FxSUMIF(result, Union);          // SUMIF
        5: result := FxSUMIFS(result, Union);         // SUMIFS
        6: result := FxAVERAGEIF(result, Union);      // AVERAGEIF
        7: result := FxAVERAGEIFS(result, Union);     // AVERAGEIFS
        8: result := FxVLOOKUP(result);               // VLOOKUP
        9: result := FxHLOOKUP(result);               // HLOOKUP
        10: result := FxMAX(result, Recall);          // MAX
        11: result := FxMIN(result, Recall);          // MIN
        12..17,26..27: result := XlEvaluate(result);  // DATE, DAY, MONTH, YEAR, DAYS, EOMONTH
//        18..20: result := XlNonunion(result);         // NONUNION: SUM, COUNT, AVERAGE
        21: result := FxIFERROR(result,'#N/A');       // IFNA
        22: result := FxIFERROR(result,'#ERROR');     // IFERROR
//        23: result := ExtractVars(result)[1];         // HYPERLINK
        24: result := FxROW(result);                  // ROW
        25: result := FxCOLUMN(result);               // COLUMN
      end;
  end;
var
  Variant, Tempory: string;
begin
  result := Formula;
  if (By = '') then
    By := Copy(Formula, 1, Pos('(', Formula) -1);
  if Copy(result, 1, 1) = '=' then
    result := Copy(result, 2, Length(result));
  for Formula in ExtractFormulas(result) do
    begin
      Tempory := Formula;
      for Variant in ExtractVars(Formula) do
        try
          Tempory := StringReplace(Tempory, Variant, ConvertFx(Variant, Union, True, By), [rfReplaceAll]);
        except
          Tempory := Variant;
        end;
      try
        result := StringReplace(result, Formula, ReplaceFx(Tempory, Union, Recall), [rfReplaceAll]);
      except
        result := Formula;
      end;
    end;
    result := StringReplace(result, sLineBreak, '', [rfReplaceAll]);
end;

procedure TFxConversion.CheckFx(Formula: string; Recall: boolean = false);
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

procedure TFxConversion.FormCreate(Sender: TObject);
begin
  AFormula := XlApp.ActiveCell.Formula;
  AFormula := StringReplace(StringReplace(AFormula, #10, '', [rfReplaceAll]), #13, '', [rfReplaceAll]);
  ARange := XlApp.ActiveCell.Address[False, False, xlA1, True, False];
  ACell := SplitString(ARange, '!')[1];
  ASheet := SplitString(ARange, '!')[0];
  ASheet := StringReplace(ASheet, '''','', [rfReplaceAll, rfIgnoreCase]);
  AWorkbook := Copy(ASheet, AnsiPos('[', ASheet) + 1, AnsiPos(']', ASheet) - AnsiPos('[', ASheet) - 1);
  ASheet := StringReplace(ASheet, '[' + AWorkbook + ']','', [rfReplaceAll, rfIgnoreCase]);

  FxInput.Items.Add('Entire formula');
  FxInput.ItemIndex := 0;
  CheckFx(AFormula);
  FxListClickCheck(Self.FxList);
  FxInputChange(Self.FxInput);
end;

procedure TFxConversion.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Shift = [ssCtrl]) then
    case Key of
      187: FxOutput.Font.Size := FxOutput.Font.Size + 1;
      189: FxOutput.Font.Size := FxOutput.Font.Size - 1;
    end
  else
    case Key of
      13:
        if ConvertButton.Enabled then
          ConvertButtonClick(Self.ConvertButton);
      27:
        FxConversion.Close;
    end;
end;

procedure TFxConversion.FxInputChange(Sender: TObject);
var
  Formula, Replace: string;
  Output: array[1..2] of string;
begin
  case FxInput.ItemIndex of
    0:
      begin
        Formula := ConvertFx(Copy(AFormula,2,Length(AFormula)));
        Replace := Copy(AFormula,2,Length(AFormula));
      end;
    else
      begin
        Formula := ConvertFx(FxInput.Text);
        Replace := FxInput.Text;
      end;
  end;
  Formula := StringReplace(Formula, '[' + AWorkbook + ']','', [rfReplaceAll]);
  Formula := StringReplace(Formula, '''' + ASheet + '''!','', [rfReplaceAll]);
  FxOutput.Text := StringReplace(AFormula, Replace, Formula, [rfReplaceAll, rfIgnoreCase]);
  Output[1] := Copy(FxOutput.Text, 0, AnsiPos(Formula, FxOutput.Text) - 1);
  Output[2] := Copy(FxOutput.Text, AnsiPos(Formula, FxOutput.Text) + Length(Formula), Length(FxOutput.Text));
  FxOutput.Clear;
  FxOutput.SelAttributes.Color := clBlack;
  FxOutput.SelText := Output[1];
  FxOutput.SelAttributes.Color := clRed;
  FxOutput.SelAttributes.Style := [fsUnderline];
  FxOutput.SelText := Formula;
  FxOutput.SelAttributes.Color := clBlack;
  FxOutput.SelAttributes.Style := [];
  FxOutput.SelText := Output[2];
end;

function TFxConversion.FxListCheckedCount(): integer;
var
  Index: integer;
begin
  result := 0;
  for Index := 0 to FxList.Items.Count - 1 do
    if FxList.Checked[Index] then
      Inc(result);
end;

procedure TFxConversion.FxListClickCheck(Sender: TObject);
begin
  FxInputChange(Self.FxList);
end;

procedure TFxConversion.ConvertButtonClick(Sender: TObject);
var
  OldFormula: string;
begin
  if (FxOutput.Text <> '') then
    try
      OldFormula := AFormula;
    finally
      try
        XlApp.ActiveCell.Formula := FxOutput.Text
      finally
        if (OldFormula = XlApp.ActiveCell.Formula) then
          MessageDlg('Formula conversion failed.',mtError, [mbOK],0)
        else
          RegWrite(Registry, XlApp.ActiveCell.Address[False, False, xlA1, True, False], OldFormula, False);
      end;
    end;
  FxConversion.Close;
end;

end.
