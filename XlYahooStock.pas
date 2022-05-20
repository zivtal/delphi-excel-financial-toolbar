unit XlYahooStock;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Classes,
  System.DateUtils,
  System.UITypes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.Grids,
  Vcl.StdCtrls,
  Vcl.CheckLst,
  Vcl.ComCtrls,
  Vcl.ExtCtrls,
  XlApplication,
  StrUtils,
  StringCtrl,
  WebCtrl,
  RegistryCtrl;

  procedure YahooStockGUI(); stdcall;

type
  TYahooStock = class(TForm)
    DataEx: TStringGrid;
    ExportButton: TButton;
    StockCode: TComboBox;
    Label1: TLabel;
    StartDate: TDateTimePicker;
    EndDate: TDateTimePicker;
    Label2: TLabel;
    EndDateCheck: TCheckBox;
    GetData: TButton;
    ValueCheck: TCheckListBox;
    Label3: TLabel;
    IncHeadersCheck: TCheckBox;
    AllRows: TRadioButton;
    SelectedRows: TRadioButton;
    procedure ClearStringGrid(const Grid: TStringGrid);
    function CheckListCount(CLB: TCheckListBox): integer;
    procedure FormCreate(Sender: TObject);
    procedure ExportButtonClick(Sender: TObject);
    procedure GetDataClick(Sender: TObject);
    procedure StartDateExit(Sender: TObject);
    procedure EndDateCheckClick(Sender: TObject);
    procedure StockCodeChange(Sender: TObject);
    procedure ValueCheckClickCheck(Sender: TObject);
    procedure DataExExit(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure StartDateChange(Sender: TObject);
    procedure StartDateKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure EndDateKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  end;

var
  YahooStock: TYahooStock;

implementation

{$R *.dfm}

const
  Registry: string = '\SOFTWARE\Ziv Tal\YahooStock';

function UnixDate(DateIn: double): double; stdcall;
begin
  result := (DateIn - 25569) * 86400
end;

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

function Quarter(Input: TDate; Offset: integer; MaxQuarterEnded: boolean = True): TDate;
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
  if MaxQuarterEnded and (result > Now) then
    result := Input;
end;

procedure YahooStockGUI(); stdcall;
begin
  Application.CreateForm(TYahooStock, YahooStock);
  try
    YahooStock.ShowModal;
  finally
    YahooStock.Destroy;
  end;
end;

procedure TYahooStock.ClearStringGrid(const Grid: TStringGrid);
var
  Col, Row: Integer;
begin
  for Col := 0 to Pred(Grid.ColCount) do
    for Row := 0 to Pred(Grid.RowCount) do
      Grid.Cells[Col, Row] := '';
end;

function TYahooStock.CheckListCount(CLB: TCheckListBox): integer;
var
  Index: integer;
begin
  result := 0;
  for Index := 0 to CLB.Items.Count - 1 do
    if CLB.Checked[Index] then
      inc(result);
end;

procedure TYahooStock.DataExExit(Sender: TObject);
begin
//  SelectedRows.Checked := DataEx.Selection.Bottom <> DataEx.Selection.Top;
//  AllRows.Checked := not SelectedRows.Checked;
end;

procedure TYahooStock.GetDataClick(Sender: TObject);
var
  Row, Col: integer;
  Rows, Cells: TStringArray;
  Data, RowValue, ColValue, CellValue: string;
begin
  if GetData.Caption = 'Get data' then
    begin
      Data := HttpGet('https://query1.finance.yahoo.com/v7/finance/download/' + StockCode.Text + '?period1=' + floattostr(UnixDate(StartDate.Date)) + '&period2=' + floattostr(UnixDate(EndDate.Date + 1)) + '&interval=1d&events=history&includeAdjustedClose=true');
      Rows := StrSplit(Chr(10), Data);
      if StartDate.Date = EndDate.Date then
        EndDateCheck.Checked := false;
        if (Data = '') or ContainsText(LowerCase(Data), '404 not found') then
          MessageDlg(Data,mtInformation, [mbOK],0)
        else
          begin
            Row := 0;
            ValueCheck.Items.Clear;
            DataEx.RowCount := high(Rows) + 1;
            for RowValue in Rows do
            begin
              Col := 0;
              Cells := StrSplit(',', RowValue);
              if DataEx.ColCount < (high(Cells) + 1) then
                DataEx.ColCount := (high(Cells) + 1);
              for CellValue in Cells do
                begin
                  DataEx.Cells[Col, Row] := CellValue;
                  if Row = 0 then
                    ValueCheck.AddItem(CellValue,nil);
                  Col := Col + 1;
                end;
              Row := Row + 1;
            end;

            for Col := 0 to (DataEx.ColCount - 1) do
              begin
                if (DataEx.DefaultRowHeight * DataEx.RowCount) > DataEx.Height then
                  DataEx.ColWidths[Col] := round((DataEx.Width - 26) / DataEx.ColCount)
                else
                  DataEx.ColWidths[Col] := round(DataEx.Width / DataEx.ColCount)
              end;

            if StockCode.Items.IndexOf(StockCode.Text) = -1 then
              StockCode.Items.Add(StockCode.Text);

            RegWrite(Registry + '\LastSettings','Start',DateToStr(StartDate.Date));
            RegWrite(Registry + '\LastSettings','End',DateToStr(EndDate.Date));
            RegWrite(Registry + '\LastSettings','Range',BoolToStr(EndDateCheck.Checked));
            RegWrite(Registry + '\LastSettings','Headers',BoolToStr(IncHeadersCheck.Checked));
            RegWrite(Registry + '\LastSettings','StockCode', StockCode.Items.Text);
            RegWrite(Registry + '\LastSettings','LastCode', StockCode.Text);
            try
              for Col := 0 to ValueCheck.Items.Count - 1 do
                for ColValue in StrSplit(',', RegRead(Registry + '\LastSettings','Col')) do
                  if ColValue = ValueCheck.Items[Col] then
                    ValueCheck.Checked[Col] := true;
            except end;
            ValueCheckClickCheck(Self.ValueCheck);
            GetData.Caption := 'Clear data';
          end;
    end
  else
    begin
      ValueCheck.Items.Clear;
      DataEx.ColCount := 7;
      DataEx.RowCount := 13;
      ClearStringGrid(DataEx);
      ValueCheckClickCheck(Self.GetData);
      GetData.Caption := 'Get data';
    end;
end;

procedure TYahooStock.StartDateChange(Sender: TObject);
begin
  EndDate.MinDate := StartDate.Date;
  if not EndDateCheck.Checked then EndDate.Date := StartDate.Date;
end;

procedure TYahooStock.StartDateExit(Sender: TObject);
begin
  EndDate.MinDate := StartDate.Date;
end;

procedure TYahooStock.StockCodeChange(Sender: TObject);
begin
  GetData.Enabled := Length(StockCode.Text) > 0;
end;

procedure TYahooStock.ValueCheckClickCheck(Sender: TObject);
begin
  ExportButton.Enabled := CheckListCount(ValueCheck) > 0;
  IncHeadersCheck.Enabled := ExportButton.Enabled;
  SelectedRows.Enabled := ExportButton.Enabled;
  AllRows.Enabled := ExportButton.Enabled;
end;

procedure TYahooStock.EndDateCheckClick(Sender: TObject);
begin
  EndDate.Enabled := EndDateCheck.Checked;
  if EndDateCheck.Checked then
    if IsQuarter(StartDate.Date) then
      StartDate.Date := Quarter(EndDate.Date, -1)
  else
    StartDate.Date := EndDate.Date;
end;

procedure TYahooStock.ExportButtonClick(Sender: TObject);
var
  Range: TStringGrid;
  Row, Col, ICol, LRow, ERow: integer;
begin
    Range := TStringGrid.Create(nil);
    Range.RowCount := 0;
    Range.ColCount := 0;
    if AllRows.Checked then
      begin
        LRow := 1;
        ERow := DataEx.RowCount - 1;
      end
    else
      begin
        LRow := DataEx.Selection.Top;
        ERow := DataEx.Selection.Bottom;
      end;
    if IncHeadersCheck.Checked then
      begin
        ICol := 0;
        for Col := 0 to ValueCheck.Items.Count -1 do
          if ValueCheck.Checked[Col] then
          begin
            ICol := ICol + 1;
            Range.Cells[ICol - 1, 0] := ValueCheck.Items[Col];
          end;
        Range.RowCount := Range.RowCount + 1;
      end;
    for Row := LRow to ERow do
      begin
        ICol := 0;
        for Col := 0 to DataEx.ColCount - 1 do
          if ValueCheck.Checked[Col] then
            begin
              ICol := ICol + 1;
              if (Range.ColCount < ICol) and (Range.ColCount < DataEx.ColCount) then
                Range.ColCount := ICol;
              Range.Cells[ICol - 1, Range.RowCount - 1] := DataEx.Cells[Col, Row];
            end;
        Range.RowCount := Range.RowCount + 1;
      end;
    Range.ColCount := Range.ColCount - 1;
    Range.RowCount := Range.RowCount - 2;
    try
      for Row := 0 to Range.RowCount do
        for Col := 0 to Range.ColCount do
          if Range.Cells[Col,Row] <> '' then
            XlApp.ActiveCell.Offset[Row, Col].Value2 := Range.Cells[Col, Row];
    except end;
    RegWrite(Registry + '\LastSettings','Selection',BoolToStr(SelectedRows.Checked));
    RegDelete(Registry + '\LastSettings', 'Col');
    for Col := 0 to ValueCheck.Items.Count -1 do
      if ValueCheck.Checked[Col] then
        RegWrite(Registry + '\LastSettings', 'Col', RegRead(Registry + '\LastSettings','Col') + ',' + ValueCheck.Items[Col]);
    RegWrite(Registry + '\LastSettings', 'Col', Copy(RegRead(Registry + '\LastSettings','Col'), 2, Length(RegRead(Registry + '\LastSettings','Col'))));
    YahooStock.Close;
end;

procedure TYahooStock.FormCreate(Sender: TObject);
begin
  try
    StartDate.Date := StrToDate(RegRead(Registry + '\LastSettings','Start'));
  except
    StartDate.Date := now;
  end;

  try
    EndDate.Date := StrToDate(RegRead(Registry + '\LastSettings','End'));
  except
    EndDate.Date := StartDate.Date;
  end;

  try
    EndDateCheck.Checked := StrToBool(RegRead(Registry + '\LastSettings','Range'));
  except
    EndDateCheck.Checked := true;
  end;

  try
    IncHeadersCheck.Checked := StrToBool(RegRead(Registry + '\LastSettings','Headers'));
  except end;

  try
    StockCode.Items.Text := RegRead(Registry + '\LastSettings','StockCode');
  except end;

  try
    StockCode.Text := RegRead(Registry + '\LastSettings','LastCode');
    StockCodeChange(Self.StockCode);
  except end;

  try
    AllRows.Checked := not StrToBool(RegRead(Registry + '\LastSettings','Selection', '-1'));
  except end;

  try
    SelectedRows.Checked := StrToBool(RegRead(Registry + '\LastSettings','Selection', '-1'));
  except end;

end;

procedure TYahooStock.StartDateKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    32:
      StartDate.Date := Now;
    33:
      StartDate.Date := Quarter(StartDate.Date, 1);
    34:
      StartDate.Date := Quarter(StartDate.Date, -1);
    35:
      StartDate.Date := Quarter(Date, 1);
    192:
      begin
        EndDateCheck.Checked := not EndDateCheck.Checked;
        EndDateCheckClick(Self.EndDateCheck);
      end;
    else
      FormKeyDown(Sender, Key, Shift);
  end;
  StartDateChange(Self.StartDate);
end;

procedure TYahooStock.EndDateKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    32:
      EndDate.Date := Now;
    33:
      EndDate.Date := Quarter(EndDate.Date, 1);
    34:
      EndDate.Date := Quarter(EndDate.Date, -1);
    35:
      EndDate.Date := Quarter(Date, 1);
    192:
      begin
        EndDateCheck.Checked := not EndDateCheck.Checked;
        EndDateCheckClick(Self.EndDateCheck);
      end;
    else
      FormKeyDown(Sender, Key, Shift);
  end;
end;

procedure TYahooStock.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    13:
      if ExportButton.Enabled then
        ExportButtonClick(Self.ExportButton)
      else
        GetDataClick(Self.GetData);
    27:
      if ExportButton.Enabled then
        GetDataClick(Self.GetData)
      else
        YahooStock.Close;
  end;
end;

end.

