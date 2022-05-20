unit XlSheetHeader;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls, Excel2010, XlApplication, XlUtilites;

  procedure SheetHeaderGUI(); stdcall;

type
  TSheetHeader = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    PeriodEnd: TDateTimePicker;
    CurrencyUnit: TEdit;
    Title: TEdit;
    Client: TComboBox;
    Add: TButton;
    procedure AddClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SheetHeader: TSheetHeader;

implementation

{$R *.dfm}

procedure SheetHeaderGUI(); stdcall;
begin
  Application.CreateForm(TSheetHeader, SheetHeader);
  try
    SheetHeader.ShowModal;
  finally
    SheetHeader.Destroy;
  end;
end;


procedure TSheetHeader.AddClick(Sender: TObject);
var
  XlRange: ExcelRange;
begin
  XlWorkbook := XlApp.ActiveWorkbook;
  XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;
  XlRange := XlSheet.UsedRange[0];
  XlRange := XlSheet.Range['A1', XlSheet.Range['A1', EmptyParam].Offset[3, XlRange.Columns.Count-1].Address[False, False, xlA1, False, False]];
  XlRange.EntireRow.Insert(xlShiftDown, Null);
  XlRange := XlSheet.Range['A1', XlSheet.Range['A1', EmptyParam].Offset[3, XlRange.Columns.Count-1].Address[False, False, xlA1, False, False]];
  XlRange.Interior.Color := xlNone;
  XlRange.Name := 'ZH_' + XlSheet.CodeName;
  XlSheet.Range['C1', XlSheet.Range['A1', EmptyParam].Offset[3, XlRange.Columns.Count-3].Address[False, False, xlA1, False, False]].Merge(EmptyParam);
  ShowMessage( XlRange.Address[False, False, xlA1, False, False] );
end;

end.
