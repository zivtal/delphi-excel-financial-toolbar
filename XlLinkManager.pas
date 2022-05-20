unit XlLinkManager;

interface

uses
  Winapi.Windows, Winapi.Messages, System.Types, System.IOUtils, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Math, StrUtils, Excel2010, Vcl.Grids,
  XlApplication, XlUtilites, XlProgress, Vcl.ComCtrls;

  procedure LinkManagerGUI(); stdcall;

type
  TArrayString = array of string;
  TLinkManager = class(TForm)
    BtnBreak: TButton;
    BtnRelink: TButton;
    BtnChange: TButton;
    Links: TStringGrid;
    BtnUpdate: TButton;
    Status: TStatusBar;
    BtnCheckStatus: TButton;
    function CheckLink(Link: string; OkReturn: string = 'OK'): string;
    procedure CheckAllLinks(StartIndex: integer; EndIndex: integer);
    procedure RemoveRow(Index: integer);
    procedure ChangeLink(Index: integer);
    procedure BtnBreakClick(Sender: TObject);
    procedure BtnRelinkClick(Sender: TObject);
    procedure BtnChangeClick(Sender: TObject);
    procedure LinksDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtnUpdateClick(Sender: TObject);
    procedure LinksSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure BtnCheckStatusClick(Sender: TObject);
  private
    { Private declarations }
  public
    XlLinks: TStringList;
    XlSourceLinks: Variant;
  end;

var
  LinkManager: TLinkManager;

implementation

{$R *.dfm}

function StringCut(Input: string; Char: string; Count: integer = 1): string; stdcall;
var
  Index: Integer;
begin
  result := Input;
  if Count < 0 then
    Count := (High(SplitString(Input,Char)) + Count + 1);
  if not (Char = '') then
    for Index := 1 to Count do
      if ContainsText(result, Char) then
        begin
          if ContainsStr(result, Char) then
            result := Copy(result, Pos(Char, result) + Length(Char), Length(result))
          else
            exit;
        end;
end;

function SplitPath(Input: string; out Filename, Path: string): boolean;
begin
  if ContainsText(Input, '//') then
    begin
      Filename := StringCut(Input,'/',-1);
      Path := Copy(Input, 1, Length(Input) - Length(Filename));
    end
  else
    begin
      Filename := StringCut(Input,'\',-1);
      Path := Copy(Input, 1, Length(Input) - Length(Filename));
    end;
  result := (Filename <> '') and (Path <> '');
end;

function IsEmptyOrNull(Value: Variant): Boolean;
begin
  result := VarIsClear(Value) or VarIsEmpty(Value) or VarIsNull(Value) or (VarCompareValue(Value, Unassigned) = vrEqual);
end;

procedure LinkManagerGUI(); stdcall;
begin
  Application.CreateForm(TLinkManager, LinkManager);
  try
    LinkManager.XlSourceLinks := XlApp.ActiveWorkbook.LinkSources(xlExcelLinks, 0);
    if not IsEmptyOrNull(LinkManager.XlSourceLinks) then
      begin
        XlApp.Application.DisplayAlerts[0] := False;
        XlApp.ActiveWorkbook.UpdateLinks := xlUpdateLinksNever;
        LinkManager.ShowModal;
      end
    else
      MessageDlg('The workbook does not contain link sources.',mtError, [mbOk], 0);
  finally
    try
      XlApp.Application.DisplayAlerts[0] := True;
      XlApp.ActiveWorkbook.UpdateLinks := xlUpdateLinksUserSetting;
    except end;
    LinkManager.Destroy;
  end;
end;

procedure TLinkManager.RemoveRow(Index: Integer);
var
  I: integer;
begin
  for I := Index to Links.RowCount - 2 do
    Links.Rows[I].Assign(Links.Rows[I + 1]);
  Links.RowCount := Links.RowCount -1;
end;

procedure TLinkManager.FormCreate(Sender: TObject);
var
  Index: integer;
  Temp: variant;
  Filename, Path: string;
begin
  XlLinks := TStringList.Create;
  Links.Cells[0,0] := 'Filename:';
  Links.Cells[1,0] := 'Status';
  Links.ColWidths[0] := Round((Links.Width-2) * (69/100));
  Links.ColWidths[1] := Round((Links.Width-2) * (32/100));
  Status.Font.Size := 8;
  try
    Temp := XlApp.ActiveWorkbook.LinkSources(xlExcelLinks, 0);
//    if not IsEmptyOrNull(Temp) then
      for Index := VarArrayLowBound(Temp, 1) to VarArrayHighBound(Temp, 1) do
        if SplitPath(Temp[Index], Filename, Path) then
          begin
            Links.RowCount := Max(Links.RowCount, Index + 1);
            Links.Cells[0, Index] := Filename;
            XlLinks.Add(Temp[Index]);
          end;
    Links.FixedRows := 1;
    CheckAllLinks(1,  XlLinks.Count);
  except
    LinkManager.Close;
  end;
end;

function TLinkManager.CheckLink(Link: string; OkReturn: string = 'OK'): string;
var
  Index: integer;
begin
  case XlApp.ActiveWorkbook.LinkInfo(Link, xlLinkInfoStatus, xlLinkTypeExcelLinks, EmptyParam, 0) of
    0:  exit(OkReturn);
    1:  exit('Error: Source not found.');
    2:  exit('Error: Sheet missing.');
    3:  exit('Error: Status may be out of date.');
    4:  exit('Error: Not yet calculated.');
    5:  exit('Error: Unable to determine status.');
    6:  exit('Error: Not started.');
    7:  exit('Error: Invalid name.');
    8:  exit('Error: Not open.');
    9:  exit('Error: Source document is open.');
    10: exit('Error: Copied values.');
  end;
end;

procedure TLinkManager.ChangeLink(Index: Integer);
var
  OpenDialog: TOpenDialog;
begin
  OpenDialog := TOpenDialog.Create(self);
  with OpenDialog do
  begin
    Options := [ofFileMustExist];
    Filter := 'Excel files (*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xlts;*.xls;*.xla;*.xlt;*.xlm;*.xlw)|*.xl*;*.xlsx;*.xlsm;*.xlsb;*.xlam;*.xltx;*.xlts;*.xls;*.xla;*.xlt;*.xlm;*.xlw';
    FilterIndex := 1;
    Title := Links.Cells[0,Index];
  end;
  if OpenDialog.Execute and (OpenDialog.FileName <> Links.Cells[2,Index] + Links.Cells[0,Index]) then
    try
      XlApp.ActiveWorkbook.ChangeLink(Links.Cells[2,Index] + Links.Cells[0,Index], OpenDialog.FileName, xlLinkTypeExcelLinks,0);
      XlLinks[Index-1] := OpenDialog.FileName;
      Links.Cells[0,Index] := StringCut(OpenDialog.FileName,'\',-1);
      Links.Cells[1, Index] := CheckLink(XlLinks[Index-1], 'Link has been changed.');
    except
      MessageDlg('Error, Could not change source.', mtError, [mbOk], 0);
    end;
  LinkManager.SetFocus;
end;

procedure TLinkManager.BtnBreakClick(Sender: TObject);
var
  Currect,Index: integer;
begin
  try
    Currect := Links.Selection.Top;
    for Index := Links.Selection.Top to Links.Selection.Bottom do
      begin
        XlApp.ActiveWorkbook.BreakLink(XlLinks[Currect-1], xlLinkTypeExcelLinks);
        XlLinks.Delete(Currect-1);
        RemoveRow(Currect);
      end;
  finally
    if Links.RowCount = 1 then
      FormCreate(Self.BtnBreak);
  end;
end;

procedure TLinkManager.BtnChangeClick(Sender: TObject);
var
  Index: integer;
begin
  for Index := Links.Selection.Top to Links.Selection.Bottom do
    ChangeLink(Index);
end;

procedure TLinkManager.BtnCheckStatusClick(Sender: TObject);
begin
  CheckAllLinks(Links.Selection.Top, Links.Selection.Bottom);
end;

procedure TLinkManager.CheckAllLinks(StartIndex: integer; EndIndex: integer);
var
  Index: integer;
begin
  try
    XlApp.DisplayAlerts[0] := False;
      for Index := StartIndex to EndIndex do
        begin
          if StartIndex <> EndIndex then
            if ProgressBar.Process(Index-StartIndex+1, EndIndex-StartIndex+1, Links.Cells[0, Index], 'Checking link ' + IntToStr(Index-StartIndex+1) + ' of ' + IntToStr(EndIndex-StartIndex+1)) = false then
              exit;
          Links.Cells[1, Index] := CheckLink(XlLinks[Index-1]);
        end;
    XlApp.DisplayAlerts[0] := True;
  except
  end;
end;

procedure TLinkManager.BtnRelinkClick(Sender: TObject);
var
  Filename: string;
  Index: integer;
  OpenDialog: TFileOpenDialog;
  SelectedFolder: string;
begin
  OpenDialog := TFileOpenDialog.Create(nil);
  try
    OpenDialog.Options := [fdoPickFolders];
    if TDirectory.Exists(GetEnvironmentVariable('USERPROFILE') + '\AppData\Local\CanvasDocHelper\Evidence\') then
      OpenDialog.DefaultFolder := GetEnvironmentVariable('USERPROFILE') + '\AppData\Local\CanvasDocHelper\Evidence\';
    if OpenDialog.Execute then
      begin
        ExcelForeground;
        LinkManager.SetFocus;
        SelectedFolder := OpenDialog.FileName;
        for Index := 1 to (Links.RowCount - 1) do
          try
            for Filename in TDirectory.GetFiles(SelectedFolder, Links.Cells[0,Index], TSearchOption.soAllDirectories) do
              if (Filename <> XlLinks[Index-1]) then
                begin
                  XlApp.ActiveWorkbook.ChangeLink(XlLinks[Index-1], Filename, xlLinkTypeExcelLinks,0);
                  XlLinks[Index-1] := Filename;
                  Links.Cells[1, Index] := CheckLink(XlLinks[Index-1], 'Link has been changed.');
                end;
          except
          end;
      end
    else
      begin
        ExcelForeground;
        LinkManager.SetFocus;
      end;
  finally
    OpenDialog.Free;
    ExcelForeground;
    LinkManager.SetFocus;
  end;
end;

procedure TLinkManager.BtnUpdateClick(Sender: TObject);
var
  Index: integer;
begin
  try
    for Index := Links.Selection.Top to Links.Selection.Bottom do
      begin
        XlApp.ActiveWorkbook.UpdateLink(XlLinks[Index-1], xlLinkTypeExcelLinks, 0);
        Links.Cells[1, Index] := CheckLink(XlLinks[Index-1]);
      end;
  except
  end;
end;

procedure TLinkManager.LinksDblClick(Sender: TObject);
begin
  BtnChangeClick(Self.Links);
end;

procedure TLinkManager.LinksSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  Status.Panels[0].Text := XlLinks[ARow-1];
end;

end.
