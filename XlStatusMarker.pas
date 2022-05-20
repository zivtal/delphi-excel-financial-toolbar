unit XlStatusMarker;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, Registry,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, Excel2010, StringCtrl, XlApplication, XlUtilites;

  procedure StatusMarkerGUI(); stdcall;

var
  XlRange: OleVariant;
  XlFound: OleVariant;
  XlStart: OleVariant;
  XlWorkbook: ExcelWorkbook;
  XlSheet: ExcelWorksheet;

type
  TStatusMarker = class(TForm)
    Progress: TButton;
    Attention: TButton;
    Complete: TButton;
    Clear: TButton;
    EntireSheet: TCheckBox;
    Correct: TButton;
    Incorrect: TButton;
    Note: TMemo;
    SetNote: TCheckBox;
    TabMarking: TCheckBox;
    function ColorCheck(Value: integer): boolean;
    procedure SetTabColor;
    procedure SetColor(Color: OleVariant; Entire, Tab: boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ProgressEnter(Sender: TObject);
    procedure AttentionClick(Sender: TObject);
    procedure CompleteClick(Sender: TObject);
    procedure ClearClick(Sender: TObject);
    procedure ProgressClick(Sender: TObject);
    procedure AttentionEnter(Sender: TObject);
    procedure CompleteEnter(Sender: TObject);
    procedure ClearEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CorrectEnter(Sender: TObject);
    procedure CorrectExit(Sender: TObject);
    procedure IncorrectExit(Sender: TObject);
    procedure IncorrectEnter(Sender: TObject);
    procedure CorrectClick(Sender: TObject);
    procedure IncorrectClick(Sender: TObject);
    procedure SetNoteClick(Sender: TObject);
    procedure NoteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure NoteExit(Sender: TObject);
  public
    procedure ActiveControlChanged(Sender: TObject);
  private
    EntireCheck: boolean;
    EntireEnable: boolean;
    wcActive, wcPrevious : TWinControl;
  end;

var
  StatusMarker: TStatusMarker;

implementation

{$R *.dfm}

const
  Colors: array[1..5] of integer = (65534,49406,14348255,16773300,13158655);
  Registry: string = '\SOFTWARE\Ziv Tal\StatusMarker';

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

procedure StatusMarkerGUI(); stdcall;
begin
  Application.CreateForm(TStatusMarker, StatusMarker);
  try
    StatusMarker.ShowModal;
  finally
    StatusMarker.Destroy;
  end;
end;

procedure TStatusMarker.ActiveControlChanged(Sender: TObject);
begin
  wcPrevious := wcActive;
  wcActive := StatusMarker.ActiveControl;
end;

function TStatusMarker.ColorCheck(Value: integer): boolean;
const
  Support: array[0..7] of integer = (0,65534,49406,14348255,65535,49407,5296273,5296274);
var
  Block: integer;
begin
  for Block in Support do
    if Value = Block then
      exit(true);
  result := false;
end;

procedure TStatusMarker.SetTabColor;
var
  LastCell, Cell: string;
begin
  LastCell := XlSheet.UsedRange[0].Offset[1,1].Address[False, False, xlA1, False, False];
  LastCell := Copy(LastCell, AnsiPos(':', LastCell) + 1, Length(LastCell)-AnsiPos(':', LastCell));
  XlRange := XlSheet.Range['A1', LastCell];
  XlSheet.Tab.ColorIndex := xlColorIndexNone;
  XlApp.FindFormat.Clear;

  XlApp.FindFormat.Interior.Color := 14348255;
  if XlFind(XlRange, '', xlPart, False, True, Cell) then
    XlSheet.Tab.Color := 5296273; // Complete

  XlApp.FindFormat.Interior.Color := 49406;
  if XlFind(XlRange, '', xlPart, False, True, Cell) then
    XlSheet.Tab.Color := 49406; // Attention

  XlApp.FindFormat.Interior.Color := 65534;
  if XlFind(XlRange, '', xlPart, False, True, Cell) then
    if XlSheet.Tab.ColorIndex = 4294963154 then
      XlSheet.Tab.Color := 65534 // In progress
    else
      XlSheet.Tab.Color := 49406; // Attention

  XlApp.FindFormat.Clear;
  RegWrite(Registry, 'Tab', BoolToStr(TabMarking.Checked));
end;

procedure TStatusMarker.SetColor(Color: OleVariant; Entire, Tab: boolean);
var
  LastCell, Cell: string;
  XlColorReplace, XlSelection: OleVariant;
begin
  LastCell := XlSheet.UsedRange[0].Offset[1,1].Address[False, False, xlA1, False, False];
  LastCell := Copy(LastCell, AnsiPos(':', LastCell) + 1, Length(LastCell)-AnsiPos(':', LastCell));
  XlRange := XlSheet.Range['A1', LastCell];
  XlSelection := XlApp.Selection[0];
  XlColorReplace := XlApp.ActiveCell.Interior.Color;
//  XlApp.FindFormat.Clear;
//  XlApp.FindFormat.Interior.Color := XlColorReplace;
  try
    XlSelection.Interior.Color := Color;
    if SetNote.Checked then
      begin
        if Copy(SetNote.Caption, 0, Pos(' ',SetNote.Caption) -1) = 'Change' then
          XlSelection.ClearNotes;
        if not (Note.Text = '') then
          XlSelection.AddComment(Note.Text);
      end
    else
      try
        if (XlSelection.Comment.Text(EmptyParam, EmptyParam, EmptyParam) <> '') and (Color = xlNone) and (MessageDlg('Would you like to clear note(s)?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes) then
          XlSelection.ClearNotes;
      except
      end;
  finally
    if TabMarking.Checked then
      SetTabColor();
    XlApp.FindFormat.Clear;
  end;
end;

procedure TStatusMarker.SetNoteClick(Sender: TObject);
begin
  Note.Enabled := SetNote.Checked;
  if Note.Enabled then
    begin
      StatusMarker.Width := StatusMarker.Width + Note.Width + 24;
      Note.SetFocus;
    end
  else
    StatusMarker.Width := StatusMarker.Width - Note.Width - 24;
end;

procedure TStatusMarker.FormCreate(Sender: TObject);
begin
  try
    TabMarking.Checked := StrToBool(RegRead(Registry, 'Tab'));
  except
  end;
  XlWorkbook := XlApp.ActiveWorkbook;
  XlSheet := XlWorkbook.ActiveSheet as ExcelWorksheet;

  if (XlApp.ActiveCell.Interior.Color = XlNone) or (XlApp.ActiveCell.Interior.Color = 16777215) then
    EntireSheet.Enabled := False;

  if Length(XlSelection(True)) = 1 then
    try
      Note.Text := XlApp.ActiveCell.Comment.Text(EmptyParam, EmptyParam, EmptyParam);
      SetNote.Caption := 'Change &note';
    except
    end
  else
    SetNote.Enabled := False;
end;

procedure TStatusMarker.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Shift = [ssAlt]) then
    case Key of
      48,67:
        ClearClick(Self.Clear);
      49:
        ProgressClick(Self.Progress);
      50:
        AttentionClick(Self.Attention);
      51:
        CompleteClick(Self.Complete);
      52:
        CorrectClick(Self.Correct);
      53:
        IncorrectClick(Self.Incorrect);
    end
  else
    case Key of
      27:
        StatusMarker.Close;
      106,192:
        if EntireSheet.Enabled then
          EntireSheet.Checked := not EntireSheet.Checked;
    end;
end;

procedure TStatusMarker.IncorrectClick(Sender: TObject);
begin
  SetColor(Colors[5],False,False);
  StatusMarker.Close;
end;

procedure TStatusMarker.IncorrectEnter(Sender: TObject);
begin
  StatusMarker.Color := Colors[5];
  with EntireSheet do
  begin
    EntireCheck := Checked;
    EntireEnable := Enabled;
    Checked := False;
    Enabled := False;
  end;
end;

procedure TStatusMarker.IncorrectExit(Sender: TObject);
begin
  with EntireSheet do
  begin
    Checked := EntireCheck;
    Enabled := EntireEnable;
  end;
end;

procedure TStatusMarker.NoteExit(Sender: TObject);
begin
  case StatusMarker.Color of
    65534: Progress.SetFocus;
    49406: Attention.SetFocus;
    14348255: Complete.SetFocus;
    16773300: Correct.SetFocus;
    13158655: Incorrect.SetFocus;
    else
      Clear.SetFocus;
  end;
end;

procedure TStatusMarker.NoteKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    27: exit;
    else
      FormKeyDown(Self.Note, Key, Shift);
  end;

end;

procedure TStatusMarker.ProgressClick(Sender: TObject);
begin
  SetColor(Colors[1], EntireSheet.Checked, True);
  StatusMarker.Close;
end;

procedure TStatusMarker.ProgressEnter(Sender: TObject);
begin
  StatusMarker.Color := Colors[1];
end;

procedure TStatusMarker.AttentionClick(Sender: TObject);
begin
  SetColor(Colors[2], EntireSheet.Checked, True);
  StatusMarker.Close;
end;

procedure TStatusMarker.AttentionEnter(Sender: TObject);
begin
  StatusMarker.Color := Colors[2];
end;

procedure TStatusMarker.CompleteClick(Sender: TObject);
begin
  SetColor(Colors[3], EntireSheet.Checked, True);
  StatusMarker.Close;
end;

procedure TStatusMarker.CompleteEnter(Sender: TObject);
begin
  StatusMarker.Color := Colors[3];
end;

procedure TStatusMarker.CorrectClick(Sender: TObject);
begin
  SetColor(Colors[4], EntireSheet.Checked, False);
  StatusMarker.Close;
end;

procedure TStatusMarker.CorrectEnter(Sender: TObject);
begin
  StatusMarker.Color := Colors[4];
  with EntireSheet do
  begin
    EntireCheck := Checked;
    EntireEnable := Enabled;
    Checked := False;
    Enabled := False;
  end;
end;

procedure TStatusMarker.CorrectExit(Sender: TObject);
begin
  with EntireSheet do
  begin
    Checked := EntireCheck;
    Enabled := EntireEnable;
  end;
end;

procedure TStatusMarker.ClearClick(Sender: TObject);
begin
  SetColor(xlNone, EntireSheet.Checked, False);
  StatusMarker.Close;
end;

procedure TStatusMarker.ClearEnter(Sender: TObject);
begin
  StatusMarker.Color := clWhite;
end;

end.
