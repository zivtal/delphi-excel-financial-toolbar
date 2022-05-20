unit XlProgress;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TProgressBar = class(TForm)
    Status: TLabel;
    Complieted: TLabel;
    Cancel: TButton;
    Uncomplieted: TLabel;
    Precents: TLabel;
    Run: TLabel;
    procedure Update(Status: string = '');
    procedure CancelClick(Sender: TObject);
    function Process(Position, Count: Integer; Status: string = ''; Title: string = 'Progress'): boolean;
  end;

var
  ProgressBar: TProgressBar;
  Process: Boolean;
  RunTime: TDateTime;

implementation

{$R *.dfm}

procedure TProgressBar.Update(Status: string = '');
begin
  ProgressBar.Status.Caption := Status;
  ProgressBar.Run.Caption := TimeToStr(Now - RunTime);
  ProgressBar.Refresh;
end;

function TProgressBar.Process(Position, Count: Integer; Status: string = ''; Title: string = 'Progress'): boolean;
begin
  result := XlProgress.Process;
  if (Position = Count) or (result = False) then
    begin
      XlProgress.Process := True;
      ProgressBar.Hide;
    end
  else
    begin
      if Round(Position / Count * 100) = 0 then
        RunTime := Now;
      Application.ProcessMessages;
      ProgressBar.Show;
      ProgressBar.Caption := Title;
      ProgressBar.Complieted.Width := Round(ProgressBar.Uncomplieted.Width * (Position / Count));
      ProgressBar.Precents.Caption := IntToStr(Round(Position / Count * 100)) + '%';
      Update(Status);
    end;
end;

procedure TProgressBar.CancelClick(Sender: TObject);
begin
  XlProgress.Process := False;
end;

begin
  Application.CreateForm(TProgressBar, ProgressBar);
  XlProgress.Process := True;
end.
