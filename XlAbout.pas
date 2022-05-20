unit XlAbout;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, XlActivation,
  Vcl.Imaging.pngimage, Vcl.ExtCtrls;

  procedure AboutGUI();

type
  TAboutForm = class(TForm)
    Title: TLabel;
    ActivateButton: TButton;
    CloseButton: TButton;
    Logo: TImage;
    procedure FormCreate(Sender: TObject);
    procedure ActivateButtonClick(Sender: TObject);
    procedure CloseButtonClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AboutForm: TAboutForm;
  Desc: string = 'ZTools Add-in for Excel 0.21 Beta'#13#10 + 'Copyright (c) 2019-2021 by Ziv Tal. All Rights Reserved.'#13#10#13#10;

implementation

{$R *.dfm}

// create activation form
procedure AboutGUI();
begin
  Application.CreateForm(TAboutForm, AboutForm);
  try
    AboutForm.ShowModal;
  finally
    AboutForm.Destroy;
  end;
end;

procedure TAboutForm.ActivateButtonClick(Sender: TObject);
begin
  if ActiveNow() then
    FormCreate(Self.ActivateButton);
end;

procedure TAboutForm.CloseButtonClick(Sender: TObject);
begin
  AboutForm.CloseModal;
  AboutForm.Close;
end;

procedure TAboutForm.FormCreate(Sender: TObject);
begin
  Title.Caption := Desc;
  if Activated(false) then
    begin
      Title.Caption := Title.Caption + 'This product is licensed to:'#13#10 + ActName() + ' (' + ActEmail() + ')'#13#10#13#10'Activation will expire in ' + inttostr(ActDays()) + ' days.' ;
      ActivateButton.Visible := false;
    end
  else
    begin
      Title.Caption := Title.Caption + 'UNREGISTERED VERSION'#13#10#13#10;
      ActivateButton.Visible := true;
    end;
  Title.Width := AboutForm.Width - 120;
end;

procedure TAboutForm.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    27: AboutForm.Close;
  end;
end;

end.
