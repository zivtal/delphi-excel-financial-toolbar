program ztxl32kg;

uses
  Vcl.Forms,
  KeyGen in 'KeyGen.pas' {KeyGenerator};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TKeyGenerator, KeyGenerator);
  Application.Run;
end.
