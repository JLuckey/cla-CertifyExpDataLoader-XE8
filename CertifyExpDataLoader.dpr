program CertifyExpDataLoader;

uses
  Vcl.Forms,
  uCertifyExpDataLoader in 'uCertifyExpDataLoader.pas' {ufrmCertifyExpDataLoader};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TufrmCertifyExpDataLoader, ufrmCertifyExpDataLoader);
  Application.Run;
end.
