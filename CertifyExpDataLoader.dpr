program CertifyExpDataLoader;

uses
  Vcl.Forms,
  uCertifyExpDataLoader in 'uCertifyExpDataLoader.pas' {ufrmCertifyExpDataLoader},
//  dmCertifyGroupClass_Corporate in 'dmCertifyGroupClass_Corporate.pas' {CertifyGroupClass_Corporate: TDataModule},
  uPushToCertify in 'uPushToCertify.pas' {frmPushToCertify};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TufrmCertifyExpDataLoader, ufrmCertifyExpDataLoader);
//Application.CreateForm(TfrmPushToCertify, frmPushToCertify);
  Application.Run;
end.
