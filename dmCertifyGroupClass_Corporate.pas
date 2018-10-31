unit dmCertifyGroupClass_Corporate;

interface

uses
  System.SysUtils, System.Classes;

type
  TCertifyGroupClass_Corporate = class(TDataModule)
  private
    Ffrab: integer;
    procedure Setfrab(const Value: integer);
    Property frab: integer read Ffrab write Setfrab;

    Procedure foop;



  public




  published






  end;

var
  CertifyGroupClass_Corporate: TCertifyGroupClass_Corporate;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TCertifyGroupClass_Corporate }

procedure TCertifyGroupClass_Corporate.foop;
begin

end;

procedure TCertifyGroupClass_Corporate.Setfrab(const Value: integer);
begin
  Ffrab := Value;
end;

end.
