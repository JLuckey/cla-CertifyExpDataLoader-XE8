unit dmCertifyGroupClass_Corporate;

interface

uses
  System.SysUtils, System.Classes;

type
  TCertifyGroupClass_Corporate = class(TDataModule)
  private
    Ffrab: integer;
    FAccountantEMail: String;
    FHasCorpCC: Boolean;
    FApprover2Email: String;
    FApprover1Email: String;
    procedure Setfrab(const Value: integer);
    procedure SetAccountantEMail(const Value: String);
    procedure SetApprover1Email(const Value: String);
    procedure SetApprover2Email(const Value: String);
    procedure SetHasCorpCC(const Value: Boolean);
    Property frab: integer read Ffrab write Setfrab;
    Property AccountantEMail: String read FAccountantEMail write SetAccountantEMail;
    Property Approver1Email: String read FApprover1Email write SetApprover1Email;
    Property Approver2Email: String read FApprover2Email write SetApprover2Email;
    Property HasCorpCC: Boolean read FHasCorpCC write SetHasCorpCC;                                    // Has Corporate Credit Card


    Procedure foop;

    Procedure ProcessRecord;
    Procedure CalcAccountant;

  public




  published






  end;

var
  CertifyGroupClass_Corporate: TCertifyGroupClass_Corporate;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TCertifyGroupClass_Corporate }

procedure TCertifyGroupClass_Corporate.CalcAccountant;
begin
{
  if DataSet.FieldByName('HasCreditCard').AsBoolean = True then
    DataSet.FieldByName('Accountant').AsString := 'CorporateCC@ClayLacy.com'
  else
    DataSet.FieldByName('Accountant').AsString := 'Corporate@ClayLacy.com'

  end;
}

end;


procedure TCertifyGroupClass_Corporate.foop;
begin

end;

procedure TCertifyGroupClass_Corporate.ProcessRecord;
begin
  CalcAccountant;
//  CalcApprover1Email;

//  DataSet.Post
end;


procedure TCertifyGroupClass_Corporate.SetAccountantEMail(const Value: String);
begin
  FAccountantEMail := Value;
end;

procedure TCertifyGroupClass_Corporate.SetApprover1Email(const Value: String);
begin
  FApprover1Email := Value;
end;

procedure TCertifyGroupClass_Corporate.SetApprover2Email(const Value: String);
begin
  FApprover2Email := Value;
end;

procedure TCertifyGroupClass_Corporate.Setfrab(const Value: integer);
begin
  Ffrab := Value;
end;

procedure TCertifyGroupClass_Corporate.SetHasCorpCC(const Value: Boolean);
begin
  FHasCorpCC := Value;
end;

end.
