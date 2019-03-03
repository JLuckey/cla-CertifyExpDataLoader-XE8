unit uPushToCertify;       // JSON to Crew Log Record
(*  Dev Notes

  Verb      Base URL   Resource( HTTP params)     Body

*)


(*  to-dos:


*)


interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls

  , IPPeerClient
  , REST.Client
  , Data.Bind.Components
  , Data.Bind.ObjectScope
  , System.JSON
  , REST.Types

  ;

type
  TfrmPushToCertify = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    RESTClient: TRESTClient;
    RESTRequest: TRESTRequest;
    RESTResponse: TRESTResponse;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  private
    FTailNumber: String;
    FUploadStatus: String;
    FCrewMemberVendorNum: String;
    FUploadStatusMessage: String;
    FDataAction: String;
    FHTTPReturnCode: Integer;
    FAPIKey: String;
    FAPISecret: String;
    FDataSetName: String;
    FCertifyDimension: Integer;
    FtheBaseURL: String;
    FTripNumber: String;
    FLogNumber: String;
    Procedure LoadDataStruct();
    Procedure SaveRecord(Const stlDataIn : TStringList);

    Procedure WriteDataStruct;

    Procedure Refresh_TailTripLog_ViaAPI;
    procedure SetHTTPReturnCode(const Value: Integer);
    procedure SetCrewMemberVendorNum(const Value: String);
    procedure SetDataAction(const Value: String);
    procedure SetTailNumber(const Value: String);
    procedure SetUploadStatus(const Value: String);
    procedure SetUploadStatusMessage(const Value: String);


    Procedure DeleteRec;
    Procedure Add_CrewTail_Rec;

    Function  GetIndexForText(Const stlToSearch: TStringList; Const strToFind: String): Integer;
    Function  DecodeFieldVal(Const JSONFieldString : String): String;
    Function  ExtractID(Const strmIn : TMemoryStream) : String;
    Function  GetCertifyRecKey(Const CodeFieldValIn: String; DimenIn: Integer): String;
    procedure SetAPIKey(const Value: String);
    procedure SetAPISecret(const Value: String);

    Procedure Main;
    procedure SetDataSetName(const Value: String);
    procedure SetCertifyDimension(const Value: Integer);
    procedure SetTheBaseURL(const Value: String);
    procedure SetLogNumber(const Value: String);
    procedure SetTripNumber(const Value: String);

  public



  published
    Procedure Push;

    Property DataSetName : String read FDataSetName write SetDataSetName;
    Property DataAction : String read FDataAction write SetDataAction;
    Property TailNumber : String read FTailNumber write SetTailNumber;
    Property CrewMemberVendorNum : String read FCrewMemberVendorNum write SetCrewMemberVendorNum;
    Property TripNumber : String read FTripNumber write SetTripNumber;
    Property LogNumber  : String read FLogNumber write SetLogNumber;


    Property UploadStatus : String read FUploadStatus write SetUploadStatus;
    Property UploadStatusMessage : String read FUploadStatusMessage write SetUploadStatusMessage;
    Property HTTPReturnCode : Integer read FHTTPReturnCode write SetHTTPReturnCode;

    Property theBaseURL : String read FtheBaseURL   write SetTheBaseURL;
    Property APIKey     : String read FAPIKey    write SetAPIKey;
    Property APISecret  : String read FAPISecret write SetAPISecret;

    Property CertifyDimension : Integer read FCertifyDimension write SetCertifyDimension;    // Should this be read-only?

  end;

  type
    TCrewTailRec = record        // Certify Data Structure:  exprptglds/1
      ExpRptGLDIndex : Integer;
      ExpRptGLDLabel : String;
      ID             : String;
      Name           : String;
      Code           : String;
      Description    : String;
      Data           : String;
      Active         : String;
  end;


  type
    TCertifyCrewLogRec = record         // Certify Data Structure:  exprptglds/3
      ExpRptGLDIndex : Integer;
      ExpRptGLDLabel : String;
      ID             : String;
      Name           : String;
      Code           : String;
      Description    : String;
      Data           : String;
      Active         : String;
  end;




var
  frmPushToCertify: TfrmPushToCertify;

  CLR : TCertifyCrewLogRec;
  CLR_Recs : Array of TCertifyCrewLogRec;

  CrewTailRec : TCrewTailRec;
  CrewTailRecs : Array of TCrewTailRec;

  gloMainIdx : Integer;


implementation

{$R *.dfm}

{ TForm5 }


procedure TfrmPushToCertify.FormCreate(Sender: TObject);
begin
  // Set basic comm params
  APIKey    := 'qQjBp9xVQ36b7KPRVmkAf7kXqrDXte4k6PxrFQSv' ;
  APISecret := '4843793A-6326-4F92-86EB-D34070C34CDC' ;

  RESTClient.Params.AddItem('X-Api-Key', APIKey, pkHTTPHEADER);
  RESTClient.Params.AddItem('X-Api-secret', APISecret, pkHTTPHEADER);
  RESTClient.ContentType  := 'application/json';

end;


procedure TfrmPushToCertify.Button1Click(Sender: TObject);
begin

  //  Setting properties that will be set in calling program
  DataSetName         := 'crew_tail';
  TailNumber          := 'N800KS';
  CrewMemberVendorNum := '15213' ;
  theBaseURL          := 'https://api.certify.com/v1/exprptglds';

  DeleteRec;
//  Add_CrewTail_Rec;
end;


procedure TfrmPushToCertify.Main;
begin


end;


procedure TfrmPushToCertify.LoadDataStruct;
var
  FileIn : TextFile;
  s : String;
  stlRecData : TStringList;
  i: Integer;

begin
//  open file
//    readln for '{'
//    read for '}' while accum fields & vals
//      create new rec

  stlRecData := TStringList.Create;
  stlRecData.Sorted := True;
  AssignFile(FileIn, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\API_Interface\crew_log_short.json' ) ;
  Reset(FileIn);

  for i := 1 to 5 do           // skip the first 5 recs of input file
    Readln(FileIn, s);

  gloMainIdx := 0;

  while not Eof(FileIn) do begin

    if Pos('{', s) > 0 then begin
      Readln(FileIn, s);
      while Pos('}', s) = 0 do begin
        stlRecData.Add(Trim(s));
        Readln(FileIn, s);
      end; { while }
      SaveRecord(stlRecData);
      Inc(gloMainIdx);
      stlRecData.Clear;
      Readln(FileIn, s);

    end else begin
      Readln(FileIn, s);

    end;

  end;  { while }

  CloseFile(FileIn);
  stlRecData.Free;

end;  { LoadDataStruct }



procedure TfrmPushToCertify.SaveRecord(const stlDataIn: TStringList);
var
  i : Integer;
  TargetIdx : Integer;

begin

  TargetIdx := GetIndexForText(stlDataIn, 'ExpRptGLDIndex') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].ExpRptGLDIndex := StrToInt(DecodeFieldVal(stlDataIn[TargetIdx]));

  TargetIdx := GetIndexForText(stlDataIn, 'ExpRptGLDLabel') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].ExpRptGLDLabel := DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'ID') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].ID := DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'Name') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].Name := DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'Code') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].Code:= DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'Description') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].Description := DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'Data') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].Data := DecodeFieldVal(stlDataIn[TargetIdx]);

  TargetIdx := GetIndexForText(stlDataIn, 'Active') ;
  if TargetIdx > -1 then
    CLR_Recs[gloMainIdx].Active := DecodeFieldVal(stlDataIn[TargetIdx]);


end;  { SaveRecord() }


function TfrmPushToCertify.DecodeFieldVal(const JSONFieldString: String): String;
var
  FieldSepPos : Integer;                  // Field Separtor Position
  OpenQuotePos : Integer;
  CloseQuotePos : Integer;
  CommaPos      : Integer;

begin
  // Get quoted field values
  FieldSepPos   := Pos(':', JSONFieldString);
  OpenQuotePos  := Pos('"', JSONFieldString, FieldSepPos + 1);       // looking at text after colon (which is data value)
  CloseQuotePos := Pos('"', JSONFieldString, OpenQuotePos + 1);
  CommaPos      := Pos(',', JSONFieldString, CloseQuotePos + 1);

  if OpenQuotePos = 0 then begin
    if CommaPos = 0 then
      Result := Trim(Copy(JSONFieldString, FieldSepPos + 1, 5))      // no trailing comma, ie end of field list in JSON rec
    else
      Result := Trim(Copy(JSONFieldString, FieldSepPos + 1, CommaPos - FieldSepPos - 1 ) );        // the data is not a string surrounded by quotes, just an integer

  end else
    Result := Copy(JSONFieldString, OpenQuotePos + 1, CloseQuotePos - OpenQuotePos - 1); // the data is a quoted string

end;  { DecodeFieldVal }


procedure TfrmPushToCertify.WriteDataStruct;
var
  i : integer;

begin

  for i := 0 to gloMainIdx - 1 do begin
   Memo1.Lines.Add(IntToStr(CLR_Recs[i].ExpRptGLDIndex));
   Memo1.Lines.Add(CLR_Recs[i].ExpRptGLDLabel);
   Memo1.Lines.Add(CLR_Recs[i].ID);
   Memo1.Lines.Add(CLR_Recs[i].Name);
   Memo1.Lines.Add(CLR_Recs[i].Code);
   Memo1.Lines.Add(CLR_Recs[i].Description);
   Memo1.Lines.Add(CLR_Recs[i].Data);
   Memo1.Lines.Add(CLR_Recs[i].Active);
   Memo1.Lines.Add('-----------------');
  end;  {for}

end;



function TfrmPushToCertify.GetIndexForText(const stlToSearch: TStringList; const strToFind: String): Integer;
var
  i : Integer;

begin

  Result := -1;
  for i := 0 to stlToSearch.count - 1 do begin
    if Pos(strToFind, stlToSearch[i]) > 0 then begin             // Make case Insensitive??
      Result := i;
      Break;
    end;
  end;

end;


procedure TfrmPushToCertify.Refresh_TailTripLog_ViaAPI;
begin

//  BuildCrewTailFile;  // generates file
//  BuildCrewTripFile;
//  BuildCrewLogFile;

end;


function TfrmPushToCertify.ExtractID(const strmIn: TMemoryStream): String;
var
  stlCertifyDataStruct : TStringList;

begin
  stlCertifyDataStruct := TStringList.Create;
  stlCertifyDataStruct.LoadFromStream(strmIn);

  ShowMessage(stlCertifyDataStruct.Text);

  Result := 'GUID_string';

  stlCertifyDataStruct.Free;

end;






procedure TfrmPushToCertify.Push;
var
  stsJSON  : TStringStream;
  strmResp : TMemoryStream;
  stlResp  : TStringList;

  RecID :  String;
  GetURL:  String;
  PostURL: String;
  PutURL:  String;

  KeyID : String;

begin

  if FDataAction = 'added' then Begin


  End else if FDataAction = 'deleted' then begin
    DeleteRec;

  end else begin
    // Error

  end;

(*

  Verb, URL, Dimension, Resource, Data/body
  Convert file name to Certify Dimension: crew_tail = 1; crew_trip = 2; crew_log = 3;

  Delete (sets Active field to 0) :
    KeyID := GetCertifyKey(CrewVendNum|TailNum)
    Verb:       POST
    URL & Dim:  https://api.certify.com/v1/exprptglds/1
    Resource:   null
    Body:       { "ID": "00390410-43f2-4004-8d25-1061a752cd50", "Active": 0 }

  Add:


  use Format() function

*)

  RESTClient.BaseURL   := 'https://api.certify.com/v1/exprptglds/1';
  RESTRequest.Resource := '?code=15213|N800KS';
  RESTRequest.Method   := rmGET;

  RESTRequest.Execute;
  ShowMessage('1' + #13 + RESTResponse.JSONValue.ToJSON);



  ShowMessage('2' + #13 + RESTResponse.JSONText);



  strmResp := TMemoryStream.Create;


//  To "delete" a record set Active = 0 using POST (update):
//    1. Find the existing record with GET and concat value in Code field
//    2. Extract its ID
//    3. Create Update JSON using ID & Active = 0
//    4. POST
//    5. Record result of POST


//  IdHTTP_Certify.Post(PostURL,
//                      '{"ID": "' + RecID + '","Active": 0}' ,
//                      strmResp);

  GetURL  := 'https://api.certify.com/v1/exprptglds/1?code='  ;  // <CrewVendorNum>|<TailNumber>';
  PostURL := 'https://api.certify.com/v1/exprptglds/1';
  PutURL  := 'https://api.certify.com/v1/exprptglds/1';

  strmResp.Free;

end;  {Push}


(* ***************************************************

  Verb, URL, Dimension, Resource, Data/body
  Convert file name to Certify Dimension: crew_tail = 1; crew_trip = 2; crew_log = 3;

  Delete (sets Active field to 0) :
    KeyID := GetCertifyKey(CrewVendNum|TailNum)
    Verb:       POST
    URL & Dim:  https://api.certify.com/v1/exprptglds/1
    Resource:   null
    Body:       { "ID": "00390410-43f2-4004-8d25-1061a752cd50", "Active": 0 }


******************************************************* *)
procedure TfrmPushToCertify.DeleteRec;
var
  KeyID : String;
  strBody : String;
  strSearch : String;

begin
  case FCertifyDimension of
    1: strSearch := FCrewMemberVendorNum + '|' + FTailNumber;        // crew_tail
    2: strSearch := FCrewMemberVendorNum + '|' + FTripNumber;        // crew_trip
    3: strSearch := FCrewMemberVendorNum + '|' + FLogNumber;         // crew_log
  end;

  KeyID := GetCertifyRecKey( strSearch, FCertifyDimension );    // ('15213|N800KS', 1);

  RESTClient.BaseURL := FtheBaseURL + '/' + IntToStr(FCertifyDimension) ;      // 'https://api.certify.com/v1/exprptglds/1';
  RESTRequest.Method := rmPOST;
  strBody := Format('{ "ID" : "%s", "Active": 0 }', [KeyID] );                 // Setting "Active" flag to 0 disables record; Certify does not allow us to actually delete record
  ShowMessage(strBody);

  RESTRequest.AddBody( strBody );
  RESTRequest.Params.Items[0].ContentType := ctAPPLICATION_JSON;
  RESTRequest.Execute;

  // Need to set Return properties
  ShowMessage(IntToStr(RESTRequest.Response.StatusCode) + '||' + RESTRequest.Response.StatusText + #13 +
                       'JSONText: ' + RESTRequest.Response.JSONText + #13 +
                       'ErrorMsg: ' + RESTRequest.Response.Content );

end;  { DeleteRec }


procedure TfrmPushToCertify.Add_CrewTail_Rec;
var
  stlBody: TStringList;

begin
  stlBody := TStringList.Create;

  RESTClient.BaseURL := FtheBaseURL + '/' + IntToStr(FCertifyDimension) ;      // 'https://api.certify.com/v1/exprptglds/1';
  RESTRequest.Method := rmPUT;

  ShowMessage(RESTClient.BaseURL);

  // Format JSON data packet
  stlBody.Add('{"ExpRptGLDIndex": ' + IntToStr(FCertifyDimension) + ',' );
  stlBody.Add(' "ExpRptGLDLabel": "Tail #", ');
  stlBody.Add(Format(' "Name": "%s",',     [FTailNumber]));
  stlBody.Add(Format(' "Code": "%s|%s", ', [FCrewMemberVendorNum, FTailNumber] ) );
  stlBody.Add(Format(' "Data": "%s", ',    [FTailNumber ] ) );
  stlBody.Add(' "Active": 1 }');

  ShowMessage(stlBody.Text);

  RESTRequest.AddBody( stlBody.Text );
  RESTRequest.Params.Items[0].ContentType := ctAPPLICATION_JSON;
//  RESTRequest.Execute;

  stlBody.Free;

end;  { Add_Crew_Tail_Rec }



//  get https://api.certify.com/v1/exprptglds/1?code=13748|N113CS
function TfrmPushToCertify.GetCertifyRecKey(const CodeFieldValIn: String; DimenIn: Integer): String;
// DimenIn is Certify ExpRptGLDs Dimension [1,2,3]
Var
  theParam : TRESTRequestParameter;

Begin
  RESTClient.BaseURL := FtheBaseURL + '/' + IntToStr(DimenIn);
  theParam := RESTRequest.Params.AddItem('code', CodeFieldValIn);
  RESTRequest.Method := rmGET;
  RESTRequest.Execute;
  Result := RESTResponse.JSONValue.GetValue<string>('ExpRptGLDs[0].ID');

  RESTRequest.Params.Delete(theParam);  // This param interferes w/ subsequent execution of RESTRequest hence removing it
  // How to handle errors:
  //   1. case where search string is not found
  //   2. case where this code raises an exception  ???JL


  // Return value when string found:
    (*
    {
      "Page": 1,
      "PageCount": 1,
      "Records": 1,
      "RecordCount": 1,
      "ExpRptGLDs": [
          {
              "ExpRptGLDIndex": 1,
              "ExpRptGLDLabel": "Tail #",
              "ID": "006305cd-7bda-4295-b353-96e099712f87",
              "Name": "N955MB",
              "Code": "14967|N955MB",
              "Description": "",
              "Data": "N955MB",
              "Active": 0
          }
      ]
    }
    *)

  // Return value when search string not found:
  {
    "ExpRptGLDs": []
  }

  //  ShowMessage( RESTResponse.JSONText );

end;  { GetCertifyRecKey }



// -------------------  Setters & Getters  ---------------------
procedure TfrmPushToCertify.SetDataSetName(const Value: String);
begin

  // Matching file types w/ Certify Dimension. Which is the integer (i) at end of URL: https://api.certify.com/v1/i  [1,2,3]

  if Value = 'crew_tail' then
    FCertifyDimension := 1
  else if Value = 'crew_trip' then
    FCertifyDimension := 2
  else if Value = 'crew_log' then
    FCertifyDimension := 3;

  FDataSetName := Value;
end;

procedure TfrmPushToCertify.SetAPIKey(const Value: String);
begin
  FAPIKey := Value;
end;

procedure TfrmPushToCertify.SetAPISecret(const Value: String);
begin
  FAPISecret := Value;
end;

procedure TfrmPushToCertify.SetTheBaseURL(const Value: String);
begin
  FtheBaseURL := Value;
end;

procedure TfrmPushToCertify.SetTripNumber(const Value: String);
begin
  FTripNumber := Value;
end;

procedure TfrmPushToCertify.SetCertifyDimension(const Value: Integer);
begin
  FCertifyDimension := Value;
end;

procedure TfrmPushToCertify.SetCrewMemberVendorNum(const Value: String);
begin
  FCrewMemberVendorNum := Value;
end;

procedure TfrmPushToCertify.SetDataAction(const Value: String);
begin
  FDataAction := Value;
end;


procedure TfrmPushToCertify.SetHTTPReturnCode(const Value: Integer);
begin
  FHTTPReturnCode := Value;
end;

procedure TfrmPushToCertify.SetLogNumber(const Value: String);
begin
  FLogNumber := Value;
end;

procedure TfrmPushToCertify.SetTailNumber(const Value: String);
begin
  FTailNumber := Value;
end;

procedure TfrmPushToCertify.SetUploadStatus(const Value: String);
begin
  FUploadStatus := Value;
end;

procedure TfrmPushToCertify.SetUploadStatusMessage(const Value: String);
begin
  FUploadStatusMessage := Value;
end;


end.
