unit uPushToCertify;       // JSON to Crew Log Record

(*  to-dos:


*)


interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP;

type
  TfrmPushToCertify = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    IdHTTP_Avinode: TIdHTTP;
    procedure Button1Click(Sender: TObject);

  private
    Procedure LoadDataStruct();
    Procedure SaveRecord(Const stlDataIn : TStringList);

    Procedure WriteDataStruct;

    Procedure Refresh_TailTripLog_ViaAPI;


    Function  DecodeFieldVal(Const JSONFieldString : String): String;
    Function  GetIndexForText(Const stlToSearch: TStringList; Const strToFind: String): Integer;


  public
    Procedure SendNewCrewTailToCertify(Const stlNewCrewTail: TStringList);

  end;


  type
    TCertifyCrewLogRec = record
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

  gloMainIdx : Integer;



implementation

{$R *.dfm}

{ TForm5 }


procedure TfrmPushToCertify.Button1Click(Sender: TObject);
begin
  SetLength(CLR_Recs, 2000);
  LoadDataStruct;
  WriteDataStruct;
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



procedure TfrmPushToCertify.SendNewCrewTailToCertify(Const stlNewCrewTail: TStringList);
begin



end;



(*

Function TfoAvinodeScheduleUploader.SendDataViaREST : String;
var
  stsJSON  : TStringStream;
  strmResp : TMemoryStream;
  stlResp  : TStringList;

begin
  memSOAPResults.Clear;
  BuildAvinodeRESTHeaders;
  stsJson := TStringStream.Create( memXMLOut.Text );

  strmResp := TMemoryStream.Create;
  stlResp := TStringList.create;
  try
    try
      IdHTTP_Avinode.Put( edURL.Text, stsJson, strmResp );
      memSOAPResults.Lines.Add( IntToStr(IdHTTP_Avinode.ResponseCode) );       // 200
      memSOAPResults.Lines.Add( IdHTTP_Avinode.ResponseText );                 // HTTP/1.1 200 OK

      strmResp.Position := 0;
      stlResp.LoadFromStream(strmResp);
      memSOAPResults.Lines.add(stlResp.Text);
      memSOAPResults.Lines.Add( '' );

    except
      on E: EIdHTTPProtocolException do begin
        memSOAPResults.Lines.Add( 'Error Message: ');
        memSOAPResults.Lines.Add( IntToStr( IdHTTP_Avinode.ResponseCode) );
        memSOAPResults.Lines.Add( E.Message );
        memSOAPResults.Lines.Add( E.ErrorMessage );
        memSOAPResults.Lines.Add( '' );
      end;

      on E: Exception do begin
        memSOAPResults.Lines.Add( 'Unknown Exception from IdHTTP_Avinode: ');
        memSOAPResults.Lines.Add( IntToStr( IdHTTP_Avinode.ResponseCode) );
        memSOAPResults.Lines.Add( E.Message + ' - ' + IdHTTP_Avinode.ResponseText );
        memSOAPResults.Lines.Add( '' );
      end;
    end; { except }

    Result := DecodeRESTRetVal(memSOAPResults.Text);
    LogIt('AVINODE');

  Finally
    stsJSON.free;
    strmResp.Free;
    stlResp.Free;
  End;

end;  { SendDataViaREST }


*)


end.
