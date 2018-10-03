(*

DevNotes:

Lifecycle for recs in CertifyExp_PayComHistory recorded by the record_status & error_text fields:

imported -> OK -> exported
imported -> error               [missing data for required field, duplicate email]
imported -> non-Certify employee


Depricated Minimum LogSheet queries/calcs


To-Do:


12 Sep 2018

  1. Add contractors that are not in Paycom to the Certify Employee file
C 2. Add the "Trips back" count as a paramater to the UI
  3. Review error logging in the Validation-file-generation process
  4. Transmit error list via email
  5. Clean-up obsolete tables
  6. ShowMessage(' and '       + ParamStr(1) );   // command line params


26 Sep 2016
  7. Add .ini
  8. Read command line param at start-up for auto-start
  9.



Notes:

1. "TripNumber" & "QuoteNumber" are use interchangeably in this system  (also "TripNum" & "QuoteNum" & occasionally "QuoteNo")




--  WHERE record_status = 'imported'
--    AND imported_on = '2018-08-16 16:41:03.817'

select work_email, COUNT(* as counter
from CertifyExp_PayComHistory
where record_status = 'non-certify' or record_status = 'imported'
group by work_email
having COUNT(* > 1


select ID, employee_name, work_email
from CertifyExp_PayComHistory
where work_email = 'achaidez@claylacy.com'




update CertifyExp_PayComHistory
set record_status = 'non-certify', status_timestamp = CURRENT_TIMESTAMP
-- from CertifyExp_PayComHistory
where (certify_gp_vendornum is null or certify_gp_vendornum = '')
  and (certify_department is null or certify_department = '' )
  and (certify_role is null or certify_role = '')

*)

unit uCertifyExpDataLoader;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,   System.DateUtils,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, UniProvider, SQLServerUniProvider, Data.DB, MemDS, DBAccess, Uni, Vcl.ComCtrls,
  IdMessage, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase,
  IdSMTP, IdAttachment, IdAttachmentFile, DAScript, UniScript , IniFiles;

type
  TufrmCertifyExpDataLoader = class(TForm)
    UniConnection1: TUniConnection;
    qryGetEmployees: TUniQuery;
    SQLServerUniProvider1: TSQLServerUniProvider;
    edPayComInputFile: TEdit;
    Label1: TLabel;
    btnGenerateFile: TButton;
    StatusBar1: TStatusBar;
    tblPaycomHistory: TUniTable;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    btnTestEmail: TButton;
    Label2: TLabel;
    edOutputFileName: TEdit;
    qryIdentifyNonCertifyRecs: TUniQuery;
    qryGetDBServerTime: TUniQuery;
    btnMain: TButton;
    qryGetDupeEmails: TUniQuery;
    qryUpdateDupeEmailRecStatus: TUniQuery;
    scrLoadTripData: TUniScript;
    qryLoadTripData: TUniQuery;
    qryBuildValFile: TUniQuery;
    edDaysBack: TEdit;
    Label3: TLabel;
    qryGetAirCrewVendorNum: TUniQuery;
    qryEmptyStartBucket: TUniQuery;
    edOutputDirectory: TEdit;
    Label4: TLabel;
    qryGetImportedRecs: TUniQuery;
    qryGetApproverEmail: TUniQuery;
    qryGetTripAccountantRec: TUniQuery;
    scrLoadTripStopData: TUniScript;
    qryGetTripStopRecs: TUniQuery;
    qryGetStartBucketSorted: TUniQuery;
    scrGetCrewLogData: TUniScript;
    qryPilotsNotInPaycom: TUniQuery;
    qryGetPilotsNotInPaycom: TUniQuery;
    qryEmptyPilotsNotInPaycom: TUniQuery;
    qryDeleteTrips: TUniQuery;
    qryContractorsNotInPaycom_Step1: TUniQuery;
    qryContractorsNotInPaycom_Step2: TUniQuery;
    qryDropWorkingTable: TUniQuery;
    qryGetPilotDetails: TUniQuery;
    qryGetEmployeeErrors: TUniQuery;
    procedure btnGenerateFileClick(Sender: TObject);
    procedure btnTestEmailClick(Sender: TObject);
    procedure btnMainClick(Sender: TObject);
    procedure qryLoadTripDataBeforeExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    Procedure Main();
    Procedure ImportPayrollData(Const BatchTimeIn : TDateTime);
    Procedure InsertIntoHistoryTable(Const slInputFileRec: TStringList; BatchTimeIn: TDateTime);
    Procedure BuildEmployeeFile(Const BatchTimeIn: TDateTime)  ;
    Procedure WriteToCertifyEmployeeFile(Const BatchTimeIn: TDateTime)   ;
    Procedure SplitEmployeeName( Const FullNameIn: String; Var LastNameOut, FirstNameOut : String );
    Procedure IdentifyNonCertifyRecs( Const BatchTimeIn : TDateTime );
    Procedure ValidateRecords(Const BatchTimeIn: Tdatetime);
    Procedure UpdateDupeEmailRecs( Const EMailIn: String; BatchTimeIn: TDateTime);
    Procedure BuildCrewLogFile();
    Procedure BuildCrewTailFile();
    Procedure BuildCrewTripFile();
    Procedure BuildTripLogFile();
    Procedure BuildTailTripFile();
    Procedure BuildTailLogFile();

    Procedure BuildGenericValidationFile(const TargetFileName, SQLIn: String) ;
    Procedure BuildGenericValidationFile2(const TargetFileName, SQLIn: String) ;
    Procedure LoadTripsIntoStartBucket;
    Procedure BuildValidationFiles;
    Procedure BuildTripAccountantFile(Const FileNameIn: String);
    Procedure CalculateApproverEmail(Const BatchTimeIn: TDateTime) ;
    Procedure FilterTripsByCount;
    Procedure FindPilotsNotInPaycom(Const BatchTimeIn : TDateTime);
    Procedure DeleteTrip(Const LogSheetIn, CrewMemberIDIn, QuoteNumIn : Integer);

    Procedure AddContractorsNotInPaycom(Const BatchTimeIn: TDateTime);
    Procedure WriteToPaycomTable(Const BatchTimeIn: TDateTime);

    Procedure ConnectToDB();
    Procedure SendStatusEmail;
    Procedure CreateEmployeeErrorReport(Const BatchTimeIn : TDateTime) ;

    Procedure AppendExtraEmployees(Const FileToAppend: TextFile);
    
    Function  GetApproverEmail(Const SupervisorCode: String; BatchTimeIn: TDateTime): String;
    Function  CalcDepartmentName(Const GroupValIn: String): String;
    Function  GetTimeFromDBServer(): TDateTime;
    Function  RecIsValid(Const TimeStampIn:TDateTime): Boolean ;

    Function  CalcPilotName: String;
    Function  CalcDeptDescrip: String;

    Function  CalcPaycomErrorFileName(Const BatchTimeIn: TDateTime): String;
    
  public
    { Public declarations }
  end;

var
  ufrmCertifyExpDataLoader: TufrmCertifyExpDataLoader;
  CertifyEmployeeFile : TextFile;
  CertifyEmployeeFileName : String;
  myIni : TIniFile;


implementation

{$R *.dfm}


procedure TufrmCertifyExpDataLoader.Main;
var
  i: Integer;
  BatchTime : TDateTime;

begin
  BatchTime := GetTimeFromDBServer;

//  ImportPayrollData(BatchTime);               // rec status: imported or error

//  AddContractorsNotInPaycom(BatchTime);  // Add Contractors that are Not-In Paycom
  // Add Tom's two testing recs

(*
  IdentifyNonCertifyRecs(BatchTime);          // rec status: non-certify;     non-certify records flagged in record_status field

  ValidateRecords(BatchTime);                 // rec status: OK
  CalculateApproverEmail(BatchTime);          // rec Status: exported
  BuildEmployeeFile(BatchTime);

  LoadTripsIntoStartBucket;

  FilterTripsByCount;

  BuildValidationFiles;
*)

  CreateEmployeeErrorReport(BatchTime);

(*
  SendStatusEmail;

  StatusBar1.Panels[1].Text := 'Current Task:  All Done!';
  Application.ProcessMessages;
*)
end;  { Main }



procedure TufrmCertifyExpDataLoader.FormCreate(Sender: TObject);
var
  HomeDirectory: String;

begin
  myIni := TIniFile.Create( ExtractFilePath(ParamStr(0)) + 'CertifyExpDataLoader.ini' );

  ConnectToDB;

  edPayComInputFile.Text := myIni.ReadString('Startup', 'PaycomFileName',   '') ;
  edOutputFileName.Text  := myIni.ReadString('Startup', 'CertifyEmployeeFileName', '') ;
  edOutputDirectory.Text := myIni.ReadString('Startup', 'OutputDirectory', '') ;

//  ShowMessage(ParamStr(1));

  if ParamStr(1) = '-autorun' then begin
    StatusBar1.Panels[2].Text := '*AutoRun: Active';
    Self.Show;
    Application.ProcessMessages;
    Sleep(2000);
    Main;
    Sleep(2000);
    Application.Terminate;
  end;



end;


procedure TufrmCertifyExpDataLoader.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  UniConnection1.Close;
  myIni.Free;

end;


procedure TufrmCertifyExpDataLoader.btnMainClick(Sender: TObject);
var
  TargetDirectory : string;

begin
  Main;

end;


procedure TufrmCertifyExpDataLoader.ConnectToDB;
var
  errorMsg: String;
  dbName: ShortString;

begin
  try
    // Connect to Data Warehouse (SQL Server) database
    UniConnection1.Connected := false;
    UniConnection1.Server    := myIni.ReadString('DBConfig_Warehouse', 'Server', '') ;
    UniConnection1.Database  := myIni.ReadString('DBConfig_Warehouse', 'DatabaseName', '') ;
    dbName := UniConnection1.Database;
    UniConnection1.Username  := myIni.ReadString('DBConfig_Warehouse', 'UserName', '') ;
    UniConnection1.Password  := myIni.ReadString('DBConfig_Warehouse', 'Password', '') ;
    UniConnection1.Connected := true;
  except on E: Exception do begin
    errorMsg := 'Error! from ConnectToDBs();  ' +
                'Problem connecting to database: ' + dbName + '; ' +
                E.ClassName + '; ' +
                E.Message ;

//    LogIt(errorMsg);
    MessageDlg(errorMsg, mtError, [mbOK], 0);
    StatusBar1.Panels[0].Text :=  'DB: Error!' ;

    Raise;
  end;
  end;

  StatusBar1.Panels[0].Text :=  'DB: ' + UniConnection1.Database;

end;




procedure TufrmCertifyExpDataLoader.BuildValidationFiles;
var
  TargetDirectory : string;

begin

  TargetDirectory :=  edOutputDirectory.Text;  // 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\Source_XE8\';

  BuildCrewTailFile;
  BuildCrewTripFile;
  BuildCrewLogFile;

  BuildTripLogFile;
  BuildTailTripFile;
  BuildTailLogFile;

  //  BuildGenericValidationFile(TargetDirectory + 'crew_log.csv',
//                             'select distinct LogSheet, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and LogSheet is not null' );
//
//  //  BuildGenericValidationFile(TargetDirectory + 'crew_tail.csv',
////                             'select distinct TailNum as TailNumber, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and TailNumber is not null' );
//
//  BuildGenericValidationFile(TargetDirectory + 'crew_trip.csv',
//                             'select distinct QuoteNum as TripNumber, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where QuoteNum is not null and CrewMemberVendorNum is not null' );
//
//  BuildGenericValidationFile2(TargetDirectory + 'tail_log.csv',
//                             'select distinct TailNum as TailNumber, LogSheet from CertifyExp_Trips_StartBucket where TailNumber is not null and LogSheet is not null' );
//
//  BuildGenericValidationFile2(TargetDirectory + 'tail_trip.csv',
//                             'select distinct TailNum as TailNumber, QuoteNum as TripNumber from CertifyExp_Trips_StartBucket where QuoteNum is not null and TailNumber is not null' );
//
//  BuildGenericValidationFile2(TargetDirectory + 'trip_log.csv',
//                             'select distinct QuoteNum as TripNumber, min( LogSheet ) as LogSheet from CertifyExp_Trips_StartBucket where QuoteNum is not null group by QuoteNum' );
//

  BuildTripAccountantFile(TargetDirectory + 'trip_accountant.csv');

  scrLoadTripStopData.Execute;
  BuildGenericValidationFile2(TargetDirectory + 'trip_stop.csv',
                             'select distinct TripNum, AirportID from CertifyExp_TripStop_Step1' );

end;  { BuildValidationFiles }


procedure TufrmCertifyExpDataLoader.CalculateApproverEmail(Const BatchTimeIn : TDateTime);
var
  ApproverEmail : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Calculating Approver Email ';
  Application.ProcessMessages;

  qryGetImportedRecs.Close;
  qryGetImportedRecs.ParamByName('parmBatchTimeIn').AsDateTime := BatchTimeIn ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn').AsString   := 'OK' ;
  qryGetImportedRecs.Open ;
  while not qryGetImportedRecs.eof do begin
    qryGetImportedRecs.Edit;
    if qryGetImportedRecs.FieldByName('certify_department').AsString = 'Corporate' then begin
      qryGetImportedRecs.FieldByName('accountant_email').AsString := 'QA-AP@ClayLacy.com';

      ApproverEmail := GetApproverEmail(qryGetImportedRecs.FieldByName('supervisor_primary_code').AsString, BatchTimeIn);

      if ApproverEmail = 'error' then begin
        qryGetImportedRecs.FieldByName('record_status').AsString := 'error';
        qryGetImportedRecs.FieldByName('error_text').AsString    := qryGetImportedRecs.FieldByName('error_text').AsString + ' missing supervisor email;';
      end else begin
        qryGetImportedRecs.FieldByName('approver_email').AsString := ApproverEmail;
      end;

      end else begin
        qryGetImportedRecs.FieldByName('accountant_email').AsString := 'QA-135@ClayLacy.com';
    end;

    qryGetImportedRecs.Post;
    qryGetImportedRecs.Next;
  end;  { While }

  qryGetImportedRecs.Close;

end;  { CalculateApproverEmail }



procedure TufrmCertifyExpDataLoader.FilterTripsByCount;
var
  PriorCrewID : String;
  Counter : Integer;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Filtering Trips by Max Count per person';
  Application.ProcessMessages;

  qryGetStartBucketSorted.close;
  qryGetStartBucketSorted.Open;

  Counter := 1;
  PriorCrewID := qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString;
  while not qryGetStartBucketSorted.eof do begin
    if qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString = PriorCrewID then begin
      counter := counter + 1;
      if Counter > 15 then
        DeleteTrip(qryGetStartBucketSorted.FieldByName('LogSheet').AsInteger,
                   qryGetStartBucketSorted.FieldByName('CrewMemberID').AsInteger,
                   qryGetStartBucketSorted.FieldByName('QuoteNum').AsInteger );

    end else begin
      PriorCrewID := qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString;
      Counter := 1;
    end;

    qryGetStartBucketSorted.Next;

  end;  { While }

end;



procedure TufrmCertifyExpDataLoader.btnGenerateFileClick(Sender: TObject);
begin

  ShowMessage(DateTimeToStr(ISO8601ToDate('2018-08-21T16:40:45.723')));

end;


procedure TufrmCertifyExpDataLoader.btnTestEmailClick(Sender: TObject);
Var
  mySMTP    : TIdSMTP;
  myMessage : TIDMessage;
  myAttach  : TIdAttachment;

begin

//  SendStatusEmail;

ShowMessage(DateToStr(StrToDate('00/00/2000')));

(*
  myMessage := TIdMessage.Create(nil);
  myMessage.From.Address := 'NoReply@claylacy.com';
  myMessage.Recipients.EMailAddresses := 'jeff@dcsit.com,jluckey@pacbell.net';
  myMessage.Body.Text := 'fubar ask not why it is fubar, ask who is at fault...' ;
  myMessage.Subject   := 'Test Email from Certify Data Loader 2';

  TIDAttachmentFile.Create(myMessage.MessageParts, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\OutputFiles\crew_log.csv' );
  TIDAttachmentFile.Create(myMessage.MessageParts, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\OutputFiles\trip_log.csv' );

  mySMTP := TIdSMTP.Create(nil);

  mySMTP.Host     := '192.168.1.73';
  mySMTP.Username := 'tkvassay@claylacy.com';                // 'lmirakian@claylacy.com' ;
  mySMTP.Password := '';                                     //'28lalal37';

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
    mySMTP.Disconnect();
  Except on E:Exception Do
    ShowMessage( 'Email Error: ' + E.Message);
  End;

  mySMTP.free;
  myMessage.Free;
*)

end;



procedure TufrmCertifyExpDataLoader.BuildEmployeeFile(Const BatchTimeIn: TDateTime)  ;
var
  slOutRec : TStringList;
  WriteTime : TDateTime;


begin
(*  Get newly-imported records

    work thru records with validation & write to output file
*)
  StatusBar1.Panels[1].Text := 'Current Task:  Writing ' + edOutputFileName.Text ;
  Application.ProcessMessages;

  slOutRec := TStringList.Create;

  // Prep Output File
  AssignFile(CertifyEmployeeFile, edOutputDirectory.Text + edOutputFileName.Text);
  Rewrite(CertifyEmployeeFile);

  // write file header
  WriteLn(CertifyEmployeeFile, 'work_email,first_name,last_name,employee_id,employee_type,group,department_name,approver_email,accountant_email') ;

  qryGetEmployees.Close;
  qryGetEmployees.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn;
  qryGetEmployees.ParamByName('parmRecordStatusIn').AsString := 'OK';
  qryGetEmployees.Open;
  WriteTime := GetTimeFromDBServer();
  while not qryGetEmployees.eof do begin
    WriteToCertifyEmployeeFile(WriteTime);
    qryGetEmployees.Next;
  end;  { While }

  slOutRec.free;                     // put in Try Finallys   ???JL
  CloseFile(CertifyEmployeeFile);

  AppendExtraEmployees(CertifyEmployeeFile);

end;  { BuildEmployeeFile }


procedure TufrmCertifyExpDataLoader.WriteToCertifyEmployeeFile( Const BatchTimeIn: TDateTime )  ;
Var
  slEmpRec : TStringList;          // Employee Record
  LNameOut, FNameOut : String;
  DeptName : String;

begin
  SplitEmployeeName( qryGetEmployees.FieldByName('employee_name').AsString, LNameOut, FNameOut );
  slEmpRec := TStringList.Create;
  try
    slEmpRec.add( qryGetEmployees.FieldByName('work_email').AsString ) ;
    slEmpRec.add( FNameOut ) ;
    slEmpRec.add( LNameOut ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certify_gp_vendornum').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certify_role').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certify_department').AsString ) ;  // aka group

    DeptName := CalcDepartmentName( qryGetEmployees.FieldByName('certify_department').AsString );
    if DeptName <> 'error' then begin
      slEmpRec.add(DeptName);

      slEmpRec.add( qryGetEmployees.FieldByName('approver_email').AsString ) ;
      slEmpRec.add( qryGetEmployees.FieldByName('accountant_email').AsString ) ;

      WriteLn(CertifyEmployeeFile, slEmpRec.CommaText) ;
    end;

    // update status fields
    qryGetEmployees.Edit;
    if DeptName = 'error' then begin
      qryGetEmployees.FieldByName('record_status').AsString := 'error';
      qryGetEmployees.FieldByName('error_text').AsString    := qryGetEmployees.FieldByName('error_text').AsString + ' cannot find: ' + qryGetEmployees.FieldByName('certify_department').AsString + ' in Department_Name lookup; ';
    end else
      qryGetEmployees.FieldByName('record_status').AsString := 'exported';

    qryGetEmployees.FieldByName('status_timestamp').AsDateTime := BatchTimeIn;
    qryGetEmployees.Post;

  finally
    slEmpRec.Free;
  end;

end;  { WriteToCertifyEmployeeFile }



procedure TufrmCertifyExpDataLoader.IdentifyNonCertifyRecs( Const BatchTimeIn : TDateTime );
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Flaggin Non-Certify Records in PaycomHistory ';
  Application.ProcessMessages;

  qryIdentifyNonCertifyRecs.close;
  qryIdentifyNonCertifyRecs.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
  qryIdentifyNonCertifyRecs.Execute;

end;


procedure TufrmCertifyExpDataLoader.ImportPayrollData(Const BatchTimeIn : TDateTime);  // rename to PayCom instead of Payroll ???JL
var
  FileIn: TextFile;
  sl : TStringList;
  s: string;
  i: Integer;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Importing Payroll Data' ;
  Application.ProcessMessages;

  sl := TStringList.Create;
  sl.StrictDelimiter := true;      { tell stringList to not use space as delimeter }

  AssignFile(FileIn, edPayComInputFile.Text) ;
  Reset(FileIn);
  tblPayComHistory.open;

  while not Eof(FileIn) do begin
    Readln(FileIn, s);
    sl.CommaText := s;
    InsertIntoHistoryTable(sl, BatchTimeIn);
  end;

  tblPayComHistory.close;
  CloseFile(FileIn);
  sl.Free;

end;  {PaycomImportMain }


procedure TufrmCertifyExpDataLoader.InsertIntoHistoryTable(const slInputFileRec: TStringList; BatchTimeIn: TDateTime);
var
  recStatus : String;

begin

(*
Paycom file columns:
0  Employee_Code,
1  Employee_Name,
2  Termination_Date,
3  Work_Email,
4  Position,
5  Department_Desc,
6  Job_Code_Desc,
7  Supervisor_Primary_Code,
8  CertifyDepartment,
9  CertifyGPVendor,
10 CertifyRole
*)

  try
    tblPayComHistory.Insert;
    recStatus := 'imported';
    tblPaycomHistory.FieldByName('employee_code').AsString           := slInputFileRec[0];
    tblPaycomHistory.FieldByName('employee_name').AsString           := slInputFileRec[1];

    if slInputFileRec[2] <> '00/00/0000' then
      tblPaycomHistory.FieldByName('termination_date').AsDateTime    := StrToDate(slInputFileRec[2]);


    tblPaycomHistory.FieldByName('work_email').AsString              := slInputFileRec[3];
    tblPaycomHistory.FieldByName('position').AsString                := slInputFileRec[4];
    tblPaycomHistory.FieldByName('department_descrip').AsString      := slInputFileRec[5];
    tblPaycomHistory.FieldByName('job_code_descrip').AsString        := slInputFileRec[6];
    tblPaycomHistory.FieldByName('supervisor_primary_code').AsString := slInputFileRec[7];

    if slInputFileRec[9] <> '' then begin    //
      try
        tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger := StrToInt(slInputFileRec[9]);
      except on E1: Exception do begin
        recStatus := 'error';
        tblPaycomHistory.FieldByName('error_text').AsString    := tblPaycomHistory.FieldByName('error_text').AsString + '; Field: certify_gp_vendornum - ' + E1.Message;
      end;
      end;
    end else begin
//      tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger := '' ;
    end;

    tblPaycomHistory.FieldByName('certify_department').AsString  := slInputFileRec[8];
    tblPaycomHistory.FieldByName('certify_role').AsString        := slInputFileRec[10];
    tblPaycomHistory.FieldByName('record_status').AsString       := recStatus ;
    tblPaycomHistory.FieldByName('status_timestamp').AsDateTime  := BatchTimeIn;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime       := BatchTimeIn;
    tblPaycomHistory.post;

  except on E: Exception do begin
    tblPaycomHistory.Edit;
    tblPaycomHistory.FieldByName('record_status').AsString := 'error';
    tblPaycomHistory.FieldByName('error_text').AsString    := tblPaycomHistory.FieldByName('error_text').AsString + '; ' + E.Message;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime := BatchTimeIn;
    tblPaycomHistory.post;
  end;

  end;  { Try/Except }

end;  { InsertIntoHistoryTable() }




procedure TufrmCertifyExpDataLoader.LoadTripsIntoStartBucket;
begin
(*  1. Empty Start Bucket
    2. Load trips into Start Bucket from Trip tables
    3. Add Vendor number for air crew
*)
  StatusBar1.Panels[1].Text := 'Current Task:  Loading Trips into StartBucket ' ;
  Application.ProcessMessages;

  scrLoadTripData.Execute;
  qryGetAirCrewVendorNum.Execute;

end;  { LoadTripsIntoStartBucket }


procedure TufrmCertifyExpDataLoader.qryLoadTripDataBeforeExecute(Sender: TObject);
begin
  scrLoadTripData.Params.ParamByName('parmDaysBack').AsInteger := StrToInt(edDaysBack.Text);

end;



function TufrmCertifyExpDataLoader.RecIsValid(Const TimeStampIn:TDateTime): Boolean;
var
  strErrorText : String;

begin
  strErrorText := '';
  Result := True;
  if qryGetEmployees.FieldByName('certify_gp_vendornum').AsString = '' then
    strErrorText := strErrorText + 'missing certify_gp_vendornum; ';

  if qryGetEmployees.FieldByName('certify_department').AsString = '' then
    strErrorText := strErrorText + 'missing certify_department; ';

  if qryGetEmployees.FieldByName('certify_role').AsString = '' then
    strErrorText := strErrorText + 'missing certify_role; ';

  qryGetEmployees.Edit;
  qryGetEmployees.FieldByName('status_timestamp').AsDateTime := TimeStampIn;  

  if strErrorText <> '' then begin
    qryGetEmployees.FieldByName('record_status').AsString := 'error';
    qryGetEmployees.FieldByName('error_text').AsString    := qryGetEmployees.FieldByName('error_text').AsString + '; ' + strErrorText;
    Result := False;
  end else begin
    qryGetEmployees.FieldByName('record_status').AsString := 'OK';
    Result := True;
  end;

  qryGetEmployees.Post;

end;  { RecIsValid }



procedure TufrmCertifyExpDataLoader.SplitEmployeeName(const FullNameIn: String; var LastNameOut, FirstNameOut: String);
var
  slFullName : TStringList;

begin
  if Pos(',', FullNameIn) = 0 then begin
    LastNameOut  := FullNameIn;
    FirstNameOut := '';
    Exit;
  end;

  slFullName := TStringList.Create;
  slFullName.StrictDelimiter := True;    // don't use space as delimeter
  try
    slFullName.CommaText := FullNameIn;
    LastNameOut  := Trim(slFullName[0]);
    FirstNameOut := Trim(slFullName[1]);

  finally
    slFullName.Free;
  end;


end;  { SplitEmployeeName }


procedure TufrmCertifyExpDataLoader.ValidateRecords(const BatchTimeIn: Tdatetime);
var
  Time_Stamp :  TDateTime;
  
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Validating Employee Records';
  Application.ProcessMessages;

  Time_Stamp := GetTimeFromDBServer();
  // check for valid Certify recs
  qryGetEmployees.Close;
  qryGetEmployees.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn;
  qryGetEmployees.ParamByName('parmRecordStatusIn').AsString := 'imported';
  qryGetEmployees.Open;

  while not qryGetEmployees.eof do begin
    RecIsValid(Time_Stamp);
    qryGetEmployees.Next;
  end;
  qryGetEmployees.Close;

  //  Check for duplicate work_emails
  qryGetDupeEmails.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
  qryGetDupeEmails.Open;
  while not qryGetDupeEmails.eof do begin
    UpdateDupeEmailRecs( qryGetDupeEmails.FieldByName('work_email').AsString, BatchTimeIn);
    qryGetDupeEmails.Next;
  end;

end;  { ValidateRecords }



procedure TufrmCertifyExpDataLoader.UpdateDupeEmailRecs(const EMailIn: String; BatchTimeIn: TDateTime);
begin

// set record_status = 'error', error_text = :parmErrorTextIn     /* 'dupe work_email foo@bar.com' */
// where work_email  = :parmEmailIn                               /* 'jrosen@claylacy.com' */
//   and imported_on = :parmImportDateIn                          /* '2018-08-22 12:34:18.780' */

  qryUpdateDupeEmailRecStatus.Close;
  qryUpdateDupeEmailRecStatus.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn;
  qryUpdateDupeEmailRecStatus.ParamByName('parmEmailIn').AsString        := EMailIn;
  qryUpdateDupeEmailRecStatus.ParamByName('parmErrorTextIn').AsString    := 'dupe work_email: ' + EMailIn;
  qryUpdateDupeEmailRecStatus.Execute;

end;  { UpdateDupeEmailRecs }



function TufrmCertifyExpDataLoader.GetApproverEmail(const SupervisorCode: String; BatchTimeIn: TDateTime): String;
var
  SC : String;

begin

  // hard-coded special case, per specs - disabled, per revised specs dated 31Aug
//  if Pos('N113CS', qryGetImportedRecs.FieldByName('department_descrip').AsString) > 0 then begin
//    Result := 'rdragoo@claylacy.com';
//    Exit;
//  end;

  SC := SupervisorCode;
  If Length(SC) = 3 then
    SC := '0' + SupervisorCode;

  // use different technique to avoid repeated execution of qryGetApproverEmail;  ???JL
  qryGetApproverEmail.Close;
  qryGetApproverEmail.ParamByName('parmEmpCode').AsString       := SC ;
  qryGetApproverEmail.ParamByName('parmBatchTimeIn').AsDateTime := BatchTimeIn ;
  qryGetApproverEmail.Open;

  if  qryGetApproverEmail.FieldByName('work_email').AsString <> '' then
    result := qryGetApproverEmail.FieldByName('work_email').AsString
  else begin
    result := 'error'
  end;

end;



function TufrmCertifyExpDataLoader.GetTimeFromDBServer: TDateTime;
begin
  qryGetDBServerTime.Close;
  qryGetDBServerTime.Open;
  Result := qryGetDBServerTime.FieldByName('DateTimeOut').AsDateTime;
  qryGetDBServerTime.Close;

end;


procedure TufrmCertifyExpDataLoader.BuildCrewLogFile;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  AssignFile(WorkFile, edOutputDirectory.Text + 'crew_log.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Text := 'select distinct LogSheet, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null order by LogSheet';
  qryBuildValFile.Open ;

(*  Disabled because requrirement for minimum LogSheet per Quote has been depricated  -- 11Sep2018 JL
  qryBuildValFile.SQL.Add( ' select QuoteNum, min(LogSheet) as MinLogSheet ' );
  qryBuildValFile.SQL.Add( ' into #CertifyExp_Work20 '  );
  qryBuildValFile.SQL.Add( ' from CertifyExp_Trips_StartBucket '  );
  qryBuildValFile.SQL.Add( ' where QuoteNum is not null '  );
  qryBuildValFile.SQL.Add( ' group by QuoteNum ' );
  qryBuildValFile.SQL.Add( ' order by QuoteNum ' );

  qryBuildValFile.Execute;

  qryBuildValFile.Close ;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Add( ' select distinct CrewMemberVendorNum, LogSheet ' );
  qryBuildValFile.SQL.Add( ' from CertifyExp_Trips_StartBucket '  );
  qryBuildValFile.SQL.Add( ' where CrewMemberVendorNum is not null '  );
  qryBuildValFile.SQL.Add( '      and LogSheet in ' );
  qryBuildValFile.SQL.Add( '       (select distinct MinLogSheet ' );
  qryBuildValFile.SQL.Add( '        from #CertifyExp_Work20)' );
  qryBuildValFile.Open ;
*)

  RowOut := 'LogSheet,CrewMemberVendorNum';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('LogSheet').AsString) + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + Trim(qryBuildValFile.FieldByName('LogSheet').AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

(*  Disable code per Tom's request 8 Sep 2018
  // Writing Future & Non Trip rows for each Crew Member
  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null' ;
  qryBuildValFile.Open ;

  while not qryBuildValFile.eof do begin
    RowOut := 'Future-Trip,' + qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + 'Future-Trip';
    WriteLn(WorkFile, RowOut) ;
    RowOut := 'Non-Trip,' + qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + 'Non-Trip';
    WriteLn(WorkFile, RowOut) ;

    qryBuildValFile.Next;
  end;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;


//**************************************************************
//
//Var
//  RowOut : String;
//  CrewLogFile : TextFile;
//
//begin
//  AssignFile(CrewLogFile, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\Source_XE8\CrewLog.csv');
//  Rewrite(CrewLogFile);
//
//  qryBuildValFile.Close;
//  qryBuildValFile.SQL.Text := 'select distinct LogSheet, CrewMemberID from CertifyExp_Trips_StartBucket';
//  qryBuildValFile.Open ;
//
//  RowOut := 'LogSheet,CrewMemberID';
//  WriteLn(CrewLogFile, RowOut) ;
//  while not qryBuildValFile.eof do begin
//    RowOut := qryBuildValFile.FieldByName('LogSheet').AsString + ',' +
//              qryBuildValFile.FieldByName('CrewMemberID').AsString + '|' + qryBuildValFile.FieldByName('LogSheet').AsString;
//
//    WriteLn(CrewLogFile, RowOut) ;
//    qryBuildValFile.Next;
//  end;
//
//  CloseFile(CrewLogFile);
//  qryBuildValFile.Close;

end;  { BuildCrewLogFile }



procedure TufrmCertifyExpDataLoader.BuildCrewTailFile;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing crew_tail.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'crew_tail.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct TailNum as TailNumber, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and TailNum is not null and CrewMemberVendorNum > 0' ;
  qryBuildValFile.Open ;

  RowOut := 'TailNumber,CrewMemberID';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNumber').AsString) + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + Trim(qryBuildValFile.FieldByName('TailNumber').AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;


  //Disabled code per Tom's request dated 8 Sep 2018
(*  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null' ;
  qryBuildValFile.Open ;

  while not qryBuildValFile.eof do begin
    RowOut := 'Future-Trip,' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + 'Future-Trip';

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildCrewTailFile }



procedure TufrmCertifyExpDataLoader.BuildCrewTripFile;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing crew_trip.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'crew_trip.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Text := 'select distinct QuoteNum, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and QuoteNum is not null and CrewMemberVendorNum > 0' ;
  qryBuildValFile.Open ;

  RowOut := 'TripNumber,CrewMemberVendorNum';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('QuoteNum').AsString) + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + Trim(qryBuildValFile.FieldByName('QuoteNum').AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

(* disabled per Tom's request dated 8 Sep 2018

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null' ;
  qryBuildValFile.Open ;

  while not qryBuildValFile.eof do begin
    RowOut := 'Future-Trip,' + qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + 'Future-Trip';
    WriteLn(WorkFile, RowOut) ;
    RowOut := 'Non-Trip,' + qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + 'Non-Trip';
    WriteLn(WorkFile, RowOut) ;

    qryBuildValFile.Next;
  end;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;


//***************************************************************************
//
//Var
//  RowOut : String;
//  WorkFile : TextFile;
//
//begin
//  AssignFile(WorkFile, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\Source_XE8\CrewTrip.csv');
//  Rewrite(WorkFile);
//
//  qryBuildValFile.Close;
//  qryBuildValFile.SQL.Text := 'select distinct QuoteNum, CrewMemberID from CertifyExp_Trips_StartBucket where QuoteNum is not null';
//  qryBuildValFile.Open ;
//
//  RowOut := 'TripNumber,CrewMemberID';
//  WriteLn(WorkFile, RowOut) ;
//  while not qryBuildValFile.eof do begin
//    RowOut := qryBuildValFile.FieldByName('QuoteNum').AsString + ',' +
//              qryBuildValFile.FieldByName('CrewMemberID').AsString + '|' + qryBuildValFile.FieldByName('QuoteNum').AsString;
//
//    WriteLn(WorkFile, RowOut) ;
//    qryBuildValFile.Next;
//  end;
//
//  ShowMessage(qryBuildValFile.Fields[0].FieldName);
//
//
//  CloseFile(WorkFile);
//  qryBuildValFile.Close;

end;  { BuildCrewTripFile }


procedure TufrmCertifyExpDataLoader.BuildTripLogFile;
Var
  RowOut : String;
  WorkFile : TextFile;
  strTripNum : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing trip_log.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'trip_log.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close ;
  qryBuildValFile.SQL.Clear;
//  qryBuildValFile.SQL.Add( ' select QuoteNum, min(LogSheet) as LogSheet ' ); Disabled because requrirement for minimum LogSheet per Quote has been depricated  -- 11Sep2018 JL

  qryBuildValFile.SQL.Add( ' select distinct QuoteNum, LogSheet ' ) ;
  qryBuildValFile.SQL.Add( ' from   CertifyExp_Trips_StartBucket '  ) ;
  qryBuildValFile.SQL.Add( ' order by LogSheet ' );
  qryBuildValFile.Open ;

  RowOut := 'LogSheet,TripNumber';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    strTripNum := qryBuildValFile.FieldByName('QuoteNum').AsString;
    if strTripNum = '' then
      strTripNum := 'No-Trip-Num';

    RowOut := Trim(qryBuildValFile.FieldByName('LogSheet').AsString) + ',' + strTripNum ;

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

(* disabled per Tom's request dated 8 Sep 2018
  RowOut := 'Future-Trip,Future-Trip' ;
  WriteLn(WorkFile, RowOut) ;
  RowOut := 'Non-Trip,Non-Trip' ;
  WriteLn(WorkFile, RowOut) ;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;   { BuildTripLogFile }



procedure TufrmCertifyExpDataLoader.BuildTailTripFile;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing trip_tail.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'trip_tail.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct TailNum as TailNumber, QuoteNum as TripNumber from CertifyExp_Trips_StartBucket where QuoteNum is not null and TailNum is not null' ;
  qryBuildValFile.Open ;

  RowOut := 'TailNumber,TripNumber';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNumber').AsString) + ',' + qryBuildValFile.FieldByName('TripNumber').AsString ;
    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

(* disabled per Tom's request dated 8 Sep 2018

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct TailNum from CertifyExp_Trips_StartBucket where QuoteNum is not null' ;
  qryBuildValFile.Open ;

  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',' + 'Future-Trip';
    WriteLn(WorkFile, RowOut) ;
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',' + 'Non-Trip';
    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildTailTripFile }



procedure TufrmCertifyExpDataLoader.BuildGenericValidationFile(const TargetFileName, SQLIn: String);
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  AssignFile(WorkFile, TargetFileName);
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := SQLIn;
  qryBuildValFile.Open ;

  RowOut := qryBuildValFile.Fields[0].FieldName +',' + qryBuildValFile.Fields[1].FieldName ;
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.Fields[0].AsString) + ',' +
              Trim(qryBuildValFile.Fields[1].AsString) + '|' + Trim(qryBuildValFile.Fields[0].AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end; { BuildGenericValidationFile }


procedure TufrmCertifyExpDataLoader.BuildGenericValidationFile2(const TargetFileName, SQLIn: String);
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  AssignFile(WorkFile, TargetFileName);
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := SQLIn;
  qryBuildValFile.Open ;

  RowOut := qryBuildValFile.Fields[0].FieldName +',' + qryBuildValFile.Fields[1].FieldName ;
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.Fields[0].AsString) + ',' +
              Trim(qryBuildValFile.Fields[1].AsString) ;

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildGenericValidationFile2 }



function TufrmCertifyExpDataLoader.CalcDepartmentName(const GroupValIn: String): String;
var
  slGroupDefault_Index : TStringList;
  slGroupDefault_Value : TStringList;
  index : Integer;

begin

  slGroupDefault_Index := TStringList.Create();
  slGroupDefault_Value := TStringList.Create();
  slGroupDefault_Index.CaseSensitive := false;
  slGroupDefault_Value.CaseSensitive := false;
  try
    slGroupDefault_Index.CommaText := '"Corporate","Flight Crew",        "Charter - KVNY", "Charter - KOXC", "Hybrid",             "All" ' ;
    slGroupDefault_Value.CommaText := '"Corporate","Flight Crew - Trip", "Charter - KVNY", "Charter - KOXC", "Flight Crew - Trip", "Corporate" ' ;

    index := slGroupDefault_Index.IndexOf(GroupValIn);

    If index > -1 Then
      Result := slGroupDefault_Value[index]
    else
      Result := 'error';

  finally
    slGroupDefault_Index.Free;
    slGroupDefault_Value.Free;
  end;


end;




procedure TufrmCertifyExpDataLoader.BuildTripAccountantFile(Const FileNameIn: String);
Var
  RowOut : String;
  WorkFile : TextFile;
  AccountantEmail : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing ' + ExtractFileName(FileNameIn) ;
  Application.ProcessMessages;


  AssignFile(WorkFile, FileNameIn);
  Rewrite(WorkFile);

  qryGetTripAccountantRec.Close;
  qryGetTripAccountantRec.Open ;

  RowOut := 'TripNumber,Accountant';
  WriteLn(WorkFile, RowOut) ;
  while not qryGetTripAccountantRec.eof do begin
    if Trim(qryGetTripAccountantRec.Fields[1].AsString) = '91' then
      AccountantEmail := 'QA-91@ClayLacy.com'
    else
      AccountantEmail := 'QA-135@ClayLacy.com';

    RowOut := Trim(qryGetTripAccountantRec.Fields[0].AsString) + ',' + AccountantEmail ;

    WriteLn(WorkFile, RowOut) ;
    qryGetTripAccountantRec.Next;
  end;

  CloseFile(WorkFile);
  qryGetTripAccountantRec.Close;

end;  { BuildTripAccountantFile }



procedure TufrmCertifyExpDataLoader.BuildTailLogFile;
Var
  RowOut : String;
  WorkFile : TextFile;


begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing tail_log.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'tail_log.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Add( ' select distinct TailNum, LogSheet ' );
  qryBuildValFile.SQL.Add( ' from CertifyExp_Trips_StartBucket '  );
  qryBuildValFile.SQL.Add( ' where TailNum is not null '  );
  qryBuildValFile.SQL.Add( ' order by TailNum ' );
//  ShowMessage(qryBuildValFile.SQL.Text);
  qryBuildValFile.Open ;


(*  Disabled because requrirement for "minimum LogSheet per Quote" has been depricated  -- 11Sep2018 JL

  qryBuildValFile.SQL.Add( ' select QuoteNum, min(LogSheet) as MinLogSheet ' );
  qryBuildValFile.SQL.Add( ' into #CertifyExp_Work30 '  );
  qryBuildValFile.SQL.Add( ' from CertifyExp_Trips_StartBucket '  );
  qryBuildValFile.SQL.Add( ' where QuoteNum is not null '  );
  qryBuildValFile.SQL.Add( ' group by QuoteNum ' );
  qryBuildValFile.SQL.Add( ' order by QuoteNum ' );

  qryBuildValFile.Execute;

  qryBuildValFile.Close ;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Add( ' select distinct TailNum, LogSheet ' );
  qryBuildValFile.SQL.Add( ' from CertifyExp_Trips_StartBucket '  );
  qryBuildValFile.SQL.Add( ' where TailNum is not null '  );
  qryBuildValFile.SQL.Add( '      and LogSheet in ' );
  qryBuildValFile.SQL.Add( '       (select distinct MinLogSheet ' );
  qryBuildValFile.SQL.Add( '        from #CertifyExp_Work30)' );
  qryBuildValFile.Open ;
*)

  RowOut := 'TailNumber,LogSheet';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',' + qryBuildValFile.FieldByName('LogSheet').AsString ;
    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;


(* disabled per Tom's request dated 8 Sep 2018

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Text := 'select distinct TailNum from CertifyExp_Trips_StartBucket where TailNum is not null' ;
  qryBuildValFile.Open ;

  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',Future-Trip' ;
    WriteLn(WorkFile, RowOut) ;
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',Non-Trip' ;
    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;
*)

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildTailLogFile }



procedure TufrmCertifyExpDataLoader.FindPilotsNotInPaycom(Const BatchTimeIn : TDateTime);
begin

//  Notice: this flight crew member was not in Paycom but flew trips during this period and was added from the PilotMaster
//  Notice: flight-crew-member was not in Paycom but flew trips during this period and therefore was added from PilotMaster


  qryEmptyPilotsNotInPaycom.Execute;

  qryPilotsNotInPaycom.Close;
  qryPilotsNotInPaycom.paramByName('parmBatchTimeIn').AsDateTime := BatchTimeIn;
  qryPilotsNotInPaycom.Execute;

  qryGetPilotsNotInPaycom.Close;
  qryGetPilotsNotInPaycom.Open;
  while not qryBuildValFile.eof do begin

//    LoadPilotIntoPaycom;
    qryGetPilotsNotInPaycom.Next;
  end;

end;  { FindPilotsNotInPaycom }



procedure TufrmCertifyExpDataLoader.DeleteTrip(const LogSheetIn, CrewMemberIDIn, QuoteNumIn: Integer);
begin
  qryDeleteTrips.Close;
  qryDeleteTrips.ParamByName('parmLogSheetIn').AsInteger     := LogSheetIn;
  qryDeleteTrips.ParamByName('parmCrewMemberIDIn').AsInteger := CrewMemberIDIn;
  qryDeleteTrips.ParamByName('parmQuoteNumIn').AsInteger     := QuoteNumIn;

  try
    qryDeleteTrips.Execute;

  except on E: Exception do
(*
     Memo1.Lines.Add('CrewID: '     + IntToStr(CrewMemberIDIn) +
                     ' LogSheet: '  + IntToStr(LogSheetIn) +
                     ' QuoteNum: '  + IntToStr(QuoteNumIn) +
                     ' ErrorMsg: '  + E.Message) ;
*)
  end;

end;  { DeleteTrips() }



procedure TufrmCertifyExpDataLoader.AddContractorsNotInPaycom(Const BatchTimeIn: TDateTime);
begin

  qryDropWorkingTable.Execute;

  qryContractorsNotInPaycom_Step1.Execute;    //  What date params should be added to these queries, if any  ???JL   when employee/contractor terminated - new scope

  qryContractorsNotInPaycom_Step2.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn ;
  qryContractorsNotInPaycom_Step2.Execute;    //  Contractor Pilot IDs now in #Contractors45 table

  qryGetPilotDetails.close;
  qryGetPilotDetails.open;

  tblPaycomHistory.Open;

  while not qryGetPilotDetails.eof do begin
    WriteToPaycomTable(BatchTimeIn);
    qryGetPilotDetails.Next;
  end ;

  tblPaycomHistory.Close;

end;



procedure TufrmCertifyExpDataLoader.WriteToPaycomTable( Const BatchTimeIn: TDateTime ) ;
begin

  tblPaycomHistory.Insert;
  tblPaycomHistory.FieldByName('employee_code').AsString           := 'contractor-test-85';
  tblPaycomHistory.FieldByName('employee_name').AsString           := CalcPilotName;
  tblPaycomHistory.FieldByName('work_email').AsString              := qryGetPilotDetails.FieldByName('EMail').AsString;
  tblPaycomHistory.FieldByName('position').AsString                := qryGetPilotDetails.FieldByName('JobTitle').AsString;
  tblPaycomHistory.FieldByName('department_descrip').AsString      := CalcDeptDescrip ;
  tblPaycomHistory.FieldByName('job_code_descrip').AsString        := 'Pilot-Designated' ;
  tblPaycomHistory.FieldByName('supervisor_primary_code').AsString := '';
  tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger   := qryGetPilotDetails.FieldByName('VendorNumber').AsInteger;
  tblPaycomHistory.FieldByName('approver_email').AsString          := '';
  tblPaycomHistory.FieldByName('accountant_email').AsString        := 'QA-91@ClayLacy.com';

  tblPaycomHistory.FieldByName('certify_department').AsString      := 'Flight Crew';
  tblPaycomHistory.FieldByName('certify_role').AsString            := 'Contractor';
  tblPaycomHistory.FieldByName('record_status').AsString           := 'imported';
  tblPaycomHistory.FieldByName('status_timestamp').AsDateTime      := BatchTimeIn;
  tblPaycomHistory.FieldByName('imported_on').AsDateTime           := BatchTimeIn;

  tblPaycomHistory.Post;

end;  { WriteToPaycomTable }




function TufrmCertifyExpDataLoader.CalcPilotName: String;
begin
  Result := qryGetPilotDetails.FieldByName('LastName').AsString + ',' + qryGetPilotDetails.FieldByName('FirstName').AsString

end;



function TufrmCertifyExpDataLoader.CalcDeptDescrip: String;
begin
  Result :=  'Designated-' + qryGetPilotDetails.FieldByName('AssignedAC').AsString    ;

//  Result := 'Designated-N1234X';
end;



procedure TufrmCertifyExpDataLoader.SendStatusEmail;
Var
  mySMTP    : TIdSMTP;
  myMessage : TIDMessage;
  myAttach  : TIdAttachment;
  OutPutFileDir : String;
  stlOutputFiles : TStringList;
  i : Integer;

begin

  myMessage := TIdMessage.Create(nil);
  myMessage.Subject      := 'CLA Certify Data Loader Status Report';
  myMessage.From.Address := 'NoReply@claylacy.com';
  myMessage.Body.Text    := 'See attached files for Employee processing errors and uploaded data files:' ;
  myMessage.Recipients.EMailAddresses := myIni.ReadString('OutputFiles', 'EMailRecipientList', '');   //jeff@dcsit.com,jluckey@pacbell.net';   //,thomasfduffy@gmail.com';

  //  Load Attachments
  OutPutFileDir := myIni.ReadString('OutputFiles', 'OutputFileDir', '');
  stlOutputFiles := TStringList.Create();
  stlOutputFiles.CommaText :=myIni.ReadString('OutputFiles', 'EmailAttachFileList', '');
  for i := 0 to stlOutputFiles.Count - 1 do begin
   if FileExists( OutPutFileDir + stlOutputFiles[i] ) then
     TIDAttachmentFile.Create(myMessage.MessageParts, OutPutFileDir + stlOutputFiles[i] );

  end ; { for }
  
  mySMTP := TIdSMTP.Create(nil);
  mySMTP.Host     := '192.168.1.73';
  mySMTP.Username := 'tkvassay@claylacy.com';                // 'lmirakian@claylacy.com' ;
  mySMTP.Password := '';                                     //'28lalal37';

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
//    mySMTP.Disconnect();
  Except on E:Exception Do
    ShowMessage( 'Email Error: ' + E.Message);
  End;

  mySMTP.free;
  myMessage.Free;

end;


procedure TufrmCertifyExpDataLoader.CreateEmployeeErrorReport(Const BatchTimeIn : TDateTime);
Var
  PaycomErrorFile : TextFile;
  i, j : Integer;
  slErrorRec : TStringList;

begin
  slErrorRec := TStringList.Create;

  // Prep Output File
  AssignFile(PaycomErrorFile, edOutputDirectory.Text + CalcPaycomErrorFileName(BatchTimeIn));

  ShowMessage( edOutputDirectory.Text + CalcPaycomErrorFileName(BatchTimeIn) );

  Rewrite(PaycomErrorFile);

  qryGetImportedRecs.Close;
//  qryGetImportedRecs.ParamByName('parmBatchTimeIn').AsDateTime := BatchTimeIn ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn').AsString   := 'error' ;
  qryGetImportedRecs.Open ;

  // Write header record
  for j := 0 to qryGetImportedRecs.FieldCount - 1  do begin
    slErrorRec.Add(qryGetImportedRecs.Fields[j].FieldName) ;
  end;
  Writeln(PaycomErrorFile, slErrorRec.CommaText);
  slErrorRec.Clear;

  while not qryGetImportedRecs.eof do begin
    for i := 0 to qryGetImportedRecs.FieldCount - 1 do begin
      slErrorRec.Add(qryGetImportedRecs.Fields[i].AsString) ;
    end;
    Writeln(PaycomErrorFile, slErrorRec.CommaText);

    slErrorRec.Clear;
    qryGetImportedRecs.Next;
  end;  { While }

  qryGetImportedRecs.Close;
  CloseFile(PaycomErrorFile);
  slErrorRec.Free;

end;  { CreateEmployeeErrorReport }


function TufrmCertifyExpDataLoader.CalcPaycomErrorFileName(const BatchTimeIn: TDateTime): String;
var
  myMonth, myDay, myYear: word;
  strMonth, strDay, strYear : String;

begin
  DecodeDate(Trunc(BatchTimeIn), myYear, myMonth, myday);

  if myMonth < 10 then
    strMonth := '0' + IntToStr(myMonth)
  else
    strMonth := IntToStr(myMonth);

  if myDay < 10 then
    strDay := '0' + IntToStr(myDay)
  else
    strDay := IntToStr(myDay);

  Result := 'PaycomErrors_' + IntToStr(myYear) + strMonth + strDay + '.csv';

end;  { CalcPaycomErrorFileName }



procedure TufrmCertifyExpDataLoader.AppendExtraEmployees(Const FileToAppend: TextFile);
var
  ExtraEmployeeFile : TextFile;
  EmpRec : String;

begin
  AssignFile(ExtraEmployeeFile, 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\InputFiles\certify_employee_extra.csv');
  Reset(ExtraEmployeeFile);
  Append(FileToAppend);

  while not eof(ExtraEmployeeFile) do begin
    ReadLn(ExtraEmployeeFile, EmpRec);
    Writeln(FileToAppend, EmpRec);
  end;

  CloseFile(FileToAppend);
  CloseFile(ExtraEmployeeFile);

end;




end.
