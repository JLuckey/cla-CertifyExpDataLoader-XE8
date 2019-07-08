(*

DevNotes:

Lifecycle for recs in CertifyExp_PayComHistory recorded by the record_status & error_text fields:

imported -> OK -> exported
imported -> error               [missing data for required field, duplicate email]
imported -> non-Certify employee

ShowMessage(DateTimeToStr(ISO8601ToDate('2018-08-21T16:40:45.723')));


Depricated Minimum LogSheet queries/calcs


To-Do:

5 Feb 2019
Possible Refactorings:

1. qryGetImportedRecs & qryGetEmployees appear to be duplicate, can we get rid of one?
2. make ErrorType a Type?





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
  IdSMTP, IdAttachment, IdAttachmentFile, DAScript, UniScript , IniFiles
  , System.UITypes
  , Vcl.Grids
  , Vcl.DBGrids
  , uPushToCertify
  ;



type
  TufrmCertifyExpDataLoader = class(TForm)
    UniConnection1: TUniConnection;
    qryGetEmployees: TUniQuery;
    SQLServerUniProvider1: TSQLServerUniProvider;
    edPayComInputFile: TEdit;
    Label1: TLabel;
    StatusBar1: TStatusBar;
    tblPaycomHistory: TUniTable;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    Label2: TLabel;
    edOutputFileName: TEdit;
    qryIdentifyNonCertifyRecs: TUniQuery;
    qryGetDBServerTime: TUniQuery;
    btnMain: TButton;
    qryGetDupeEmails: TUniQuery;
    qryUpdateDupeEmailRecStatus: TUniQuery;
    qryLoadTripData: TUniQuery;
    qryBuildValFile: TUniQuery;
    edDaysBack: TEdit;
    Label3: TLabel;
    qryGetAirCrewVendorNum: TUniQuery;
    edOutputDirectory: TEdit;
    Label4: TLabel;
    qryGetImportedRecs: TUniQuery;
    qryGetApproverEmail: TUniQuery;
    qryGetTripAccountantRec: TUniQuery;
    scrLoadTripStopData: TUniScript;
    qryGetTripStopRecs: TUniQuery;
    qryGetStartBucketSorted: TUniQuery;
    qryPilotsNotInPaycom: TUniQuery;
    qryGetPilotsNotInPaycom: TUniQuery;
    qryEmptyPilotsNotInPaycom: TUniQuery;
    qryDeleteTrips: TUniQuery;
    qryContractorsNotInPaycom_Step1: TUniQuery;
    qryContractorsNotInPaycom_Step2: TUniQuery;
    qryPurgeWorkingTable: TUniQuery;
    qryGetPilotDetails: TUniQuery;
    qryGetEmployeeErrors: TUniQuery;
    edSpecialUsersFile: TEdit;
    Label5: TLabel;
    qryGetFutureTrips: TUniQuery;
    tblStartBucket: TUniTable;
    qryUpdateHasCCField: TUniQuery;
    qryGetTailTripLog: TUniQuery;
    qryValidateVendorNum: TUniQuery;
    qryFlagMissingFlightCrews: TUniQuery;
    qryFlagTerminatedEmployees: TUniQuery;
    edLastNTrips: TEdit;
    edTerminatedDaysBack: TEdit;
    edContractorDaysBack: TEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    qryInsertCharterVisa: TUniQuery;
    btnGoHourly: TButton;
    edCharterVisaUsers: TEdit;
    Label16: TLabel;
    qryGetDOMEmployees: TUniQuery;
    edTailLeadPilotFile: TEdit;
    Label17: TLabel;
    tblTailLeadPilot: TUniTable;
    btnLoadTailLeadPilotTable: TButton;
    qryFindLeadPilotEmail: TUniQuery;
    qryGetTerminationDate: TUniQuery;
    qryInsertIFS: TUniQuery;
    qryGetIFS: TUniQuery;
//<<<<<<< HEAD
//    DBGrid1: TDBGrid;
//    DataSource1: TDataSource;
//=======
    qryInsertCrewTailHist: TUniQuery;
    qryGetCrewTailBatchDates: TUniQuery;
    qryGetFailedRecs_CrewTail: TUniQuery;
    qryGetDeletedRecs_CrewTail: TUniQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    qryGetCrewTailRecs: TUniQuery;
    qryUpdateRecStatus_CrewTail: TUniQuery;
    edNewDate: TEdit;
    edPreviousDate: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    qryUpdateRecStatus_CrewTrip: TUniQuery;
    qryGetCrewTripRecs: TUniQuery;
    qryInsertCrewTripHist: TUniQuery;
    qryGetCrewTripBatchDates: TUniQuery;
    qryInsertCrewLogHist: TUniQuery;
    qryGetCrewLogBatchDates: TUniQuery;
    qryUpdateRecStatus_CrewLog: TUniQuery;
    qryGetCrewLogRecs: TUniQuery;
    btnFixer: TButton;
    qryPruneHistoryTables: TUniQuery;
    qryFindVendorNumInStartBucket: TUniQuery;
    tblTripStop: TUniTable;
    qryGetNewHireRecs: TUniQuery;
    qryGetFlightCrewNewHire: TUniQuery;

    procedure btnMainClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnGoHourlyClick(Sender: TObject);
    procedure btnLoadTailLeadPilotTableClick(Sender: TObject);
    procedure btnFixerClick(Sender: TObject);

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
    Procedure BuildCrewTailFile(Const CurrBatchTimeIn: TDatetime);
    Procedure BuildCrewTripFile();
    Procedure BuildTripLogFile();
    Procedure BuildTailTripFile();
    Procedure BuildTailLogFile();

    Procedure BuildGenericValidationFile(const TargetFileName, SQLIn: String) ;
    Procedure BuildGenericValidationFile2(const TargetFileName, SQLIn: String) ;
    Procedure LoadTripsIntoStartBucket;
    Procedure BuildValidationFiles(Const BatchTimeIn : TDateTime);
    Procedure BuildTripAccountantFile(Const FileNameIn: String);
    Procedure CalculateApproverEmail(Const BatchTimeIn: TDateTime) ;
    Procedure FilterTripsByCount;
    Procedure FindPilotsNotInPaycom(Const BatchTimeIn : TDateTime);               // de-cruft, appears not to be called  ???JL  3 dec 2018
    Procedure DeleteTrip(Const LogSheetIn, QuoteNumIn : Integer; CrewMemberIDIn: String);

    Procedure AddContractorsNotInPaycom(Const BatchTimeIn: TDateTime);
    Procedure WriteContractorsToPaycomTable(Const BatchTimeIn: TDateTime; NewHireFlag: Boolean);

    Procedure ConnectToDB();
    Procedure SendStatusEmail;
    Procedure CreateEmployeeErrorReport(Const BatchTimeIn : TDateTime) ;

    Procedure AppendSpecialUsers(Const FileToAppend: TextFile);

    Procedure UpdateCCField(Const BatchTimeIn: TDateTime);    // rename CC to CreditCard here & query  - ???JL

    Procedure LoadCertFileFields(Const BatchTime: TDateTime);

    Procedure BuildTailTripLogFile;


    Procedure Load_CharterVisa_IntoStartBucket;
    Procedure Load_DOM_IntoStartBucket(Const BatchTimeIn: TDatetime);
    Procedure InsertCrewTail(Const TailNumIn:String; VendorNumIn: Integer);

    Procedure LoadTailLeadPilot;
    Procedure InsertTailLeadPilot(Const TailNumIn, EMailIn: String);
    Procedure FlagRecordAsError(Const ErrorType, ErrorMsgIn : String);

    Procedure Load_IFS_IntoStartBucket(BatchTime: TDateTime);

    Procedure HourlyPushMain;

    Procedure LoadData(Const BatchTimeIn: TDateTime);

    Procedure SendNewCrewTailToCertify(stlNewCrewTailIn: TStringList);


    Procedure SendToCertify(Const WorkingQueryIn: TUniQuery; Const BatchTimeIn : TDateTime; DataSetNameIn: String);

    Procedure GetBatchDates_CrewTail(Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);
    Procedure GetBatchDates_CrewTrip(Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);
    Procedure GetBatchDates_CrewLog( Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);

    Procedure UpdateRecordStatus_CrewTail(Const RecStatusIn: String; Const NewBatchDateIn, PreviousBatchDateIn: TDateTime);

    Procedure LogIt(ErrorMsgIn: String);

    Procedure Do_CrewTail_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);
    Procedure Do_CrewTrip_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);
    Procedure Do_CrewLog_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);

    Procedure UpdateRecordStatus_CrewTrip(const RecStatusIn: String; const NewBatchDateIn, PreviousBatchDateIn: TDateTime);
    Procedure UpdateRecordStatus_CrewLog(const RecStatusIn: String; const NewBatchDateIn, PreviousBatchDateIn: TDateTime);

    Procedure Load_CrewTail_HistoryTable(Const BatchTimeIn: TDateTime);
    Procedure Load_CrewTrip_HistoryTable(Const BatchTimeIn: TDateTime);
    Procedure Load_CrewLog_HistoryTable(const BatchTimeIn: TDateTime);

    Procedure PruneHistoryTables(Const TableNameIn : String);

    // 7 Jun 2019
    Procedure Process_NewHire_Contractors(Const BatchTimeIn: TDateTime);
    Procedure CalcSQLForQryGetPilotDetails(Const QryKind : String);
    Procedure AddDummyTripToStartBucket(Const VendorNumIn: Integer; NewHireKind: String) ;
    Procedure InsertDummyNewHireTripStop(Const QuoteNumIn : Integer);
    Procedure Process_NewHire_Employees_FlightCrew(Const BatchTimeIn: TDateTime);


    Function  GetApproverEmail(Const SupervisorCode: String; BatchTimeIn: TDateTime): String;
    Function  CalcDepartmentName(Const GroupValIn: String): String;
    Function  GetTimeFromDBServer(): TDateTime;
    Function  RecIsValid(Const TimeStampIn:TDateTime): Boolean ;

    Function  CalcPilotName: String;
    Function  CalcDeptDescrip: String;
    Function  ScrubFARPart(Const FarPartIn: String): String;
    Function  ScrubCertifyDept(Const DepartmentIn : String) : String;
    Function  CalcPaycomErrorFileName(Const BatchTimeIn: TDateTime): String;


    Function CalcCertfileDepartmentName(Const GroupNameIn: String): String;

    Function FindLeadPilot(Const AssignedACString: String; BatchTimeIn: TDateTime): String;
    Function EmployeeTerminated(Const EMailIn: String; ImportDateIn: TDateTime): Boolean;

    Function GroupIsIn(Const GroupIn, GroupSetIn: String): Boolean;

    Function CalcCrewTailFileName(Const BatchTimeIn: TDateTime): String;

    // 7 Jun 2019
    Function FindVendorNumInStartBucket(Const VendorNumIn: Integer): Boolean;

  public
    { Public declarations }
  end;

var
  ufrmCertifyExpDataLoader: TufrmCertifyExpDataLoader;
  CertifyEmployeeFile : TextFile;
  CertifyEmployeeFileName : String;
  myIni : TIniFile;
  myLog : TIniFile;
  gloPaycomErrorFile: String;

  gloPusher : TfrmPushToCertify;
  gloNewHireDummyQuoteNum : Integer;

implementation

{$R *.dfm}

procedure TufrmCertifyExpDataLoader.FormCreate(Sender: TObject);
begin
  myIni := TIniFile.Create( ExtractFilePath(ParamStr(0)) + 'CertifyExpDataLoader.ini' );

  ConnectToDB;

  edPayComInputFile.Text  := myIni.ReadString('Startup', 'PaycomFileName',   '') ;
  edOutputFileName.Text   := myIni.ReadString('Startup', 'CertifyEmployeeFileName', '') ;
  edOutputDirectory.Text  := myIni.ReadString('Startup', 'OutputDirectory', '') ;
  edSpecialUsersFile.Text := myIni.ReadString('Startup', 'SpecialUsersFileName', '') ;

  //  ShowMessage(ParamStr(1));
//  if ParamStr(1) = '-autorun' then begin
  if ParamStr(1) = '-hourly' then begin
    StatusBar1.Panels[2].Text := '*AutoRun: Hourly, Active';
    Self.Show;
    Application.ProcessMessages;
    HourlyPushMain;
    Application.Terminate;

  end else if ParamStr(1) = '-nightly' then begin
    StatusBar1.Panels[2].Text := '*AutoRun: Nightly, Active';
    Self.Show;
    Application.ProcessMessages;
    Main;
    Application.Terminate;

  end;

end;    {FormCreate}



procedure TufrmCertifyExpDataLoader.Main;
var
  BatchTime : TDateTime;

begin
  BatchTime := GetTimeFromDBServer;

  LoadData(BatchTime);

  BuildEmployeeFile(BatchTime);

  BuildValidationFiles(BatchTime);

  CreateEmployeeErrorReport(BatchTime);

  SendStatusEmail;

  StatusBar1.Panels[1].Text := 'Current Task:  All Done!';
  Application.ProcessMessages;


end;  { Main }


(*
  1. Upload added recs
  2. Upload deleted recs
*)
procedure TufrmCertifyExpDataLoader.HourlyPushMain;
var
  BatchTime : TDateTime;
  stlNewCrewTail : TStringList;
  PreviousBatchDate, NewBatchDate : TDateTime;

begin
  BatchTime := GetTimeFromDBServer();          //StrToDateTime('03/25/2019 22:15:00');
  LoadData(BatchTime);                         // loads data into StartBucket & PaycomHistory tables

  gloPusher := TfrmPushToCertify.Create(self);
  gloPusher.theBaseURL := 'https://api.certify.com/v1/exprptglds';
  gloPusher.APIKey     := 'qQjBp9xVQ36b7KPRVmkAf7kXqrDXte4k6PxrFQSv' ;
  gloPusher.APISecret  := '4843793A-6326-4F92-86EB-D34070C34CDC' ;

  // crew_tail
  PruneHistoryTables('CrewTail');
  Load_CrewTail_HistoryTable(BatchTime);                       // puts latest batch into CrewTailHistory table
  GetBatchDates_CrewTail(PreviousBatchDate, NewBatchDate);     // identifies "added" & "deleted" recs; (those params are output)
  Do_CrewTail_API(Now(), PreviousBatchDate, NewBatchDate);     // sends data to Certify via API

  // crew_trip
  PruneHistoryTables('CrewTrip');
  Load_CrewTrip_HistoryTable(BatchTime);
  GetBatchDates_CrewTrip(PreviousBatchDate, NewBatchDate);
  Do_CrewTrip_API(Now(), PreviousBatchDate, NewBatchDate);

  // crew_log
  PruneHistoryTables('CrewLog');
  Load_CrewLog_HistoryTable(BatchTime);
  GetBatchDates_CrewLog(PreviousBatchDate, NewBatchDate);
  Do_CrewLog_API(Now(), PreviousBatchDate, NewBatchDate);

  gloPusher.free ;
  StatusBar1.Panels[1].Text := 'Current Task:  All done!'  ;

//  //  Get Failed Recs from previous Pushes                     ???JL unnecessary, remove
//  qryGetFailedRecs_CrewTail.Close;
//  qryGetFailedRecs_CrewTail.ParamByName('parmNewBatchDateIn').AsDateTime := NewBatchDate;
//  qryGetFailedRecs_CrewTail.Open;
//  SendToCertify(qryGetFailedRecs_CrewTail, BatchTime);
//  qryGetFailedRecs_CrewTail.Close;

end;  { HourlyPushMain }


procedure TufrmCertifyExpDataLoader.LoadData(Const BatchTimeIn: TDateTime);
begin

  ImportPayrollData(BatchTimeIn);               // rec status: imported or error

  LoadTripsIntoStartBucket;

    Load_CharterVisa_IntoStartBucket;

(*    Load_DOM_IntoStartBucket(BatchTimeIn); *) // disabled per Phase 3 Tasks & Estimates item # 14;   14 May 2019  Jeff@dcsit.com

  AddContractorsNotInPaycom(BatchTimeIn);       // writes Contractors to PaycomHistory table

    Process_NewHire_Contractors(BatchTimeIn);     // writes New-Hire contractors to PaycomHistory, if they have not yet flown any trips
    Process_NewHire_Employees_FlightCrew(BatchTimeIn);

  UpdateCCField(BatchTimeIn);                   // Update Credit Card Field

  IdentifyNonCertifyRecs(BatchTimeIn);          // rec status: non-certify;     non-certify records flagged in record_status field

  ValidateRecords(BatchTimeIn);                 // rec status: OK

  LoadCertFileFields(BatchTimeIn);

    Load_IFS_IntoStartBucket(BatchTimeIn);

  FilterTripsByCount;                           // removes selected recs from StartBucket


end;  { LoadData() }


procedure TufrmCertifyExpDataLoader.btnFixerClick(Sender: TObject);
begin

{
  gloPusher := TfrmPushToCertify.Create(self);
  gloPusher.theBaseURL := 'https://api.certify.com/v1/exprptglds';
  gloPusher.APIKey     := 'qQjBp9xVQ36b7KPRVmkAf7kXqrDXte4k6PxrFQSv' ;
  gloPusher.APISecret  := '4843793A-6326-4F92-86EB-D34070C34CDC' ;

  (*
  qryGetCrewTripRecs.Close;
  qryGetCrewTripRecs.ParamByName('parmCreatedOnIn').AsString := '2019-04-05 07:42:00.823';
//  qryGetCrewTripRecs.ParamByName('parmRecStatusIn').AsString := 'added';
  qryGetCrewTripRecs.Open;

  DataSource1.DataSet := qryGetCrewTripRecs;          // Update data grid in UI for test/dev purposes
  Application.ProcessMessages;
  Sleep(5000);                                        // wait so user can see data in grid

  SendToCertify(qryGetCrewTripRecs, Now(), 'crew_trip');
  qryGetCrewTripRecs.Close;
*)

  qryGetCrewLogRecs.Close;
  qryGetCrewLogRecs.ParamByName('parmCreatedOnIn').AsString := '2019-04-05 07:42:00.823';
//  qryGetCrewTripRecs.ParamByName('parmRecStatusIn').AsString := 'added';
  qryGetCrewLogRecs.Open;

  DataSource1.DataSet := qryGetCrewLogRecs;          // Update data grid in UI for test/dev purposes
  Application.ProcessMessages;
  Sleep(5000);                                        // wait so user can see data in grid

  SendToCertify(qryGetCrewLogRecs, Now(), 'crew_log');
  qryGetCrewLogRecs.Close;

  gloPusher.free;
}

end;




procedure TufrmCertifyExpDataLoader.btnGoHourlyClick(Sender: TObject);
begin

  HourlyPushMain;

end;


procedure TufrmCertifyExpDataLoader.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  try
    UniConnection1.Close;              //  ???JL  this still may be cause the program to hang on exist under some unknown conditions  14 may 2019

  except

  end;

  myIni.Free;

end;


procedure TufrmCertifyExpDataLoader.btnMainClick(Sender: TObject);

begin
  Main;

end;


procedure TufrmCertifyExpDataLoader.ConnectToDB;
var
  errorMsg: String;
  dbName: String;

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


procedure TufrmCertifyExpDataLoader.BuildValidationFiles(Const BatchTimeIn : TDateTime);
var
  TargetDirectory : string;

begin
  TargetDirectory :=  edOutputDirectory.Text;  // 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\Source_XE8\';

  BuildCrewTailFile(BatchTimeIn);
  BuildCrewTripFile;
  BuildCrewLogFile;

  BuildTripLogFile;
  BuildTailTripFile;
  BuildTailLogFile;

  BuildTailTripLogFile;

  BuildTripAccountantFile(TargetDirectory + 'trip_accountant.csv');

  // Build Trip/Stop records
  scrLoadTripStopData.Execute;    // puts recs into working table CertifyExp_TripStop_Step1

    InsertDummyNewHireTripStop(818181);   // adds NewHire records to working table CertifyExp_TripStop_Step1
    InsertDummyNewHireTripStop(828282);   // adds NewHire records to working table CertifyExp_TripStop_Step1

  BuildGenericValidationFile2(TargetDirectory + 'trip_stop.csv',
                             'select distinct TripNum, AirportID from CertifyExp_TripStop_Step1 where AirportID is not null' );

end;  { BuildValidationFiles }


procedure TufrmCertifyExpDataLoader.InsertDummyNewHireTripStop(Const QuoteNumIn : Integer);
begin

  tblTripStop.Open;
  tblTripStop.Insert;
  tblTripStop.FieldByName('TripNum').AsInteger  := QuoteNumIn;
  tblTripStop.FieldByName('AirportID').AsString := 'KVNY';
  tblTripStop.Post;
  tblTripStop.Insert;
  tblTripStop.FieldByName('TripNum').AsInteger  := QuoteNumIn;
  tblTripStop.FieldByName('AirportID').AsString := 'KTEB';
  tblTripStop.Post;

  tblTripStop.Close;

end;  { InsertDummyNewHireTripStop }





procedure TufrmCertifyExpDataLoader.LoadCertFileFields(const BatchTime: TDateTime);
Var
  LNameOut, FNameOut : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Loading Certify fields in PaycomHistory ';
  Application.ProcessMessages;

  qryGetImportedRecs.Close;
  qryGetImportedRecs.ParamByName('parmBatchTimeIn').AsDateTime := BatchTime ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn').AsString   := 'OK' ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn2').AsString  := 'warning' ;

  qryGetImportedRecs.Open ;
  while not qryGetImportedRecs.eof do begin
    qryGetImportedRecs.Edit;

    SplitEmployeeName( qryGetImportedRecs.FieldByName('employee_name').AsString, LNameOut, FNameOut );
    with qryGetImportedRecs do begin
      FieldByName('certfile_work_email').AsString      := FieldByName('work_email').AsString;
      FieldByName('certfile_last_name').AsString       := LNameOut;
      FieldByName('certfile_first_name').AsString      := FNameOut;
      FieldByName('certfile_employee_id').AsString     := FieldByName('certify_gp_vendornum').AsString;
      FieldByName('certfile_employee_type').AsString   := FieldByName('certify_role').AsString;
      FieldByName('certfile_group').AsString           := ScrubCertifyDept(FieldByName('certify_department').AsString);
      FieldByName('certfile_department_name').AsString := CalcCertfileDepartmentName( FieldByName('certfile_group').AsString );                                  // - depricated FieldByName('department_descrip').AsString;
    end;

    CalculateApproverEmail(BatchTime);

    qryGetImportedRecs.Post;
    qryGetImportedRecs.Next;
//    application.ProcessMessages;
  end;

  qryGetImportedRecs.Close;

end; { LoadCertFileFields }



procedure TufrmCertifyExpDataLoader.CalculateApproverEmail(Const BatchTimeIn : TDateTime);
var
  strAccountantEmail : String;
  strCertifyGroup : String;
  strAssignedAC : String;
  PaycomApprover1, PaycomApprover2: String;
  PilotEmail : String;

begin
    strCertifyGroup := qryGetImportedRecs.FieldByName('certfile_group').AsString ;

    if GroupIsIn(strCertifyGroup, '|Corporate|FBO|Maintenance|Marketing|AircraftManagement|HR|') then begin

      // Assign Accountant Email
      if qryGetImportedRecs.FieldByName('has_credit_card').AsString = 'T' then
        strAccountantEmail := 'CorporateCC@ClayLacy.com'
      else
        strAccountantEmail := 'Corporate@ClayLacy.com';

      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := strAccountantEmail;


      // Assign Approver1 Email
      if qryGetImportedRecs.FieldByName('paycom_approver1_email').AsString <> '' then
        qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := qryGetImportedRecs.FieldByName('paycom_approver1_email').AsString
      else begin
        FlagRecordAsError('error', 'error-approver1_email_not_provided ' + strCertifyGroup);
      end;


      // Assign Approver2 Email
      if qryGetImportedRecs.FieldByName('paycom_approver2_email').AsString <> '' then
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := qryGetImportedRecs.FieldByName('paycom_approver2_email').AsString
      else
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := strAccountantEmail;


    end else if GroupIsIn(strCertifyGroup, '|DOM|') then begin


      strAssignedAC   := qryGetImportedRecs.FieldByName('paycom_assigned_ac').AsString;
      PaycomApprover1 := qryGetImportedRecs.FieldByName('paycom_approver1_email').AsString;
      PaycomApprover2 := qryGetImportedRecs.FieldByName('paycom_approver2_email').AsString;

      // Assign Accountant Email
      if UpperCase(qryGetImportedRecs.FieldByName('has_credit_card').AsString) = 'T' then
        strAccountantEmail := 'FlightCrewCC@ClayLacy.com'
      else
        strAccountantEmail := 'FlightCrew@ClayLacy.com';

      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := strAccountantEmail;


      //  Assign Approver1
      if PaycomApprover1 = '' then
        FlagRecordAsError('error', 'paycom_approver1_email is blank')

      else if UpperCase(PaycomApprover1) = 'LEADPILOT' then begin
        PilotEmail := FindLeadPilot(strAssignedAC, BatchTimeIn);
        if PilotEmail = 'NotFound' then
          qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := strAccountantEmail
        else
          qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := PilotEmail ;

      end else   // Contains a email address
        qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := PaycomApprover1;


      //  Assign Approver2
      if PaycomApprover2 = '' then
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := strAccountantEmail   // qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString

      else if UpperCase(PaycomApprover2) = 'LEADPILOT' then begin
        PilotEmail := FindLeadPilot(strAssignedAC, BatchTimeIn);
        if PilotEmail = 'NotFound' then
          qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := strAccountantEmail
        else
          qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := PilotEmail;

      end else  // Contains an email address
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := PaycomApprover2 ;


    end else if GroupIsIn(strCertifyGroup, '|FlightCrew|PoolPilot|PoolFA|FlightCrewCorp|FlightCrewNonPCal|IFS|') then begin


      // Assign Accountant Email
      if qryGetImportedRecs.FieldByName('has_credit_card').AsString = 'T' then
        strAccountantEmail := 'FlightCrewCC@ClayLacy.com'
      else
        strAccountantEmail := 'FlightCrew@ClayLacy.com';

      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := strAccountantEmail;

(* 15May2019 - added per "Phase3 Tasks & Estimates" item #21 *)
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString  := strAccountantEmail;
      qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString  := strAccountantEmail;



(*  4Feb2019 -JL
    Note 13: This logic is now handled by Certify "Workflows" within the Certify system
      //  Assign Approver1 Email
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := strAccountantEmail;

      //  Assign Approver2 Email
      if qryGetImportedRecs.FieldByName('employee_code').AsString = 'contractor'  then
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := strAccountantEmail
      else
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := '';
*)


    end else if GroupIsIn(strCertifyGroup, '|CharterVisa|') then begin


      // Assign Accountant Email
      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := 'FlightCrewCC@ClayLacy.com';

(* See Note 13
      //  Assign Approver1 Email
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := 'CorporateCC@ClayLacy.com';

      //  Assign Approver2 Email
      qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := 'CorporateCC@ClayLacy.com';
*)


    end else begin       // error, unknown Certify Department


      FlagRecordAsError('error', 'unknown certify_department/Group: ' + strCertifyGroup);


    end;

end;  { CalculateApproverEmail }




function TufrmCertifyExpDataLoader.GroupIsIn(const GroupIn, GroupSetIn: String): Boolean;
begin

  Result := False;

  If Pos(UpperCase('|' + GroupIn + '|'), UpperCase(GroupSetIn) ) > 0 Then
    Result := True;

end;


procedure TufrmCertifyExpDataLoader.FilterTripsByCount;
var
  PriorCrewID : String;
  Counter : Integer;
  Limit : Integer;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Filtering Trips by Max Count per person';
  Application.ProcessMessages;

  Limit := StrToInt(edLastNTrips.Text);
  qryGetStartBucketSorted.close;
  qryGetStartBucketSorted.Open;              // do not change the sort order of this query

  Counter := 1;
  PriorCrewID := qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString;
  while not qryGetStartBucketSorted.eof do begin
    if qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString = PriorCrewID then begin
      counter := counter + 1;
      if Counter > Limit then
        DeleteTrip(qryGetStartBucketSorted.FieldByName('LogSheet').AsInteger,
                   qryGetStartBucketSorted.FieldByName('QuoteNum').AsInteger, 
                   qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString );

    end else begin
      PriorCrewID := qryGetStartBucketSorted.FieldByName('CrewMemberID').AsString;
      Counter := 1;
    end;

    qryGetStartBucketSorted.Next;

  end;  { While }

end;  {FilterTripsByCount}



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
  WriteLn(CertifyEmployeeFile, 'work_email,first_name,last_name,employee_id,employee_type,group,department_name,approver_email_1,approver_email_2,accountant_email') ;

  qryGetEmployees.Close;
  qryGetEmployees.ParamByName('parmImportDateIn').AsDateTime  := BatchTimeIn;
  qryGetEmployees.ParamByName('parmRecordStatusIn').AsString  := 'OK';
  qryGetEmployees.ParamByName('parmRecordStatusIn2').AsString := 'warning';
  qryGetEmployees.Open;
  WriteTime := GetTimeFromDBServer();
  while not qryGetEmployees.eof do begin
    WriteToCertifyEmployeeFile(WriteTime);
    qryGetEmployees.Next;
  end;  { While }

  slOutRec.free;                     // put in Try Finallys   ???JL
  CloseFile(CertifyEmployeeFile);

  AppendSpecialUsers(CertifyEmployeeFile);

end;  { BuildEmployeeFile }


procedure TufrmCertifyExpDataLoader.WriteToCertifyEmployeeFile( Const BatchTimeIn: TDateTime )  ;
Var
  slEmpRec : TStringList;          // Employee Record

begin
  slEmpRec := TStringList.Create;
  try
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_work_email').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_first_name').AsString  ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_last_name').AsString  ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_employee_id').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_employee_type').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_group').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_department_name').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_approver1_email').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_approver2_email').AsString ) ;
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_accountant_email').AsString ) ;

    WriteLn(CertifyEmployeeFile, slEmpRec.CommaText) ;

    // update status fields
    qryGetEmployees.Edit;
    if qryGetEmployees.FieldByName('record_status').AsString = 'warning' then
      qryGetEmployees.FieldByName('record_status').AsString := 'warning-exported'
    else
      qryGetEmployees.FieldByName('record_status').AsString := 'exported';

    qryGetEmployees.FieldByName('status_timestamp').AsDateTime := BatchTimeIn;
    qryGetEmployees.Post;

  finally
    slEmpRec.Free;
  end;

end;  { WriteToCertifyEmployeeFile }



procedure TufrmCertifyExpDataLoader.IdentifyNonCertifyRecs( Const BatchTimeIn : TDateTime );
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Flagging Non-Certify Records in PaycomHistory ';
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
  i : integer;

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

  // Scrub data a little to eliminate leading & trialing spaces
  for i := 0 to slInputFileRec.Count - 1 do begin
    slInputFileRec[i] := Trim(slInputFileRec[i]);
  end;

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

    if slInputFileRec[9] <> '' then begin    //  refactoring - move this test into ValidateRecords()?    ???JL
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

    tblPaycomHistory.FieldByName('certify_department').AsString     := slInputFileRec[8];
    tblPaycomHistory.FieldByName('certify_role').AsString           := slInputFileRec[10];
    tblPaycomHistory.FieldByName('paycom_approver1_email').AsString := slInputFileRec[11];
    tblPaycomHistory.FieldByName('paycom_approver2_email').AsString := slInputFileRec[12];
    tblPaycomHistory.FieldByName('paycom_assigned_ac').AsString     := slInputFileRec[13];
    tblPaycomHistory.FieldByName('record_status').AsString          := recStatus ;
    tblPaycomHistory.FieldByName('status_timestamp').AsDateTime     := BatchTimeIn;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime          := BatchTimeIn;
    tblPaycomHistory.FieldByName('data_source').AsString            := 'paycom_file';
    tblPaycomHistory.FieldByName('hire_date').AsString              := slInputFileRec[14];
    tblPaycomHistory.post;

  except on E: Exception do begin
    tblPaycomHistory.Edit;
    tblPaycomHistory.FieldByName('record_status').AsString := 'error';
    tblPaycomHistory.FieldByName('error_text').AsString    := tblPaycomHistory.FieldByName('error_text').AsString + '; ' + E.Message;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime := BatchTimeIn;
    tblPaycomHistory.FieldByName('data_source').AsString   := 'paycom_file';
    tblPaycomHistory.post;
  end;

  end;  { Try/Except }

end;  { InsertIntoHistoryTable() }



procedure TufrmCertifyExpDataLoader.btnLoadTailLeadPilotTableClick(Sender: TObject);
begin
  LoadTailLeadPilot;
end;



procedure TufrmCertifyExpDataLoader.LoadTailLeadPilot;
var
  FileIn: TextFile;
  sl : TStringList;
  s: string;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Importing Tail_LeadPilot Data' ;
  Application.ProcessMessages;

  sl := TStringList.Create;
  sl.StrictDelimiter := true;      { tell stringList to not use space as delimeter }

  AssignFile(FileIn, edTailLeadPilotFile.Text) ;
  Reset(FileIn);
  tblTailLeadPilot.open;
  tblTailLeadPilot.EmptyTable;

  Readln(FileIn, s);  // skip first/header row
  while not Eof(FileIn) do begin
    Readln(FileIn, s);
    sl.CommaText := s;
    InsertTailLeadPilot(sl[0], sl[1]);
  end;

  tblTailLeadPilot.close;
  CloseFile(FileIn);
  sl.Free;

end;  { LoadTailLeadPilot }


procedure TufrmCertifyExpDataLoader.InsertTailLeadPilot(const TailNumIn, EMailIn: String);
begin
//  ShowMessage( TailNumIn + #13 + EMailIn );
  tblTailLeadPilot.Insert;
  tblTailLeadPilot.FieldByName('Tail').AsString  := Trim(TailNumIn);
  tblTailLeadPilot.FieldByName('Email').AsString := Trim(EmailIn);
  tblTailLeadPilot.Post;

end;  { InsertTailLeadPilot }




procedure TufrmCertifyExpDataLoader.LoadTripsIntoStartBucket;
var
  strDaysBack: String;
  i: Integer;

begin
(*  1. Empty Start Bucket & Load trips into Start Bucket from Trip tables
    3. Add Vendor number for air crew
*)

  StatusBar1.Panels[1].Text := 'Current Task:  Loading Trips into StartBucket ' ;
  Application.ProcessMessages;

  qryLoadTripData.SQL.Clear;
  strDaysBack := edDaysBack.Text;

  // Empty the working table
  qryLoadTripData.SQL.Append('delete from CertifyExp_Trips_StartBucket');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load PIC data into StartBucket
  qryLoadTripData.SQL.Clear;
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.PICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, T.TR_DEPART');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.PICPILOTNO > 0');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load SIC data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.SICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, T.TR_DEPART');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.SICPILOTNO > 0');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load TIC data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.TICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, T.TR_DEPART');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.TICPILOTNO > 0');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load FA (Flight Attendant) data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.FANO, T.QUOTENO, L.ACREGNO, FARPART, 0, T.TR_DEPART');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.FANO > 0');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  qryGetAirCrewVendorNum.Execute;


  // Load default crew & tail pairs regardless of whether any trips flown recently. per  Phase3 Tasks & Estimates item #4    15May2019
  qryLoadTripData.SQL.Append('INSERT INTO CertifyExp_Trips_StartBucket (CrewMemberVendorNum, TailNum, CrewMemberID)');
  qryLoadTripData.SQL.Append('SELECT DISTINCT VendorNumber, AssignedAC, ' + QuotedStr('DefaultCrewTail') );
  qryLoadTripData.SQL.Append('FROM [Warehouse].[dbo].[QuoteSys_PilotMaster] ' );
  qryLoadTripData.SQL.Append('WHERE AssignedAC   LIKE ' + QuotedStr('N%'));                           // Valid Tail Numbers begin w/ N ...
  qryLoadTripData.SQL.Append('  AND AssignedAC   NOT IN (' + QuotedStr( 'NORDSTROM' ) + ', ' + QuotedStr('NORDSTROMS') + ')' );  // ... except for a few goofy records
  qryLoadTripData.SQL.Append('  AND VendorNumber IS NOT NULL' );                                      // Need to have Vendor Number
  qryLoadTripData.SQL.Append('  AND ArchiveFlag  IS NULL' );                                          // Get only non-terminated flight crew
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;


//  T&E-25    ???JL

//  Query PaycomHistory:
//
//    SELECT certify_gp_vendornum, certify_department, certfile_group, paycom_assigned_ac
//    FROM [Warehouse].[dbo].[CertifyExp_PayComHistory]
//    where imported_on = '2019-07-01 22:15:01.877'
//      and certfile_group in ('FlightCrew','PoolPilot','PoolFA','FlightCrewCorp','FlightCrewNonPCal','IFS')
//      and paycom_assigned_ac not blank


//   Write recs to StartBucket
//      while not qry.eof do begin
//        stlAssignedAC := ParseACList(qry.fieldByName('paycom_assigned_ac').AsString);
//        for i := 0 to stlAssignedAC.Count - 1 do begin
//          InsertCrewTailRec(qry37.FieldByName('certify_gP_vendornum').asString, stlAssignedAC[i]);
//        end;
//
//        qry.next;
//      end;


end;  { LoadTripsIntoStartBucket }


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

  if Pos( '|' + qryGetEmployees.FieldByName('certify_role').AsString + '|', '|Accountant|Employee|Executive|Manager|') = 0 then
    strErrorText := strErrorText + 'invalid certify_role; ';

//  if qryGetEmployees.FieldByName('certify_role').AsString = '' then
//    strErrorText := strErrorText + 'missing certify_role; ';


  if qryGetEmployees.FieldByName('work_email').AsString = '' then
    strErrorText := strErrorText + 'missing work_email; ';


  qryGetEmployees.Edit;
  qryGetEmployees.FieldByName('status_timestamp').AsDateTime := TimeStampIn;

  if strErrorText <> '' then begin
    qryGetEmployees.FieldByName('error_text').AsString    := qryGetEmployees.FieldByName('error_text').AsString + '; ' + strErrorText;
    qryGetEmployees.FieldByName('record_status').AsString := 'error';
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

  // Set status = 'terminated' for employees terminated more than edTerminatedDaysBack.Text days ago so that they don't get procesed
  //   Test for Terminated employee first because if Terminated then other validation is moot.
  qryFlagTerminatedEmployees.Close;
  qryFlagTerminatedEmployees.ParamByName('parmDaysBackTerminated').AsInteger := StrToInt(edTerminatedDaysBack.Text);
  qryFlagTerminatedEmployees.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
  qryFlagTerminatedEmployees.Execute;   // Sets RecordStatus = 'terminated'

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

  // Validate Vendor Number w/ an Update query
  qryValidateVendorNum.Close;     // checks if vendornum exists in Great Plains and sets ...
  qryValidateVendorNum.ParamByName('parmImportedOn').AsDateTime := BatchTimeIn;
  qryValidateVendorNum.Execute;   //        ... record_status = 'error' & error_text = 'certify_gp_vendornum not found' if not

  // Set status = 'error' for flight crews that are missing values for important Certify fields, see query
  qryFlagMissingFlightCrews.Close;
  qryFlagMissingFlightCrews.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
  qryFlagMissingFlightCrews.Execute;

end;  { ValidateRecords }


procedure TufrmCertifyExpDataLoader.UpdateCCField(const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Updating has_credit_card field';
  Application.ProcessMessages;

  qryUpdateHasCCField.Close;
  qryUpdateHasCCField.ParamByName('parmImportedOn').AsDateTime := BatchTimeIn;
  qryUpdateHasCCField.Execute;

end;


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
//var
//  SC : String;

begin

  // hard-coded special case, per specs - disabled, per revised specs dated 31Aug
//  if Pos('N113CS', qryGetImportedRecs.FieldByName('department_descrip').AsString) > 0 then begin
//    Result := 'rdragoo@claylacy.com';
//    Exit;
//  end;

//  SC := SupervisorCode;
//  If Length(SC) = 3 then
//    SC := '0' + SupervisorCode;

  // use different technique to avoid repeated execution of qryGetApproverEmail;  ???JL
  qryGetApproverEmail.Close;
  qryGetApproverEmail.ParamByName('parmEmpCode').AsString       := SupervisorCode ;
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
  qryBuildValFile.SQL.Text := 'select distinct LogSheet, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null AND LogSheet is not null order by LogSheet';
  qryBuildValFile.Open ;

  RowOut := 'LogSheet,CrewMemberVendorNum';
  WriteLn(WorkFile, RowOut) ;
  while ( not qryBuildValFile.eof ) do begin
    RowOut := Trim(qryBuildValFile.FieldByName('LogSheet').AsString) + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + Trim(qryBuildValFile.FieldByName('LogSheet').AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildCrewLogFile }



procedure TufrmCertifyExpDataLoader.Load_CrewTail_HistoryTable(Const BatchTimeIn: TDateTime);
Var
  RowOut : String;
  WorkFile : TextFile;                // ??? superfluous?
  CurrentBatchDateTime : TDateTime;
  CrewTailFileName: String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Loading data into CrewTail_History table'  ;
  Application.ProcessMessages;

  qryInsertCrewTailHist.Close;
  qryInsertCrewTailHist.ParamByName('parmBatchDateTime').AsDateTime := BatchTimeIn ;
  qryInsertCrewTailHist.Execute;

end;  { LoadCrewTailHistoryTable }


procedure TufrmCertifyExpDataLoader.Load_CrewTrip_HistoryTable(const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Loading data into CrewTrip_History table'  ;
  Application.ProcessMessages;

  qryInsertCrewTripHist.Close;
  qryInsertCrewTripHist.ParamByName('parmBatchDateTime').AsDateTime := BatchTimeIn ;
  qryInsertCrewTripHist.Execute;

end;  { Load_Crew_Trip_HistoryTable }


procedure TufrmCertifyExpDataLoader.Load_CrewLog_HistoryTable(const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Loading data into CrewLog_History table'  ;
  Application.ProcessMessages;

  qryInsertCrewLogHist.Close;
  qryInsertCrewLogHist.ParamByName('parmBatchDateTime').AsDateTime := BatchTimeIn ;
  qryInsertCrewLogHist.Execute;

end;  { Load_Crew_Trip_HistoryTable }


procedure TufrmCertifyExpDataLoader.BuildCrewTailFile(Const CurrBatchTimeIn: TDatetime);
Var
  RowOut : String;
  WorkFile : TextFile;
//  CurrentBatchDateTime : TDateTime;

begin
//  qryBuildValFile.Close;
//  qryBuildValFile.SQL.Clear;
//  qryBuildValFile.SQL.Add('select TailNumber, CrewMemberVendorNum ');
//  qryBuildValFile.SQL.Add('from CertifyExp_CrewTail_History ');
//  qryBuildValFile.SQL.Add('where CrewMemberVendorNum is not null and TailNumber is not null ');
//  qryBuildValFile.SQL.Add('  and CrewMemberVendorNum > 0 and CreatedOn =  ' + QuotedStr(DateTimeToStr(CurrentBatchDateTime))) ;
//  qryBuildValFile.Open ;


  qryBuildValFile.Close;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Text := 'select distinct TailNum, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and TailNum is not null and CrewMemberVendorNum > 0' ;
  qryBuildValFile.Open ;

  AssignFile(WorkFile, CalcCrewTailFileName(CurrBatchTimeIn));
//  AssignFile(WorkFile, edOutputDirectory.Text + 'crew_tail.csv');

  Rewrite(WorkFile);
  RowOut := 'TailNumber,CrewMemberID';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + Trim(qryBuildValFile.FieldByName('TailNum').AsString);

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

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

  CloseFile(WorkFile);
  qryBuildValFile.Close;

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
  qryBuildValFile.SQL.Add( ' where  LogSheet is not null '  ) ;
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

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildTailTripFile }


procedure TufrmCertifyExpDataLoader.BuildTailTripLogFile;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing tail_trip_log.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'tail_trip_log.csv');
  Rewrite(WorkFile);
  RowOut := 'TailNum,TripNum,LogSheet';
  WriteLn(WorkFile, RowOut) ;

  qryGetTailTripLog.Close;
  qryGetTailTripLog.Open;
  while Not qryGetTailTripLog.eof do Begin
    RowOut := Trim(qryGetTailTripLog.FieldByName('TailNum').AsString) + ',' +
                   qryGetTailTripLog.FieldByName('QuoteNum').AsString + ',' +
                   qryGetTailTripLog.FieldByName('LogSheet').AsString;
    WriteLn(WorkFile, RowOut) ;
    qryGetTailTripLog.Next;
  end;

  CloseFile(WorkFile);
  qryGetTailTripLog.Close;

end;  { BuildTailTripLogFile }



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


function TufrmCertifyExpDataLoader.CalcCertfileDepartmentName(const GroupNameIn: String): String;
var
  slGroupToDropdown : TStringList;
  FoundIndex : Integer;

begin
  slGroupToDropdown := TStringList.Create;    // converting Group value to default value for Certify department_code
  slGroupToDropdown.StrictDelimiter := true;  // don't use space as delimeter;
  try
    slGroupToDropdown.CommaText := myIni.ReadString('Lookups', 'DepartmentName', '');
    FoundIndex := slGroupToDropdown.IndexOf(GroupNameIn);
    if FoundIndex > -1 then
      Result := slGroupToDropdown[FoundIndex + 1]
    else
      Result := 'Error! - ' + GroupNameIn + ' not found in Department Name Lookup (in .ini)';

  finally
    slGroupToDropdown.Free;
  end;

end;


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
  qryBuildValFile.SQL.Add( '   and LogSheet is not null '  );
  qryBuildValFile.SQL.Add( ' order by TailNum ' );
  qryBuildValFile.Open ;

  RowOut := 'TailNumber,LogSheet';
  WriteLn(WorkFile, RowOut) ;
  while not qryBuildValFile.eof do begin
    RowOut := Trim(qryBuildValFile.FieldByName('TailNum').AsString) + ',' + qryBuildValFile.FieldByName('LogSheet').AsString ;
    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

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


procedure TufrmCertifyExpDataLoader.DeleteTrip(const LogSheetIn, QuoteNumIn: Integer; CrewMemberIDIn : String);
begin

//  LogIt('LogSheet: '  + IntToStr(LogSheetIn) + '; ' +
//        'TripNum: '   + IntToStr(QuoteNumIn) + '; ' +
//        'VendorNum: ' + qryGetStartBucketSorted.FieldByName('CrewMemberVendorNum').AsString + '; ' +
//        'TailNum: '   + qryGetStartBucketSorted.FieldByName('TailNum').AsString );

  qryDeleteTrips.Close;
  qryDeleteTrips.ParamByName('parmLogSheetIn').AsInteger     := LogSheetIn;
  qryDeleteTrips.ParamByName('parmCrewMemberIDIn').AsString  := CrewMemberIDIn;

//  qryDeleteTrips.ParamByName('parmCrewMemberIDIn').AsInteger := CrewMemberIDIn;

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

  qryPurgeWorkingTable.Execute;

  qryContractorsNotInPaycom_Step1.ParamByName('parmDaysBack').AsInteger := StrToInt(edContractorDaysBack.text);
  qryContractorsNotInPaycom_Step1.Execute;    //  ???JL   when employee/contractor terminated - new scope

  qryContractorsNotInPaycom_Step2.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn ;
  qryContractorsNotInPaycom_Step2.Execute;    //  remove contractors from Contractors45 table that are common w/ Paycom File

  tblPaycomHistory.Open;

  qryGetPilotDetails.close;
  CalcSQLForQryGetPilotDetails('FlownRecent');   // calculate the Where clause for qryGetPilotDetails; refactoring for T&E-22
  qryGetPilotDetails.open;

  while not qryGetPilotDetails.eof do begin
    WriteContractorsToPaycomTable(BatchTimeIn, False);
    qryGetPilotDetails.Next;
  end ;

  tblPaycomHistory.Close;
  qryGetPilotDetails.Close;

end;  { AddContractorsNotInPaycom }


(* ------------------------------------------------------
  Adding dummy records for newly-hired contractors so that they can
   be trained on Certify before they have flown any trips.

  The process requires that records be added to the following tables/files
    1. PaycomHistory table
    2. StartBucket table
    3. trip_stop.csv file

  Items 1 & 2 are performed here.
  Item 3 is done in BuildValidationFiles() by calling InsertDummyNewHireTripStop()

  -------------------------------------------------------- *)
procedure TufrmCertifyExpDataLoader.Process_NewHire_Contractors(const BatchTimeIn: TDateTime);
begin
  gloNewHireDummyQuoteNum := 818181;

  tblPaycomHistory.Open;

  qryGetPilotDetails.close;
  CalcSQLForQryGetPilotDetails('NewHireContractor');         // Get NewHire contractors from the PilotMaster table
  qryGetPilotDetails.open;

  while not qryGetPilotDetails.eof do begin
    if Not FindVendorNumInStartBucket(qryGetPilotDetails.FieldByName('VendorNumber').AsInteger) then begin
      WriteContractorsToPaycomTable(BatchTimeIn, True);                                            // should we use parameter to signify Contractor New Hire  ???JL
      AddDummyTripToStartBucket(qryGetPilotDetails.FieldByName('VendorNumber').AsInteger, 'ConNewHire' );
    end;

    qryGetPilotDetails.Next;
  end;

  qryGetPilotDetails.Close;
  tblPaycomHistory.close;

end;  { ProcessNewHireContractors }


procedure TufrmCertifyExpDataLoader.Process_NewHire_Employees_FlightCrew(const BatchTimeIn: TDateTime);
begin
  gloNewHireDummyQuoteNum := 828282;

  qryGetFlightCrewNewHire.Close;    // contains definition of Flight Crew; refactor to pull that definition up & store in one place  ???JL  13 Jun 2019
  qryGetFlightCrewNewHire.ParamByName('parmImportedOn').AsDateTime := BatchTimeIn;
  qryGetFlightCrewNewHire.Open;

  while not qryGetFlightCrewNewHire.eof do  begin
    if Not FindVendorNumInStartBucket(qryGetFlightCrewNewHire.FieldByName('certify_gp_vendornum').AsInteger) then begin
      AddDummyTripToStartBucket(qryGetFlightCrewNewHire.FieldByName('certify_gp_vendornum').AsInteger, 'EmpNewHire');
    end;
    qryGetFlightCrewNewHire.Next;
  end;

  qryGetFlightCrewNewHire.Close;

end;


procedure TufrmCertifyExpDataLoader.AddDummyTripToStartBucket(const VendorNumIn: Integer; NewHireKind: String);
begin
  // Set values for Dummy trip for New Hire
  tblStartBucket.Open;
  tblStartBucket.Insert;
  tblStartBucket.FieldByName('LogSheet').AsInteger            := 123456;
  tblStartBucket.FieldByName('CrewMemberID').AsString         := NewHireKind;
  tblStartBucket.FieldByName('QuoteNum').AsInteger            := gloNewHireDummyQuoteNum;
  tblStartBucket.FieldByName('TailNum').AsString              := 'NTEST1';
  tblStartBucket.FieldByName('CrewMemberVendorNum').AsInteger := VendorNumIn;
  tblStartBucket.Post;
  tblStartBucket.Close;

end;  { AddDummyTripToStartBucket }


procedure TufrmCertifyExpDataLoader.WriteContractorsToPaycomTable( Const BatchTimeIn: TDateTime; NewHireFlag: Boolean ) ;
begin

  tblPaycomHistory.Insert;
  tblPaycomHistory.FieldByName('employee_code').AsString           := 'contractor';     // should we use parameter to signify Contractor New Hire  ???JL
  tblPaycomHistory.FieldByName('employee_name').AsString           := CalcPilotName;
  tblPaycomHistory.FieldByName('work_email').AsString              := qryGetPilotDetails.FieldByName('EMail').AsString;
  tblPaycomHistory.FieldByName('position').AsString                := qryGetPilotDetails.FieldByName('JobTitle').AsString;
  tblPaycomHistory.FieldByName('department_descrip').AsString      := CalcDeptDescrip ;
  tblPaycomHistory.FieldByName('job_code_descrip').AsString        := 'Pilot-Designated' ;
  tblPaycomHistory.FieldByName('supervisor_primary_code').AsString := '';
  tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger   := qryGetPilotDetails.FieldByName('VendorNumber').AsInteger;
  tblPaycomHistory.FieldByName('certify_department').AsString      := 'FlightCrew';
  tblPaycomHistory.FieldByName('certify_role').AsString            := 'Manager';

  tblPaycomHistory.FieldByName('certfile_approver1_email').AsString  := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('certfile_approver2_email').AsString  := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('accountant_email').AsString          := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('certfile_accountant_email').AsString := 'FlightCrew@ClayLacy.com';

  tblPaycomHistory.FieldByName('paycom_assigned_ac').AsString      := qryGetPilotDetails.FieldByName('AssignedAC').AsString ;

  tblPaycomHistory.FieldByName('record_status').AsString           := 'imported';
  tblPaycomHistory.FieldByName('status_timestamp').AsDateTime      := BatchTimeIn;
  tblPaycomHistory.FieldByName('imported_on').AsDateTime           := BatchTimeIn;

  tblPaycomHistory.FieldByName('data_source').AsString             := 'PilotMaster';         // how about pilot_masterNewHire
  If NewHireFlag then
    tblPaycomHistory.FieldByName('data_source').AsString           := 'PilotMaster-NewHire';         // how about pilot_masterNewHire

  tblPaycomHistory.Post;

end;  { WriteContractorsToPaycomTable }



function TufrmCertifyExpDataLoader.FindVendorNumInStartBucket(const VendorNumIn: Integer): Boolean;
begin
  Result := False;

  qryFindVendorNumInStartBucket.Close;
  qryFindVendorNumInStartBucket.ParamByName('parmVendorNumIn').AsInteger := VendorNumIn;
  qryFindVendorNumInStartBucket.Open;

  if qryFindVendorNumInStartBucket.RecordCount > 0 then
    Result := True;

end;


function TufrmCertifyExpDataLoader.CalcPilotName: String;
begin
  Result := qryGetPilotDetails.FieldByName('LastName').AsString + ',' + qryGetPilotDetails.FieldByName('FirstName').AsString

end;


procedure TufrmCertifyExpDataLoader.CalcSQLForQryGetPilotDetails(const QryKind: String);
begin
  qryGetPilotDetails.SQL.Clear;
  qryGetPilotDetails.SQL.Add('select PilotID,LastName,FirstName,VendorNumber,UpdatedInQuoteSys,UpdatedBy,Base,ArchiveFlag,JobTitle,EmployeeStatus,Status,EMail,AssignedAC ');
  qryGetPilotDetails.SQL.Add('from QuoteSys_PilotMaster ');

  if QryKind = 'FlownRecent' then
    qryGetPilotDetails.SQL.Add('where PilotID in (select PilotID from CertifyExp_Contractors45) ')

  else if QryKind ='NewHireContractor' then begin       // Refactor!! we have definition of contractor here & in qryContractorsNotInPaycom should only be 1 place   ???JL
    qryGetPilotDetails.SQL.Add('where EmployedOn > (CURRENT_TIMESTAMP - 45) ') ;  // param for days back?   ???JL
    qryGetPilotDetails.SQL.Add('  and Status         in ( ' + QuotedStr('Agent of CLA') + ',' + QuotedStr('Cabin Server') + ',' + QuotedStr('Parttime - CLA') + ' ) ') ;  // Definition of "Contractor"
    qryGetPilotDetails.SQL.Add('  and EmployeeStatus in ( ' + QuotedStr('Part 91') + ',' + QuotedStr('Part 135') + ',' + QuotedStr('Cabin Serv') + ') ' );       // Definition of "Contractor"

  end else
    ShowMessage('CalcSQLForQryGetPilotDetails(): Unknown Value for param QryKind: ' + QryKind);

//  ShowMessage(QryKind + #13 + qryGetPilotDetails.SQL.Text);        // debugging code

end;  { CalcSQLForQryGetPilotDetails }


function TufrmCertifyExpDataLoader.CalcDeptDescrip: String;
begin
  Result :=  'Designated-' + qryGetPilotDetails.FieldByName('AssignedAC').AsString    ;

//  Result := 'Designated-N1234X';
end;



procedure TufrmCertifyExpDataLoader.SendStatusEmail;
Var
  mySMTP    : TIdSMTP;
  myMessage : TIDMessage;
  OutPutFileDir : String;
  stlOutputFiles : TStringList;
  i : Integer;

begin

  myMessage := TIdMessage.Create(nil);
  myMessage.Subject      := 'CLA Certify Data Loader Status Report';
  myMessage.From.Address := 'CertifyDataLoader@claylacy.com';
  myMessage.Body.Text    := 'See attached files for Employee processing errors and uploaded data files:' ;
  myMessage.Recipients.EMailAddresses := myIni.ReadString('OutputFiles', 'EMailRecipientList', '');   //jeff@dcsit.com,jluckey@pacbell.net';   //,thomasfduffy@gmail.com';

  //  Load Attachments
  OutPutFileDir  := edOutputDirectory.Text;          // myIni.ReadString('OutputFiles', 'OutputFileDir', '');
  stlOutputFiles := TStringList.Create();
  stlOutputFiles.CommaText := myIni.ReadString('OutputFiles', 'EmailAttachFileList', '');
  for i := 0 to stlOutputFiles.Count - 1 do begin
   if FileExists( OutPutFileDir + stlOutputFiles[i] ) then
     TIDAttachmentFile.Create(myMessage.MessageParts, OutPutFileDir + stlOutputFiles[i] );

  end ; { for }

  // Add the Paycom Error file to attachemnt list
  if FileExists( gloPaycomErrorFile ) then
    TIDAttachmentFile.Create(myMessage.MessageParts, gloPaycomErrorFile );

  mySMTP := TIdSMTP.Create(nil);
  mySMTP.Host     := '192.168.1.73';
  mySMTP.Username := 'tkvassay@claylacy.com';
  mySMTP.Password := '';

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
//    mySMTP.Disconnect();
  Except on E:Exception Do
    //ShowMessage( 'Email Error: ' + E.Message);   need to log this
  End;

  mySMTP.free;
  myMessage.Free;

end;  { SendStatusEmail }


procedure TufrmCertifyExpDataLoader.CreateEmployeeErrorReport(Const BatchTimeIn : TDateTime);
Var
  PaycomErrorFile : TextFile;
  i, j : Integer;
  slErrorRec : TStringList;

begin
  slErrorRec := TStringList.Create;

  // Prep Output File
  AssignFile(PaycomErrorFile, CalcPaycomErrorFileName(BatchTimeIn));

//  ShowMessage( edOutputDirectory.Text + CalcPaycomErrorFileName(BatchTimeIn) );

  Rewrite(PaycomErrorFile);

  qryGetImportedRecs.Close;
  qryGetImportedRecs.ParamByName('parmBatchTimeIn').AsDateTime := BatchTimeIn ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn').AsString   := 'error' ;
  qryGetImportedRecs.ParamByName('parmRecStatusIn2').AsString  := 'warning-exported' ;
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
  strMonth, strDay: String;

begin
  DecodeDate(Trunc(BatchTimeIn), myYear, myMonth, myDay);

  strMonth := IntToStr(myMonth);
  if myMonth < 10 then
    strMonth := '0' + strMonth ;

  strDay := IntToStr(myDay);
  if myDay < 10 then
    strDay := '0' + strDay;

  Result :=   'C:\CertifyExpense\OutputFiles\PaycomErrors_' + IntToStr(myYear) + strMonth + strDay + '.csv';
               // get this prefix from .ini  ???JL 2 Jan 2019

  gloPaycomErrorFile := Result;

end;  { CalcPaycomErrorFileName }


function TufrmCertifyExpDataLoader.CalcCrewTailFileName(const BatchTimeIn: TDateTime): String;
var
  myMonth, myDay, myYear: word;
  strMonth, strDay: String;
  FileDest : String;

begin
  DecodeDate(Trunc(BatchTimeIn), myYear, myMonth, myDay);

  strMonth := IntToStr(myMonth);
  if myMonth < 10 then
    strMonth := '0' + strMonth ;

  strDay := IntToStr(myDay);
  if myDay < 10 then
    strDay := '0' + strDay;


  Result := myIni.ReadString('OutputFiles', 'CrewTailDestination', ''); // + '.csv'  ;

(*
  FileDest := myIni.ReadString('OutputFiles', 'CrewTailDestination', '');

  ShowMessage(FileDest + #13 +  Copy(FileDest, Length(FileDest), 1) );

  if Copy(FileDest, Length(FileDest), 1) = '_' then        // If file destination string ends with an underbar, "_"
    Result := myIni.ReadString('OutputFiles', 'CrewTailDestination', '') + IntToStr(myYear) + strMonth + strDay + '.csv'
  else

  //  add feature: if last char of CrewTailDestination is "_" then append time stamp         ???JL
*)
end;  { CalcCrewTailFileName }



procedure TufrmCertifyExpDataLoader.AppendSpecialUsers(Const FileToAppend: TextFile);
var
  ExtraEmployeeFile : TextFile;
  EmpRec : String;

begin
  AssignFile(ExtraEmployeeFile, edSpecialUsersFile.Text );
  Reset(ExtraEmployeeFile);
  ReadLn(ExtraEmployeeFile, EmpRec);  // Read first row which contains file header & ignore it
  Append(FileToAppend);

  while not eof(ExtraEmployeeFile) do begin
    ReadLn(ExtraEmployeeFile, EmpRec);
    Writeln(FileToAppend, EmpRec);
  end;

  CloseFile(FileToAppend);
  CloseFile(ExtraEmployeeFile);

end;  { AppendSpecialUsers }


function TufrmCertifyExpDataLoader.ScrubCertifyDept(const DepartmentIn: String): String;
begin

  Result := DepartmentIn;     // need to implement ???JL

end;


function TufrmCertifyExpDataLoader.ScrubFARPart(const FarPartIn: String): String;
begin
  Result := 'Error_FARPart';

  if Pos('135', FarPartIn) > 0 then
    Result := '135'
  else if Pos('91', FarPartIn) > 0 then
    Result := '91' ;

end;  { ScrubFARPart }



procedure TufrmCertifyExpDataLoader.Load_CharterVisa_IntoStartBucket;
var
  stlCharterVisaUsers : TStringList;
  i : Integer;

begin
  stlCharterVisaUsers           := TStringList.Create;
  stlCharterVisaUsers.CommaText := edCharterVisaUsers.Text;       // comma-seperated string of Vendor Numbers for CharterVisa Group
  try
    for i := 0 to stlCharterVisaUsers.Count - 1 do Begin
      qryInsertCharterVisa.Close;
      qryInsertCharterVisa.ParamByName('parmCrewMemberVendorNum').AsInteger := StrToInt(stlCharterVisaUsers[i]);
      qryInsertCharterVisa.Execute;
    End;

  finally
    stlCharterVisaUsers.Free;
  end;

end;  { LoadCharterVisaTripsIntoStartBucket }




(*
must handle the case where multiple aircraft could be listed in the paycom_assigned_ac field, separated by a forward slash '/'
  for example:  N225MC/N8241W  or  N225MC/N8241W/N550WT

However most of the time there is only one aircraft listed              
*)
procedure TufrmCertifyExpDataLoader.Load_DOM_IntoStartBucket(Const BatchTimeIn: TDatetime);
var
  stlACList : TStringList;
  i : Integer;

begin
  stlACList := TStringList.Create;
  stlACList.Delimiter := '/';
  Try
    tblStartBucket.Open;

    qryGetDOMEmployees.Close;
    qryGetDOMEmployees.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
    qryGetDOMEmployees.Open;

    while not qryGetDOMEmployees.eof do begin
      stlACList.DelimitedText := qryGetDOMEmployees.FieldByName('paycom_assigned_ac').AsString;
      for i := 0 to stlACList.Count - 1 do begin
        InsertCrewTail(Trim(stlACList[i]), qryGetDOMEmployees.FieldByName('certify_gp_vendornum').AsInteger);
      end;  {for}
      qryGetDOMEmployees.Next;
    end;  { while }

    qryGetDOMEmployees.Close;
    tblStartBucket.CLose;

  Finally
    stlACList.Free;
  End;  { Try/finally }
  
end;  { LoadDOMsIntoStartBucket }



// **********************************************
//
//   1. Get all IFS recs from PaycomHistory for given batch
//   2. Add all tails for each IFS employee into StartBucket Crew & Tail
//
// **********************************************
procedure TufrmCertifyExpDataLoader.Load_IFS_IntoStartBucket(BatchTime: TDateTime);
begin

  qryGetIFS.Close;
  qryGetIFS.ParamByName('parmImportedOn').AsDateTime := BatchTime;
  qryGetIFS.Open;

  while not qryGetIFS.eof do begin
    qryInsertIFS.Close;
    qryInsertIFS.ParamByName('parmCrewMemberVendorNum').AsInteger := qryGetIFS.FieldByName('certify_gp_vendornum').AsInteger;
    qryInsertIFS.Execute;
    qryGetIFS.Next;
  end;

  qryGetIFS.Close;

end;  { LoadIFSIntoStartBucket }



procedure TufrmCertifyExpDataLoader.InsertCrewTail(const TailNumIn: String; VendorNumIn: Integer);
begin
  tblStartBucket.Insert;
  tblStartBucket.FieldByName('TailNum').AsString              := TailNumIn;
  tblStartBucket.FieldByName('CrewMemberVendorNum').AsInteger := VendorNumIn;
  tblStartBucket.FieldByName('CrewMemberID').AsString         := 'DOM_processing';
  tblStartBucket.Post;

end;



function TufrmCertifyExpDataLoader.FindLeadPilot(const AssignedACString: String; BatchTimeIn: TDateTime): String;
var
  stlACList : TStringList;
  AssignedAC: String;        // Assigned Aircraft
  LeadPilotEMail : String;

begin
  Result := 'NotFound';

  if AssignedACString = '' then Begin
    FlagRecordAsError('warning', 'Assigned Aircraft (paycom_assigned_ac) is missing');
    Result := 'NotFound';
    Exit;
  End;

  // step 1: parse AssignedACString, it could contain multiple aircraft seperated by forward slash: 'N855LD/N5504B'
  stlACList := TStringList.Create;
  stlACList.Delimiter := '/';
  stlACList.DelimitedText := AssignedACString ;
  AssignedAC := stlACList[0];   // per Specs, we only care about the first tail number in the string
  stlACList.Free;

  // step 2: Find email for that AC's lead pilot
  qryFindLeadPilotEMail.Close;
  qryFindLeadPilotEmail.ParamByName('parmTailNumIn').AsString := AssignedAC;
  qryFindLeadPilotEmail.Open;
  LeadPilotEMail := qryFindLeadPilotEMail.FieldByName('EMail').AsString;

  if (qryFindLeadPilotEmail.RecordCount > 0) and (LeadPilotEMail <> '' ) then begin            // If LeadPilot not found
    if Not EmployeeTerminated(LeadPilotEMail, BatchTimeIn) then                                //   If LeadPilot is terminated
      Result := LeadPilotEMail
    else begin
      FlagRecordAsError('warning', 'Lead Pilot: ' + LeadPilotEMail + ' terminated');
    end;

  end else begin
    FlagRecordAsError('warning', 'Lead Pilot: ' + LeadPilotEMail + ' or TailNumber: ' + AssignedAC + ' not found in Tail_LeadPilot table');
  end;

end;  { FindLeadPilot }


function TufrmCertifyExpDataLoader.EmployeeTerminated(const EMailIn: String; ImportDateIn: TDateTime): Boolean;
begin

  Result := False;

  qryGetTerminationDate.Close;
  qryGetTerminationDate.ParamByName('parmEMail').AsString        := EMailIn;
  qryGetTerminationDate.ParamByName('parmImportedOn').AsDateTime := ImportDateIn;
  qryGetTerminationDate.Open;

  if (qryGetTerminationDate.RecordCount > 0) then begin
    If qryGetTerminationDate.FieldByName('termination_date').IsNull then
      Result := False
    else
      Result := True;

  end else
    FlagRecordAsError('error', 'Cannot find ' + EMailIn + ' in PaycomHistory table - (check for differences in Tail_LeadPilot table)');

  qryGetTerminationDate.Close;

end;  { EmployeeTerminated }


procedure TufrmCertifyExpDataLoader.FlagRecordAsError(const ErrorType, ErrorMsgIn: String);
Var
  AdditionalErrorMsg : String;

begin
   AdditionalErrorMsg := '';

  // assuming we are on the current record in PaycomHistory table!  All calling procs should be in this context
  if ErrorType = 'error' then
    qryGetImportedRecs.FieldByName('record_status').AsString := 'error'
  else if ErrorType = 'warning' then
    qryGetImportedRecs.FieldByName('record_status').AsString := 'warning'
  else
    AdditionalErrorMsg := '-- Unknown ErrorType passed to FlagRecordAsError(): ' + ErrorType;

  qryGetImportedRecs.FieldByName('error_text').AsString    := ErrorMsgIn + AdditionalErrorMsg;

end;  { FlagRecordAsError }


procedure TufrmCertifyExpDataLoader.Do_CrewTail_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);
begin

  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving added CrewTail_History recs'  ;
  Application.ProcessMessages;

  // Get Added recs from this new batch
  UpdateRecordStatus_CrewTail('added', PreviousBatchDateIn, NewBatchDateIn);
  qryGetCrewTailRecs.ParamByName('parmCreatedOnIn').AsDateTime := NewBatchDateIn;
  qryGetCrewTailRecs.ParamByName('parmRecStatusIn').AsString   := 'added';
  qryGetCrewTailRecs.Open;

  DataSource1.DataSet := qryGetCrewTailRecs;
  Application.ProcessMessages;
  Sleep(5000);

  SendToCertify(qryGetCrewTailRecs, BatchTimeIn, 'crew_tail');
  qryGetCrewTailRecs.Close;

  //  Get Deleted recs from this new batch
  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving deleted CrewTail_History recs'  ;
  Application.ProcessMessages;
  UpdateRecordStatus_CrewTail('deleted', NewBatchDateIn, PreviousBatchDateIn);
  qryGetCrewTailRecs.ParamByName('parmCreatedOnIn').AsDateTime := PreviousBatchDateIn;
  qryGetCrewTailRecs.ParamByName('parmRecStatusIn').AsString   := 'deleted';
  qryGetCrewTailRecs.Open;

  DataSource1.DataSet := qryGetCrewTailRecs;
  Application.ProcessMessages;
  Sleep(5000);

  SendToCertify(qryGetCrewTailRecs, BatchTimeIn, 'crew_tail');
  qryGetCrewTailRecs.Close;

end;  { DoCrewTailAPI }


procedure TufrmCertifyExpDataLoader.Do_CrewTrip_API(const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn: TDateTime);
begin
  // Get Added recs from this new batch
  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving added CrewTrip_History recs'  ;
  Application.ProcessMessages;

  UpdateRecordStatus_CrewTrip('added', PreviousBatchDateIn, NewBatchDateIn);
  qryGetCrewTripRecs.Close;
  qryGetCrewTripRecs.ParamByName('parmCreatedOnIn').AsDateTime := NewBatchDateIn;
  qryGetCrewTripRecs.ParamByName('parmRecStatusIn').AsString   := 'added';
  qryGetCrewTripRecs.Open;

  DataSource1.DataSet := qryGetCrewTripRecs;          // Update data grid in UI for test/dev purposes
  Application.ProcessMessages;
  Sleep(5000);                                        // wait so user can see data in grid

  SendToCertify(qryGetCrewTripRecs, BatchTimeIn, 'crew_trip');
  qryGetCrewTripRecs.Close;

  //  Get Deleted recs from this new batch
  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving deleted CrewTrip_History recs'  ;
  Application.ProcessMessages;

  UpdateRecordStatus_CrewTrip('deleted', NewBatchDateIn, PreviousBatchDateIn);
  qryGetCrewTripRecs.Close;
  qryGetCrewTripRecs.ParamByName('parmCreatedOnIn').AsDateTime := PreviousBatchDateIn;
  qryGetCrewTripRecs.ParamByName('parmRecStatusIn').AsString   := 'deleted';
  qryGetCrewTripRecs.Open;

  DataSource1.DataSet := qryGetCrewTripRecs;
  Application.ProcessMessages;
  Sleep(5000);

  SendToCertify(qryGetCrewTripRecs, BatchTimeIn, 'crew_trip');
  qryGetCrewTripRecs.Close;

end;  { Do_CrewTrip_API }


procedure TufrmCertifyExpDataLoader.Do_CrewLog_API(const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn: TDateTime);
begin
  // Get Added recs from this new batch
  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving added CrewLog_History recs'  ;
  Application.ProcessMessages;

  // Get added recs from this new batch
  UpdateRecordStatus_CrewLog('added', PreviousBatchDateIn, NewBatchDateIn);
  qryGetCrewLogRecs.Close;
  qryGetCrewLogRecs.ParamByName('parmCreatedOnIn').AsDateTime := NewBatchDateIn;    // modify this query to exclude already uploaded recs  ???JL
  qryGetCrewLogRecs.ParamByName('parmRecStatusIn').AsString   := 'added';
  qryGetCrewLogRecs.Open;

  // Update data grid in UI for test/dev purposes
  DataSource1.DataSet := qryGetCrewLogRecs;
  Application.ProcessMessages;
  Sleep(5000);

  SendToCertify(qryGetCrewLogRecs, BatchTimeIn, 'crew_log');
  qryGetCrewLogRecs.Close;


  //  Get Deleted recs from this new batch
  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving deleted CrewTrip_History recs'  ;
  Application.ProcessMessages;

  UpdateRecordStatus_CrewTrip('deleted', NewBatchDateIn, PreviousBatchDateIn);
  qryGetCrewLogRecs.Close;
  qryGetCrewLogRecs.ParamByName('parmCreatedOnIn').AsDateTime := PreviousBatchDateIn;
  qryGetCrewLogRecs.ParamByName('parmRecStatusIn').AsString   := 'deleted';
  qryGetCrewLogRecs.Open;

  // Update data grid in UI for test/dev purposes
  DataSource1.DataSet := qryGetCrewLogRecs;
  Application.ProcessMessages;
  Sleep(5000);

  SendToCertify(qryGetCrewLogRecs, BatchTimeIn, 'crew_log');
  qryGetCrewLogRecs.Close;

end;  {Do_CrewLog_API}



procedure TufrmCertifyExpDataLoader.SendToCertify(Const WorkingQueryIn: TUniQuery; Const BatchTimeIn : TDateTime; DataSetNameIn: String);
begin

  StatusBar1.Panels[1].Text := 'Current Task:  Sending recs to Certify'  ;
  Application.ProcessMessages;

  while Not WorkingQueryIn.eof do begin
    gloPusher.DataSetName := DataSetNameIn;  // DataSetName values are: 'crew_tail' or 'crew_trip' or 'crew_log';  sets Certify "Dimension" in Pusher.  See DataSetName setter
    case gloPusher.CertifyDimension of
      1: gloPusher.TailNumber := WorkingQueryIn.FieldByName('TailNumber').AsString;
      2: gloPusher.TripNumber := WorkingQueryIn.FieldByName('TripNumber').AsString;
      3: gloPusher.LogNumber  := WorkingQueryIn.FieldByName('LogNumber').AsString;
    end;

    gloPusher.DataAction          := WorkingQueryIn.FieldByName('RecordStatus').AsString;
    gloPusher.CrewMemberVendorNum := WorkingQueryIn.FieldByName('CrewMemberVendorNum').AsString;
    gloPusher.Push;    // ???JL

    WorkingQueryIn.Edit;
    WorkingQueryIn.FieldByName('HTTPResultCode').AsString      := IntToStr(gloPusher.HTTPReturnCode);
    WorkingQueryIn.FieldByName('UploadStatus').AsString        := gloPusher.UploadStatus;
    WorkingQueryIn.FieldByName('UploadStatusMessage').AsString := gloPusher.UploadStatusMessage;
    WorkingQueryIn.FieldByName('UploadedOn').AsDateTime        := BatchTimeIn;

    WorkingQueryIn.Post;
    WorkingQueryIn.Next;
    application.ProcessMessages;
  end;  { While }

end;  { SendToCertify() }


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewTail(const RecStatusIn: String; const NewBatchDateIn, PreviousBatchDateIn: TDateTime);
begin
  qryUpdateRecStatus_CrewTail.close;
  qryUpdateRecStatus_CrewTail.ParamByName('parmRecStatus').AsString   := RecStatusIn;
  qryUpdateRecStatus_CrewTail.ParamByName('parmOlderDate').AsDateTime := PreviousBatchDateIn;
  qryUpdateRecStatus_CrewTail.ParamByName('parmNewerDate').AsDateTime := NewBatchDateIn;
  qryUpdateRecStatus_CrewTail.Execute;

end;


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewTrip(const RecStatusIn: String; const NewBatchDateIn, PreviousBatchDateIn: TDateTime);
begin
  qryUpdateRecStatus_CrewTrip.close;
  qryUpdateRecStatus_CrewTrip.ParamByName('parmRecStatus').AsString   := RecStatusIn;
  qryUpdateRecStatus_CrewTrip.ParamByName('parmOlderDate').AsDateTime := PreviousBatchDateIn;
  qryUpdateRecStatus_CrewTrip.ParamByName('parmNewerDate').AsDateTime := NewBatchDateIn;
  qryUpdateRecStatus_CrewTrip.Execute;

end;


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewLog(const RecStatusIn: String; const NewBatchDateIn, PreviousBatchDateIn: TDateTime);
begin
  qryUpdateRecStatus_CrewLog.close;
  qryUpdateRecStatus_CrewLog.ParamByName('parmRecStatus').AsString   := RecStatusIn;
  qryUpdateRecStatus_CrewLog.ParamByName('parmOlderDate').AsDateTime := PreviousBatchDateIn;
  qryUpdateRecStatus_CrewLog.ParamByName('parmNewerDate').AsDateTime := NewBatchDateIn;
  qryUpdateRecStatus_CrewLog.Execute;

end;


procedure TufrmCertifyExpDataLoader.SendNewCrewTailToCertify(stlNewCrewTailIn: TStringList);
begin

  sleep(1);
  // qryGetCrewTailRecs

end;


procedure TufrmCertifyExpDataLoader.GetBatchDates_CrewTail(var PreviousBatchDateOut, NewBatchDateOut: TDateTime);
begin
  qryGetCrewTailBatchDates.Close;
  qryGetCrewTailBatchDates.Open;
  NewBatchDateOut := qryGetCrewTailBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTailBatchDates.Next;
  PreviousBatchDateOut := qryGetCrewTailBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTailBatchDates.Close;

  //  Test/Debug Code   ???JL
  edPreviousDate.Text := DateTimeToStr( PreviousBatchDateOut);
  edNewDate.Text      := DateTimeToStr( NewBatchDateOut);

//  ShowMessage('NewBatchDate: '      + DateTimeToStr(NewBatchDateOut) + #13 +
//              'PreviousBatchDate: ' + DateTimeToStr(PreviousBatchDateOut) );

end;


procedure TufrmCertifyExpDataLoader.GetBatchDates_CrewTrip(var PreviousBatchDateOut, NewBatchDateOut: TDateTime);
begin
  qryGetCrewTripBatchDates.Close;
  qryGetCrewTripBatchDates.Open;
  NewBatchDateOut := qryGetCrewTripBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTripBatchDates.Next;
  PreviousBatchDateOut := qryGetCrewTripBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTripBatchDates.Close;

  //  Test/Debug Code   ???JL
  edPreviousDate.Text := DateTimeToStr( PreviousBatchDateOut );
  edNewDate.Text      := DateTimeToStr( NewBatchDateOut );

end;  { GetBatchDates_CrewTrip }


procedure TufrmCertifyExpDataLoader.GetBatchDates_CrewLog(var PreviousBatchDateOut, NewBatchDateOut: TDateTime);
begin
  qryGetCrewLogBatchDates.Close;
  qryGetCrewLogBatchDates.Open;
  NewBatchDateOut := qryGetCrewLogBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewLogBatchDates.Next;
  PreviousBatchDateOut := qryGetCrewLogBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewLogBatchDates.Close;

  //  Test/Debug Code   ???JL
  edPreviousDate.Text := DateTimeToStr( PreviousBatchDateOut );
  edNewDate.Text      := DateTimeToStr( NewBatchDateOut );

end;  { GetBatchDates_CrewLog }



procedure TufrmCertifyExpDataLoader.PruneHistoryTables(Const TableNameIn : String);
Var
  strDaysBack  : String;
  strTableName : String;

begin
  strTableName := 'CertifyExp_' + TableNameIn + '_History' ;    // must match table name exactly
  strDaysBack  := '3';                                       // for performance reasons, keep only the last 3 days of data (and for T/S purposes)
                                                             //   (the more records in these tables the slower the 'NOT IN' queries run)
  qryPruneHistoryTables.SQL.Clear;
  qryPruneHistoryTables.SQL.Add( ' delete from ' + strTableName );
  qryPruneHistoryTables.SQL.Add( ' where CreatedOn < CAST(CURRENT_TIMESTAMP - ' + strDaysBack + ' AS DATE) ' );
  qryPruneHistoryTables.Execute;

end;  { PruneHistoryTables }



procedure TufrmCertifyExpDataLoader.LogIt(ErrorMsgIn: String);
var
  YearMonth       : ShortString;
  TargetDate      : TDateTime;
  TargetYearMonth : ShortString;

begin
  Sleep(1000);      // keeps Identifier strings for ini file entries unique. Klunky, but it works  -JL
  myLog := TIniFile.Create( ExtractFilePath(ParamStr(0)) + 'CertifyDataLoader.log');
  YearMonth := IntToStr(YearOf(Now)) + '-' + IntToStr(MonthOf(Now));
  myLog.WriteString(YearMonth, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now()), ErrorMsgIn) ;

  // prune log file - purge log entries older than 180 days
  TargetDate := Now() - 60;
  TargetYearMonth := IntToStr(YearOf(TargetDate)) + '-' + IntToStr(MonthOf(TargetDate));
  if myLog.SectionExists(TargetYearMonth) Then
    myLog.EraseSection(TargetYearMonth);

end;  { LogIt }


end.
