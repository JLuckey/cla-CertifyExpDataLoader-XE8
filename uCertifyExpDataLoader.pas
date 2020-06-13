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
    scrLoadTripStopData: TUniScript;
    qryGetTripStopRecs: TUniQuery;
    qryGetStartBucketSorted: TUniQuery;
    qryPilotsNotInPaycom: TUniQuery;
    qryGetPilotsNotInPaycom: TUniQuery;
    qryEmptyPilotsNotInPaycom: TUniQuery;
    qryDeleteTrips: TUniQuery;
    qryPurgeWorkingTable: TUniQuery;
    qryGetEmployeeErrors: TUniQuery;
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
    btnGoHourly: TButton;
    edCharterVisaUsers: TEdit;
    Label16: TLabel;
    tblTailLeadPilot: TUniTable;
    qryFindLeadPilotEmail: TUniQuery;
    qryGetTerminationDate: TUniQuery;
    qryGetIFSMembers: TUniQuery;
    qryInsertCrewTailHist: TUniQuery;
    qryGetCrewTailBatchDates: TUniQuery;
    qryGetFailedRecs_CrewTail: TUniQuery;
    qryGetDeletedRecs_CrewTail: TUniQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    qryGetCrewTailRecs: TUniQuery;
    edNewDate: TEdit;
    edPreviousDate: TEdit;
    Label18: TLabel;
    Label19: TLabel;
    qryGetCrewTripRecs: TUniQuery;
    qryInsertCrewTripHist: TUniQuery;
    qryGetCrewTripBatchDates: TUniQuery;
    qryInsertCrewLogHist: TUniQuery;
    qryGetCrewLogBatchDates: TUniQuery;
    qryGetCrewLogRecs: TUniQuery;
    btnFixer: TButton;
    qryPruneHistoryTables: TUniQuery;
    qryFindVendorNumInStartBucket: TUniQuery;
    tblTripStop: TUniTable;
    qryGetNewHireRecs: TUniQuery;
    qryGetFlightCrewNewHire: TUniQuery;
    qryUpdtStatus_CrewTrip_1: TUniQuery;
    qryUpdtStatus_CrewTrip_2: TUniQuery;
    qryUpdtStatus_CrewTail_1: TUniQuery;
    qryUpdtStatus_CrewTail_2: TUniQuery;
    qryUpdtStatus_CrewLog_1: TUniQuery;
    qryUpdtStatus_CrewLog_2: TUniQuery;
    connOnBase: TUniConnection;
    qryGetNewTailLeadPilotRecs: TUniQuery;
    edIFSPseudoUsers: TEdit;
    Label20: TLabel;
    qrySpecialUserOverride: TUniQuery;
    qryLoadCertifyEmployeesTable: TUniQuery;
    qryGetSpecialUsers: TUniQuery;
    qryInsertTripsForGroup: TUniQuery;
    qryInsertTailsForIFS: TUniQuery;
    qryGetTailTripDepartdate: TUniQuery;
    qryGetMissingFlightCrew: TUniQuery;
    qryGetJobCodeDescrips: TUniQuery;
    qryGetGroups: TUniQuery;
    qryGetCertifyDeptName: TUniQuery;
    qrySpecialUserDupes: TUniQuery;

    procedure btnMainClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnGoHourlyClick(Sender: TObject);
    procedure btnFixerClick(Sender: TObject);

  private
    Procedure Main();
    Procedure ImportPaycomData(Const BatchTimeIn : TDateTime);
    Procedure InsertIntoHistoryTable(Const slInputFileRec: TStringList; BatchTimeIn: TDateTime);
    Procedure BuildEmployeeFile(Const BatchTimeIn: TDateTime)  ;
    Procedure WriteToCertifyEmployeeFile(Const BatchTimeIn: TDateTime)   ;
    Procedure SplitEmployeeName( Const FullNameIn: String; Var LastNameOut, FirstNameOut : String );
    Procedure IdentifyNonCertifyRecs( Const BatchTimeIn : TDateTime );
    Procedure ValidateEmployeeRecords(Const BatchTimeIn: Tdatetime);
    Procedure UpdateDupeEmailRecs( Const EMailIn: String; BatchTimeIn: TDateTime);
    Procedure BuildCrewLogFile();
    Procedure BuildCrewTailFile(Const CurrBatchTimeIn: TDatetime);
    Procedure BuildCrewTripFile();
    Procedure BuildTripLogFile();

    Procedure BuildCrewDepartDateAirportFile();
    Procedure BuildTailTripDepartTimeAirport();

    Procedure BuildGenericValidationFile(const TargetFileName, SQLIn: String) ;
    Procedure BuildGenericValidationFile2(const TargetFileName, SQLIn: String) ;
    Procedure LoadTripsIntoStartBucket(Const BatchTimeIn : TDateTime);
    Procedure BuildValidationFiles(Const BatchTimeIn : TDateTime);
    Procedure CalculateApproverEmail(Const BatchTimeIn: TDateTime) ;
    Procedure FilterTripsByCount;
    Procedure FindPilotsNotInPaycom(Const BatchTimeIn : TDateTime);               // de-cruft, appears not to be called  ???JL  3 dec 2018
    Procedure DeleteTrip(Const LogSheetIn, QuoteNumIn : Integer; CrewMemberIDIn: String);

    Procedure WriteContractorsToPaycomTable(Const BatchTimeIn: TDateTime; NewHireFlag: Boolean; SourceQry: TUniQuery );

    Procedure ConnectToDB();
    Procedure SendStatusEmail;
    Procedure Send_NoIni_ErrorViaEmail(Const ErrorMsgIn: String);
    Procedure SendWarningViaEmail(Const ErrorMsgIn: String);
    Procedure CreateEmployeeErrorReport(Const BatchTimeIn : TDateTime) ;

    Procedure UpdateCCField(Const BatchTimeIn: TDateTime);

    Procedure LoadCertFileFields(Const BatchTime: TDateTime);

    Procedure BuildTailTripLogFile;


    Procedure Load_CharterVisa_IntoStartBucket;
    Procedure InsertCrewTail(Const TailNumIn:String; VendorNumIn: Integer; DataSourceIn: String);

    Procedure LoadTailLeadPilot2;

    Procedure FlagRecordAsError(Const ErrorType, ErrorMsgIn : String);

    Procedure Load_IFS_IntoStartBucket(BatchTime: TDateTime);

    Procedure HourlyPushMain;

    Procedure LoadData(Const BatchTimeIn: TDateTime);

    Procedure SendNewCrewTailToCertify(stlNewCrewTailIn: TStringList);


    Procedure SendToCertify(Const WorkingQueryIn: TUniQuery; Const BatchTimeIn : TDateTime; DataSetNameIn: String);

    Procedure GetBatchDates_CrewTail(Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);
    Procedure GetBatchDates_CrewTrip(Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);
    Procedure GetBatchDates_CrewLog( Var PreviousBatchDateOut, NewBatchDateOut : TDateTime);


    Procedure LogIt(ErrorMsgIn: String);

    Procedure Do_CrewTail_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);
    Procedure Do_CrewTrip_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);
    Procedure Do_CrewLog_API(Const BatchTimeIn, PreviousBatchDateIn, NewBatchDateIn : TDateTime);

    Procedure UpdateRecordStatus_CrewTail(Const RecStatusIn: String; Const A_BatchDateIn, B_BatchDateIn: TDateTime);
    Procedure UpdateRecordStatus_CrewTrip(const RecStatusIn: String; const A_BatchDateIn, B_BatchDateIn: TDateTime);
    Procedure UpdateRecordStatus_CrewLog(const  RecStatusIn: String; const A_BatchDateIn, B_BatchDateIn: TDateTime);

    Procedure Load_CrewTail_HistoryTable(Const BatchTimeIn: TDateTime);
    Procedure Load_CrewTrip_HistoryTable(Const BatchTimeIn: TDateTime);
    Procedure Load_CrewLog_HistoryTable(const BatchTimeIn: TDateTime);

    Procedure PruneHistoryTables(Const TableNameIn : String);

    // 7 Jun 2019
    Procedure AddDummyTripToStartBucket(Const VendorNumIn: Integer; NewHireKind: String) ;
    Procedure InsertDummyNewHireTripStop(Const QuoteNumIn : Integer);
    Procedure Process_NewHire_Employees_FlightCrew(Const BatchTimeIn: TDateTime);

    // 10 Jul 2019
    Procedure WriteToStartBucket(DataSetIn: TUniQuery; FieldNameIn: String);

    Procedure PurgeTable(Const TableNameIn: String);

    Procedure LoadNewTailLeadPilotRecs();
    Procedure WriteTailLeadPilotToFile;

    // 17 Dec 2019
    Procedure ImportSpecialUsers(Const BatchTimeIn: TDateTime);
    Procedure InsertSUIntoHistoryTable(BatchTimeIn: TDateTime);

    Procedure OverrideWithSpecialUsers(Const BatchTimeIn: TDateTime);

    Procedure FlagTerminatedEmployees(Const BatchTimeIn: TDateTime);

    Procedure  LoadCertifyEmployeesTable(Const BatchTimeIn: TDateTime);



    Function  GetTimeFromDBServer(): TDateTime;
    Function  RecIsValid(Const TimeStampIn:TDateTime; Const strValid_Roles, strValid_JobCodeDescrips, strValidGroups: String ): Boolean ;

    Function  CalcPilotName(SourceQueryIn: TUniQuery): String;
    Function  ScrubFARPart(Const FarPartIn: String): String;
    Function  ScrubCertifyDept(Const DepartmentIn : String) : String;
    Function  CalcPaycomErrorFileName(Const BatchTimeIn: TDateTime): String;


    Function CalcCertfileDepartmentName(Const GroupNameIn: String; Var DeptNameOut: String): Boolean ;

    Function FindLeadPilot(Const AssignedACString: String; BatchTimeIn: TDateTime): String;
    Function EmployeeTerminated(Const EMailIn: String; ImportDateIn: TDateTime): Boolean;

    Function GroupIsIn(Const GroupIn, GroupSetIn: String): Boolean;

    Function CalcCrewTailFileName(Const BatchTimeIn: TDateTime): String;

    // 7 Jun 2019
    Function FindVendorNumInStartBucket(Const VendorNumIn: Integer): Boolean;

    Function CalcInClause(Const GroupStrIn: String): String;

    Function ScrubVendorNum(Const strVendorNumIn: String; Var ErrorTxtOut: String): Integer;

  // TID:1112 15 May 2020
    Procedure GenerateMissingFlightCrewReport(Const BatchTimeIn: TDateTime);
    Procedure WriteQueryResultsToFile(SourceQueryIn: TUniQuery; FileNameOut: String);
    Procedure InsertSpecialUsersHistoryTable(const DataSetIn: TUniQuery; BatchTimeIn: TDateTime);

    Function GetValidRoles(): String;
    Function GetValidJobCodeDescrips(): String;
    Function GetValidGroups(): String;
    Function StripTrailingPipe(Const strEmpIDin : string): String ;

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

  gloFlightCrewGroup: String = 'FlightCrew|PoolPilot|PoolFA|FlightCrewCorp|FlightCrewNonPCal|IFS';

implementation

{$R *.dfm}

procedure TufrmCertifyExpDataLoader.FormCreate(Sender: TObject);
var
  iniFileName : String;

begin
  iniFileName := ExtractFilePath(ParamStr(0)) + 'CertifyExpDataLoader.ini';

  if not FileExists(iniFileName) then Begin
    LogIt('Error from FormCreate(): .ini file (' + iniFileName + ') not found!!');
    Send_NoIni_ErrorViaEmail('Cannot Find ini file: ' + iniFileName);
    MessageDlg( 'Ini File is missing! ' + #13 + 'Cannot find: ' + #13 + iniFileName, mtError, [mbOK], 0);
    Exit;
  End;

  myIni := TIniFile.Create( iniFileName );

  ConnectToDB;

  // Initialize form fields from ini
  edPayComInputFile.Text  := myIni.ReadString('Startup', 'PaycomFileName',   '') ;
  edOutputFileName.Text   := myIni.ReadString('Startup', 'CertifyEmployeeFileName', '') ;
  edOutputDirectory.Text  := myIni.ReadString('Startup', 'OutputDirectory', '') ;
//  edSpecialUsersFile.Text := myIni.ReadString('Startup', 'SpecialUsersFileName', '') ;
  edCharterVisaUsers.Text := myIni.ReadString('Startup', 'CharterVisaVendorNumbers', '') ;
  edIFSPseudoUsers.Text   := myIni.ReadString('Startup', 'IFSPseudoUsers', '') ;

  edDaysBack.Text           := myIni.ReadString('Startup', 'EmployeeFlightCrewDaysBack', '') ;
  edContractorDaysBack.Text := myIni.ReadString('Startup', 'ContractFlightCrewDaysBack', '') ;
  edLastNTrips.Text         := myIni.ReadString('Startup', 'MostRecentTripCount', '') ;
  edTerminatedDaysBack.Text := myIni.ReadString('Startup', 'TerminatedEmployeeGracePeriodDays', '') ;


  //  ShowMessage(ParamStr(1));
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

  LoadTailLeadPilot2;

  PruneHistoryTables('CertifyExp_PaycomHistory');

  LoadData(BatchTime);      // Loads PaycomHistory & StartBucket tables; also imports SpecialUsers file

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
  PreviousBatchDate, NewBatchDate : TDateTime;

begin
  BatchTime := GetTimeFromDBServer();          //StrToDateTime('03/25/2019 22:15:00');
  LoadData(BatchTime);                         // loads data into StartBucket & PaycomHistory tables

  gloPusher := TfrmPushToCertify.Create(self);
  gloPusher.theBaseURL := myIni.ReadString('CertifyAPI', 'BaseURL', '');    // 'https://api.certify.com/v1/exprptglds';
  gloPusher.APIKey     := myIni.ReadString('CertifyAPI', 'APIKey', '');     // 'qQjBp9xVQ36b7KPRVmkAf7kXqrDXte4k6PxrFQSv' ;
  gloPusher.APISecret  := myIni.ReadString('CertifyAPI', 'APISecret', '');  // '4843793A-6326-4F92-86EB-D34070C34CDC' ;

  // crew_tail
  PruneHistoryTables('CrewTail');
  Load_CrewTail_HistoryTable(BatchTime);                       // puts latest batch into CrewTailHistory table
  GetBatchDates_CrewTail(PreviousBatchDate, NewBatchDate);     // identifies "added" & "deleted" recs;
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

end;  { HourlyPushMain }


procedure TufrmCertifyExpDataLoader.LoadData(Const BatchTimeIn: TDateTime);
begin

  ImportPaycomData(BatchTimeIn);               // rec status: imported or error  6 Jan 2020

    FlagTerminatedEmployees(BatchTimeIn);      // checks Termination Date & updates record_status in PaycomHistory, was part of ValidateEmployeeRecord


  ImportSpecialUsers(BatchTimeIn);            //


  LoadTripsIntoStartBucket(BatchTimeIn);

    Load_CharterVisa_IntoStartBucket;

//  AddContractorsNotInPaycom(BatchTimeIn);       // writes Contractors to PaycomHistory table

//    Process_NewHire_Contractors(BatchTimeIn);     // writes New-Hire contractors to PaycomHistory, if they have not yet flown any trips
    Process_NewHire_Employees_FlightCrew(BatchTimeIn);

  UpdateCCField(BatchTimeIn);                   // Update Credit Card Field

  IdentifyNonCertifyRecs(BatchTimeIn);          // rec status: non-certify;     non-certify records flagged in record_status field

  ValidateEmployeeRecords(BatchTimeIn);         // rec status: OK


  LoadCertFileFields(BatchTimeIn);

//    ImportSpecialUsers(BatchTimeIn);            //  <---- new 6 Jan 2020  ???JL

  Load_IFS_IntoStartBucket(BatchTimeIn);

  FilterTripsByCount;                           // removes selected recs from StartBucket


  // Should this go here?  What is the chronology?
  GenerateMissingFlightCrewReport(BatchTimeIn);


end;  { LoadData() }


procedure TufrmCertifyExpDataLoader.btnFixerClick(Sender: TObject);
var
  BatchTime, PreviousBatchDate, NewBatchDate : TDateTime;

begin

  BatchTime := StrToDateTime('06/04/2020 09:45:01.500');

//  ImportSpecialUsers(BatchTime);
//  ValidateEmployeeRecords(BatchTime);

  LoadCertFileFields(BatchTime);

//  GenerateMissingFlightCrewReport(BatchTime);


end;


procedure TufrmCertifyExpDataLoader.btnGoHourlyClick(Sender: TObject);
begin

  HourlyPushMain;

end;


procedure TufrmCertifyExpDataLoader.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  try
    UniConnection1.Close;              //  ???JL  this still may be cause the program to hang on exit under some unknown conditions  14 may 2019

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

    LogIt(errorMsg);
    StatusBar1.Panels[0].Text :=  'DB: Error! - ' + dbName ;
//   Raise;
    end;
  end;

  // Connect to the OnBase (SQL Server) database
  try
    connOnBase.Connected := false;
    connOnBase.Server    := myIni.ReadString('DBConfig_OnBase', 'Server', '') ;
    connOnBase.Database  := myIni.ReadString('DBConfig_OnBase', 'DatabaseName', '') ;
    dbName := connOnBase.Database;
    connOnBase.Username  := myIni.ReadString('DBConfig_OnBase', 'UserName', '') ;
    connOnBase.Password  := myIni.ReadString('DBConfig_OnBase', 'Password', '') ;
    connOnBase.Connected := true;
  except on E: Exception do begin
    errorMsg := 'Error! from ConnectToDBs();  ' +
                'Problem connecting to database: ' + dbName + '; ' +
                E.ClassName + '; ' +
                E.Message ;

    LogIt(errorMsg);
    StatusBar1.Panels[0].Text :=  'DB: Error! - ' + dbName ;
    end;
  end;

  StatusBar1.Panels[0].Text :=  'Main DB: ' + UniConnection1.Database;

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
//  BuildTailTripFile;               depricated 24 Oct 2019
//  BuildTailLogFile;                          "

  BuildTailTripLogFile;

//  BuildTripAccountantFile(TargetDirectory + 'trip_accountant.csv');

  // Build Trip/Stop records
  scrLoadTripStopData.Execute;    // puts recs into working table CertifyExp_TripStop_Step1

    InsertDummyNewHireTripStop(818181);   // adds NewHire records for new contractors to working table CertifyExp_TripStop_Step1
    InsertDummyNewHireTripStop(828282);   // adds NewHire records for new employees to working table CertifyExp_TripStop_Step1

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
  LNameOut, FNameOut, DeptDisplayNameOut : String;

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
      FieldByName('certfile_group').AsString           := FieldByName('certify_department').AsString;     // already validated by RecIsValid()

      If CalcCertfileDepartmentName( FieldByName('certfile_group').AsString, DeptDisplayNameOut ) Then
        FieldByName('certfile_department_name').AsString := DeptDisplayNameOut
      else begin
        FieldByName('record_status').AsString := 'error';
        FieldByName('error_text').AsString   := DeptDisplayNameOut;
      end;

    end;

    CalculateApproverEmail(BatchTime);

    qryGetImportedRecs.Post;
    qryGetImportedRecs.Next;
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

    if GroupIsIn(strCertifyGroup, 'Corporate|FBO|Maintenance|Marketing|AircraftManagement|HR') then begin

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


    end else if GroupIsIn(strCertifyGroup, 'DOM') then begin


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


//    end else if GroupIsIn(strCertifyGroup, '|FlightCrew|PoolPilot|PoolFA|FlightCrewCorp|FlightCrewNonPCal|IFS|') then begin
    end else if GroupIsIn(strCertifyGroup, gloFlightCrewGroup) then begin

(* Note 13: This logic is now handled by Certify "Workflows" within the Certify system      4Feb2019 -JL
      //  Assign Approver1 Email
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := strAccountantEmail;

      //  Assign Approver2 Email
      if qryGetImportedRecs.FieldByName('employee_code').AsString = 'contractor'  then
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := strAccountantEmail
      else
        qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := '';
*)


      // Assign Accountant Email
      if qryGetImportedRecs.FieldByName('has_credit_card').AsString = 'T' then
        strAccountantEmail := 'FlightCrewCC@ClayLacy.com'
      else
        strAccountantEmail := 'FlightCrew@ClayLacy.com';

      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := strAccountantEmail;

(* 15May2019 - added per "Phase3 Tasks & Estimates" item #21 *)
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString  := strAccountantEmail;
      qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString  := strAccountantEmail;


    end else if GroupIsIn(strCertifyGroup, 'CharterVisa') then begin

      // Assign Accountant Email
      qryGetImportedRecs.FieldByName('certfile_accountant_email').AsString := 'FlightCrewCC@ClayLacy.com';

(* See Note 13, above
      //  Assign Approver1 Email
      qryGetImportedRecs.FieldByName('certfile_approver1_email').AsString := 'CorporateCC@ClayLacy.com';   ???JL just get whatever is in paycom

      //  Assign Approver2 Email
      qryGetImportedRecs.FieldByName('certfile_approver2_email').AsString := 'CorporateCC@ClayLacy.com';
*)


    end else begin       // error, unknown Certify Department


      FlagRecordAsError('error', 'unknown certify_department/Group: ' + strCertifyGroup);


    end;

end;  { CalculateApproverEmail }




function TufrmCertifyExpDataLoader.GroupIsIn(const GroupIn, GroupSetIn: String): Boolean;
var
  stlGroupSet : TStringList;

begin
  Result := False;

  stlGroupSet := TStringList.Create;
  try
    stlGroupSet.Delimiter := '|';
    stlGroupSet.DelimitedText := GroupSetIn;    // GroupSetIn looks like: 'Corporate|FBO|Maintenance|Marketing|AircraftManagement|HR'
    Result := stlGroupSet.IndexOf(GroupIn) > -1;
  finally
    stlGroupSet.Free;
  end;
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

  LoadCertifyEmployeesTable(BatchTimeIn);

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
    slEmpRec.add( qryGetEmployees.FieldByName('certfile_employee_id').AsString + '|') ;    // adding pipe to fix "1398" bug - 8 Jul 2019  JL
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

end; { WriteToCertifyEmployeeFile }


procedure TufrmCertifyExpDataLoader.IdentifyNonCertifyRecs( Const BatchTimeIn : TDateTime );
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Flagging Non-Certify Records in PaycomHistory ';
  Application.ProcessMessages;

  qryIdentifyNonCertifyRecs.close;
  qryIdentifyNonCertifyRecs.ParamByName('parmImportDate').AsDateTime := BatchTimeIn;
  qryIdentifyNonCertifyRecs.Execute;

end;


procedure TufrmCertifyExpDataLoader.ImportPaycomData(Const BatchTimeIn : TDateTime);
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

  Readln(FileIn, s);            // skip the first line of the file which is the field names header, not data
  while not Eof(FileIn) do begin
    Readln(FileIn, s);
    sl.CommaText := s;
    InsertIntoHistoryTable(sl, BatchTimeIn);
  end;

  tblPayComHistory.close;
  CloseFile(FileIn);
  sl.Free;

end;  {PaycomImportMain }



procedure TufrmCertifyExpDataLoader.InsertIntoHistoryTable(const slInputFileRec: TStringList; BatchTimeIn: TDateTime);  // ???JL remove?
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
    tblPaycomHistory.FieldByName('paycom_assigned_ac_2').AsString   := slInputFileRec[14];
    tblPaycomHistory.FieldByName('record_status').AsString          := recStatus ;
    tblPaycomHistory.FieldByName('status_timestamp').AsDateTime     := BatchTimeIn;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime          := BatchTimeIn;
    tblPaycomHistory.FieldByName('data_source').AsString            := 'paycom_file';
    tblPaycomHistory.FieldByName('hire_date').AsString              := slInputFileRec[15];
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



procedure TufrmCertifyExpDataLoader.InsertSpecialUsersHistoryTable(const DataSetIn: TUniQuery; BatchTimeIn: TDateTime);
var
  recStatus : String;
  i : integer;
  strEmpID: String;

begin

  try
    tblPayComHistory.Insert;
    recStatus := 'imported';
    tblPaycomHistory.FieldByName('employee_name').AsString := DataSetIn.FieldByName('last_name').AsString + ', ' + DataSetIn.FieldByName('first_name').AsString ;
    tblPaycomHistory.FieldByName('work_email').AsString    := DataSetIn.FieldByName('work_email').AsString;

    if DataSetIn.FieldByName('employee_id').AsString <> '' then begin
      try
        strEmpID := StripTrailingPipe(DataSetIn.FieldByName('employee_id').AsString);
        tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger := StrToInt(strEmpID);
      except on E1: Exception do begin
        recStatus := 'error';
        tblPaycomHistory.FieldByName('error_text').AsString := tblPaycomHistory.FieldByName('error_text').AsString + '; Field: certify_gp_vendornum - ' + E1.Message;
      end;
      end;
    end else begin
//      tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger := '' ;
    end;

    tblPaycomHistory.FieldByName('certify_department').AsString     := DataSetIn.FieldByName('group').AsString;
    tblPaycomHistory.FieldByName('certify_role').AsString           := DataSetIn.FieldByName('employee_type').AsString;
    tblPaycomHistory.FieldByName('paycom_approver1_email').AsString := DataSetIn.FieldByName('approver_email_1').AsString;
    tblPaycomHistory.FieldByName('paycom_approver2_email').AsString := DataSetIn.FieldByName('approver_email_2').AsString;
    tblPaycomHistory.FieldByName('paycom_assigned_ac').AsString     := DataSetIn.FieldByName('assigned_aircraft').AsString;
    tblPaycomHistory.FieldByName('accountant_email').AsString       := DataSetIn.FieldByName('accountant_email').AsString;

    tblPaycomHistory.FieldByName('record_status').AsString          := recStatus ;
    tblPaycomHistory.FieldByName('status_timestamp').AsDateTime     := BatchTimeIn;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime          := BatchTimeIn;
    tblPaycomHistory.FieldByName('data_source').AsString            := 'special_users';
    tblPaycomHistory.post;

  except on E: Exception do begin
    tblPaycomHistory.Edit;
    tblPaycomHistory.FieldByName('record_status').AsString := 'error';
    tblPaycomHistory.FieldByName('error_text').AsString    := tblPaycomHistory.FieldByName('error_text').AsString + '; ' + E.Message;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime := BatchTimeIn;
    tblPaycomHistory.FieldByName('data_source').AsString   := 'special_users';
    tblPaycomHistory.post;
  end;

  end;  { Try/Except }

end;  { InsertSpecialUsersHistoryTable() }



{  This new procedure gets Tail/LeadPilot data from OnBase/Workbench instead of from a manually-imported .csv file
  1. open local tail_leadpilot table
  2. query OnBase to get latest Tail/LeadPilot data
  3. if record counts within tolerance then empty CertifyExp_Tail_LeadPilot & add new recs
  4. write tail_leadpilot.csv to output directory
}
procedure TufrmCertifyExpDataLoader.LoadTailLeadPilot2;
var
  CurrentCount, NewCount : Integer;
  TolerancePercent : Real;

begin
  tblTailLeadPilot.Open;               // This is the local (Warehouse) table
  CurrentCount := tblTailLeadPilot.RecordCount;

  qryGetNewTailLeadPilotRecs.Close;    // This is the external (OnBase) table
  qryGetNewTailLeadPilotRecs.Open;
  NewCount := qryGetNewTailLeadPilotRecs.RecordCount;

  TolerancePercent := myIni.ReadInteger('Startup','TailLeadPilotCountTolerance', 20) / 100;

  // If change in record count is within Tolerance
  if Abs(CurrentCount - NewCount) < Round((CurrentCount * TolerancePercent)) then begin
    LoadNewTailLeadPilotRecs();
  end else begin
    SendWarningViaEmail('WARNING - Large delta in Tail_LeadPilot record count!' + #13#13 +
                      'Existing count: ' + IntToStr(CurrentCount) + ', New count: ' + IntToStr(NewCount) + #13#13 +
                      'If this is OK and you want to import this data, then increase the ' + QuotedStr('TailLeadPilotCountTolerance') + ' parameter in the .ini file.' + #13 +
                      'The changed value will be used on the next Nightly or Morning run. ' + #13#13 +
                      '(The Upload process ran using the existing data and did not import the new data.)');
  end;

  WriteTailLeadPilotToFile;

  tblTailLeadPilot.Close;
  qryGetNewTailLeadPilotRecs.Close;

end; { LoadTailLeadPilot2 }


procedure TufrmCertifyExpDataLoader.LoadNewTailLeadPilotRecs;
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Loading New Tail/LeadPilot recs from OnBase' ;

  tblTailLeadPilot.EmptyTable;
  qryGetNewTailLeadPilotRecs.First;
  while not qryGetNewTailLeadPilotRecs.eof do begin
    tblTailLeadPilot.Append;
    tblTailLeadPilot.FieldByName('Tail').AsString  := qryGetNewTailLeadPilotRecs.FieldByName('tail_number').AsString;
    tblTailLeadPilot.FieldByName('Email').AsString := qryGetNewTailLeadPilotRecs.FieldByName('lead_pilot_email').AsString;
    tblTailLeadPilot.Post;
    qryGetNewTailLeadPilotRecs.Next;
  end;

end;  { LoadNewTailLeadPilotRecs }


procedure TufrmCertifyExpDataLoader.WriteTailLeadPilotToFile;
var
  WorkFile: TextFile;
  RowOut : string;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing tail_leadpilot.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'tail_leadpilot.csv');
  Rewrite(WorkFile);

  RowOut := 'Tail,EMail';      // write header record
  WriteLn(WorkFile, RowOut) ;
  tblTailLeadPilot.First;
  while not tblTailLeadPilot.eof do begin
    RowOut := Trim(tblTailLeadPilot.FieldByName('Tail').AsString) + ',' + tblTailLeadPilot.FieldByName('EMail').AsString ;
    WriteLn(WorkFile, RowOut) ;
    tblTailLeadPilot.Next;
  end;

  CloseFile(WorkFile);

end;  { WriteTailLeadPilotToFile }



procedure TufrmCertifyExpDataLoader.LoadTripsIntoStartBucket(Const BatchTimeIn : TDateTime);
var
  strDaysBack: String;

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
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.PICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, L.DEPARTURE, L.ARRIVEID, L.LEGNO');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('  and L.PICPILOTNO > 0');
  qryLoadTripData.SQL.Append('  and L.LEGNO = 1 ');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load SIC data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.SICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, L.DEPARTURE, L.ARRIVEID, L.LEGNO');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.SICPILOTNO > 0');
//  qryLoadTripData.SQL.Append('  and L.LEGNO = 1 ');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load TIC data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.TICPILOTNO, T.QUOTENO, L.ACREGNO, FARPART, 0, L.DEPARTURE, L.ARRIVEID, L.LEGNO');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.TICPILOTNO > 0');
//  qryLoadTripData.SQL.Append('  and L.LEGNO = 1 ');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  // Load FA (Flight Attendant) data into StartBucket
  qryLoadTripData.SQL.Append('insert into CertifyExp_Trips_StartBucket');
  qryLoadTripData.SQL.Append('select distinct L.LOGSHEET, L.FANO, T.QUOTENO, L.ACREGNO, FARPART, 0, L.DEPARTURE, L.ARRIVEID, L.LEGNO');
  qryLoadTripData.SQL.Append('from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)');
  qryLoadTripData.SQL.Append('where L.DEPARTURE > (CURRENT_TIMESTAMP - ' + strDaysBack + ')');
  qryLoadTripData.SQL.Append('and L.FANO > 0');
//  qryLoadTripData.SQL.Append('  and L.LEGNO = 1 ');
  qryLoadTripData.Execute;
  qryLoadTripData.SQL.Clear;

  qryGetAirCrewVendorNum.Execute;

  //  T&E-25  Adding default crew_tail recs for employees & T&E:25b
(*
 D 1. import new field from paycom_employees.csv
   2. run WriteToStartBucket twice w/ field name as param
 D 3. add paycom_assigned_ac_2 field to PaycomHistory table
*)
  qryLoadTripData.SQL.Clear;
  qryLoadTripData.SQL.Append('SELECT certify_gp_vendornum, certify_department, certfile_group, paycom_assigned_ac, paycom_assigned_ac_2');
  qryLoadTripData.SQL.Append('FROM CertifyExp_PaycomHistory' );
  qryLoadTripData.SQL.Append('WHERE imported_on = ' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:nn:ss.zzz', BatchTimeIn) ) );   // get current Batch
  qryLoadTripData.SQL.Append('  AND certify_department IN ( ' + CalcInClause(gloFlightCrewGroup) + ' )' );                    // build IN clause
  qryLoadTripData.SQL.Append('  AND not (paycom_assigned_ac = ' + QuotedStr('') + ' or paycom_assigned_ac is null ) ');       // assigned_ac field not blank
  qryLoadTripData.SQL.Append('  AND termination_date IS NULL' );                                                              // only currently employed people
//  ShowMessage(qryLoadTripData.SQL.Text);
  DataSource1.DataSet := qryLoadTripData;
  qryLoadTripData.Open;

  WriteToStartBucket(qryLoadTripData, 'paycom_assigned_ac');
  WriteToStartBucket(qryLoadTripData, 'paycom_assigned_ac_2');
  qryLoadTripData.Close;

end;  { LoadTripsIntoStartBucket }



// The data is a little goofy here.
//   The assigned_ac field (in the PaycomHistory table) can contain multiple aircraft separated by '/'
//    like this:  'N880JM/N4D/N8241W/N345CL' or just a single aircraft: 'N880MJ'
// This code creates a new record in the StartBucket table for each crew_member, tail combination (crew_member identified by VendorNum)
//                                                                                       - 10 Jul 2019  JL, Jeff@dcsit.com
procedure TufrmCertifyExpDataLoader.WriteToStartBucket(DataSetIn: TUniQuery; FieldNameIn: String);
var
  stlAssignedAC : TStringList;
  i : Integer;

begin
  tblStartBucket.Open;
  stlAssignedAC := TStringList.Create;
  try
    stlAssignedAC.Delimiter := '/';
    stlAssignedAC.StrictDelimiter := true;      { tell stringList to not use space as delimeter }

    DataSetIn.First;
    while not DataSetIn.eof do begin
      stlAssignedAC.Clear;
      stlAssignedAC.DelimitedText := DataSetIn.fieldByName(FieldNameIn).AsString;   // parse multiple tails (if any) into StringList

      for i := 0 to stlAssignedAC.Count - 1 do begin
        InsertCrewTail(stlAssignedAC[i], DataSetIn.FieldByName('certify_gp_vendornum').AsInteger, 'DefaultCrewTail');   // Writes to tblStartBucket
      end;

      DataSetIn.next;
    end;

  finally
    stlAssignedAC.Free;
  end;
  tblStartBucket.Close;

end;  { WriteToStartBucket }



function TufrmCertifyExpDataLoader.CalcInClause(const GroupStrIn: String): String;
var
  stlGroupString: TStringList;
  i : Integer;
  strOut : String;

begin
  stlGroupString := TStringList.Create;
  stlGroupString.Delimiter := '|';
  stlGroupString.DelimitedText := GroupStrIn;

  for i := 0 to stlGroupString.Count - 1 do begin
    strOut := StrOut + QuotedStr(stlGroupString[i]) + ', ';
  end;

  Result := Copy(strOut, 0, Length(strOut) - 2);  // get rid of trailing comma & space

  stlGroupString.Free;

end;  { CaldInClause }



function TufrmCertifyExpDataLoader.RecIsValid(Const TimeStampIn:TDateTime; Const strValid_Roles, strValid_JobCodeDescrips, strValidGroups: String): Boolean;
var
  strErrorText : String;

begin
  strErrorText := '';
  Result := True;

  // don't validate job_code_descrip for Special Users, since it is not supplied from the Special Users data source
  if qryGetEmployees.FieldByName('data_source').AsString <> 'special_users' then Begin
    if Pos( '|' + qryGetEmployees.FieldByName('job_code_descrip').AsString + '|', strValid_JobCodeDescrips) = 0 then
      strErrorText := strErrorText + 'invalid job_code_descrip; ';
  End;


  if qryGetEmployees.FieldByName('certify_gp_vendornum').AsString = '' then
    strErrorText := strErrorText + 'missing certify_gp_vendornum; ';

  if Pos( '|' + qryGetEmployees.FieldByName('certify_role').AsString + '|', strValid_Roles) = 0 then
    strErrorText := strErrorText + 'invalid certify_role; ';

  if Pos( '|' + qryGetEmployees.FieldByName('certify_department').AsString + '|', strValidGroups) = 0 then     // certify_department aka "group"
    strErrorText := strErrorText + 'invalid certify_department(group); ';



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


procedure TufrmCertifyExpDataLoader.ValidateEmployeeRecords(const BatchTimeIn: Tdatetime);
var
  Time_Stamp :  TDateTime;
  strValid_Roles : String;
  strValid_JobCodeDescrips : String;
  strValidGroups : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Validating Employee Records';
  Application.ProcessMessages;

  Time_Stamp := GetTimeFromDBServer();

  strValid_Roles           := GetValidRoles();
  strValid_JobCodeDescrips := GetValidJobCodeDescrips();
  strValidGroups           := GetValidGroups();

  // check for valid Certify recs
  qryGetEmployees.Close;
  qryGetEmployees.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn;
  qryGetEmployees.ParamByName('parmRecordStatusIn').AsString := 'imported';
  qryGetEmployees.Open;

  while not qryGetEmployees.eof do begin
    RecIsValid(Time_Stamp, strValid_Roles, strValid_JobCodeDescrips, strValidGroups);
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

end;  { ValidateEmployeeRecords }


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



procedure TufrmCertifyExpDataLoader.BuildCrewDepartDateAirportFile;
Var
  RowOut : String;
  WorkFile : TextFile;
  strDateDest : String;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing crew_departdate_airport.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'crew_departdate_airport.csv');
  Rewrite(WorkFile);

  qryBuildValFile.Close;
  qryBuildValFile.SQL.Clear;
  qryBuildValFile.SQL.Text := 'select distinct TripDepartDate, FirstDestination, CrewMemberVendorNum from CertifyExp_Trips_StartBucket where CrewMemberVendorNum is not null and TripDepartDate is not null order by TripDepartDate';
  qryBuildValFile.Open ;

  RowOut := 'DepartDate_Airport, CrewMemberVendorNum';
  WriteLn(WorkFile, RowOut) ;
  while ( not qryBuildValFile.eof ) do begin
    strDateDest := Trim(FormatDateTime('mm/dd/yy', qryBuildValFile.FieldByName('TripDepartDate').AsDateTime)) + '_' + qryBuildValFile.FieldByName('FirstDestination').AsString;
    RowOut := strDateDest + ',' +
              qryBuildValFile.FieldByName('CrewMemberVendorNum').AsString + '|' + strDateDest;

    WriteLn(WorkFile, RowOut) ;
    qryBuildValFile.Next;
  end;

  CloseFile(WorkFile);
  qryBuildValFile.Close;

end;  { BuildCrewDepartDateAirportFile }





procedure TufrmCertifyExpDataLoader.Load_CrewTail_HistoryTable(Const BatchTimeIn: TDateTime);
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

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing crew_tail.csv'  ;

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



procedure TufrmCertifyExpDataLoader.BuildTailTripDepartTimeAirport;
Var
  RowOut : String;
  WorkFile : TextFile;

begin
  StatusBar1.Panels[1].Text := 'Current Task:  Writing tail_trip_departdate_airport.csv'  ;
  Application.ProcessMessages;

  AssignFile(WorkFile, edOutputDirectory.Text + 'tail_trip_departdate_airport.csv');
  Rewrite(WorkFile);
  RowOut := 'TailNum,TripNum,Departdate_Airport';
  WriteLn(WorkFile, RowOut) ;

  qryGetTailTripDepartdate.Close;
  qryGetTailTripDepartdate.Open;
  while Not qryGetTailTripDepartdate.eof do Begin
    RowOut := Trim(qryGetTailTripDepartdate.FieldByName('TailNum').AsString) + ',' +
                   qryGetTailTripDepartdate.FieldByName('QuoteNum').AsString + ',' +
                   FormatDateTime('mm/dd/yy', qryGetTailTripDepartdate.FieldByName('TripDepartDate').AsDateTime) + '_' +qryGetTailTripDepartdate.FieldByName('FirstDestination').AsString;
//                   qryGetTailTripDepartdate.FieldByName('TripDepartDate').AsString + '_' +qryGetTailTripDepartdate.FieldByName('FirstDestination').AsString;

    WriteLn(WorkFile, RowOut) ;
    qryGetTailTripDepartdate.Next;
  end;

  CloseFile(WorkFile);
  qryGetTailTripDepartdate.Close;

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


procedure TufrmCertifyExpDataLoader.WriteQueryResultsToFile(SourceQueryIn: TUniQuery; FileNameOut: String);
var
  WorkFile: TextFile;
  RowOut: String;
  i: Integer;

begin
  RowOut := '';
  AssignFile(WorkFile, FileNameOut);
  Rewrite(WorkFile);

  // Write column header row
  for i := 0 to SourceQueryIn.Fields.Count - 1 do Begin
    RowOut := RowOut + SourceQueryIn.Fields[i].FieldName + ',';
  end;
  RowOut := Copy(RowOut, 1, RowOut.Length - 1);   // Remove trailing comma
  WriteLn(WorkFile, RowOut) ;
  RowOut := '';

  // Write data rows
  while not SourceQueryIn.eof do begin
    for i := 0 to SourceQueryIn.Fields.Count - 1 do Begin
      RowOut := RowOut + SourceQueryIn.Fields[i].AsString + ',';
    end;
    RowOut := Copy(RowOut, 1, RowOut.Length - 1);   // Remove trailing comma

    WriteLn(WorkFile, RowOut) ;
    RowOut := '';
    SourceQueryIn.Next;
  end;

  CloseFile(WorkFile);

end;  { WriteQueryResultsToFile }



function TufrmCertifyExpDataLoader.CalcCertfileDepartmentName(const GroupNameIn: String; Var DeptNameOut: String): Boolean;
var
  slGroupToDropdown : TStringList;
  FoundIndex : Integer;

begin
  Result := True;

  qryGetCertifyDeptName.Close;
  qryGetCertifyDeptName.ParamByName('parmGroupIn').AsString := GroupNameIn;
  qryGetCertifyDeptName.Open;

  if qryGetCertifyDeptName.RecordCount > 0 then
    DeptNameOut := qryGetCertifyDeptName.FieldByName('certify_department_display_name').AsString
  else Begin
    Result      := False;
    DeptNameOut := QuotedStr(GroupNameIn) + ' - certify_deparment value not found in JobCode Lookup table';
  End;

  qryGetCertifyDeptName.Close;

(*
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
*)

end;


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


procedure TufrmCertifyExpDataLoader.GenerateMissingFlightCrewReport(Const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Generating Missing Flight Crew Report.. ' ;
  Application.ProcessMessages;

  qryGetMissingFlightCrew.Close;
  qryGetMissingFlightCrew.ParamByName('parmImportDateIn').AsDateTime := BatchTimeIn ;
  qryGetMissingFlightCrew.ParamByName('parmDaysBack').AsInteger       := StrToInt(edContractorDaysBack.text);
  qryGetMissingFlightCrew.Open;

  WriteQueryResultsToFile(qryGetMissingFlightCrew, 'MissingFlightCrewTest.txt');
  qryGetMissingFlightCrew.Close;

end;  { GenerateMissingFlightCrewReport }





procedure TufrmCertifyExpDataLoader.Process_NewHire_Employees_FlightCrew(const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Processing New Hire Employees' ;
  Application.ProcessMessages;

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


procedure TufrmCertifyExpDataLoader.WriteContractorsToPaycomTable( Const BatchTimeIn: TDateTime; NewHireFlag: Boolean; SourceQry: TUniQuery ) ;
begin

  tblPaycomHistory.Insert;
  tblPaycomHistory.FieldByName('employee_code').AsString           := 'contractor';     // should we use parameter to signify Contractor New Hire  ???JL
  tblPaycomHistory.FieldByName('employee_name').AsString           := CalcPilotName(SourceQry);
  tblPaycomHistory.FieldByName('work_email').AsString              := SourceQry.FieldByName('EMail').AsString;
  tblPaycomHistory.FieldByName('position').AsString                := SourceQry.FieldByName('JobTitle').AsString;
  tblPaycomHistory.FieldByName('department_descrip').AsString      := 'Designated-' + SourceQry.FieldByName('AssignedAC').AsString ;  //CalcDeptDescrip ;
  tblPaycomHistory.FieldByName('job_code_descrip').AsString        := 'Pilot-Designated' ;
  tblPaycomHistory.FieldByName('supervisor_primary_code').AsString := '';
  tblPaycomHistory.FieldByName('certify_gp_vendornum').AsInteger   := SourceQry.FieldByName('VendorNumber').AsInteger;
  tblPaycomHistory.FieldByName('certify_department').AsString      := 'FlightCrew';
  tblPaycomHistory.FieldByName('certify_role').AsString            := 'Manager';

  tblPaycomHistory.FieldByName('certfile_approver1_email').AsString  := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('certfile_approver2_email').AsString  := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('accountant_email').AsString          := 'FlightCrew@ClayLacy.com';
  tblPaycomHistory.FieldByName('certfile_accountant_email').AsString := 'FlightCrew@ClayLacy.com';

  tblPaycomHistory.FieldByName('paycom_assigned_ac').AsString      := SourceQry.FieldByName('AssignedAC').AsString ;

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


function TufrmCertifyExpDataLoader.CalcPilotName(SourceQueryIn: TUniQuery): String;
begin
  Result := SourceQueryIn.FieldByName('LastName').AsString + ',' + SourceQueryIn.FieldByName('FirstName').AsString  ;

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
  myMessage.Recipients.EMailAddresses := myIni.ReadString('OutputFiles', 'EMailRecipientList', '');

  //  Load Attachments
  OutPutFileDir  := edOutputDirectory.Text;          // myIni.ReadString('OutputFiles', 'OutputFileDir', '');
  stlOutputFiles := TStringList.Create();
  stlOutputFiles.CommaText := myIni.ReadString('OutputFiles', 'EmailAttachFileList', '');

  // Add main attachments
  for i := 0 to stlOutputFiles.Count - 1 do begin
   if FileExists( OutPutFileDir + stlOutputFiles[i] ) then
     TIDAttachmentFile.Create(myMessage.MessageParts, OutPutFileDir + stlOutputFiles[i] );
  end ; { for }


  // Add the Paycom Data file to attachemnt list
  stlOutputFiles.Clear;
  stlOutputFiles.CommaText := myIni.ReadString('OutputFiles', 'EmailAttachFileList_Additional', '');
  for i := 0 to stlOutputFiles.Count - 1 do begin
   if FileExists( stlOutputFiles[i] ) then
     TIDAttachmentFile.Create(myMessage.MessageParts, stlOutputFiles[i] );
  end ; { for }


  // Add the Paycom Error file to attachemnt list
  if FileExists( gloPaycomErrorFile ) then
    TIDAttachmentFile.Create(myMessage.MessageParts, gloPaycomErrorFile );

  mySMTP := TIdSMTP.Create(nil);
  mySMTP.Host     := myIni.ReadString('Startup', 'EmailServer', '192.168.1.73') ;
  mySMTP.Username := myIni.ReadString('Startup', 'EmailServerLogin', 'tkvassay@claylacy.com') ;
  mySMTP.Password := myIni.ReadString('Startup', 'EmailServerLoginPW', '') ;

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
   //  mySMTP.Disconnect();
  Except on E:Exception Do
    LogIt('Email Error: ' + E.Message);
  End;

  mySMTP.free;
  myMessage.Free;
  stlOutputFiles.Free;

end;  { SendStatusEmail }


procedure TufrmCertifyExpDataLoader.Send_NoIni_ErrorViaEmail(const ErrorMsgIn: String);
Var
  mySMTP    : TIdSMTP;
  myMessage : TIDMessage;

begin
  myMessage := TIdMessage.Create(nil);
  myMessage.Subject      := 'CLA Certify Data Loader - Error!!';
  myMessage.From.Address := 'CertifyDataLoader@claylacy.com';
  myMessage.Body.Text    := ErrorMsgIn ;

  //  Hard-coding params because the error being trapped is 'cannot find .ini file' which contains these params
  myMessage.Recipients.EMailAddresses := 'Jeff@dcsit.com,LTaylor@ClayLacy.com,DLittlefield@ClayLacy.com,thomasfduffy@gmail.com' ;
  mySMTP := TIdSMTP.Create(nil);
  mySMTP.Host     :=  '192.168.1.73' ;
  mySMTP.Username :=  'tkvassay@claylacy.com' ;
  mySMTP.Password :=  '' ;

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
  Except on E:Exception Do
    LogIt('Email Error: ' + E.Message);
  End;

  mySMTP.free;
  myMessage.Free;

end;  { SendErrorViaEmail }


procedure TufrmCertifyExpDataLoader.SendWarningViaEmail(const ErrorMsgIn: String);
Var
  mySMTP    : TIdSMTP;
  myMessage : TIDMessage;

begin
  myMessage := TIdMessage.Create(nil);
  myMessage.Subject      := 'CLA Certify Data Loader - Warning!';
  myMessage.From.Address := 'CertifyDataLoader@claylacy.com';
  myMessage.Body.Text    := ErrorMsgIn ;
  myMessage.Recipients.EMailAddresses := myIni.ReadString('OutputFiles', 'EMailRecipientList', '');

  mySMTP := TIdSMTP.Create(nil);
  mySMTP.Host     := myIni.ReadString('Startup', 'EmailServer', '192.168.1.73') ;
  mySMTP.Username := myIni.ReadString('Startup', 'EmailServerLogin', 'tkvassay@claylacy.com') ;
  mySMTP.Password := myIni.ReadString('Startup', 'EmailServerLoginPW', '') ;

  Try
    mySMTP.Connect;
    mySMTP.Send(myMessage);
  Except on E:Exception Do
    LogIt('Email Error: ' + E.Message);
  End;

  mySMTP.free;
  myMessage.Free;

end;  { SendWarningViaEmail }


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

begin
  DecodeDate(Trunc(BatchTimeIn), myYear, myMonth, myDay);

  strMonth := IntToStr(myMonth);
  if myMonth < 10 then
    strMonth := '0' + strMonth ;

  strDay := IntToStr(myDay);
  if myDay < 10 then
    strDay := '0' + strDay;


  Result := myIni.ReadString('OutputFiles', 'CrewTailDestination', ''); // + '.csv'  ;

end;  { CalcCrewTailFileName }


procedure TufrmCertifyExpDataLoader.ImportSpecialUsers(const BatchTimeIn: TDateTime);
begin
  StatusBar1.Panels[1].Text := 'Current Task:  Importing Special Users file' ;
  Application.ProcessMessages;

  tblPaycomHistory.Open;
  try
    qryGetSpecialUsers.Close;
    qryGetSpecialUsers.Open;

    while not qryGetSpecialUsers.eof do begin
//      InsertSUIntoHistoryTable(BatchTimeIn);

      InsertSpecialUsersHistoryTable(qryGetSpecialUsers, BatchTimeIn);

      qryGetSpecialUsers.Next;
    end;

  finally
    tblPaycomHistory.Close;
    qryGetSpecialUsers.Close;

  end;
//  OverrideWithSpecialUsers(BatchTimeIn);  depricated

end;  {ImportSpecialUsers}



procedure TufrmCertifyExpDataLoader.OverrideWithSpecialUsers(const BatchTimeIn: TDateTime);
begin

  qrySpecialUserOverride.Close;
  qrySpecialUserOverride.ParamByName('parmImportedOn').AsDateTime := BatchTimeIn;
  qrySpecialUserOverride.Execute;

end;




(*
special_users file columns:

0   work_email,           [certfile_work_email]
1   first_name,           [certfile_first_name]
2   last_name,            [certfile_last_name]
3   employee_id,          [certfile_employee_id]
4   employee_type,        [certfile_employee_type]
5   group,                [certfile_group]
6   department_name,      [certfile_department_name]
7   approver_email - 1,   [certfile_approver1_email]
8   approver_email - 2,   [certfile_approver2_email]
9   accountant_email      [certfile_accountant_email]


PaycomHistory Table Field Names:

0   ,[certfile_work_email]
1   ,[certfile_first_name]
2   ,[certfile_last_name]
3   ,[certfile_employee_id]
4   ,[certfile_employee_type]
5   ,[certfile_group]
6   ,[certfile_department_name]
7   ,[certfile_approver1_email]
8   ,[certfile_approver2_email]
9   ,[certfile_accountant_email]

*)

procedure TufrmCertifyExpDataLoader.InsertSUIntoHistoryTable(BatchTimeIn: TDateTime);
var
  strRecStatus : String;
  strErrorTextOut : String;
  VendorNum: Integer;

begin
  try
    tblPayComHistory.Insert;
    tblPaycomHistory.FieldByName('data_source').AsString        := 'special_users';
    tblPaycomHistory.FieldByName('status_timestamp').AsDateTime := BatchTimeIn;
    tblPaycomHistory.FieldByName('imported_on').AsDateTime      := BatchTimeIn;

    strRecStatus := 'OK';    // set to OK so no other subsequent validation happens to these recs

    tblPaycomHistory.FieldByName('certfile_work_email').AsString       := qryGetSpecialUsers.FieldByName('work_email').AsString;
    tblPaycomHistory.FieldByName('certfile_first_name').AsString       := qryGetSpecialUsers.FieldByName('first_name').AsString;
    tblPaycomHistory.FieldByName('certfile_last_name').AsString        := qryGetSpecialUsers.FieldByName('last_name').AsString;
    tblPaycomHistory.FieldByName('certfile_employee_type').AsString    := qryGetSpecialUsers.FieldByName('employee_type').AsString;
    tblPaycomHistory.FieldByName('certfile_group').AsString            := qryGetSpecialUsers.FieldByName('group').AsString;
    tblPaycomHistory.FieldByName('certfile_department_name').AsString  := qryGetSpecialUsers.FieldByName('department_name').AsString;   //???JL should I call CalcCertfileDeptName()
    tblPaycomHistory.FieldByName('certfile_approver1_email').AsString  := qryGetSpecialUsers.FieldByName('approver_email_1').AsString;
    tblPaycomHistory.FieldByName('certfile_approver2_email').AsString  := qryGetSpecialUsers.FieldByName('approver_email_2').AsString;
    tblPaycomHistory.FieldByName('certfile_accountant_email').AsString := qryGetSpecialUsers.FieldByName('accountant_email').AsString;

    VendorNum := ScrubVendorNum(qryGetSpecialUsers.FieldByName('employee_id').AsString, strErrorTextOut);
    if VendorNum = 0 then begin
      tblPaycomHistory.FieldByName('error_text').AsString := tblPaycomHistory.FieldByName('error_text').AsString + '; ' + strErrorTextOut;
      strRecStatus := 'error';
    end;
    tblPaycomHistory.FieldByName('certfile_employee_id').AsInteger := VendorNum;
    tblPaycomHistory.FieldByName('record_status').AsString         := strRecStatus ;
    tblPaycomHistory.post;

  except on E: Exception do begin
    tblPaycomHistory.Edit;
    tblPaycomHistory.FieldByName('record_status').AsString := 'error';
    tblPaycomHistory.FieldByName('error_text').AsString    := tblPaycomHistory.FieldByName('error_text').AsString + '; ' + E.Message;
    tblPaycomHistory.post;
  end;

  end;  { Try/Except }

end;  {InsertSUIntoHistoryTable}


function TufrmCertifyExpDataLoader.ScrubVendorNum(const strVendorNumIn: String; Var ErrorTxtOut: String): Integer;
Var
  intVendorNum : Integer;
  strVendorNum : String;

begin
  ErrorTxtOut  := '';
  Result       := 0;
  strVendorNum := strVendorNumIn;

  if strVendorNum <> '' then begin

    // Strip-off trailing '|' if exists
    if Pos('|', strVendorNum) = Length(strVendorNum) then
      strVendorNum := Copy(strVendorNum, 1, Length(strVendorNum) - 1);

    try
      intVendorNum := StrToInt(strVendorNum);
      Result       := intVendorNum ;

    except on E1: Exception do begin
      Result      := 0;
      ErrorTxtOut := ' Field: employee_id - ' + E1.Message;  // employee_id in special_users file is Vendor Number
    end;

    end;
  end else begin
    Result := 0;
    ErrorTxtOut := 'employee_id is blank';
  end;

end;  {ScrubAndValidateVendorNum}


function TufrmCertifyExpDataLoader.ScrubCertifyDept(const DepartmentIn: String): String;  // ???JL depricate
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
  StatusBar1.Panels[1].Text := 'Current Task:  Loading CharterVisa into StartBucket ' ;
  Application.ProcessMessages;

  stlCharterVisaUsers := TStringList.Create;

  // the edCharterVisaUsers field initialized in "Initialize form fields from ini" section of FormCreate()
  stlCharterVisaUsers.CommaText := edCharterVisaUsers.Text;       // comma-seperated string of Vendor Numbers for CharterVisa Group
  try
    for i := 0 to stlCharterVisaUsers.Count - 1 do Begin
      qryInsertTripsForGroup.Close;
      qryInsertTripsForGroup.ParamByName('parmCrewMemberVendorNumIn').AsInteger := StrToInt(stlCharterVisaUsers[i]);
      qryInsertTripsForGroup.ParamByName('parmGroupNameIn').AsString            := 'CharterVisa';
      qryInsertTripsForGroup.ParamByName('parmDaysBackIn').AsInteger            := strToInt(edDaysBack.Text);
      qryInsertTripsForGroup.Execute;
    End;

  finally
    stlCharterVisaUsers.Free;
  end;

end;  { LoadCharterVisaTripsIntoStartBucket }


// **********************************************
//   1. Get all IFS recs from PaycomHistory for given batch
//   2. Add all tails for each IFS employee into StartBucket Crew & Tail
//
// **********************************************
procedure TufrmCertifyExpDataLoader.Load_IFS_IntoStartBucket(BatchTime: TDateTime);
var
  stlAdditionalIFSUsers : TStringList;
  i : Integer;

begin
  // Get a list of IFS members for this batch
  qryGetIFSMembers.Close;
  qryGetIFSMembers.ParamByName('parmImportedOn').AsDateTime := BatchTime;
  qryGetIFSMembers.Open;

  // Load all tails for each IFS member ( at this time there are approx 100 tails & 15 IFS members, so we are adding ~1500 recs to StartBucket )
  while not qryGetIFSMembers.eof do begin
    qryInsertTailsForIFS.Close;
    qryInsertTailsForIFS.ParamByName('parmCrewMemberVendorNumIn').AsInteger := qryGetIFSMembers.FieldByName('certify_gp_vendornum').AsInteger;
    qryInsertTailsForIFS.Execute;
    qryGetIFSMembers.Next;
  end;

  // Load actual Trips that IFS members flew on
  qryGetIFSMembers.First;
  while not qryGetIFSMembers.eof do begin
    qryInsertTripsForGroup.Close;
    qryInsertTripsForGroup.ParamByName('parmCrewMemberVendorNumIn').AsInteger := qryGetIFSMembers.FieldByName('certify_gp_vendornum').AsInteger;
    qryInsertTripsForGroup.ParamByName('parmGroupNameIn').AsString            := 'IFS';
    qryInsertTripsForGroup.ParamByName('parmDaysBackIn').AsInteger            := strToInt(edDaysBack.Text);
    qryInsertTripsForGroup.Execute;
    qryGetIFSMembers.Next;
  end;

  qryGetIFSMembers.Close;

  // TE-43  Adding Pseudo-users for IFS group,   9 Dec 2019
  stlAdditionalIFSUsers := TStringList.Create;
  Try
    stlAdditionalIFSUsers.CommaText := edIFSPseudoUsers.Text;
    for i := 0 to stlAdditionalIFSUsers.Count - 1 do begin
      qryInsertTripsForGroup.Close;
      qryInsertTripsForGroup.ParamByName('parmCrewMemberVendorNumIn').AsInteger := strToInt(stlAdditionalIFSUsers[i]);
      qryInsertTripsForGroup.ParamByName('parmGroupNameIn').AsString            := 'IFS';
      qryInsertTripsForGroup.ParamByName('parmDaysBackIn').AsInteger            := strToInt(edDaysBack.Text);
      qryInsertTripsForGroup.Execute;
    end;
  Finally
    stlAdditionalIFSUsers.Free;
  End;

end;  { LoadIFSIntoStartBucket }


procedure TufrmCertifyExpDataLoader.InsertCrewTail(const TailNumIn: String; VendorNumIn: Integer; DataSourceIn: String);
begin
  tblStartBucket.Insert;
  tblStartBucket.FieldByName('TailNum').AsString              := TailNumIn;
  tblStartBucket.FieldByName('CrewMemberVendorNum').AsInteger := VendorNumIn;
  tblStartBucket.FieldByName('CrewMemberID').AsString         := DataSourceIn;
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

  if (qryFindLeadPilotEmail.RecordCount > 0) and (LeadPilotEMail <> '' ) then begin            // If LeadPilot found
    if Not EmployeeTerminated(LeadPilotEMail, BatchTimeIn) then                                //   If LeadPilot is not terminated
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

  // Special case logic for this email address; skip the employee-terminated test & return False
  if Pos(LowerCase(EMailIn), 'flightcrew@claylacy.com|flightcrewcc@claylacy.com') > 0 then begin
    LogIt('EmployeeTerminated() Skipped Email: ' + EMailIn);
    Exit;
  end;


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

(*  Depricating this process. Not necessary to deal with Deleted recs on an hourly basis.
    The nightly full-file-refresh handles records that would be Deleted by this process -  28 oct 2019  JL

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
*)

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

  (*  Depricating this process. Not necessary to deal with Deleted recs on an hourly basis.
    The nightly full-file-refresh handles records that would be Deleted by this process -  28 oct 2019  JL

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
*)

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

 (*  Depricating this process. Not necessary to deal with Deleted recs on an hourly basis.
    The nightly full-file-refresh handles records that would be Deleted by this process -  28 oct 2019  JL

  StatusBar1.Panels[1].Text := 'Current Task:  Retrieving deleted CrewLog_History recs'  ;
  Application.ProcessMessages;

  UpdateRecordStatus_CrewLog('deleted', NewBatchDateIn, PreviousBatchDateIn);
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
 *)

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
    gloPusher.Push;

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


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewTail(const RecStatusIn: String; const A_BatchDateIn, B_BatchDateIn: TDateTime);
begin
  PurgeTable('CertifyExp_CrewWork1');

  qryUpdtStatus_CrewTail_1.Close;
  qryUpdtStatus_CrewTail_1.ParamByName('parmCreateDate').AsDateTime := A_BatchDateIn;       // PreviousBatchDateIn;
  qryUpdtStatus_CrewTail_1.Execute;   // puts data into working table CertifyExp_CrewWork1

  qryUpdtStatus_CrewTail_2.Close;
  qryUpdtStatus_CrewTail_2.ParamByName('parmCreateDate').AsDateTime := B_BatchDateIn;
  qryUpdtStatus_CrewTail_2.ParamByName('parmNewStatus').AsString    := RecStatusIn;
  qryUpdtStatus_CrewTail_2.Execute;   // uses working table created above

end;


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewTrip(const RecStatusIn: String; const A_BatchDateIn, B_BatchDateIn: TDateTime);
begin
  PurgeTable('CertifyExp_CrewWork1');

  qryUpdtStatus_CrewTrip_1.Close;
  qryUpdtStatus_CrewTrip_1.ParamByName('parmCreateDate').AsDateTime := A_BatchDateIn;       // PreviousBatchDateIn;
  qryUpdtStatus_CrewTrip_1.Execute;   // puts data into working table CertifyExp_CrewWork1

  qryUpdtStatus_CrewTrip_2.Close;
  qryUpdtStatus_CrewTrip_2.ParamByName('parmCreateDate').AsDateTime := B_BatchDateIn;
  qryUpdtStatus_CrewTrip_2.ParamByName('parmNewStatus').AsString    := RecStatusIn;
  qryUpdtStatus_CrewTrip_2.Execute;   // uses working table created above

end;  { UpdateRecordStatus_CrewTrip }


procedure TufrmCertifyExpDataLoader.UpdateRecordStatus_CrewLog(const RecStatusIn: String; const A_BatchDateIn, B_BatchDateIn: TDateTime);
begin
  PurgeTable('CertifyExp_CrewWork1');

  qryUpdtStatus_CrewLog_1.Close;
  qryUpdtStatus_CrewLog_1.ParamByName('parmCreateDate').AsDateTime := A_BatchDateIn;       // PreviousBatchDateIn;
  qryUpdtStatus_CrewLog_1.Execute;   // puts data into working table CertifyExp_CrewWork1

  qryUpdtStatus_CrewLog_2.Close;
  qryUpdtStatus_CrewLog_2.ParamByName('parmCreateDate').AsDateTime := B_BatchDateIn;
  qryUpdtStatus_CrewLog_2.ParamByName('parmNewStatus').AsString    := RecStatusIn;
  qryUpdtStatus_CrewLog_2.Execute;   // uses working table created above

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

  edPreviousDate.Text := DateTimeToStr( PreviousBatchDateOut);
  edNewDate.Text      := DateTimeToStr( NewBatchDateOut);

end;


procedure TufrmCertifyExpDataLoader.GetBatchDates_CrewTrip(var PreviousBatchDateOut, NewBatchDateOut: TDateTime);
begin
  qryGetCrewTripBatchDates.Close;
  qryGetCrewTripBatchDates.Open;
  NewBatchDateOut := qryGetCrewTripBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTripBatchDates.Next;
  PreviousBatchDateOut := qryGetCrewTripBatchDates.FieldByName('CreatedOn').AsDateTime;
  qryGetCrewTripBatchDates.Close;

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

  edPreviousDate.Text := DateTimeToStr( PreviousBatchDateOut );
  edNewDate.Text      := DateTimeToStr( NewBatchDateOut );

end;  { GetBatchDates_CrewLog }



procedure TufrmCertifyExpDataLoader.PruneHistoryTables(Const TableNameIn : String);
Var
  strDaysBack  : String;
  strTableName : String;

begin

  if TableNameIn = 'CertifyExp_PaycomHistory' then Begin
    strDaysBack := myIni.ReadString('Startup', 'RetainPaycomHistoryDays',  '') ;

    qryPruneHistoryTables.SQL.Clear;
    qryPruneHistoryTables.SQL.Add( ' delete from CertifyExp_PaycomHistory ' );
    qryPruneHistoryTables.SQL.Add( ' where imported_on < CAST(CURRENT_TIMESTAMP - ' + strDaysBack + ' AS DATE) ' );
    qryPruneHistoryTables.Execute;

  End Else Begin

    strTableName := 'CertifyExp_' + TableNameIn + '_History' ;    // must match table name exactly
    strDaysBack  := myIni.ReadString('Startup', 'RetainHourlyDeltaTablesHistoryDays',  '') ;
                                                                  // for performance reasons, keep only the last n days of data
                                                                  //   (the more records in these tables the slower the 'NOT IN' queries run)
    qryPruneHistoryTables.SQL.Clear;
    qryPruneHistoryTables.SQL.Add( ' delete from ' + strTableName );
    qryPruneHistoryTables.SQL.Add( ' where CreatedOn < CAST(CURRENT_TIMESTAMP - ' + strDaysBack + ' AS DATE) ' );
    qryPruneHistoryTables.Execute;

  End;
end;  { PruneHistoryTables }


procedure TufrmCertifyExpDataLoader.PurgeTable(const TableNameIn: String);
begin

  qryPurgeWorkingTable.Close;
  qryPurgeWorkingTable.SQL.Text := 'delete from ' + TableNameIn;
  qryPurgeWorkingTable.Execute;

end;


procedure TufrmCertifyExpDataLoader.FlagTerminatedEmployees(const BatchTimeIn: TDateTime);
begin

  // Set status = 'terminated' for employees terminated more than edTerminatedDaysBack.Text days ago so that they don't get procesed
  //   Test for Terminated employee after import because if Terminated then other validation is moot.
  qryFlagTerminatedEmployees.Close;
  qryFlagTerminatedEmployees.ParamByName('parmDaysBackTerminated').AsInteger := StrToInt(edTerminatedDaysBack.Text);
  qryFlagTerminatedEmployees.ParamByName('parmImportDate').AsDateTime        := BatchTimeIn;
  qryFlagTerminatedEmployees.Execute;   // Sets RecordStatus = 'terminated'

end;  { FlagTerminatedEmployees() }



procedure TufrmCertifyExpDataLoader.LoadCertifyEmployeesTable(const BatchTimeIn: TDateTime);
begin

  PurgeTable('CertifyExp_Certify_Employees');
  qryLoadCertifyEmployeesTable.ParamByName('parmBatchDateIn').AsDateTime := BatchTimeIn;
  qryLoadCertifyEmployeesTable.Execute;

end;



procedure TufrmCertifyExpDataLoader.LogIt(ErrorMsgIn: String);
var
  YearMonth       : String;
  TargetDate      : TDateTime;
  TargetYearMonth : String;

begin
  Sleep(1000);      // keeps Identifier strings for ini file entries unique. Klunky, but it works  -JL
  myLog := TIniFile.Create( ExtractFilePath(ParamStr(0)) + 'CertifyDataLoader.log');
  YearMonth := IntToStr(YearOf(Now)) + '-' + IntToStr(MonthOf(Now));
  myLog.WriteString(YearMonth, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now()), ErrorMsgIn) ;

  // prune log file - purge log entries older than 180 days
  TargetDate := Now() - 180;
  TargetYearMonth := IntToStr(YearOf(TargetDate)) + '-' + IntToStr(MonthOf(TargetDate));
  if myLog.SectionExists(TargetYearMonth) Then
    myLog.EraseSection(TargetYearMonth);

end;  { LogIt }


function TufrmCertifyExpDataLoader.GetValidGroups: String;
var
  strAccum : String;

begin
  strAccum := '|';
  qryGetGroups.Close;
  qryGetgroups.Open;

  While not qryGetgroups.Eof do begin
    strAccum := strAccum + qryGetgroups.FieldByName('certify_group').AsString + '|';
    qryGetgroups.Next;
  end;
  qryGetgroups.Close;

  Result := strAccum;

end;  {GetValidGroups}



function TufrmCertifyExpDataLoader.GetValidJobCodeDescrips: String;
var
  strAccum : String;

begin
  strAccum := '|';
  qryGetJobCodeDescrips.Close;
  qryGetJobCodeDescrips.Open;

  While not qryGetJobCodeDescrips.Eof do begin
    strAccum := strAccum + qryGetJobCodeDescrips.FieldByName('job_code_descrip').AsString + '|';
    qryGetJobCodeDescrips.Next;
  end;
  qryGetJobCodeDescrips.Close;

  Result := strAccum;

end;  { GetValidJobCodeDescrips }


function TufrmCertifyExpDataLoader.GetValidRoles: String;
begin
  Result := '|Accountant|Employee|Executive|Manager|';              // get this from .ini   ???JL

end;


function TufrmCertifyExpDataLoader.StripTrailingPipe(const strEmpIDin: string): String;
var
  pipePos : Integer;
begin
  Result := strEmpIDin;

  pipePos := Pos('|', strEmpIDin);
  if pipePos > 0 then
    Result :=  Copy(strEmpIDin, 1, pipePos - 1);

end;  {StripTrailingPipe}



end.
