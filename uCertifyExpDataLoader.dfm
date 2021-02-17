object ufrmCertifyExpDataLoader: TufrmCertifyExpDataLoader
  Left = 0
  Top = 0
  Caption = 'Certify Data Loader - ver 3.X - FX Testing'
  ClientHeight = 601
  ClientWidth = 716
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 257
    Top = 8
    Width = 146
    Height = 13
    Caption = 'PayCom Employee File (input):'
  end
  object Label2: TLabel
    Left = 257
    Top = 54
    Width = 148
    Height = 13
    Caption = 'Certify Employee File (output):'
  end
  object Label3: TLabel
    Left = 13
    Top = 33
    Width = 127
    Height = 13
    Caption = 'Include trips from the past'
  end
  object Label4: TLabel
    Left = 257
    Top = 99
    Width = 85
    Height = 13
    Caption = 'Output Directory:'
  end
  object Label6: TLabel
    Left = 14
    Top = 109
    Width = 48
    Height = 13
    Caption = 'Show the '
  end
  object Label7: TLabel
    Left = 14
    Top = 141
    Width = 216
    Height = 13
    Caption = 'Process expenses for terminated Employees:'
  end
  object Label8: TLabel
    Left = 13
    Top = 60
    Width = 103
    Height = 13
    Caption = 'Contract Flight Crew:'
  end
  object Label9: TLabel
    Left = 187
    Top = 33
    Width = 27
    Height = 13
    Caption = 'days.'
  end
  object Label10: TLabel
    Left = 104
    Top = 109
    Width = 140
    Height = 13
    Caption = 'most recent trips per person.'
  end
  object Label11: TLabel
    Left = 14
    Top = 159
    Width = 122
    Height = 13
    Caption = 'terminated within the last'
  end
  object Label12: TLabel
    Left = 185
    Top = 157
    Width = 27
    Height = 13
    Caption = 'days.'
  end
  object Label13: TLabel
    Left = 13
    Top = 77
    Width = 127
    Height = 13
    Caption = 'Include trips from the past'
  end
  object Label14: TLabel
    Left = 187
    Top = 77
    Width = 27
    Height = 13
    Caption = 'days.'
  end
  object Label15: TLabel
    Left = 13
    Top = 16
    Width = 107
    Height = 13
    Caption = 'Employee Flight Crew:'
  end
  object Label16: TLabel
    Left = 14
    Top = 192
    Width = 172
    Height = 13
    Caption = 'CharterVisa Users Vendor Numbers:'
  end
  object Label18: TLabel
    Left = 202
    Top = 413
    Width = 99
    Height = 13
    Caption = 'New Batch DateTime'
  end
  object Label19: TLabel
    Left = 382
    Top = 413
    Width = 119
    Height = 13
    Caption = 'Previous Batch DateTime'
  end
  object Label20: TLabel
    Left = 257
    Top = 192
    Width = 176
    Height = 13
    Caption = 'Inflight Services (IFS) Pseudo Users:'
  end
  object edPayComInputFile: TEdit
    Left = 257
    Top = 24
    Width = 451
    Height = 21
    TabOrder = 0
    Text = 
      'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\InputFiles\PaycomFi' +
      'le_withErrors.csv'
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 582
    Width = 716
    Height = 19
    Panels = <
      item
        Width = 150
      end
      item
        Width = 400
      end
      item
        Width = 200
      end>
  end
  object edOutputFileName: TEdit
    Left = 257
    Top = 70
    Width = 451
    Height = 21
    TabOrder = 2
    Text = 'certify_employees.csv'
  end
  object btnMain: TButton
    Left = 174
    Top = 304
    Width = 117
    Height = 25
    Caption = 'Run Nightly Process'
    TabOrder = 3
    OnClick = btnMainClick
  end
  object edDaysBack: TEdit
    Left = 149
    Top = 30
    Width = 32
    Height = 21
    Alignment = taRightJustify
    TabOrder = 4
    Text = '60'
  end
  object edOutputDirectory: TEdit
    Left = 257
    Top = 115
    Width = 451
    Height = 21
    TabOrder = 5
    Text = 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\OutputFiles\'
  end
  object edLastNTrips: TEdit
    Left = 63
    Top = 106
    Width = 35
    Height = 21
    Alignment = taRightJustify
    TabOrder = 6
    Text = '10'
  end
  object edTerminatedDaysBack: TEdit
    Left = 142
    Top = 156
    Width = 37
    Height = 21
    Alignment = taRightJustify
    TabOrder = 7
    Text = '14'
  end
  object edContractorDaysBack: TEdit
    Left = 149
    Top = 74
    Width = 32
    Height = 21
    Alignment = taRightJustify
    TabOrder = 8
    Text = '45'
  end
  object btnGoHourly: TButton
    Left = 431
    Top = 304
    Width = 116
    Height = 25
    Caption = 'Run Hourly Process'
    TabOrder = 9
    OnClick = btnGoHourlyClick
  end
  object edCharterVisaUsers: TEdit
    Left = 14
    Top = 208
    Width = 172
    Height = 21
    Alignment = taRightJustify
    TabOrder = 10
    Text = '11930,12779,14220'
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 456
    Width = 690
    Height = 110
    DataSource = DataSource1
    TabOrder = 11
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object edNewDate: TEdit
    Left = 185
    Top = 429
    Width = 152
    Height = 21
    TabOrder = 12
  end
  object edPreviousDate: TEdit
    Left = 361
    Top = 429
    Width = 158
    Height = 21
    TabOrder = 13
  end
  object btnFixer: TButton
    Left = 633
    Top = 425
    Width = 75
    Height = 25
    Caption = 'Fixer'
    TabOrder = 14
    OnClick = btnFixerClick
  end
  object edIFSPseudoUsers: TEdit
    Left = 257
    Top = 208
    Width = 172
    Height = 21
    Alignment = taRightJustify
    TabOrder = 15
  end
  object cbShowSQL: TCheckBox
    Left = 8
    Top = 433
    Width = 160
    Height = 17
    Caption = 'Show SQL Before Queries Run'
    Color = clBtnFace
    ParentColor = False
    TabOrder = 16
  end
  object UniConnection1: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WarehouseDEV'
    Username = 'sa'
    Server = '192.168.1.122'
    LoginPrompt = False
    Left = 43
    Top = 65528
    EncryptedPassword = '9CFF93FF9EFF8CFF8EFF93FF8CFF8DFF89FFCDFFCFFFCEFFC9FF'
  end
  object qryGetEmployees: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT [ID]'
      '      ,[employee_code]'
      '      ,[employee_name]'
      '      ,[work_email]'
      '      ,[position]'
      '      ,[department_descrip]'
      '      ,[job_code_descrip]'
      '      ,[supervisor_primary_code]'
      '      ,[certify_gp_vendornum]'
      '      ,[certify_department]'
      '      ,[certify_role]'
      '      ,[record_status]'
      '      ,[status_timestamp]'
      '      ,[imported_on]'
      '      ,[error_text]'
      '      ,[approver_email]'
      '      ,[accountant_email]'
      '      ,[termination_date]'
      '      ,[data_source]'
      '      ,[certfile_work_email]'
      '      ,[certfile_first_name]'
      '      ,[certfile_last_name]'
      '      ,[certfile_employee_id]'
      '      ,[certfile_employee_type]'
      '      ,[certfile_group]'
      '      ,[certfile_department_name]'
      '      ,[certfile_approver1_email]'
      '      ,[certfile_approver2_email]'
      '      ,[certfile_accountant_email]'
      '      ,[certfile_record_status]'
      '      ,[certfile_record_status_text]'
      '      ,[has_credit_card]'
      '      ,[paycom_approver1_email]'
      '      ,[paycom_approver2_email]'
      '      ,[paycom_assigned_ac]'
      '  FROM  CertifyExp_PayComHistory'
      
        '  WHERE (record_status = :parmRecordStatusIn  or  record_status ' +
        '= :parmRecordStatusIn2)'
      '    AND imported_on = :parmImportDateIn'
      ''
      '')
    Left = 122
    Top = 234
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmRecordStatusIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmRecordStatusIn2'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmImportDateIn'
        Value = nil
      end>
  end
  object SQLServerUniProvider1: TSQLServerUniProvider
    Left = 46
    Top = 217
  end
  object tblPaycomHistory: TUniTable
    TableName = 'CertifyExp_PayComHistory'
    Connection = UniConnection1
    Left = 121
    Top = 305
  end
  object IdSMTP1: TIdSMTP
    SASLMechanisms = <>
    Left = 667
    Top = 233
  end
  object IdMessage1: TIdMessage
    AttachmentEncoding = 'UUE'
    BccList = <>
    CCList = <>
    Encoding = meDefault
    FromList = <
      item
      end>
    Recipients = <>
    ReplyTo = <>
    ConvertPreamble = True
    Left = 666
    Top = 279
  end
  object qryIdentifyNonCertifyRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PayComHistory'
      
        'set record_status = '#39'non-certify'#39', status_timestamp = :parmImpor' +
        'tDate '
      'where imported_on = :parmImportDate'
      
        '  and (certify_gp_vendornum is null or certify_gp_vendornum = '#39#39 +
        ')'
      '  and (certify_department is null or certify_department = '#39#39' )'
      '  and (certify_role is null or certify_role = '#39#39')'
      '  and (record_status <> '#39'error'#39')'
      '')
    Left = 509
    Top = 363
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDate'
        Value = nil
      end>
  end
  object qryGetDBServerTime: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select CURRENT_TIMESTAMP as DateTimeOut;')
    Left = 288
    Top = 147
  end
  object qryGetDupeEmails: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/* Looking for duplicate emails */'
      ''
      'select work_email, COUNT(*) as counter'
      'from CertifyExp_PayComHistory'
      'where record_status = '#39'OK'#39' '
      
        '  and imported_on   = :parmImportDate    /* '#39'2018-08-22 12:34:18' +
        '.780'#39'*/'
      'group by work_email'
      'having COUNT(*) > 1')
    Left = 447
    Top = 250
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDate'
        Value = nil
      end>
  end
  object qryUpdateDupeEmailRecStatus: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PayComHistory'
      
        'set record_status = '#39'error'#39', error_text = :parmErrorTextIn     /' +
        '* '#39'dupe work_email foo@bar.com'#39' */'
      
        'where work_email  = :parmEmailIn                               /' +
        '* '#39'jrosen@claylacy.com'#39' */'
      
        '  and imported_on = :parmImportDateIn                          /' +
        '* '#39'2018-08-22 12:34:18.780'#39' */'
      '  and record_status = '#39'OK'#39)
    Left = 473
    Top = 71
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmErrorTextIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmEmailIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmImportDateIn'
        Value = nil
      end>
  end
  object qryLoadTripData: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/*'
      ''
      '*/')
    Left = 594
    Top = 264
  end
  object qryBuildValFile: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct LogSheet, CrewMemberID'
      'from CertifyExp_Trips_StartBucket')
    Left = 34
    Top = 280
  end
  object qryGetAirCrewVendorNum: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_Trips_StartBucket'
      'set CrewMemberVendorNum = QuoteSys_PilotMaster.VendorNumber'
      'from QuoteSys_PilotMaster'
      'where CrewMemberID = QuoteSys_PilotMaster.PilotID'
      '')
    Left = 182
    Top = 324
  end
  object qryGetImportedRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT [id]'
      '      ,[employee_code]'
      '      ,[employee_name]'
      '      ,[work_email]'
      '      ,[position]'
      '      ,[department_descrip]'
      '      ,[job_code_descrip]'
      '      ,[supervisor_primary_code]'
      '      ,[certify_gp_vendornum]'
      '      ,[certify_department]'
      '      ,[certify_role]'
      '      ,[record_status]'
      '      ,[status_timestamp]'
      '      ,[imported_on]'
      '      ,[error_text]'
      '      ,[approver_email]'
      '      ,[accountant_email]'
      '      ,[termination_date]'
      '      ,[data_source]'
      '      ,[certfile_work_email]'
      '      ,[certfile_first_name]'
      '      ,[certfile_last_name]'
      '      ,[certfile_employee_id]'
      '      ,[certfile_employee_type]'
      '      ,[certfile_group]'
      '      ,[certfile_department_name]'
      '      ,[certfile_approver1_email]'
      '      ,[certfile_approver2_email]'
      '      ,[certfile_accountant_email]'
      '      ,[certfile_record_status]'
      '      ,[certfile_record_status_text]'
      '      ,[has_credit_card]'
      '      ,[paycom_approver1_email]'
      '      ,[paycom_approver2_email]'
      '      ,[paycom_assigned_ac]'
      '  FROM CertifyExp_PayComHistory'
      '  where imported_on =  :parmBatchTimeIn '
      
        '    and (record_status = :parmRecStatusIn OR record_status = :pa' +
        'rmRecStatusIn2)'
      ' ')
    Left = 476
    Top = 139
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmBatchTimeIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn2'
        Value = nil
      end>
  end
  object scrLoadTripStopData: TUniScript
    SQL.Strings = (
      'delete from CertifyExp_tripStop_Step1 '
      ''
      'insert into CertifyExp_TripStop_Step1'
      'select distinct B.QuoteNum, L.DEPTID as AirportID'
      
        'from CertifyExp_Trips_StartBucket B left outer join vQuoteSys_Tr' +
        'ipLeg L on B.TailNum = L.ACREGNO and B.LogSheet = L.LOGSHEET'
      'where QuoteNum is not null'
      'order by QuoteNum'
      ''
      'insert into CertifyExp_TripStop_Step1'
      'select distinct B.QuoteNum, L.ARRIVEID as AirportID'
      
        'from CertifyExp_Trips_StartBucket B left outer join vQuoteSys_Tr' +
        'ipLeg L on B.TailNum = L.ACREGNO and B.LogSheet = L.LOGSHEET'
      'where QuoteNum is not null'
      'order by QuoteNum'
      ''
      '')
    Connection = UniConnection1
    Left = 223
    Top = 11
  end
  object qryGetTripStopRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct TripNum, AirportID'
      'from CertifyExp_TripStop_Step1')
    Left = 595
    Top = 377
  end
  object qryGetStartBucketSorted: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/* do not change the ORDER BY clause !!! */'
      'SELECT *     '
      '  FROM CertifyExp_Trips_StartBucket'
      
        '  WHERE CrewMemberID not in ('#39'CharterVisa'#39', '#39'DOM_processing'#39', '#39'I' +
        'FS'#39')'
      '  ORDER By crewmemberid , TripDepartDate desc ')
    Left = 608
    Top = 206
  end
  object qryEmptyPilotsNotInPaycom: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_PilotsNotInPaycom')
    Left = 443
    Top = 191
  end
  object qryDeleteTrips: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Trips_StartBucket'
      'where LogSheet     = :parmLogSheetIn'
      '  and CrewMemberID = :parmCrewMemberIDIn'
      '  and QuoteNum     = :parmQuoteNumIn')
    Left = 660
    Top = 165
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmLogSheetIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCrewMemberIDIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmQuoteNumIn'
        Value = nil
      end>
  end
  object qryPurgeWorkingTable: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Contractors45'
      '')
    Left = 402
    Top = 52
  end
  object qryGetEmployeeErrors: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT [ID]'
      '      ,[employee_code]'
      '      ,[employee_name]'
      '      ,[work_email]'
      '      ,[position]'
      '      ,[department_descrip]'
      '      ,[job_code_descrip]'
      '      ,[supervisor_primary_code]'
      '      ,[certify_gp_vendornum]'
      '      ,[certify_department]'
      '      ,[certify_role]'
      '      ,[record_status]'
      '      ,[status_timestamp]'
      '      ,[imported_on]'
      '      ,[error_text]'
      #9'  ,[approver_email]'
      #9'  ,[accountant_email]'
      '  FROM CertifyExp_PayComHistory'
      '  where imported_on = :parmBatchTimeIn'
      '    and record_status = :parmRecStatusIn')
    Left = 316
    Top = 378
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmBatchTimeIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn'
        Value = nil
      end>
  end
  object qryGetFutureTrips: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'select QUOTENO, PICPILOTNO, SICPILOTNO, FANO, ACREGNO, PART135, ' +
        'TRIP_START_DATE'
      'from QuoteSys_Quote'
      
        'where TRIP_START_DATE between ( CURRENT_TIMESTAMP - 1 ) and ( CU' +
        'RRENT_TIMESTAMP + :parmDaysForward )'
      '  and STATUS = '#39'confirmed'#39
      'order by QUOTENO')
    Left = 383
    Top = 365
    ParamData = <
      item
        DataType = ftInteger
        Name = 'parmDaysForward'
        Value = 30
      end>
  end
  object tblStartBucket: TUniTable
    TableName = 'CertifyExp_Trips_StartBucket'
    Connection = UniConnection1
    Left = 527
    Top = 225
  end
  object qryUpdateHasCCField: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ' update CertifyExp_PayComHistory '
      ' set has_credit_card = '#39'T'#39
      
        ' where certify_gp_vendornum in (select distinct vendorid from V_' +
        'DynamicsGP_Alternate_Payment_Vendor)'
      '    and imported_on = :parmImportedOn')
    Left = 284
    Top = 304
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryValidateVendorNum: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PayComHistory'
      
        'set record_status = '#39'error'#39', error_text = '#39'certify_gp_vendornum ' +
        'not found;'#39
      
        'where cast(certify_gp_vendornum as char(15)) not in (select dist' +
        'inct vendorid from V_DynamicsGP_Vendors)'
      '  and imported_on = :parmImportedOn'
      '  and record_status not in ('#39'error'#39', '#39'terminated'#39')')
    Left = 641
    Top = 326
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryFlagMissingFlightCrews: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PayComHistory'
      
        'set record_status = '#39'error'#39', error_text = '#39'flight crew missing C' +
        'ertify fields: gp_vendornum, department, role;'#39
      'where imported_on = :parmImportDate'
      
        '  and job_code_descrip in ('#39'Pilot Designated'#39','#39'Pilot-Designated'#39 +
        ','#39'FA Designated'#39','#39'Pilot Not-Designated'#39','#39'FA Non-Designated'#39', '#39'Pi' +
        'lot On Demand'#39', '#39'FA On Demand'#39')*/'
      '  and record_status = '#39'non-certify'#39
      
        '  and ((termination_date is null) or (termination_date > CURRENT' +
        '_TIMESTAMP - 14))'
      ''
      ''
      
        '/*  and job_code_descrip in (select job_code_descrip from Certif' +
        'yExp_JobCode_Lookup where certify_group = '#39'FlightCrew'#39')  */'
      ''
      '')
    Left = 357
    Top = 318
    ParamData = <
      item
        DataType = ftString
        Name = 'parmImportDate'
        Value = '2020-06-17 09:49:10.523'
      end>
  end
  object qryFlagTerminatedEmployees: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PayComHistory'
      
        'set record_status = '#39'terminated'#39', error_text = '#39'employee termina' +
        'ted more than n days ago'#39
      'FROM CertifyExp_PayComHistory'
      'where imported_on = :parmImportDate'
      
        '  and termination_date < (CURRENT_TIMESTAMP - :parmDaysBackTermi' +
        'nated)'
      '  and record_status = '#39'imported'#39)
    Left = 476
    Top = 336
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDate'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmDaysBackTerminated'
        Value = nil
      end>
  end
  object tblTailLeadPilot: TUniTable
    TableName = 'CertifyExp_Tail_LeadPilot'
    Connection = UniConnection1
    Left = 27
    Top = 529
  end
  object qryFindLeadPilotEmail: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT Email'
      'FROM   CertifyExp_Tail_LeadPilot'
      'WHERE Tail = :parmTailNumIn')
    Left = 210
    Top = 371
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmTailNumIn'
        Value = nil
      end>
  end
  object qryGetTerminationDate: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select termination_date'
      'from CertifyExp_PayComHistory'
      'where imported_on = :parmImportedOn'
      '  and work_email  = :parmEMail')
    Left = 117
    Top = 393
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmEMail'
        Value = nil
      end>
  end
  object qryGetIFSMembers: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select certify_gp_vendornum'
      'from CertifyExp_PayComHistory'
      'where Upper(certfile_group) = '#39'IFS'#39
      '  and certify_gp_vendornum is not null'
      
        '  and imported_on = :parmImportedOn      /* '#39'2019-01-13 13:45:30' +
        '.227'#39' */')
    Left = 284
    Top = 247
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object DataSource1: TDataSource
    DataSet = qryGetValidGroups
    Left = 191
    Top = 532
  end
  object qryInsertCrewTailHist: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'insert into CertifyExp_CrewTail_History (CrewMemberVendorNum,Tai' +
        'lNumber,CreatedOn,RecordStatus)'
      
        'select distinct CrewMemberVendorNum, TailNum, :parmBatchDateTime' +
        ', '#39'imported'#39
      'from CertifyExp_Trips_StartBucket '
      'where CrewMemberVendorNum is not null '
      '  and TailNum is not null '
      '  and CrewMemberVendorNum > 0')
    Left = 89
    Top = 452
    ParamData = <
      item
        DataType = ftString
        Name = 'parmBatchDateTime'
        Value = '2019-01-15 18:00:34.423'
      end>
  end
  object qryGetCrewTailBatchDates: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct CreatedOn'
      'from CertifyExp_CrewTail_History'
      'where RecordStatus = '#39'imported'#39'        '
      'order by CreatedOn desc')
    Left = 159
    Top = 492
  end
  object qryGetFailedRecs_CrewTail: TUniQuery
    Connection = UniConnection1
    Left = 131
    Top = 546
  end
  object qryGetDeletedRecs_CrewTail: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/* Get Deleted Recs  */'
      
        'select ID, CrewMemberVendorNum, TailNumber, CreatedOn, RecordSta' +
        'tus, UploadedOn, UploadStatus, UploadStatusMessage, UploadBatchI' +
        'D'
      'from CertifyExp_CrewTail_History'
      'where CreatedOn = :parmOldDateTime '
      '  and RecordStatus = '#39'imported'#39'        '
      '  and concat(CrewMemberVendorNum, '#39'|'#39', TailNumber) not in '
      '  (select concat(CrewMemberVendorNum, '#39'|'#39', TailNumber) '
      '   from CertifyExp_CrewTail_History'
      '   where CreatedOn = :parmNewDateTime                '
      '     and RecordStatus = '#39'imported'#39' )       '
      ''
      '   '
      ''
      ''
      '')
    Left = 587
    Top = 522
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmOldDateTime'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmNewDateTime'
        Value = nil
      end>
  end
  object qryGetCrewTailRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select ID, CrewMemberVendorNum, TailNumber, CreatedOn, '
      '       RecordStatus, UploadedOn, UploadStatus, '
      '       UploadStatusMessage, UploadBatchID, HTTPResultCode'
      'from CertifyExp_CrewTail_History'
      'where RecordStatus = :parmRecStatusIn'
      '  and CreatedOn    = :parmCreatedOnIn'
      '')
    Left = 171
    Top = 442
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedOnIn'
        Value = nil
      end>
  end
  object qryGetCrewTripRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select ID, CrewMemberVendorNum, TripNumber, CreatedOn, '
      '       RecordStatus, UploadedOn, UploadStatus, '
      '       UploadStatusMessage, UploadBatchID, HTTPResultCode'
      'from CertifyExp_CrewTrip_History'
      'where RecordStatus = :parmRecStatusIn '
      '  and CreatedOn    = :parmCreatedOnIn'
      '')
    Left = 594
    Top = 437
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedOnIn'
        Value = nil
      end>
  end
  object qryInsertCrewTripHist: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'insert into CertifyExp_CrewTrip_History (CrewMemberVendorNum,Tri' +
        'pNumber,CreatedOn,RecordStatus)'
      
        'select distinct CrewMemberVendorNum, QuoteNum, :parmBatchDateTim' +
        'e, '#39'imported'#39
      'from CertifyExp_Trips_StartBucket '
      'where CrewMemberVendorNum is not null '
      '  and QuoteNum is not null and CrewMemberVendorNum > 0')
    Left = 659
    Top = 455
    ParamData = <
      item
        DataType = ftString
        Name = 'parmBatchDateTime'
        Value = '2019-01-15 18:00:34.423'
      end>
  end
  object qryGetCrewTripBatchDates: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct CreatedOn'
      'from CertifyExp_CrewTrip_History'
      'where RecordStatus = '#39'imported'#39'        '
      'order by CreatedOn desc')
    Left = 448
    Top = 542
  end
  object qryInsertCrewLogHist: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'insert into CertifyExp_CrewLog_History (CrewMemberVendorNum,LogN' +
        'umber,CreatedOn,RecordStatus)'
      
        'select distinct CrewMemberVendorNum, LogSheet, :parmBatchDateTim' +
        'e, '#39'imported'#39
      'from CertifyExp_Trips_StartBucket '
      'where CrewMemberVendorNum is not null '
      '  and LogSheet is not null and CrewMemberVendorNum > 0')
    Left = 322
    Top = 455
    ParamData = <
      item
        DataType = ftString
        Name = 'parmBatchDateTime'
        Value = '2019-01-15 18:00:34.423'
      end>
  end
  object qryGetCrewLogBatchDates: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct CreatedOn'
      'from CertifyExp_CrewLog_History'
      'where RecordStatus = '#39'imported'#39'        '
      'order by CreatedOn desc')
    Left = 274
    Top = 522
  end
  object qryGetCrewLogRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select ID, CrewMemberVendorNum, LogNumber, CreatedOn, '
      '       RecordStatus, UploadedOn, UploadStatus, '
      '       UploadStatusMessage, UploadBatchID, HTTPResultCode'
      'from CertifyExp_CrewLog_History'
      'where RecordStatus = :parmRecStatusIn '
      '  and CreatedOn    = :parmCreatedOnIn'
      '')
    Left = 364
    Top = 513
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmRecStatusIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedOnIn'
        Value = nil
      end>
  end
  object qryPruneHistoryTables: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select certify_gp_vendornum'
      'from CertifyExp_PayComHistory'
      'where Upper(certfile_group) = '#39'IFS'#39
      
        '  and imported_on = :parmImportedOn      /* '#39'2019-01-13 13:45:30' +
        '.227'#39' */')
    Left = 542
    Top = 407
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryFindVendorNumInStartBucket: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select CrewMemberVendorNum, QuoteNum'
      'from CertifyExp_Trips_StartBucket'
      'where CrewMemberVendorNum = :parmVendorNumIn')
    Left = 598
    Top = 107
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmVendorNumIn'
        Value = nil
      end>
  end
  object tblTripStop: TUniTable
    TableName = 'CertifyExp_TripStop_Step1'
    Connection = UniConnection1
    Left = 659
    Top = 367
  end
  object qryGetNewHireRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select QuoteNum'
      'from CertifyExp_Trips_StartBucket'
      'where CrewMemberID in ('#39'ConNewHire'#39', '#39'EmpNewHire'#39')')
    Left = 596
    Top = 152
  end
  object qryGetFlightCrewNewHire: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select certify_gp_vendornum'
      'from CertifyExp_PayComHistory'
      'where imported_on = :parmImportedOn'
      
        '  and certify_department in ('#39'FlightCrew'#39', '#39'PoolPilot'#39', '#39'PoolFA'#39 +
        ', '#39'FlightCrewCorp'#39', '#39'FlightCrewNonPCal'#39', '#39'IFS'#39')  /* definition o' +
        'f Flight Crew, refactor to store in one place ???JL */'
      '  and hire_date > (CURRENT_TIMESTAMP - 45)'
      'order  by certify_gp_vendornum')
    Left = 535
    Top = 533
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewTrip_1: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'insert into CertifyExp_CrewWork1'
      
        'select concat(CrewMemberVendorNum, '#39'|'#39', TripNumber ) as VendorNu' +
        'm_TailTripLog '
      'from CertifyExp_CrewTrip_History'
      
        'where CreatedOn = :parmCreateDate    -- '#39'2019-10-09 10:06:17.533' +
        #39' '
      ''
      ''
      '   ')
    Left = 504
    Top = 457
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmCreateDate'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewTrip_2: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        '/* This query is used to detect and flag Deleted or Added Recs *' +
        '/'
      ''
      'update CertifyExp_CrewTrip_History'
      
        'set RecordStatus = :parmNewStatus                  /* '#39'deleted'#39',' +
        ' '#39'added'#39' */'
      'where CreatedOn =  :parmCreatedate     '
      '  and concat(CrewMemberVendorNum, '#39'|'#39', TripNumber )'
      
        '  not in (select VendorNum_TailTripLog from CertifyExp_CrewWork1' +
        ') ;'
      '')
    Left = 504
    Top = 496
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmNewStatus'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedate'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewTail_1: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'insert into CertifyExp_CrewWork1'
      
        'select concat(CrewMemberVendorNum, '#39'|'#39', TailNumber ) as VendorNu' +
        'm_TailTripLog '
      'from CertifyExp_CrewTail_History'
      
        'where CreatedOn = :parmCreateDate    -- '#39'2019-10-09 10:06:17.533' +
        #39' '
      ''
      ''
      '   ')
    Left = 257
    Top = 425
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmCreateDate'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewTail_2: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'update CertifyExp_CrewTail_History'
      
        'set RecordStatus = :parmNewStatus                  /* '#39'deleted'#39',' +
        ' '#39'added'#39' */'
      'where CreatedOn =  :parmCreatedate     '
      '  and concat(CrewMemberVendorNum, '#39'|'#39', TailNumber )'
      
        '  not in (select VendorNum_TailTripLog from CertifyExp_CrewWork1' +
        ') ;'
      '')
    Left = 250
    Top = 473
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmNewStatus'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedate'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewLog_1: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'insert into CertifyExp_CrewWork1'
      
        'select concat(CrewMemberVendorNum, '#39'|'#39', LogNumber ) as VendorNum' +
        '_TailTripLog '
      'from CertifyExp_CrewLog_History'
      
        'where CreatedOn = :parmCreateDate    -- '#39'2019-10-09 10:06:17.533' +
        #39' '
      ''
      ''
      '   ')
    Left = 407
    Top = 428
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmCreateDate'
        Value = nil
      end>
  end
  object qryUpdtStatus_CrewLog_2: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'update CertifyExp_CrewLog_History'
      
        'set RecordStatus = :parmNewStatus                  /* '#39'deleted'#39',' +
        ' '#39'added'#39' */'
      'where CreatedOn =  :parmCreatedate     '
      '  and concat(CrewMemberVendorNum, '#39'|'#39', LogNumber )'
      
        '  not in (select VendorNum_TailTripLog from CertifyExp_CrewWork1' +
        ') ;'
      '')
    Left = 407
    Top = 472
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmNewStatus'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCreatedate'
        Value = nil
      end>
  end
  object connOnBase: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'CLAOnBase'
    Username = 'sa'
    Server = '192.168.1.122'
    LoginPrompt = False
    Left = 153
    Top = 65520
    EncryptedPassword = '9CFF93FF9EFF8CFF8EFF93FF8CFF8DFF89FFCDFFCFFFCEFFC9FF'
  end
  object qryGetNewTailLeadPilotRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select tail_number, lead_pilot_email'
      'from V_CertifyExp_TailLeadPilot')
    Left = 52
    Top = 483
  end
  object qryLoadCertifyEmployeesTable: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'insert into CertifyExp_Certify_Employees'
      'SELECT certfile_work_email'
      '      ,certfile_first_name'
      '      ,certfile_last_name'
      '      ,certfile_employee_id'
      '      ,certfile_employee_type'
      '      ,certfile_group'
      '      ,certfile_department_name'
      '      ,certfile_approver1_email'
      '      ,certfile_approver2_email'
      '      ,certfile_accountant_email'
      '      ,imported_on'
      '  FROM CertifyExp_PayComHistory'
      '  where imported_on = :parmBatchDateIn'
      '    and record_status = '#39'exported'#39)
    Left = 62
    Top = 369
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmBatchDateIn'
        Value = nil
      end>
  end
  object qryGetSpecialUsers: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT'
      '   work_email'
      '  ,first_name'
      '  ,last_name'
      '  ,employee_id'
      '  ,employee_type'
      '  ,"group"'
      '  ,department_name'
      '  ,approver_email_1'
      '  ,approver_email_2'
      '  ,accountant_email'
      '  ,assigned_aircraft'
      '  ,comments'
      ''
      'FROM V_CertifyExp_SpecialUsers'
      '')
    Left = 641
    Top = 505
  end
  object qryInsertTripsForGroup: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ''
      'insert into CertifyExp_Trips_StartBucket'
      
        'select distinct L.LogSheet, :parmGroupNameIn, T.QUOTENO, L.acreg' +
        'no, null, :parmCrewMemberVendorNumIn, L.DEPARTURE, L.ARRIVEID, L' +
        '.LEGNO'
      
        '  from vQuoteSys_TripLeg L LEFT OUTER JOIN vQuoteSys_Trip T ON L' +
        '.ACREGNO = T.ACREGNO AND L.LOGSHEET = T.LOGSHEET'
      
        '  where ( (L.departure < CURRENT_TIMESTAMP)                     ' +
        'and (L.arrival > CURRENT_TIMESTAMP) )   -- in-progress'
      
        '     OR ( (L.departure > (CURRENT_TIMESTAMP - :parmDaysBackIn)) ' +
        'and (L.arrival < CURRENT_TIMESTAMP) )   -- ended within the past' +
        ' n days'
      '    AND L.LEGNO = 1'
      '    AND L.Source = '#39'IB'#39
      'ORDER BY DEPARTURE'
      ''
      '')
    Left = 207
    Top = 202
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmGroupNameIn'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmCrewMemberVendorNumIn'
        Value = nil
      end
      item
        DataType = ftInteger
        Name = 'parmDaysBackIn'
        Value = 60
      end>
  end
  object qryInsertTailsForIFS: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '  insert into CertifyExp_Trips_StartBucket'
      
        '  select distinct null, '#39'IFS'#39', null, Tail, null, :parmCrewMember' +
        'VendorNumIn, null, null, null'
      '  from CertifyExp_Tail_LeadPilot'
      '')
    Left = 192
    Top = 258
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmCrewMemberVendorNumIn'
        Value = nil
      end>
  end
  object qryGetTailTripDepartdate: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'select distinct TailNum, QuoteNum, TripDepartDate, FirstDestinat' +
        'ion'
      'from CertifyExp_Trips_StartBucket'
      'where QuoteNum is not Null'
      '/*  and LogSheet is not Null  */'
      'order by QuoteNum'
      '')
    Left = 394
    Top = 128
  end
  object qryCrewChange: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/* SQL for this query is calculated in Delphi code */')
    Left = 515
    Top = 187
  end
  object qryGetChangedCrew: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct O.QuoteNum, O.CrewMemberID'
      'from CertifyExp_OtherLegs O left join CertifyExp_FirstLegs F '
      '  on O.QuoteNum = F.QuoteNum and O.CrewMemberID = F.CrewMemberID'
      'where F.QuoteNum is null'
      'order by QuoteNum')
    Left = 547
    Top = 162
  end
  object qryLookUpFirstLeg: TUniQuery
    LocalUpdate = True
    Connection = UniConnection1
    SQL.Strings = (
      
        'select T.QuoteNo, L.ACREGNO, L.LOGSHEET, L.LEGNO, L.DEPTID, L.AR' +
        'RIVEID, L.PICPILOTNO, L.SICPILOTNO, '
      '       L.TICPILOTNO, L.FANO, L.DEPARTURE, T.FARPART'
      
        'from vQuoteSys_TripLeg L left outer join vQuoteSys_Trip T on L.A' +
        'CREGNO = T.ACREGNO and L.LogSheet = T.Logsheet'
      'where T.QuoteNo = :parmQuoteNumIn'
      '  and L.LEGNO = 1'
      'order by T.QuoteNo')
    Left = 250
    Top = 126
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmQuoteNumIn'
        Value = nil
      end>
  end
  object qryGetMissingFlightCrew: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        'select distinct P.LastName, P.FirstName, S.CrewMemberVendorNum, ' +
        'S.CrewMemberID, S.QuoteNum, S.TripDepartDate, S.TailNum'
      
        'from  QuoteSys_PilotMaster P left outer join CertifyExp_Trips_St' +
        'artBucket S  on S.CrewMemberID = CONVERT(varchar(20), P.PilotID)'
      'where S.TripDepartDate > (CURRENT_TIMESTAMP - :parmDaysBack)'
      '  and CrewMemberVendorNum not in ('
      
        '        select distinct certify_gp_vendornum      -- This result' +
        ' set is an "exclusion list" of regular employees who are current' +
        'ly employed'
      
        '        from CertifyExp_PayComHistory             --    - If you' +
        ' are NOT a member of this set then you are placed on the Missing' +
        ' Flight Crew Report'
      
        '        where imported_on = :parmImportDateIn     -- '#39'2020-05-12' +
        ' 08:30:00.387'#39'  -- identify the batch of records'
      '          and certify_gp_vendornum is not null'
      
        '          and data_source in ('#39'paycom_file'#39','#39'special_users'#39') -- ' +
        'this identifies Employees & Special Users'
      
        '          and record_status <> '#39'terminated'#39'       -- get only pe' +
        'ople who are currently employees'
      '      )'
      'order by P.LastName'
      '')
    Left = 497
    Top = 65523
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmDaysBack'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmImportDateIn'
        Value = nil
      end>
  end
  object qryGetJobCodeDescrips: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '/* the SQL for this query is calculated in the code'
      '     this SQL is an sample of the typical SQL       */'
      'select distinct job_code_descrip'
      'FROM [V_CertifyExp_JobCode_Lookup]'
      'where active = '#39'Y'#39
      '  and certify_group = :parmCertifyGroupIn'
      '')
    Left = 640
    Top = 62
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmCertifyGroupIn'
        Value = nil
      end>
  end
  object qryGetCertifyGroups: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct certify_group'
      'FROM [V_CertifyExp_JobCode_Lookup]'
      'where active = '#39'Y'#39
      '  and data_loader_logic_group = :parmLogicGroupIn'
      ''
      ''
      '')
    Left = 548
    Top = 41
    ParamData = <
      item
        DataType = ftString
        Name = 'parmLogicGroupIn'
        Value = ''
      end>
  end
  object qryGetCertifyDeptName: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct certify_department_display_name'
      'from V_CertifyExp_JobCode_Lookup'
      'where certify_group = :parmGroupIn'
      '  and certify_department_display_name is not null')
    Left = 560
    Top = 313
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmGroupIn'
        Value = nil
      end>
  end
  object qrySpecialUserDupes: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_PaycomHistory'
      'set record_status    = '#39'error'#39','
      '    status_timestamp = CURRENT_TIMESTAMP,'
      
        '    error_text       = '#39'Dupe certfile_employee_id (Vendor Number' +
        ') from Special User file'#39'      '
      ''
      'where imported_on = :parmImportedOn'
      '  and data_source in ('#39'special_users'#39' )                '
      '  and record_status in ( '#39'OK'#39' ) '
      '  and certify_gp_vendornum in'
      ''
      '     (select certfile_employee_id'
      '      from CertifyExp_PaycomHistory'
      '      where imported_on   = :parmImportedOn'
      '        and record_status in ( '#39'OK'#39', '#39'terminated'#39' ) '
      '        and data_source   = '#39'paycom_file'#39')')
    Left = 35
    Top = 416
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryGetValidGroups: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct certify_group'
      'FROM [V_CertifyExp_JobCode_Lookup]'
      'where active = '#39'Y'#39
      '  '
      ''
      '')
    Left = 611
    Top = 24
  end
end
