object ufrmCertifyExpDataLoader: TufrmCertifyExpDataLoader
  Left = 0
  Top = 0
  Caption = 'ufrmCertifyExpDataLoader-Phase 2B v 0.1'
  ClientHeight = 446
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
  object Label5: TLabel
    Left = 257
    Top = 142
    Width = 86
    Height = 13
    Caption = 'Special Users File:'
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
    Top = 159
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
  object Label17: TLabel
    Left = 257
    Top = 192
    Width = 88
    Height = 13
    Caption = 'Tail_LeadPilot File:'
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
    Top = 427
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
    Left = 396
    Top = 305
    Width = 99
    Height = 25
    Caption = 'btnMain'
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
  object edSpecialUsersFile: TEdit
    Left = 257
    Top = 158
    Width = 451
    Height = 21
    TabOrder = 6
    Text = 
      'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\InputFiles\certify_' +
      'special_users.csv'
  end
  object edLastNTrips: TEdit
    Left = 63
    Top = 106
    Width = 35
    Height = 21
    Alignment = taRightJustify
    TabOrder = 7
    Text = '10'
  end
  object edTerminatedDaysBack: TEdit
    Left = 142
    Top = 156
    Width = 37
    Height = 21
    Alignment = taRightJustify
    TabOrder = 8
    Text = '14'
  end
  object edContractorDaysBack: TEdit
    Left = 149
    Top = 74
    Width = 32
    Height = 21
    Alignment = taRightJustify
    TabOrder = 9
    Text = '45'
  end
  object btnTest: TButton
    Left = 633
    Top = 396
    Width = 75
    Height = 25
    Caption = 'btnTest'
    TabOrder = 10
    Visible = False
    OnClick = btnTestClick
  end
  object edCharterVisaUsers: TEdit
    Left = 14
    Top = 208
    Width = 172
    Height = 21
    Alignment = taRightJustify
    TabOrder = 11
    Text = '12779,14220'
  end
  object edTailLeadPilotFile: TEdit
    Left = 257
    Top = 208
    Width = 451
    Height = 21
    TabOrder = 12
    Text = 
      'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\InputFiles\tail_lea' +
      'dpilot_20190107.csv'
  end
  object btnLoadTailLeadPilotTable: TButton
    Left = 569
    Top = 235
    Width = 139
    Height = 25
    Caption = 'Load Tail_LeadPilot Table'
    TabOrder = 13
    OnClick = btnLoadTailLeadPilotTableClick
  end
  object UniConnection1: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WarehouseDEV'
    Username = 'sa'
    Server = '192.168.1.122'
    Connected = True
    LoginPrompt = False
    Left = 20
    Top = 256
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
      '  WHERE record_status = :parmRecordStatusIn'
      '    AND imported_on = :parmImportDateIn'
      ''
      '')
    Left = 194
    Top = 221
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmRecordStatusIn'
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
    Left = 272
    Top = 254
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
    Left = 348
    Top = 145
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
    Left = 502
    Top = 63
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
      'Note:'
      
        'This query'#39's BeforeExecute Event is used to set parameters in sc' +
        'rLoadTripData.SQL'
      'See code & scrLoadTripDate.dataset for more info;'
      ''
      '23 Aug 2018 - Jeff Luckey, Jeff@dcsit.com'
      '*/')
    Left = 399
    Top = 188
  end
  object qryBuildValFile: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct LogSheet, CrewMemberID'
      'from CertifyExp_Trips_StartBucket')
    Left = 78
    Top = 284
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
      '    and record_status = :parmRecStatusIn')
    Left = 471
    Top = 146
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
  object qryGetApproverEmail: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select work_email'
      'from CertifyExp_PayComHistory'
      'where employee_code = :parmEmpCode'
      '  and imported_on   = :parmBatchTimeIn')
    Left = 559
    Top = 162
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmEmpCode'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'parmBatchTimeIn'
        Value = nil
      end>
  end
  object qryGetTripAccountantRec: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct QuoteNum, FarPart'
      'from CertifyExp_Trips_StartBucket'
      'where QuoteNum is not null'
      '  and CrewMemberVendorNum is not null'
      '  and FarPart is not null')
    Left = 172
    Top = 268
  end
  object scrLoadTripStopData: TUniScript
    SQL.Strings = (
      'delete from CertifyExp_tripStop_Step1 '
      ''
      'insert into CertifyExp_TripStop_Step1'
      'select distinct B.QuoteNum, L.DEPTID as AirportID'
      
        'from CertifyExp_Trips_StartBucket B left outer join QuoteSys_Tri' +
        'pLeg L on B.TailNum = L.ACREGNO and B.LogSheet = L.LOGSHEET'
      'where QuoteNum is not null'
      'order by QuoteNum'
      ''
      'insert into CertifyExp_TripStop_Step1'
      'select distinct B.QuoteNum, L.ARRIVEID as AirportID'
      
        'from CertifyExp_Trips_StartBucket B left outer join QuoteSys_Tri' +
        'pLeg L on B.TailNum = L.ACREGNO and B.LogSheet = L.LOGSHEET'
      'where QuoteNum is not null'
      'order by QuoteNum'
      ''
      '')
    Connection = UniConnection1
    Left = 221
    Top = 12
  end
  object qryGetTripStopRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct TripNum, AirportID'
      'from CertifyExp_TripStop_Step1')
    Left = 282
    Top = 380
  end
  object qryGetStartBucketSorted: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT *     '
      '  FROM CertifyExp_Trips_StartBucket'
      '  WHERE CrewMemberID not in ('#39'CharterVisa'#39', '#39'DOM_processing'#39')'
      '  ORDER By crewmemberid , TripDepartDate desc ')
    Left = 467
    Top = 203
  end
  object qryPilotsNotInPaycom: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'insert into CertifyExp_PilotsNotInPaycom'
      'select distinct  CrewMemberVendorNum '
      'from CertifyExp_Trips_StartBucket'
      'where CrewMemberVendorNum not in ('
      #9'select distinct certify_gp_vendornum'
      #9'from CertifyExp_PayComHistory'
      #9'where imported_on = :parmBatchTimeIn'
      #9'  and certify_department = '#39'Flight Crew'#39
      #9'  and certify_gp_vendornum is not null    )')
    Left = 40
    Top = 325
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmBatchTimeIn'
        Value = nil
      end>
  end
  object qryGetPilotsNotInPaycom: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select * from QuoteSys_PilotMaster'
      
        'where VendorNumber in ( select CrewMemberVendorNum from CertifyE' +
        'xp_PilotsNotInPaycom )'
      
        '  and Status         in ('#39'Agent of CLA'#39', '#39'Cabin Serv'#39', '#39'Parttime' +
        '-CLA'#39')'
      '  and EmployeeStatus in ('#39'Part 91'#39', '#39'Part 135'#39', '#39'Cabin Serv'#39')'
      '  and ArchiveFlag is null'
      'order by LastName')
    Left = 365
    Top = 232
  end
  object qryEmptyPilotsNotInPaycom: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_PilotsNotInPaycom')
    Left = 593
    Top = 260
  end
  object qryDeleteTrips: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Trips_StartBucket'
      'where LogSheet     = :parmLogSheetIn'
      '  and CrewMemberID = :parmCrewMemberIDIn'
      '  and QuoteNum     = :parmQuoteNumIn')
    Left = 639
    Top = 145
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
  object qryContractorsNotInPaycom_Step1: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        '/* Step 1:  -- Contractors who have flown in last parmDaysBack d' +
        'ays by CrewMemberID: */'
      ''
      'insert into CertifyExp_Contractors45'
      
        'select distinct S.CrewMemberVendorNum, P.Status, P.EmployeeStatu' +
        's, P.PilotID'
      
        'from CertifyExp_Trips_StartBucket S left outer join QuoteSys_Pil' +
        'otMaster P on S.CrewMemberVendorNum = P.VendorNumber'
      
        'where P.Status in ( '#39'Agent of CLA'#39', '#39'Cabin Server'#39', '#39'Parttime - ' +
        'CLA'#39' )    -- Definition of "Contractor"'
      
        '  and P.EmployeeStatus in ( '#39'Part 91'#39', '#39'Part 135'#39', '#39'Cabin Serv'#39' ' +
        ')         -- Definition of "Contractor" '
      
        '  and (P.ArchiveFlag is null or P.ArchiveFlag = '#39#39')             ' +
        '          -- Currrently employed/not terminated'
      '  and S.TripDepartDate > (CURRENT_TIMESTAMP - :parmDaysBack)'
      '')
    Left = 491
    Top = 5
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmDaysBack'
        Value = nil
      end>
  end
  object qryContractorsNotInPaycom_Step2: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        '/* Step 2: -- Find VendorNums in #CertifyExp_Contractors45 that ' +
        'are not in CertifyExp_PaycomHistory'#39's current batch */'
      ''
      'delete from CertifyExp_Contractors45'
      'where CrewMemberVendorNum is not null'
      '  and CrewMemberVendorNum  in ('
      '        select distinct certify_gp_vendornum'
      '        from CertifyExp_PayComHistory'
      
        '        where imported_on = :parmImportDateIn          -- identi' +
        'fy the batch of records'
      '          and certify_gp_vendornum is not null'
      #9#9'  and data_source = '#39'paycom_file'#39
      '      )')
    Left = 609
    Top = 26
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDateIn'
        Value = nil
      end>
  end
  object qryPurgeWorkingTable: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Contractors45'
      ''
      '/*'
      'if OBJECT_ID('#39'CertifyExp_Contractors45'#39') is not null'
      '  drop table CertifyExp_Contractors45'
      '*/')
    Left = 411
    Top = 51
  end
  object qryGetPilotDetails: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select [PilotID]'
      '      ,[LastName]'
      '      ,[FirstName]'
      '      ,[VendorNumber]'
      '      ,[UpdatedInQuoteSys]'
      '      ,[UpdatedBy]'
      '      ,[Base]'
      '      ,[ArchiveFlag]'
      '      ,[JobTitle]'
      '      ,[EmployeeStatus]'
      '      ,[Status]'
      '      ,[EMail]'
      '      ,[AssignedAC]'
      ''
      'from QuoteSys_PilotMaster'
      'where PilotID in (select PilotID from CertifyExp_Contractors45)')
    Left = 636
    Top = 72
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
    Left = 360
    Top = 371
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
    Left = 442
    Top = 298
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
    Left = 508
    Top = 274
  end
  object qryUpdateHasCCField: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ' update CertifyExp_PayComHistory set has_credit_card = '#39'T'#39
      
        ' where certify_gp_vendornum in (select distinct vendorid from V_' +
        'DynamicsGP_Alternate_Payment_Vendor)'
      '    and imported_on = :parmImportedOn')
    Left = 253
    Top = 308
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportedOn'
        Value = nil
      end>
  end
  object qryGetTailTripLog: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct TailNum, QuoteNum, LogSheet'
      'from CertifyExp_Trips_StartBucket'
      'where QuoteNum is not Null'
      '  and LogSheet is not Null'
      'order by LogSheet'
      '')
    Left = 358
    Top = 277
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
      '')
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
        ','#39'FA Designated'#39','#39'Pilot Not-Designated'#39','#39'FA Non-Designated'#39')'
      '  and record_status = '#39'non-certify'#39
      
        '  and ((termination_date is null) or (termination_date > CURRENT' +
        '_TIMESTAMP - 14))')
    Left = 357
    Top = 321
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDate'
        Value = nil
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
      '  and record_status = '#39'OK'#39)
    Left = 507
    Top = 323
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
  object qryInsertCharterVisa: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'insert into CertifyExp_Trips_StartBucket'
      
        'select distinct L.logsheet, '#39'CharterVisa'#39', T.QUOTENO, L.acregno,' +
        ' null, :parmCrewMemberVendorNum, null'
      
        'from QuoteSys_TripLeg L LEFT OUTER JOIN QuoteSys_Trip T ON L.ACR' +
        'EGNO = T.ACREGNO AND L.LOGSHEET = T.LOGSHEET'
      
        'where ( (departure < CURRENT_TIMESTAMP)        and (arrival > CU' +
        'RRENT_TIMESTAMP) )   -- in-progress'
      
        '   OR ( (departure > (CURRENT_TIMESTAMP - 10)) and (arrival < CU' +
        'RRENT_TIMESTAMP) )   -- ended within the past 10 days')
    Left = 131
    Top = 251
    ParamData = <
      item
        DataType = ftInteger
        Name = 'parmCrewMemberVendorNum'
        Value = 77777
      end>
  end
  object qryGetDOMEmployees: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select certify_gp_vendornum, paycom_assigned_ac'
      'from CertifyExp_PayComHistory'
      'where imported_on = :parmImportDate'
      '  and certify_department = '#39'DOM'#39)
    Left = 48
    Top = 368
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDate'
        Value = nil
      end>
  end
  object tblTailLeadPilot: TUniTable
    TableName = 'CertifyExp_Tail_LeadPilot'
    Connection = UniConnection1
    Left = 587
    Top = 300
  end
  object qryFindLeadPilotEmail: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'SELECT Email'
      'FROM   CertifyExp_Tail_LeadPilot'
      'WHERE Tail = :parmTailNumIn')
    Left = 208
    Top = 368
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
    Left = 126
    Top = 380
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
end
