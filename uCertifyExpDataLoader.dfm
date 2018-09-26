object ufrmCertifyExpDataLoader: TufrmCertifyExpDataLoader
  Left = 0
  Top = 0
  Caption = 'ufrmCertifyExpDataLoader'
  ClientHeight = 306
  ClientWidth = 716
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
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
    Left = 257
    Top = 143
    Width = 53
    Height = 13
    Caption = 'Days Back:'
  end
  object Label4: TLabel
    Left = 257
    Top = 99
    Width = 85
    Height = 13
    Caption = 'Output Directory:'
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
  object btnGenerateFile: TButton
    Left = 13
    Top = 20
    Width = 107
    Height = 25
    Caption = 'btnGenerateFile'
    TabOrder = 1
    OnClick = btnGenerateFileClick
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 287
    Width = 716
    Height = 19
    Panels = <>
  end
  object btnTestEmail: TButton
    Left = 13
    Top = 60
    Width = 75
    Height = 25
    Caption = 'Test SMTP'
    TabOrder = 3
    OnClick = btnTestEmailClick
  end
  object edOutputFileName: TEdit
    Left = 257
    Top = 70
    Width = 451
    Height = 21
    TabOrder = 4
    Text = 'certify_employees.csv'
  end
  object btnMain: TButton
    Left = 257
    Top = 227
    Width = 99
    Height = 25
    Caption = 'btnMain'
    TabOrder = 5
    OnClick = btnMainClick
  end
  object edDaysBack: TEdit
    Left = 257
    Top = 159
    Width = 53
    Height = 21
    TabOrder = 6
    Text = '180'
  end
  object edOutputDirectory: TEdit
    Left = 257
    Top = 115
    Width = 451
    Height = 21
    TabOrder = 7
    Text = 'F:\XDrive\DCS\CLA\Certify_Expense\DataLoader\OutputFiles\'
  end
  object Memo1: TMemo
    Left = 367
    Top = 145
    Width = 341
    Height = 136
    Lines.Strings = (
      'Memo1')
    ParentColor = True
    ScrollBars = ssVertical
    TabOrder = 8
  end
  object UniConnection1: TUniConnection
    ProviderName = 'SQL Server'
    Database = 'WarehouseDEV'
    Username = 'sa'
    Server = '192.168.1.122'
    Connected = True
    LoginPrompt = False
    Left = 38
    Top = 242
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
      '  FROM [WarehouseDEV].[dbo].[CertifyExp_PayComHistory]'
      '  WHERE record_status = :parmRecordStatusIn'
      '    AND imported_on = :parmImportDateIn'
      ''
      '')
    Left = 195
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
    Left = 36
    Top = 200
  end
  object tblPaycomHistory: TUniTable
    TableName = 'CertifyExp_PayComHistory'
    Connection = UniConnection1
    Left = 131
    Top = 245
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
    Left = 622
    Top = 211
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
      '')
    Left = 263
    Top = 247
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
    Left = 364
    Top = 137
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
    Left = 443
    Top = 248
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
        '* '#39'2018-08-22 12:34:18.780'#39' */')
    Left = 572
    Top = 162
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
  object scrLoadTripData: TUniScript
    SQL.Strings = (
      'delete from CertifyExp_Trips_StartBucket'
      ''
      'insert into CertifyExp_Trips_StartBucket'
      
        '  select distinct L.LOGSHEET, L.PICPILOTNO, T.QUOTENO, L.ACREGNO' +
        ', FARPART, 0, T.TR_DEPART'
      
        '  from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.' +
        'ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)'
      '  where DEPARTURE > CURRENT_TIMESTAMP - :parmDaysBack'
      '    and PICPILOTNO > 0'
      ''
      'insert into CertifyExp_Trips_StartBucket'
      
        '  select distinct L.LOGSHEET, L.SICPILOTNO, T.QUOTENO, L.ACREGNO' +
        ', FARPART, 0, T.TR_DEPART'
      
        '  from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.' +
        'ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)'
      '  where DEPARTURE > CURRENT_TIMESTAMP - :parmDaysBack'
      '    and SICPILOTNO > 0'
      ''
      'insert into CertifyExp_Trips_StartBucket'
      
        '  select distinct L.LOGSHEET, L.FANO, T.QUOTENO, L.ACREGNO, FARP' +
        'ART, 0, T.TR_DEPART'
      
        '  from QuoteSys_TripLeg L left outer join QuoteSys_Trip T on (L.' +
        'ACREGNO = T.ACREGNO and L.LOGSHEET = T.LOGSHEET)'
      '  where DEPARTURE > CURRENT_TIMESTAMP - :parmDaysBack'
      '    and FANO > 0')
    Connection = UniConnection1
    DataSet = qryLoadTripData
    Left = 43
    Top = 10
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
    BeforeExecute = qryLoadTripDataBeforeExecute
    Left = 188
    Top = 22
  end
  object qryBuildValFile: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct LogSheet, CrewMemberID'
      'from CertifyExp_Trips_StartBucket')
    Left = 26
    Top = 116
  end
  object qryGetAirCrewVendorNum: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'update CertifyExp_Trips_StartBucket'
      'set CrewMemberVendorNum = QuoteSys_PilotMaster.VendorNumber'
      'from QuoteSys_PilotMaster'
      'where CrewMemberID = QuoteSys_PilotMaster.PilotID'
      '')
    Left = 269
    Top = 180
  end
  object qryEmptyStartBucket: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Trips_StartBucket;')
    Left = 168
    Top = 82
  end
  object qryGetImportedRecs: TUniQuery
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
      '  FROM [WarehouseDEV].[dbo].[CertifyExp_PayComHistory]'
      '  where imported_on = :parmBatchTimeIn'
      '    and record_status = '#39'OK'#39)
    Left = 500
    Top = 140
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmBatchTimeIn'
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
    Left = 80
    Top = 127
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
      '')
    Left = 130
    Top = 179
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
    Left = 40
    Top = 65
  end
  object qryGetTripStopRecs: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select distinct TripNum, AirportID'
      'from CertifyExp_TripStop_Step1')
    Left = 112
    Top = 38
  end
  object qryGetStartBucketSorted: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      '  SELECT *     '
      '  FROM [WarehouseDEV].[dbo].[CertifyExp_Trips_StartBucket]'
      '  ORDER By crewmemberid , TripDepartDate desc'
      '')
    Left = 185
    Top = 133
  end
  object scrGetCrewLogData: TUniScript
    SQL.Strings = (
      'select QuoteNum, min(LogSheet) as MinLogSheet'
      'into #CertifyExp_Work20'
      'from CertifyExp_Trips_StartBucket'
      'where QuoteNum is not null'
      'group by QuoteNum'
      'order by QuoteNum'
      ''
      'select distinct CrewMemberVendorNum, LogSheet'
      'from CertifyExp_Trips_StartBucket'
      'where CrewMemberVendorNum is not null'
      '  and LogSheet in '
      '     (select distinct MinLogSheet'
      '      from #CertifyExp_Work20)'
      ''
      'drop table #CertifyExp_Work20'
      ''
      '')
    Connection = UniConnection1
    Left = 125
    Top = 83
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
    Left = 363
    Top = 184
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
    Left = 459
    Top = 197
  end
  object qryEmptyPilotsNotInPaycom: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_PilotsNotInPaycom')
    Left = 547
    Top = 233
  end
  object qryDeleteTrips: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'delete from CertifyExp_Trips_StartBucket'
      'where LogSheet     = :parmLogSheetIn'
      '  and CrewMemberID = :parmCrewMemberIDIn'
      '  and QuoteNum     = :parmQuoteNumIn')
    Left = 642
    Top = 128
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
      
        '/* Step 1:  -- Contractors who have flown in last 30 days by Cre' +
        'wMemberID: */'
      ''
      
        'select distinct S.CrewMemberVendorNum, P.Status, P.EmployeeStatu' +
        's, P.PilotID'
      'into Contractors45'
      
        'from CertifyExp_Trips_StartBucket S left outer join QuoteSys_Pil' +
        'otMaster P on S.CrewMemberVendorNum = P.VendorNumber'
      
        'where P.Status in ( '#39'Agent of CLA'#39', '#39'Cabin Server'#39', '#39'Parttime-CL' +
        'A'#39' )    -- Definition of "Contractor"'
      
        '  and P.EmployeeStatus in ( '#39'Part 91'#39', '#39'Part 135'#39', '#39'Cabin Serv'#39' ' +
        ')       -- Definition of "Contractor" ')
    Left = 502
    Top = 12
  end
  object qryContractorsNotInPaycom_Step2: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      
        '/* Step 2: -- Find VendorNums in #Contractors45 that are not in ' +
        'CertifyExp_PaycomHistory'#39's current batch */'
      ''
      'select PilotID, CrewMemberVendorNum, Status, EmployeeStatus'
      'from Contractors45'
      'where CrewMemberVendorNum is not null'
      '  and CrewMemberVendorNum not in ('
      '        select distinct certify_gp_vendornum'
      '        from CertifyExp_PayComHistory'
      
        '        where imported_on     = :parmImportDateIn           -- '#39 +
        '2018-09-12 10:07:41.537'#39
      
        '         -- and record_status = '#39'exported'#39'                  -- s' +
        'uccessfully exported, not an Error record'
      '          and certify_gp_vendornum is not null'
      '      )')
    Left = 618
    Top = 24
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'parmImportDateIn'
        Value = nil
      end>
  end
  object qryDropWorkingTable: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'if OBJECT_ID('#39'Contractors45'#39') is not null'
      '  drop table Contractors45')
    Left = 480
    Top = 58
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
      'where PilotID in (select PilotID from Contractors45)')
    Left = 636
    Top = 73
  end
end
