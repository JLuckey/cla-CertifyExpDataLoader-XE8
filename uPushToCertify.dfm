object frmPushToCertify: TfrmPushToCertify
  Left = 0
  Top = 0
  Caption = 'frmPushToCertify'
  ClientHeight = 258
  ClientWidth = 661
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 49
    Top = 11
    Width = 20
    Height = 13
    Caption = 'Tail:'
  end
  object Label2: TLabel
    Left = 49
    Top = 53
    Width = 38
    Height = 13
    Caption = 'Vendor:'
  end
  object Label3: TLabel
    Left = 49
    Top = 97
    Width = 34
    Height = 13
    Caption = 'Action:'
  end
  object Button1: TButton
    Left = 264
    Top = 225
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 294
    Top = 8
    Width = 352
    Height = 198
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object edTail: TEdit
    Left = 49
    Top = 26
    Width = 121
    Height = 21
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    TabOrder = 2
    Text = 'N8241W'
  end
  object edVendorNum: TEdit
    Left = 49
    Top = 70
    Width = 121
    Height = 21
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    TabOrder = 3
    Text = '13748'
  end
  object edAction: TEdit
    Left = 49
    Top = 112
    Width = 121
    Height = 21
    BiDiMode = bdLeftToRight
    ParentBiDiMode = False
    TabOrder = 4
    Text = 'added'
  end
  object RESTClient: TRESTClient
    BaseURL = 'https://api.certify.com/v1/exprptglds/1'
    Params = <>
    HandleRedirects = True
    Left = 85
    Top = 180
  end
  object RESTRequest: TRESTRequest
    Client = RESTClient
    Params = <>
    Response = RESTResponse
    SynchronizedEvents = False
    Left = 20
    Top = 161
  end
  object RESTResponse: TRESTResponse
    Left = 35
    Top = 203
  end
end
