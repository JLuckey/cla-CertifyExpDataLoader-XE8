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
    Left = 136
    Top = 14
    Width = 352
    Height = 198
    Lines.Strings = (
      'Memo1')
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object RESTClient: TRESTClient
    BaseURL = 'https://api.certify.com/v1/exprptglds/1'
    Params = <>
    HandleRedirects = True
    Left = 32
    Top = 24
  end
  object RESTRequest: TRESTRequest
    Client = RESTClient
    Params = <>
    Response = RESTResponse
    SynchronizedEvents = False
    Left = 40
    Top = 77
  end
  object RESTResponse: TRESTResponse
    Left = 50
    Top = 121
  end
end
