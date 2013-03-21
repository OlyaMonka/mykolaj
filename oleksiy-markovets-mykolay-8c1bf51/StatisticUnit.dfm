object StatisticForm: TStatisticForm
  Left = 0
  Top = 0
  Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
  ClientHeight = 584
  ClientWidth = 292
  Color = clGradientActiveCaption
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCanResize = FormCanResize
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object StatisticStringGrid: TStringGrid
    Left = 8
    Top = 8
    Width = 273
    Height = 481
    TabOrder = 0
  end
  object CloseButton: TButton
    Left = 206
    Top = 522
    Width = 75
    Height = 25
    Caption = #1047#1072#1082#1088#1080#1090#1080
    TabOrder = 1
    OnClick = CloseButtonClick
  end
  object SaveXMLButton: TButton
    Left = 8
    Top = 520
    Width = 129
    Height = 25
    Caption = #1047#1073#1077#1088#1077#1075#1090#1080' '#1074' '#1092#1072#1081#1083' XML'
    TabOrder = 2
    OnClick = SaveXMLButtonClick
  end
  object SaveHTMLButton: TButton
    Left = 8
    Top = 489
    Width = 129
    Height = 25
    Caption = #1047#1073#1077#1088#1077#1075#1090#1080' '#1074' '#1092#1072#1081#1083' HTML'
    TabOrder = 3
    OnClick = SaveHTMLButtonClick
  end
  object SaveAddressListButton: TButton
    Left = 8
    Top = 551
    Width = 129
    Height = 25
    Caption = #1047#1073#1077#1088#1077#1075#1090#1080' '#1089#1087#1080'c'#1082#1080' '#1072#1076#1088#1077#1089
    TabOrder = 4
    OnClick = SaveAddressListButtonClick
  end
  object StaticticSaveXMLDialog: TSaveDialog
    DefaultExt = 'xml'
    Filter = 'xml|*.xml'
    Left = 48
    Top = 336
  end
  object StatisticXMLDocument: TXMLDocument
    Left = 48
    Top = 384
    DOMVendorDesc = 'MSXML'
  end
  object StaticticSaveHTMLDialog: TSaveDialog
    DefaultExt = 'html'
    Filter = 'html|*.html'
    Left = 48
    Top = 288
  end
end
