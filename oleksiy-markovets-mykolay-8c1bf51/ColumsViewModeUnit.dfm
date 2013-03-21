object ColumsViewModeForm: TColumsViewModeForm
  Left = 573
  Top = 202
  ClientHeight = 347
  ClientWidth = 254
  Color = clBtnFace
  Constraints.MaxHeight = 385
  Constraints.MaxWidth = 270
  Constraints.MinHeight = 380
  Constraints.MinWidth = 270
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object OkButton: TButton
    Left = 16
    Top = 318
    Width = 75
    Height = 25
    Caption = #1054#1082
    TabOrder = 0
    OnClick = OkButtonClick
  end
  object CancelButton: TButton
    Left = 126
    Top = 318
    Width = 75
    Height = 25
    Caption = #1042#1110#1076#1084#1110#1085#1080#1090#1080
    TabOrder = 1
    OnClick = CancelButtonClick
  end
  object ColumsViewCheckListBox: TCheckListBox
    Left = 16
    Top = 48
    Width = 185
    Height = 217
    ItemHeight = 13
    TabOrder = 2
  end
  object SelectAllButton: TButton
    Left = 16
    Top = 17
    Width = 75
    Height = 25
    Caption = #1042#1110#1076#1084#1110#1090#1080#1090#1080' '#1074#1089#1110
    TabOrder = 3
    OnClick = SelectAllButtonClick
  end
  object SelectNoneButton: TButton
    Left = 97
    Top = 17
    Width = 104
    Height = 25
    Caption = #1047#1085#1103#1090#1080' '#1074#1089#1110' '#1074#1110#1076#1084#1110#1090#1082#1080
    TabOrder = 4
    OnClick = SelectNoneButtonClick
  end
  object UpButton: TButton
    Left = 207
    Top = 120
    Width = 42
    Height = 25
    Caption = #1042#1074#1077#1088#1093
    TabOrder = 5
    OnClick = UpButtonClick
  end
  object DownButton: TButton
    Left = 207
    Top = 160
    Width = 42
    Height = 25
    Caption = #1042#1085#1080#1079
    TabOrder = 6
    OnClick = DownButtonClick
  end
  object GroupByAddressCheckBox: TCheckBox
    Left = 56
    Top = 279
    Width = 145
    Height = 17
    Caption = #1043#1088#1091#1087#1091#1074#1072#1090#1080' '#1079#1072' '#1072#1076#1088#1077#1089#1072#1084#1080
    Checked = True
    State = cbChecked
    TabOrder = 7
  end
  object IgnoreAppartmentCheckBox: TCheckBox
    Left = 40
    Top = 295
    Width = 161
    Height = 17
    Caption = #1030#1075#1085#1086#1088#1091#1074#1072#1090#1080' '#1085#1086#1084#1077#1088' '#1082#1074#1072#1088#1090#1080#1088#1080
    TabOrder = 8
  end
  object SaveDialog: TSaveDialog
    DefaultExt = '.xml'
    Filter = 'XML|*.xml'
    Left = 32
    Top = 232
  end
  object SaveXMLDocument: TXMLDocument
    Active = True
    NodeIndentStr = '        '
    Left = 80
    Top = 232
    DOMVendorDesc = 'MSXML'
  end
end
