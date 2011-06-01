object GetPeriod: TGetPeriod
  Left = 363
  Top = 259
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1054#1090#1095#1105#1090#1085#1099#1081' '#1087#1077#1088#1080#1086#1076
  ClientHeight = 249
  ClientWidth = 241
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 4
    Top = 56
    Width = 64
    Height = 13
    Caption = #1044#1072#1090#1072' '#1085#1072#1095#1072#1083#1072
  end
  object Label2: TLabel
    Left = 4
    Top = 80
    Width = 82
    Height = 13
    Caption = #1044#1072#1090#1072' '#1086#1082#1086#1085#1095#1072#1085#1080#1103
  end
  object Label3: TLabel
    Left = 4
    Top = 4
    Width = 35
    Height = 13
    Caption = #1056#1077#1078#1080#1084
  end
  object Label4: TLabel
    Left = 4
    Top = 104
    Width = 65
    Height = 13
    Caption = #1069#1082#1089#1087#1086#1088#1090#1099' '#1079#1072
    Enabled = False
  end
  object StartDate: TDateTimePicker
    Left = 104
    Top = 52
    Width = 133
    Height = 21
    CalAlignment = dtaLeft
    Date = 37954.6118685069
    Time = 37954.6118685069
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 0
    OnExit = StartDateExit
  end
  object EndDate: TDateTimePicker
    Left = 104
    Top = 76
    Width = 133
    Height = 21
    CalAlignment = dtaLeft
    Date = 37954.6119108333
    Time = 37954.6119108333
    ShowCheckbox = True
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 1
  end
  object Ok: TBitBtn
    Left = 164
    Top = 224
    Width = 71
    Height = 20
    Caption = 'Ok'
    ModalResult = 1
    TabOrder = 2
  end
  object ExpMode: TComboBox
    Left = 4
    Top = 20
    Width = 233
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
    OnChange = ExpModeChange
    Items.Strings = (
      #1055#1086' '#1076#1072#1090#1077' '#1074#1074#1086#1076#1072' '#1076#1072#1085#1085#1099#1093
      #1055#1086' '#1092#1080#1083#1100#1090#1088#1091
      #1042#1089#1077' '#1085#1077#1101#1082#1089#1087#1086#1088#1090#1080#1088#1086#1074#1072#1074#1096#1080#1077#1089#1103
      #1055#1086#1074#1090#1086#1088#1080#1090#1100' '#1101#1082#1089#1087#1086#1088#1090)
  end
  object ListExports: TListBox
    Left = 4
    Top = 120
    Width = 233
    Height = 97
    Enabled = False
    ItemHeight = 13
    TabOrder = 4
  end
  object Query: TQuery
    DatabaseName = 'DBPOLAND'
    SQL.Strings = (
      'SELECT DISTINCT EXPDATE'
      'FROM POLANDPL'
      'WHERE EXPDATE IS NOT NULL')
    Left = 20
    Top = 156
  end
end
