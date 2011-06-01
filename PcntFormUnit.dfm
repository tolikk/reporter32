object PcntForm: TPcntForm
  Left = 422
  Top = 204
  BorderStyle = bsDialog
  Caption = #1055#1088#1086#1094#1077#1085#1090#1099' '#1087#1086' '#1087#1086#1083#1080#1089#1072#1084
  ClientHeight = 122
  ClientWidth = 217
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 90
    Height = 13
    Caption = #1055#1077#1088#1077#1083#1086#1084#1085#1072#1103' '#1076#1072#1090#1072
  end
  object Label2: TLabel
    Left = 8
    Top = 28
    Width = 51
    Height = 13
    Caption = '% '#1076#1086' '#1076#1072#1090#1099
  end
  object Label3: TLabel
    Left = 8
    Top = 56
    Width = 89
    Height = 13
    Caption = '% '#1085#1072#1095#1080#1085#1072#1103' '#1089' '#1076#1072#1090#1099
  end
  object btnOk: TButton
    Left = 56
    Top = 88
    Width = 75
    Height = 25
    Caption = 'Ok'
    Default = True
    ModalResult = 1
    TabOrder = 0
  end
  object btnCancel: TButton
    Left = 136
    Top = 88
    Width = 75
    Height = 25
    Caption = #1054#1090#1084#1077#1085#1072
    ModalResult = 2
    TabOrder = 1
  end
  object Date: TDateTimePicker
    Left = 112
    Top = 4
    Width = 97
    Height = 21
    CalAlignment = dtaLeft
    Date = 37987.5964843056
    Time = 37987.5964843056
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 2
  end
  object PcntBefore: TRxSpinEdit
    Left = 112
    Top = 28
    Width = 97
    Height = 21
    MaxValue = 99
    ValueType = vtFloat
    Value = 7
    TabOrder = 3
  end
  object PcntAfter: TRxSpinEdit
    Left = 112
    Top = 52
    Width = 97
    Height = 21
    MaxValue = 99
    ValueType = vtFloat
    Value = 14
    TabOrder = 4
  end
  object FormStorage: TFormStorage
    IniFileName = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077
    IniSection = #1055#1088#1086#1094#1077#1085#1090#1099
    StoredProps.Strings = (
      'Date.Date'
      'PcntBefore.Value'
      'PcntAfter.Value')
    StoredValues = <>
    Left = 12
    Top = 72
  end
end
