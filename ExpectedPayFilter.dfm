object ExpectedPaysForm: TExpectedPaysForm
  Left = 495
  Top = 509
  BorderStyle = bsDialog
  Caption = #1054#1078#1080#1076#1072#1077#1084#1099#1077' '#1087#1083#1072#1090#1077#1078#1080
  ClientHeight = 112
  ClientWidth = 287
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 8
    Top = 7
    Width = 88
    Height = 13
    Caption = #1054#1090#1095#1105#1090#1085#1099#1081' '#1087#1077#1088#1080#1086#1076
  end
  object Label1: TLabel
    Left = 8
    Top = 35
    Width = 89
    Height = 13
    AutoSize = False
    Caption = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100
  end
  object Label3: TLabel
    Left = 9
    Top = 58
    Width = 89
    Height = 13
    AutoSize = False
    Caption = #1057#1090#1088#1072#1093#1086#1074#1097#1080#1082
  end
  object DtFrom: TDateTimePicker
    Left = 103
    Top = 4
    Width = 89
    Height = 21
    CalAlignment = dtaLeft
    Date = 38213.7995614005
    Time = 38213.7995614005
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 0
  end
  object DtTo: TDateTimePicker
    Left = 200
    Top = 4
    Width = 82
    Height = 21
    CalAlignment = dtaLeft
    Date = 38213.7995742361
    Time = 38213.7995742361
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 1
  end
  object OK: TButton
    Left = 189
    Top = 81
    Width = 91
    Height = 25
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 2
  end
  object Name: TEdit
    Left = 104
    Top = 32
    Width = 177
    Height = 21
    CharCase = ecUpperCase
    TabOrder = 3
  end
  object Agent: TEdit
    Left = 104
    Top = 56
    Width = 177
    Height = 21
    CharCase = ecUpperCase
    TabOrder = 4
  end
end
