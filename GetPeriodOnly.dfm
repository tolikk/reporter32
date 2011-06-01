object GetPeriodForm: TGetPeriodForm
  Left = 333
  Top = 336
  Width = 293
  Height = 94
  Caption = #1059#1082#1072#1078#1080' '#1087#1077#1088#1080#1086#1076
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
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
  object btnOk: TButton
    Left = 169
    Top = 30
    Width = 111
    Height = 25
    Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
    ModalResult = 1
    TabOrder = 2
  end
end
