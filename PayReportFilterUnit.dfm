object PayReportFilter: TPayReportFilter
  Left = 452
  Top = 395
  BorderStyle = bsDialog
  Caption = #1057#1087#1088#1072#1074#1082#1072' '#1087#1086' '#1087#1083#1072#1090#1077#1078#1072#1084
  ClientHeight = 88
  ClientWidth = 288
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
  object Label1: TLabel
    Left = 8
    Top = 35
    Width = 63
    Height = 13
    Caption = #1042#1099#1073#1086#1088#1082#1072' '#1087#1086' '
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
  object listConds: TComboBox
    Left = 104
    Top = 32
    Width = 177
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    Items.Strings = (
      #1044#1040#1058#1045' '#1054#1058#1063#1045#1058#1040
      #1044#1040#1058#1045' '#1055#1051#1040#1058#1045#1046#1040)
  end
  object OK: TButton
    Left = 189
    Top = 57
    Width = 91
    Height = 25
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 3
  end
end
