object RepParams: TRepParams
  Left = 360
  Top = 175
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1044#1072#1085#1085#1099#1077' '#1076#1083#1103' '#1086#1090#1095#1105#1090#1072
  ClientHeight = 63
  ClientWidth = 295
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
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
    Left = 108
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
    Width = 89
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
  object Button1: TButton
    Left = 180
    Top = 32
    Width = 111
    Height = 25
    Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
    ModalResult = 1
    TabOrder = 2
  end
end
