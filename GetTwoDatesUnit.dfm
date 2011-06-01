object GetTwoDates: TGetTwoDates
  Left = 517
  Top = 326
  BorderStyle = bsDialog
  Caption = #1042#1074#1077#1076#1080' '#1076#1072#1090#1099
  ClientHeight = 87
  ClientWidth = 240
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
  object Label1: TLabel
    Left = 4
    Top = 12
    Width = 64
    Height = 13
    Caption = #1044#1072#1090#1072' '#1085#1072#1095#1072#1083#1072
  end
  object Label2: TLabel
    Left = 4
    Top = 36
    Width = 82
    Height = 13
    Caption = #1044#1072#1090#1072' '#1086#1082#1086#1085#1095#1072#1085#1080#1103
  end
  object StartDate: TDateTimePicker
    Left = 104
    Top = 8
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
  end
  object EndDate: TDateTimePicker
    Left = 104
    Top = 32
    Width = 133
    Height = 21
    CalAlignment = dtaLeft
    Date = 37954.6119108333
    Time = 37954.6119108333
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 1
  end
  object Button1: TButton
    Left = 160
    Top = 60
    Width = 75
    Height = 21
    Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
    ModalResult = 1
    TabOrder = 2
  end
end
