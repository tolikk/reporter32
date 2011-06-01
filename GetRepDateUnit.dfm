object GetRepDate: TGetRepDate
  Left = 330
  Top = 240
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1054#1090#1095#1105#1090#1085#1072#1103' '#1076#1072#1090#1072
  ClientHeight = 108
  ClientWidth = 235
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
    Top = 8
    Width = 73
    Height = 13
    Caption = #1054#1090#1095#1105#1090#1085#1072#1103' '#1076#1072#1090#1072
  end
  object Label2: TLabel
    Left = 4
    Top = 32
    Width = 99
    Height = 13
    Caption = #1053#1072#1083#1086#1075' '#1085#1072' '#1092#1080#1079'. '#1083#1080#1094#1086
  end
  object Label3: TLabel
    Left = 4
    Top = 56
    Width = 93
    Height = 13
    Caption = #1053#1072#1083#1086#1075' '#1085#1072' '#1102#1088'. '#1083#1080#1094#1086
  end
  object RepDate: TDateTimePicker
    Left = 112
    Top = 4
    Width = 121
    Height = 21
    CalAlignment = dtaLeft
    Date = 37937.4000050579
    Time = 37937.4000050579
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 0
  end
  object Button1: TButton
    Left = 112
    Top = 80
    Width = 119
    Height = 25
    Caption = #1042#1099#1087#1086#1083#1080#1090#1100
    ModalResult = 1
    TabOrder = 1
  end
  object FizTax: TRxSpinEdit
    Left = 112
    Top = 28
    Width = 121
    Height = 21
    MaxValue = 50
    Value = 40
    TabOrder = 2
  end
  object UrTax: TRxSpinEdit
    Left = 112
    Top = 52
    Width = 121
    Height = 21
    MaxValue = 50
    TabOrder = 3
  end
end
