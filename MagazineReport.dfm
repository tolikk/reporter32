object MagazineForm: TMagazineForm
  Left = 372
  Top = 342
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1046#1091#1088#1085#1072#1083' '#1091#1095#1105#1090#1072' '#1079#1072#1082#1083#1102#1095#1105#1085#1085#1099#1093' '#1076#1086#1075#1086#1074#1086#1088#1086#1074
  ClientHeight = 80
  ClientWidth = 312
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
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 73
    Height = 13
    Caption = #1054#1090#1095#1105#1090#1085#1072#1103' '#1076#1072#1090#1072
  end
  object Label2: TLabel
    Left = 8
    Top = 32
    Width = 88
    Height = 13
    Caption = #1054#1090#1074#1077#1090#1089#1090#1074#1077#1085#1085#1086#1089#1090#1100
  end
  object RepDate: TDateTimePicker
    Left = 108
    Top = 4
    Width = 201
    Height = 21
    CalAlignment = dtaLeft
    Date = 37711.5085827083
    Time = 37711.5085827083
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 0
  end
  object OtvSumma: TEdit
    Left = 108
    Top = 28
    Width = 133
    Height = 21
    TabOrder = 1
  end
  object OtvCurr: TEdit
    Left = 244
    Top = 28
    Width = 65
    Height = 21
    TabOrder = 2
  end
  object btnOk: TButton
    Left = 216
    Top = 56
    Width = 95
    Height = 21
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 3
  end
end
