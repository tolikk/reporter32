object BG_Buro: TBG_Buro
  Left = 289
  Top = 189
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' '#1041#1070#1056#1054
  ClientHeight = 274
  ClientWidth = 307
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
    Left = 8
    Top = 8
    Width = 83
    Height = 13
    Caption = #1069#1082#1089#1087#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100
  end
  object Label2: TLabel
    Left = 8
    Top = 32
    Width = 126
    Height = 13
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1074' '#1076#1080#1088#1077#1082#1090#1086#1088#1080#1102
  end
  object Path: TEdit
    Left = 140
    Top = 28
    Width = 161
    Height = 21
    TabOrder = 0
    Text = 'C:\'
  end
  object OK: TButton
    Left = 200
    Top = 248
    Width = 103
    Height = 21
    Caption = #1069#1082#1089#1087#1086#1088#1090
    ModalResult = 1
    TabOrder = 1
  end
  object NExport: TComboBox
    Left = 140
    Top = 4
    Width = 161
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 2
    Items.Strings = (
      #1053#1086#1074#1099#1077' '#1076#1072#1085#1085#1099#1077
      #1055#1086#1089#1083#1077#1076#1085#1080#1081' '#1101#1082#1089#1087#1086#1088#1090
      #1055#1088#1077#1076#1087#1086#1089#1083#1077#1076#1085#1080#1081' '#1101#1082#1089#1087#1086#1088#1090)
  end
  object GroupBox1: TGroupBox
    Left = 4
    Top = 56
    Width = 298
    Height = 181
    Caption = ' '#1048#1085#1092#1072' '#1087#1086' '#1101#1082#1089#1087#1086#1088#1090#1072#1084' '
    TabOrder = 3
    object listExports: TListBox
      Left = 8
      Top = 16
      Width = 281
      Height = 133
      ItemHeight = 13
      TabOrder = 0
    end
    object Button1: TButton
      Left = 204
      Top = 152
      Width = 85
      Height = 21
      Caption = #1059#1076#1072#1083#1080#1090#1100
      TabOrder = 1
      OnClick = Button1Click
    end
    object Button2: TButton
      Left = 116
      Top = 152
      Width = 85
      Height = 21
      Caption = #1055#1086#1082#1072#1079#1072#1090#1100
      TabOrder = 2
      OnClick = Button2Click
    end
  end
  object IsTest: TCheckBox
    Left = 8
    Top = 248
    Width = 161
    Height = 17
    Caption = #1058#1077#1089#1090#1086#1074#1099#1081' '#1087#1088#1086#1075#1086#1085
    TabOrder = 4
  end
  object FormStorage: TFormStorage
    IniFileName = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077
    IniSection = #1041#1077#1083'. '#1047#1050
    Options = [fpState]
    UseRegistry = True
    StoredProps.Strings = (
      'Path.Text')
    StoredValues = <>
    Left = 24
    Top = 4
  end
end
