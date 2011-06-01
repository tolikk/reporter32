object Modifier: TModifier
  Left = 289
  Top = 285
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1059#1090#1080#1083#1080#1090#1072
  ClientHeight = 121
  ClientWidth = 447
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  ShowHint = True
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 392
    Top = 76
    Width = 49
    Height = 22
    Caption = #1047#1072#1082#1088#1099#1090#1100
    Flat = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsUnderline]
    ParentFont = False
    OnClick = SpeedButton1Click
  end
  object Label1: TLabel
    Left = 4
    Top = 8
    Width = 62
    Height = 13
    Caption = #1044#1080#1088#1077#1082#1090#1086#1088#1080#1081
  end
  object lbPath: TLabel
    Left = 248
    Top = 9
    Width = 193
    Height = 13
    AutoSize = False
  end
  object lbInfo: TLabel
    Left = 8
    Top = 81
    Width = 376
    Height = 13
    Hint = 
      #1042' '#1101#1090#1086#1090' '#1082#1072#1090#1072#1083#1086#1075' '#1073#1091#1076#1091' '#1079#1072#1087#1080#1089#1072#1085#1099' '#1092#1072#1081#1083#1099' '#1087#1077#1088#1077#1076' '#1080#1093' '#1080#1079#1084#1077#1085#1077#1085#1080#1077#1084#13#10#1042' '#1089#1083#1091#1095#1072#1077 +
      ' '#1087#1086#1074#1088#1077#1078#1076#1077#1085#1080#1103' '#1076#1072#1085#1085#1099#1093' '#1087#1088#1080' '#1074#1099#1087#1086#1083#1085#1077#1085#1080#1080' '#1086#1087#1077#1088#1072#1094#1080#1080' '#1074' '#1101#1090#1086#1084' '#13#10#1082#1072#1090#1072#1083#1086#1075#1077' '#1084#1086 +
      #1078#1085#1086' '#1085#1072#1081#1090#1080' '#1082#1086#1087#1080#1102' '#1092#1072#1081#1083#1072' ('#1080#1085#1076#1077#1082#1089#1099' '#1085#1077' '#1082#1086#1087#1080#1088#1091#1102#1090#1089#1103')'
    AutoSize = False
    Color = clMenu
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
  end
  object iniInfo: TLabel
    Left = 8
    Top = 65
    Width = 376
    Height = 13
    Hint = 
      #1042' '#1101#1090#1086#1090' '#1082#1072#1090#1072#1083#1086#1075' '#1073#1091#1076#1091' '#1079#1072#1087#1080#1089#1072#1085#1099' '#1092#1072#1081#1083#1099' '#1087#1077#1088#1077#1076' '#1080#1093' '#1080#1079#1084#1077#1085#1077#1085#1080#1077#1084#13#10#1042' '#1089#1083#1091#1095#1072#1077 +
      ' '#1087#1086#1074#1088#1077#1078#1076#1077#1085#1080#1103' '#1076#1072#1085#1085#1099#1093' '#1087#1088#1080' '#1074#1099#1087#1086#1083#1085#1077#1085#1080#1080' '#1086#1087#1077#1088#1072#1094#1080#1080' '#1074' '#1101#1090#1086#1084' '#13#10#1082#1072#1090#1072#1083#1086#1075#1077' '#1084#1086 +
      #1078#1085#1086' '#1085#1072#1081#1090#1080' '#1082#1086#1087#1080#1102' '#1092#1072#1081#1083#1072' ('#1080#1085#1076#1077#1082#1089#1099' '#1085#1077' '#1082#1086#1087#1080#1088#1091#1102#1090#1089#1103')'
    AutoSize = False
    Color = clMenu
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clNavy
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
  end
  object PackTablesBtn: TButton
    Left = 4
    Top = 32
    Width = 137
    Height = 25
    Hint = 
      #1059#1087#1072#1082#1086#1074#1099#1074#1077#1090' '#1076#1072#1085#1085#1099#1077' '#1074' '#1090#1072#1073#1083#1080#1094#1077#13#10#1069#1090#1086' '#1091#1084#1077#1085#1100#1096#1072#1077#1090' '#1088#1072#1079#1084#1077#1088' '#1090#1072#1073#1083#1080#1094#1099' '#1080' '#1074#1086#1079#1084 +
      #1086#1078#1085#1086' '#1091#1074#1077#1083#1080#1095#1080#1090' '#13#10#1089#1082#1086#1088#1086#1089#1090#1100' '#1088#1072#1073#1086#1090#1099' '#1089' '#1085#1077#1081
    Caption = #1059#1087#1086#1088#1103#1076#1086#1095#1080#1090#1100' '#1076#1072#1085#1085#1099#1077'...'
    TabOrder = 0
    OnClick = PackTablesBtnClick
  end
  object Button3: TButton
    Left = 144
    Top = 32
    Width = 137
    Height = 25
    Hint = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1086#1083#1077' '#1074' '#1090#1072#1073#1083#1080#1094#1091
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1086#1083#1077'...'
    TabOrder = 1
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 284
    Top = 32
    Width = 157
    Height = 25
    Hint = 
      #1042#1099#1087#1086#1083#1085#1080#1090#1100' '#1085#1072#1073#1086#1088' '#1082#1086#1084#1072#1085#1076':'#13#10'MSG - '#1087#1086#1082#1072#1079#1072#1090#1100' '#1089#1086#1086#1073#1097#1077#1085#1080#1077#13#10'RUSS - '#1088#1091#1089#1080#1092#1080 +
      #1094#1080#1088#1086#1074#1072#1090#1100' '#1090#1072#1073#1083#1080#1094#1091#13#10'DB - '#1091#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1090#1077#1082#1091#1097#1091#1102' '#1041#1044#13#10'INI - '#1084#1077#1085#1103#1090#1100' '#1095#1080#1089#1090#1080 +
      #1090#1100' '#1092#1072#1081#1083' '#1085#1072#1089#1090#1088#1086#1077#1082#13#10'SQL '#1047#1072#1087#1088#1086#1089' '#1082' '#1041#1044
    Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100' '#1085#1072#1073#1086#1088' '#1082#1086#1084#1072#1085#1076'...'
    TabOrder = 2
    OnClick = Button4Click
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 102
    Width = 447
    Height = 19
    Panels = <>
    SimplePanel = False
  end
  object ListAliases: TComboBox
    Left = 72
    Top = 6
    Width = 169
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 4
    OnChange = ListAliasesChange
  end
  object OpenDialog: TOpenDialog
    DefaultExt = 'db'
    Filter = #1058#1072#1073#1083#1080#1094#1099' '#1055#1072#1088#1072#1076#1086#1082#1089' (*.db)|*.db'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Title = #1042#1099#1073#1077#1088#1080' '#1086#1076#1085#1091' '#1080#1083#1080' '#1085#1077#1089#1082#1086#1083#1100#1082#1086' '#1090#1072#1073#1083#1080#1094
    Left = 128
    Top = 60
  end
  object Table: TTable
    Exclusive = True
    Left = 12
    Top = 60
  end
  object WorkQuery: TQuery
    Left = 68
    Top = 60
  end
  object BatchMove: TBatchMove
    Destination = TableDest
    Mode = batCopy
    Source = Table
    Left = 196
    Top = 60
  end
  object TableDest: TTable
    Left = 252
    Top = 60
  end
  object OpenDialogTxt: TOpenDialog
    DefaultExt = 'taskcmd'
    Filter = #1060#1072#1081#1083#1099' '#1082#1086#1084#1072#1085#1076' (*.taskcmd)|*.taskcmd'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Title = #1054#1090#1082#1088#1099#1090#1100' '#1082#1086#1084#1072#1085#1076#1085#1099#1081' '#1092#1072#1081#1083
    Left = 312
    Top = 60
  end
  object Database: TDatabase
    SessionName = 'Default'
    Left = 376
    Top = 64
  end
end
