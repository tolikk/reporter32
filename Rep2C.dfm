object Form2C: TForm2C
  Left = 250
  Top = 195
  Width = 500
  Height = 200
  Caption = #1054#1090#1095#1105#1090' 2'#1057
  Color = 14024191
  Constraints.MinHeight = 200
  Constraints.MinWidth = 500
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PopupMenu = PopupMenu
  Position = poScreenCenter
  ShowHint = True
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    492
    173)
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 394
    Top = 148
    Width = 96
    Height = 21
    Anchors = [akRight, akBottom]
    Caption = #1047#1072#1082#1088#1099#1090#1100
    Flat = True
    OnClick = SpeedButton1Click
  end
  object Panel: TPanel
    Left = 0
    Top = 0
    Width = 492
    Height = 141
    Align = alTop
    Anchors = [akLeft, akTop, akRight, akBottom]
    Color = 14024191
    TabOrder = 0
    OnResize = PanelResize
    DesignSize = (
      492
      141)
    object SpeedButton2: TSpeedButton
      Left = 4
      Top = 116
      Width = 121
      Height = 22
      Anchors = [akLeft, akBottom]
      Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100' '#1074#1089#1105' '#1089#1088#1072#1079#1091
      Flat = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clGreen
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsUnderline]
      ParentFont = False
      OnClick = SpeedButton2Click
    end
  end
  object ProgressBar: TProgressBar
    Left = 4
    Top = 148
    Width = 387
    Height = 21
    Anchors = [akLeft, akBottom]
    Min = 0
    Max = 100
    TabOrder = 1
  end
  object DatabaseRpt: TDatabase
    DatabaseName = 'DatabaseRpt'
    SessionName = 'Default'
    Left = 28
    Top = 12
  end
  object WorkSQL: TQuery
    DatabaseName = 'DatabaseRpt'
    Left = 96
    Top = 12
  end
  object MANDRPT: TTable
    DatabaseName = 'DatabaseRpt'
    Exclusive = True
    TableName = 'MANDRPT'
    Left = 164
    Top = 12
  end
  object Timer: TTimer
    Interval = 300
    OnTimer = TimerTimer
    Left = 228
    Top = 12
  end
  object PopupMenu: TPopupMenu
    Left = 352
    Top = 80
    object N1: TMenuItem
      Caption = #1058#1072#1073#1083#1080#1094#1072' '#1082#1091#1088#1089#1086#1074' '#1074#1072#1083#1102#1090'...'
      OnClick = N1Click
    end
  end
end
