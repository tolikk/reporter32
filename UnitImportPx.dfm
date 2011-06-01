object formImportPx: TformImportPx
  Left = 352
  Top = 289
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
  ClientHeight = 76
  ClientWidth = 368
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 4
    Top = 8
    Width = 65
    Height = 13
    Caption = #1054#1090' '#1082#1086#1075#1086
  end
  object Label2: TLabel
    Left = 4
    Top = 32
    Width = 100
    Height = 13
    Caption = #1042#1086#1079#1085#1072#1075#1088#1072#1078#1076#1077#1085#1080#1077', %'
  end
  object Button1: TButton
    Left = 260
    Top = 52
    Width = 103
    Height = 21
    Caption = #1048#1084#1087#1086#1088#1090
    TabOrder = 0
    OnClick = Button1Click
  end
  object AgPercent: TRxSpinEdit
    Left = 112
    Top = 28
    Width = 253
    Height = 21
    MaxValue = 50
    ValueType = vtFloat
    TabOrder = 1
  end
  object Agent: TRxDBLookupCombo
    Left = 136
    Top = 4
    Width = 229
    Height = 21
    DropDownCount = 24
    LookupField = 'Agent_code'
    LookupDisplay = 'Name'
    LookupSource = DataSource
    TabOrder = 2
    OnChange = AgentChange
  end
  object txtFilter: TEdit
    Left = 112
    Top = 4
    Width = 21
    Height = 21
    MaxLength = 2
    TabOrder = 3
    OnChange = txtFilterChange
  end
  object DataSource: TDataSource
    DataSet = Agents
    Left = 68
    Top = 32
  end
  object Agents: TRxQuery
    DatabaseName = 'DB'
    SQL.Strings = (
      'SELECT * FROM AGENT WHERE MANDPCNT>0 AND NAME LIKE %FILTER')
    Macros = <
      item
        DataType = ftString
        Name = 'FILTER'
        ParamType = ptInput
        Value = #39'%'#39
      end>
    Left = 20
    Top = 32
  end
end
