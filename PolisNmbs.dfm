object PolisNmbsComp: TPolisNmbsComp
  Left = 333
  Top = 192
  Width = 880
  Height = 431
  Caption = #1053#1086#1084#1077#1088#1072' '#1087#1086#1083#1080#1089#1086#1074', '#1074#1099#1076#1072#1085#1085#1099#1077' '#1082#1086#1084#1087#1072#1085#1080#1080
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object DBGrid: TDBGrid
    Left = 0
    Top = 0
    Width = 872
    Height = 356
    Align = alClient
    DataSource = DataSource
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    OnKeyDown = DBGridKeyDown
    Columns = <
      item
        Expanded = False
        FieldName = 'S'
        Width = 43
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'START'
        Width = 102
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'END'
        Width = 94
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Count'
        Visible = True
      end
      item
        DropDownRows = 24
        Expanded = False
        FieldName = 'AgentCombo'
        Width = 184
        Visible = True
      end>
  end
  object Panel1: TPanel
    Left = 0
    Top = 356
    Width = 872
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object StatusLine: TLabel
      Left = 5
      Top = 14
      Width = 396
      Height = 13
      AutoSize = False
    end
    object Panel2: TPanel
      Left = 535
      Top = 0
      Width = 337
      Height = 41
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      DesignSize = (
        337
        41)
      object Button3: TButton
        Left = 126
        Top = 9
        Width = 105
        Height = 25
        Anchors = [akTop, akRight]
        Caption = #1059#1076#1072#1083#1080#1090#1100
        TabOrder = 0
        OnClick = Button3Click
      end
      object Button2: TButton
        Left = 24
        Top = 9
        Width = 99
        Height = 25
        Anchors = [akTop, akRight]
        Caption = #1044#1086#1073#1072#1074#1080#1090#1100
        TabOrder = 1
        OnClick = Button2Click
      end
      object Button1: TButton
        Left = 233
        Top = 9
        Width = 99
        Height = 25
        Anchors = [akTop, akRight]
        Caption = #1047#1072#1082#1088#1099#1090#1100
        TabOrder = 2
        OnClick = Button1Click
      end
    end
  end
  object Table: TTable
    BeforePost = TableBeforePost
    AfterScroll = TableAfterScroll
    OnCalcFields = TableCalcFields
    DatabaseName = 'BASO'
    TableName = 'POLISNMB.db'
    Left = 40
    Top = 40
    object TableS: TStringField
      DisplayLabel = #1057#1077#1088#1080#1103
      FieldName = 'S'
      Required = True
      Size = 4
    end
    object TableSTART: TFloatField
      DisplayLabel = #1053#1072#1095#1072#1083#1100#1085#1099#1081' '#1085#1086#1084#1077#1088
      FieldName = 'START'
      Required = True
      DisplayFormat = '### ### ### ###'
      EditFormat = '############'
    end
    object TableEND: TFloatField
      DisplayLabel = #1050#1086#1085#1077#1095#1085#1099#1081' '#1085#1086#1084#1077#1088
      FieldName = 'END'
      Required = True
      DisplayFormat = '### ### ### ###'
      EditFormat = '############'
    end
    object TableCount: TIntegerField
      DisplayLabel = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' (5000 '#1084#1072#1082#1089')'
      FieldKind = fkCalculated
      FieldName = 'Count'
      ReadOnly = True
      DisplayFormat = '### ### ### ###'
      Calculated = True
    end
  end
  object DataSource: TDataSource
    DataSet = TableAg
    Left = 96
    Top = 40
  end
  object Agent: TTable
    Active = True
    BeforePost = TableBeforePost
    DatabaseName = 'BASO'
    TableName = 'Agent.DB'
    Left = 96
    Top = 112
    object AgentAgent_code: TStringField
      FieldName = 'Agent_code'
      Size = 4
    end
    object AgentName: TStringField
      FieldName = 'Name'
      Size = 60
    end
  end
  object TableAg: TTable
    BeforePost = TableAgBeforePost
    AfterScroll = TableAgAfterScroll
    OnCalcFields = TableAgCalcFields
    DatabaseName = 'BASO'
    TableName = 'POLISNMB.DB'
    Left = 48
    Top = 112
    object TableAgS: TStringField
      DisplayLabel = #1057#1077#1088#1080#1103
      FieldName = 'S'
      Size = 4
    end
    object TableAgSTART: TFloatField
      DisplayLabel = #1053#1072#1095#1072#1083#1100#1085#1099#1081' '#1085#1086#1084#1077#1088
      FieldName = 'START'
      Required = True
      DisplayFormat = '### ### ### ###'
      EditFormat = '############'
    end
    object TableAgEND: TFloatField
      DisplayLabel = #1050#1086#1085#1077#1095#1085#1099#1081' '#1085#1086#1084#1077#1088
      FieldName = 'END'
      Required = True
      DisplayFormat = '### ### ### ###'
      EditFormat = '############'
    end
    object TableAgAgent: TStringField
      FieldName = 'Agent'
      Size = 4
    end
    object TableAgAgentCombo: TStringField
      DisplayLabel = #1040#1075#1077#1085#1090' ('#1087#1091#1089#1090#1086' - '#1042#1057#1045', Del - '#1086#1095#1080#1089#1090#1080#1090#1100')'
      FieldKind = fkLookup
      FieldName = 'AgentCombo'
      LookupDataSet = Agent
      LookupKeyFields = 'Agent_code'
      LookupResultField = 'Name'
      KeyFields = 'Agent'
      Size = 50
      Lookup = True
    end
    object TableAgCount: TIntegerField
      DisplayLabel = #1050#1086#1083#1080#1095#1077#1089#1090#1074#1086' (5000 '#1084#1072#1082#1089')'
      FieldKind = fkCalculated
      FieldName = 'Count'
      DisplayFormat = '### ### ### ###'
      Calculated = True
    end
  end
end
