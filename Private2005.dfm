object Private2005Form: TPrivate2005Form
  Left = 279
  Top = 119
  Width = 767
  Height = 536
  Caption = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077' '#1078#1080#1079#1085#1080' 2005'
  Color = clBtnFace
  Constraints.MinHeight = 500
  Constraints.MinWidth = 550
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object StatusPanel: TPanel
    Left = 0
    Top = 453
    Width = 759
    Height = 37
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object StatusBar: TLabel
      Left = 344
      Top = 12
      Width = 237
      Height = 15
      AutoSize = False
    end
    object ExpInfo: TLabel
      Left = 344
      Top = 12
      Width = 137
      Height = 13
      AutoSize = False
      Caption = '.'
    end
    object Panel2: TPanel
      Left = 495
      Top = 0
      Width = 264
      Height = 37
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object Button1: TButton
        Left = 168
        Top = 8
        Width = 95
        Height = 25
        Caption = #1047#1072#1082#1088#1099#1090#1100
        TabOrder = 0
        OnClick = Button1Click
      end
    end
    object ProgressBar: TProgressBar
      Left = 4
      Top = 12
      Width = 333
      Height = 16
      Min = 0
      Max = 100
      TabOrder = 1
    end
  end
  object PageControl: TPageControl
    Left = 0
    Top = 0
    Width = 759
    Height = 453
    ActivePage = FilterPage
    Align = alClient
    TabIndex = 1
    TabOrder = 1
    OnChange = PageControlChange
    object DataPage: TTabSheet
      Caption = #1054#1090#1095#1105#1090
      object MainGrid: TRxDBGrid
        Left = 0
        Top = 0
        Width = 751
        Height = 417
        Align = alClient
        DataSource = DataSource
        Font.Charset = RUSSIAN_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'Seria'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Number'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RegDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Insurer'
            Width = 149
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsCnt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Address'
            Width = 126
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Birthdate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Passport'
            Width = 105
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Insurer2'
            Width = 166
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'From'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FromTm'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'To'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Period'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Duration'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UridSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Tarif'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsSumCurr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PayDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RepDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Sum1'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Sum1Curr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Sum2'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Sum2Curr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'stopdate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'retsum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'retcur'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StateText'
            Width = 76
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Name'
            Width = 162
            Visible = True
          end>
      end
    end
    object FilterPage: TTabSheet
      Caption = #1060#1080#1083#1100#1090#1088
      ImageIndex = 1
      object Panel5: TPanel
        Left = 0
        Top = 0
        Width = 751
        Height = 425
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 0
        OnResize = Panel5Resize
        object Label1: TLabel
          Left = 8
          Top = 12
          Width = 34
          Height = 13
          Caption = #1053#1086#1084#1077#1088
        end
        object Label2: TLabel
          Left = 264
          Top = 12
          Width = 16
          Height = 13
          Caption = #1055#1054
        end
        object Label3: TLabel
          Left = 8
          Top = 36
          Width = 62
          Height = 13
          Caption = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
        end
        object Label4: TLabel
          Left = 122
          Top = 12
          Width = 7
          Height = 13
          Alignment = taRightJustify
          Caption = #1057
        end
        object Label5: TLabel
          Left = 122
          Top = 36
          Width = 7
          Height = 13
          Alignment = taRightJustify
          Caption = #1057
        end
        object Label12: TLabel
          Left = 264
          Top = 36
          Width = 16
          Height = 13
          Caption = #1055#1054
        end
        object CntChecked: TLabel
          Left = 8
          Top = 100
          Width = 93
          Height = 13
          AutoSize = False
          Caption = #1042#1099#1073#1088#1072#1085#1086' 0'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clRed
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
        end
        object SpeedButton: TSpeedButton
          Left = 112
          Top = 116
          Width = 23
          Height = 21
          Caption = '!'
          Flat = True
          OnClick = SpeedButtonClick
        end
        object Label16: TLabel
          Left = 8
          Top = 60
          Width = 85
          Height = 13
          Caption = #1044#1072#1090#1072' '#1080#1079#1084#1077#1085#1077#1085#1080#1103
        end
        object Label17: TLabel
          Left = 122
          Top = 60
          Width = 7
          Height = 13
          Alignment = taRightJustify
          Caption = #1057
        end
        object Label18: TLabel
          Left = 264
          Top = 60
          Width = 16
          Height = 13
          Caption = #1055#1054
        end
        object NmbFrom: TEdit
          Left = 136
          Top = 8
          Width = 117
          Height = 21
          TabOrder = 0
        end
        object NmbTo: TEdit
          Left = 284
          Top = 8
          Width = 121
          Height = 21
          TabOrder = 1
        end
        object RepDtFrom: TDateTimePicker
          Left = 136
          Top = 32
          Width = 117
          Height = 21
          CalAlignment = dtaLeft
          Date = 37718.5493044676
          Time = 37718.5493044676
          ShowCheckbox = True
          Checked = False
          DateFormat = dfShort
          DateMode = dmComboBox
          Kind = dtkDate
          ParseInput = False
          TabOrder = 2
        end
        object RepDtTo: TDateTimePicker
          Left = 284
          Top = 32
          Width = 121
          Height = 21
          CalAlignment = dtaLeft
          Date = 37718.5493668519
          Time = 37718.5493668519
          ShowCheckbox = True
          Checked = False
          DateFormat = dfShort
          DateMode = dmComboBox
          Kind = dtkDate
          ParseInput = False
          TabOrder = 3
        end
        object IsAgents: TCheckBox
          Left = 8
          Top = 80
          Width = 69
          Height = 17
          Caption = #1040#1075#1077#1085#1090#1099
          TabOrder = 4
        end
        object ListAgents: TCheckListBox
          Left = 136
          Top = 80
          Width = 269
          Height = 197
          ItemHeight = 13
          TabOrder = 5
          OnClick = ListAgentsClick
        end
        object ListTemplates: TComboBox
          Left = 8
          Top = 116
          Width = 101
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 6
          OnChange = ListTemplatesChange
          Items.Strings = (
            #1042#1089#1077' '#1072#1075#1077#1085#1090#1099)
        end
        object DownPanel: TPanel
          Left = 0
          Top = 285
          Width = 751
          Height = 140
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 7
          object Label6: TLabel
            Left = 8
            Top = 8
            Width = 70
            Height = 13
            Caption = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100
          end
          object Label11: TLabel
            Left = 8
            Top = 32
            Width = 54
            Height = 13
            Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
          end
          object Label13: TLabel
            Left = 8
            Top = 56
            Width = 116
            Height = 13
            Caption = #1044#1072#1090#1072' '#1086#1082#1086#1085#1095#1072#1085#1080#1103' '#1089#1090#1088#1072#1093'.'
          end
          object Label15: TLabel
            Left = 264
            Top = 56
            Width = 16
            Height = 13
            Caption = #1055#1054
          end
          object Label10: TLabel
            Left = 8
            Top = 80
            Width = 98
            Height = 13
            Caption = #1044#1072#1090#1072' '#1085#1072#1095#1072#1083#1072' '#1089#1090#1088#1072#1093'.'
          end
          object Label14: TLabel
            Left = 264
            Top = 80
            Width = 16
            Height = 13
            Caption = #1055#1054
          end
          object Apply: TButton
            Left = 300
            Top = 106
            Width = 103
            Height = 25
            Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
            TabOrder = 0
            OnClick = ApplyClick
          end
          object Insurer: TEdit
            Left = 136
            Top = 4
            Width = 85
            Height = 21
            CharCase = ecUpperCase
            TabOrder = 1
          end
          object StateCombo: TComboBox
            Left = 136
            Top = 28
            Width = 269
            Height = 21
            Style = csDropDownList
            ItemHeight = 13
            TabOrder = 2
            Items.Strings = (
              #1053#1077' '#1074#1072#1078#1085#1086
              #1053#1054#1056#1052#1040#1051#1068#1053#1067#1049
              #1048#1057#1055#1054#1056#1063#1045#1053
              #1056#1040#1057#1058#1054#1056#1043#1053#1059#1058
              #1059#1058#1045#1056#1071#1053)
          end
          object EndDateFrom: TDateTimePicker
            Left = 136
            Top = 52
            Width = 117
            Height = 21
            CalAlignment = dtaLeft
            Date = 37718.5493044676
            Time = 37718.5493044676
            ShowCheckbox = True
            Checked = False
            DateFormat = dfShort
            DateMode = dmComboBox
            Kind = dtkDate
            ParseInput = False
            TabOrder = 3
          end
          object EndDateTo: TDateTimePicker
            Left = 284
            Top = 52
            Width = 121
            Height = 21
            CalAlignment = dtaLeft
            Date = 37718.5493668519
            Time = 37718.5493668519
            ShowCheckbox = True
            Checked = False
            DateFormat = dfShort
            DateMode = dmComboBox
            Kind = dtkDate
            ParseInput = False
            TabOrder = 4
          end
          object listInsType: TComboBox
            Left = 224
            Top = 4
            Width = 101
            Height = 21
            Style = csDropDownList
            ItemHeight = 13
            TabOrder = 5
            Items.Strings = (
              #1053#1077' '#1074#1072#1078#1085#1086
              #1060#1080#1079' '#1083#1080#1094#1086
              #1070#1088' '#1083#1080#1094#1086)
          end
          object InsurType: TComboBox
            Left = 328
            Top = 4
            Width = 77
            Height = 21
            Style = csDropDownList
            ItemHeight = 13
            TabOrder = 6
            Items.Strings = (
              #1053#1077' '#1074#1072#1078#1085#1086
              #1060#1080#1079#1080#1095
              #1070#1088#1080#1076#1080#1095)
          end
          object FromDateFrom: TDateTimePicker
            Left = 136
            Top = 76
            Width = 117
            Height = 21
            CalAlignment = dtaLeft
            Date = 37718.5493044676
            Time = 37718.5493044676
            ShowCheckbox = True
            Checked = False
            DateFormat = dfShort
            DateMode = dmComboBox
            Kind = dtkDate
            ParseInput = False
            TabOrder = 7
          end
          object FromDateTo: TDateTimePicker
            Left = 284
            Top = 76
            Width = 121
            Height = 21
            CalAlignment = dtaLeft
            Date = 37718.5493668519
            Time = 37718.5493668519
            ShowCheckbox = True
            Checked = False
            DateFormat = dfShort
            DateMode = dmComboBox
            Kind = dtkDate
            ParseInput = False
            TabOrder = 8
          end
        end
        object IsStopped: TCheckBox
          Left = 408
          Top = 34
          Width = 153
          Height = 17
          Caption = #1044#1086#1089#1088#1086#1095#1085#1086' '#1087#1088#1077#1082#1088#1072#1097#1077#1085#1085#1099#1077
          TabOrder = 8
        end
        object UpdateDtFrom: TDateTimePicker
          Left = 136
          Top = 56
          Width = 117
          Height = 21
          CalAlignment = dtaLeft
          Date = 37718.5493044676
          Time = 37718.5493044676
          ShowCheckbox = True
          Checked = False
          DateFormat = dfShort
          DateMode = dmComboBox
          Kind = dtkDate
          ParseInput = False
          TabOrder = 9
        end
        object UpdateDtTo: TDateTimePicker
          Left = 284
          Top = 56
          Width = 121
          Height = 21
          CalAlignment = dtaLeft
          Date = 37718.5493668519
          Time = 37718.5493668519
          ShowCheckbox = True
          Checked = False
          DateFormat = dfShort
          DateMode = dmComboBox
          Kind = dtkDate
          ParseInput = False
          TabOrder = 10
        end
      end
    end
    object StatPage: TTabSheet
      Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
      ImageIndex = 2
      object StatisticTxt: TMemo
        Left = 0
        Top = 0
        Width = 751
        Height = 417
        Align = alClient
        Color = 14024191
        Font.Charset = RUSSIAN_CHARSET
        Font.Color = clNavy
        Font.Height = -16
        Font.Name = 'Times New Roman'
        Font.Style = [fsBold]
        ParentFont = False
        ReadOnly = True
        ScrollBars = ssVertical
        TabOrder = 0
      end
    end
    object Reports: TTabSheet
      Caption = #1054#1090#1095#1105#1090#1099
      ImageIndex = 3
      object Panel3: TPanel
        Left = 0
        Top = 0
        Width = 751
        Height = 417
        Align = alClient
        BevelOuter = bvNone
        Caption = 'Panel3'
        TabOrder = 0
        DesignSize = (
          751
          425)
        object Label7: TLabel
          Left = 4
          Top = 4
          Width = 29
          Height = 13
          Caption = #1054#1090#1095#1105#1090
        end
        object Label8: TLabel
          Left = 4
          Top = 44
          Width = 337
          Height = 13
          Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' ('#1077#1089#1083#1080' '#1085#1077#1089#1082#1086#1083#1100#1082#1086', '#1090#1086' '#1087#1077#1088#1077#1095#1080#1089#1083#1080' '#1095#1077#1088#1077#1079' '#1090#1086#1095#1082#1091' '#1089' '#1079#1072#1087#1103#1090#1086#1081')'
        end
        object Label9: TLabel
          Left = 4
          Top = 84
          Width = 52
          Height = 13
          Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090
        end
        object RepList: TComboBox
          Left = 4
          Top = 20
          Width = 733
          Height = 21
          Style = csDropDownList
          Anchors = [akLeft, akTop, akRight]
          ItemHeight = 0
          TabOrder = 0
          OnChange = RepListChange
        end
        object ParamsText: TEdit
          Left = 4
          Top = 60
          Width = 641
          Height = 21
          Anchors = [akLeft, akTop, akRight]
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
        object StartButton: TButton
          Left = 648
          Top = 60
          Width = 89
          Height = 21
          Anchors = [akTop, akRight]
          Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
          TabOrder = 2
          OnClick = StartButtonClick
        end
        object RxDBGrid1: TRxDBGrid
          Left = 4
          Top = 100
          Width = 733
          Height = 303
          Anchors = [akLeft, akTop, akRight, akBottom]
          DataSource = REPDataSource
          TabOrder = 3
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'MS Sans Serif'
          TitleFont.Style = []
        end
        object Button2: TButton
          Left = 651
          Top = 406
          Width = 87
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1042' Excel'
          TabOrder = 4
          OnClick = Button2Click
        end
        object IsNumbers: TCheckBox
          Left = 552
          Top = 410
          Width = 97
          Height = 17
          Anchors = [akRight, akBottom]
          Caption = #1053#1091#1084#1077#1088#1086#1074#1072#1090#1100
          TabOrder = 5
        end
        object IsUseFilter: TCheckBox
          Left = 8
          Top = 412
          Width = 141
          Height = 17
          Anchors = [akLeft, akBottom]
          Caption = #1048#1079' '#1090#1077#1082#1091#1097#1077#1081' '#1074#1099#1073#1086#1088#1082#1080
          TabOrder = 6
        end
      end
    end
  end
  object MainMenu: TMainMenu
    Left = 64
    Top = 264
    object N1: TMenuItem
      Caption = #1054#1090#1095#1105#1090#1099
      object N3: TMenuItem
        Caption = #1060#1080#1083#1100#1090#1088'...'
        OnClick = N3Click
      end
      object N4: TMenuItem
        Caption = #1057#1086#1088#1090#1080#1088#1086#1074#1082#1072'...'
        OnClick = N4Click
      end
      object FixSerNmb: TMenuItem
        Caption = #1060#1080#1082#1089#1080#1088#1086#1074#1072#1090#1100' '#1057#1077#1088#1080#1102'/'#1053#1086#1084#1077#1088
        OnClick = FixSerNmbClick
      end
      object N7: TMenuItem
        Caption = '-'
      end
      object Excel1: TMenuItem
        Caption = #1042' Excel'
        OnClick = Excel1Click
      end
      object InputSum: TMenuItem
        Caption = #1055#1086#1083#1091#1095#1077#1085#1085#1099#1077' '#1074#1079#1085#1086#1089#1099' + '#1050#1056'...'
        OnClick = InputSumClick
      end
      object N11: TMenuItem
        Caption = #1054#1090#1095#1105#1090' '#1087#1086' '#1072#1075#1077#1085#1090#1091'...'
        OnClick = N11Click
      end
      object N10: TMenuItem
        Caption = '-'
      end
      object MenuGlobo: TMenuItem
        Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' '#1089#1080#1089#1090#1077#1084#1091' '#1055#1054#1051#1048#1057'...'
        Default = True
        OnClick = MenuGloboClick
      end
      object To1C: TMenuItem
        Caption = #1069#1082#1089#1087#1086#1088#1090' '#1074' 1'#1057'...'
        OnClick = To1CClick
      end
      object N5: TMenuItem
        Caption = '-'
      end
      object N2: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N2Click
      end
    end
    object Statictic: TMenuItem
      Caption = ' '
    end
  end
  object MainQuery: TRxQuery
    OnCalcFields = MainQueryCalcFields
    DatabaseName = 'Priv2005Database'
    SQL.Strings = (
      'SELECT * FROM PRIV2005 T'
      'LEFT OUTER JOIN AGENT A ON T.AGENT = A.AGENT_CODE'
      'WHERE %WHERE'
      '%ORDER')
    Macros = <
      item
        DataType = ftString
        Name = 'WHERE'
        ParamType = ptInput
        Value = '0=0'
      end
      item
        DataType = ftString
        Name = 'ORDER'
        ParamType = ptInput
      end>
    Left = 280
    Top = 88
    object MainQuerySeria: TStringField
      DisplayLabel = #1057#1077#1088#1080#1103
      FieldName = 'Seria'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Seria'
      Size = 2
    end
    object MainQueryNumber: TFloatField
      DisplayLabel = #1053#1086#1084#1077#1088
      FieldName = 'Number'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Number'
    end
    object MainQueryRegDt: TDateField
      DisplayLabel = #1042#1099#1076#1072#1085
      FieldName = 'RegDt'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".RegDt'
    end
    object MainQueryInsurer: TStringField
      DisplayLabel = #1047#1072#1089#1090#1088#1072#1093#1086#1074#1072#1085#1085#1099#1081
      FieldName = 'Insurer'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Insurer'
      Size = 65
    end
    object MainQueryInsCnt: TSmallintField
      DisplayLabel = #1043#1088#1091#1087#1087#1072
      FieldName = 'InsCnt'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".InsCnt'
    end
    object MainQueryAddress: TStringField
      DisplayLabel = #1040#1076#1088#1077#1089
      FieldName = 'Address'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Address'
      Size = 65
    end
    object MainQueryBirthdate: TStringField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1086#1078#1076#1077#1085#1080#1103
      FieldName = 'Birthdate'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Birthdate'
      Size = 10
    end
    object MainQueryPassport: TStringField
      DisplayLabel = #1055#1072#1089#1087#1086#1088#1090
      FieldName = 'Passport'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Passport'
      Size = 40
    end
    object MainQueryInsurer2: TStringField
      DisplayLabel = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100
      FieldName = 'Insurer2'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Insurer2'
      Size = 65
    end
    object MainQueryFrom: TDateField
      DisplayLabel = #1057
      FieldName = 'From'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".From'
    end
    object MainQueryFromTm: TStringField
      DisplayLabel = #1042#1088#1077#1084#1103
      FieldName = 'FromTm'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".FromTm'
      Size = 5
    end
    object MainQueryTo: TDateField
      DisplayLabel = #1055#1086
      FieldName = 'To'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".To'
    end
    object MainQueryPeriod: TSmallintField
      DisplayLabel = #1055#1077#1088#1080#1086#1076
      FieldName = 'Period'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Period'
    end
    object MainQueryDuration: TSmallintField
      DisplayLabel = #1055#1088#1077#1073#1099#1074#1072#1085#1080#1077
      FieldName = 'Duration'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Duration'
    end
    object MainQueryInsSum: TFloatField
      DisplayLabel = #1057#1090#1088#1072#1093'.'#1089#1091#1084#1084#1072
      FieldName = 'InsSum'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".InsSum'
    end
    object MainQueryUridSum: TFloatField
      DisplayLabel = #1070#1088'.'#1087#1086#1084#1086#1097#1100
      FieldName = 'UridSum'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".UridSum'
    end
    object MainQueryTarif: TFloatField
      DisplayLabel = #1058#1072#1088#1080#1092
      FieldName = 'Tarif'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Tarif'
    end
    object MainQueryInsSumCurr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1090#1072#1088#1080#1092#1072
      FieldName = 'InsSumCurr'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".InsSumCurr'
      Size = 5
    end
    object MainQueryPay: TFloatField
      DisplayLabel = #1050' '#1086#1087#1083#1072#1090#1077
      FieldName = 'Pay'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Pay'
    end
    object MainQueryPayDt: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1087#1083'.'
      FieldName = 'PayDt'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".PayDt'
    end
    object MainQueryRepDt: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
      FieldName = 'RepDt'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".RepDt'
    end
    object MainQuerySum1: TFloatField
      DisplayLabel = #1054#1087#1083#1072#1090#1072'1'
      FieldName = 'Sum1'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Sum1'
    end
    object MainQuerySum1Curr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' 1'
      FieldName = 'Sum1Curr'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Sum1Curr'
      Size = 3
    end
    object MainQuerySum2: TFloatField
      DisplayLabel = #1054#1087#1083#1072#1090#1072'2'
      FieldName = 'Sum2'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Sum2'
    end
    object MainQuerySum2Curr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072'2'
      FieldName = 'Sum2Curr'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Sum2Curr'
      Size = 3
    end
    object MainQuerystopdate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1072#1089#1090#1086#1088#1078#1077#1085#1080#1103
      FieldName = 'stopdate'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".stopdate'
    end
    object MainQueryretsum: TFloatField
      DisplayLabel = #1042#1086#1079#1074#1088#1072#1090
      FieldName = 'retsum'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".retsum'
    end
    object MainQueryretcur: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1074#1086#1079#1074#1088#1072#1090#1072
      FieldName = 'retcur'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".retcur'
      Size = 3
    end
    object MainQueryStateText: TStringField
      DisplayLabel = #1057#1086#1089#1090#1086#1103#1085#1080#1077
      FieldKind = fkCalculated
      FieldName = 'StateText'
      Calculated = True
    end
    object MainQueryState: TStringField
      FieldName = 'State'
      Origin = 'PRIV2005DATABASE."priv2005.DB".State'
      Visible = False
      Size = 1
    end
    object MainQueryPSeria: TStringField
      FieldName = 'PSeria'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".PSeria'
      Visible = False
      Size = 2
    end
    object MainQueryPNumber: TFloatField
      FieldName = 'PNumber'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".PNumber'
      Visible = False
    end
    object MainQueryName: TStringField
      DisplayLabel = #1040#1075#1077#1085#1090
      FieldName = 'Name'
      Origin = 'PRIV2005DATABASE."PRIV2005.DB".Seria'
      Size = 60
    end
  end
  object Priv2005Database: TDatabase
    AliasName = 'BASO'
    Connected = True
    DatabaseName = 'Priv2005Database'
    SessionName = 'Default'
    Left = 360
    Top = 80
  end
  object DataSource: TDataSource
    DataSet = MainQuery
    Left = 264
    Top = 120
  end
  object ActionList: TActionList
    Left = 348
    Top = 192
  end
  object AgQuery: TQuery
    DatabaseName = 'Priv2005Database'
    SQL.Strings = (
      
        'SELECT * FROM  AGENT WHERE LEBEN2PCNT IS NOT NULL AND LEBEN2PCNT' +
        '>=0'
      'ORDER BY NAME')
    Left = 96
    Top = 156
    object AgQueryAgent_code: TStringField
      FieldName = 'Agent_code'
      Origin = 'BELGREENDATABASE."AGENT.DB".Agent_code'
      Size = 4
    end
    object AgQueryName: TStringField
      FieldName = 'Name'
      Origin = 'BELGREENDATABASE."AGENT.DB".Name'
      Size = 60
    end
  end
  object AllAgents: TTable
    DatabaseName = 'Priv2005Database'
    TableName = 'AGENT.DB'
    Left = 464
    Top = 216
  end
  object WorkSQL: TQuery
    DatabaseName = 'Priv2005Database'
    Left = 88
    Top = 200
  end
  object LocateTbl: TTable
    DatabaseName = 'BelGreenDatabase'
    TableName = 'BELGREEN.DB'
    Left = 208
    Top = 300
  end
  object REPDataSource: TDataSource
    DataSet = REPQuery
    Left = 168
    Top = 124
  end
  object REPQuery: TRxQuery
    DatabaseName = 'Priv2005Database'
    Macros = <>
    Left = 16
    Top = 160
  end
  object PopupMenu: TPopupMenu
    Left = 168
    Top = 92
    object MenuItem1: TMenuItem
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1090#1077#1082#1091#1097#1077#1077' '#1089#1086#1089#1090#1086#1103#1085#1080#1077' '#1082#1072#1082'...'
      OnClick = MenuItem1Click
    end
    object DelMenu: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1080#1079' '#1085#1072#1073#1086#1088#1072'...'
      OnClick = DelMenuClick
    end
  end
  object OpenDialogPx: TOpenDialog
    DefaultExt = 'db'
    Filter = 'Paradox '#1092#1072#1081#1083#1099'|*.db'
    Title = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093' '#1040#1083#1100#1074#1077#1085#1099
    Left = 444
    Top = 168
  end
  object FormStorage: TFormStorage
    IniFileName = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077
    IniSection = #1046#1080#1079#1085#1100' 2005'
    Options = []
    StoredProps.Strings = (
      'NmbFrom.Text'
      'NmbTo.Text'
      'RepDtFrom.Date'
      'RepDtFrom.Checked'
      'RepDtTo.Date'
      'RepDtTo.Checked'
      'UpdateDtFrom.Date'
      'UpdateDtFrom.Checked'
      'UpdateDtTo.Date'
      'UpdateDtTo.Checked')
    StoredValues = <>
    Left = 300
    Top = 264
  end
  object WorkSQL2: TQuery
    DatabaseName = 'Priv2005Database'
    Left = 148
    Top = 200
  end
  object SaveDialog: TSaveDialog
    DefaultExt = 'unload'
    Filter = #1042#1089#1077' '#1092#1072#1081#1083#1099' (*.*)|*.*'
    InitialDir = 'C:\'
    Title = #1042#1099#1075#1088#1091#1079#1082#1072' '#1076#1072#1085#1085#1099#1093' '#1074' '#1044#1080#1088#1077#1082#1094#1080#1102
    Left = 44
    Top = 332
  end
  object SaveDialog1C: TSaveDialog
    DefaultExt = 'txt'
    Filter = #1060#1072#1081#1083#1099' '#1041#1072#1079#1099' '#1076#1072#1085#1085#1099#1093' (*.txt)|*.txt'
    InitialDir = 'C:\'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Title = #1042#1099#1075#1088#1091#1079#1082#1072' '#1076#1072#1085#1085#1099#1093' '#1074' '#1044#1080#1088#1077#1082#1094#1080#1102
    Left = 44
    Top = 388
  end
  object Table1C: TTable
    TableType = ttDBase
    Left = 120
    Top = 392
  end
end
