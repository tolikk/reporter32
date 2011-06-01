object BelGreenForm: TBelGreenForm
  Left = 244
  Top = 129
  Width = 910
  Height = 697
  Caption = #1041#1077#1083#1086#1088#1091#1089#1089#1082#1072#1103' '#1047#1077#1083#1105#1085#1072#1103' '#1082#1072#1088#1090#1072
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
    Top = 606
    Width = 902
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
      Left = 638
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
    Width = 902
    Height = 606
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
        Width = 630
        Height = 578
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
            FieldName = 'Paydate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Repdate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Autonmb'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Model'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Owner'
            Width = 100
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'OwnType'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Dtfrom'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Dtto'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Letter'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Tarif'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay1'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay1c'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay2'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay2c'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StateText'
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AgName'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Text'
            Visible = True
          end>
      end
      object Splitter: TRxSplitter
        Left = 630
        Top = 0
        Width = 5
        Height = 578
        ControlFirst = MainGrid
        ControlSecond = UbPanel
        Align = alRight
        BevelOuter = bvNone
        Color = clMoneyGreen
        Visible = False
      end
      object UbPanel: TPanel
        Left = 635
        Top = 0
        Width = 259
        Height = 578
        Align = alRight
        BevelOuter = bvNone
        TabOrder = 2
        Visible = False
        object Label10: TLabel
          Left = 0
          Top = 0
          Width = 39
          Height = 13
          Align = alTop
          Caption = #1059#1073#1099#1090#1082#1080
        end
        object RxDBGrid2: TRxDBGrid
          Left = 0
          Top = 13
          Width = 259
          Height = 536
          Align = alClient
          Color = 13697023
          DataSource = dsUb
          ReadOnly = True
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'MS Sans Serif'
          TitleFont.Style = []
        end
        object Panel4: TPanel
          Left = 0
          Top = 549
          Width = 259
          Height = 29
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 1
          object DelBtn: TButton
            Left = 124
            Top = 4
            Width = 59
            Height = 21
            Caption = #1059#1076#1072#1083#1080#1090#1100
            Font.Charset = RUSSIAN_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 0
            OnClick = DelBtnClick
          end
          object AddBtn: TButton
            Left = 0
            Top = 4
            Width = 63
            Height = 21
            Caption = #1044#1086#1073#1072#1074#1080#1090#1100
            Enabled = False
            Font.Charset = RUSSIAN_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
            OnClick = AddBtnClick
          end
          object EditBtn: TButton
            Left = 64
            Top = 4
            Width = 61
            Height = 21
            Caption = #1048#1079#1084#1077#1085#1080#1090#1100
            Font.Charset = RUSSIAN_CHARSET
            Font.Color = clWindowText
            Font.Height = -11
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 2
            OnClick = EditBtnClick
          end
        end
      end
    end
    object FilterPage: TTabSheet
      Caption = #1060#1080#1083#1100#1090#1088
      ImageIndex = 1
      object Panel5: TPanel
        Left = 0
        Top = 0
        Width = 894
        Height = 578
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
          Top = 84
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
          Top = 100
          Width = 23
          Height = 21
          Caption = '!'
          Flat = True
          OnClick = SpeedButtonClick
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
          Top = 64
          Width = 69
          Height = 17
          Caption = #1040#1075#1077#1085#1090#1099
          TabOrder = 4
        end
        object ListAgents: TCheckListBox
          Left = 136
          Top = 56
          Width = 269
          Height = 221
          ItemHeight = 13
          TabOrder = 5
          OnClick = ListAgentsClick
        end
        object IsBreak: TCheckBox
          Left = 408
          Top = 32
          Width = 97
          Height = 17
          Caption = #1056#1072#1089#1090#1086#1088#1075#1085#1091#1090#1099#1077
          TabOrder = 6
        end
        object ListTemplates: TComboBox
          Left = 8
          Top = 100
          Width = 101
          Height = 21
          Style = csDropDownList
          ItemHeight = 13
          TabOrder = 7
          OnChange = ListTemplatesChange
          Items.Strings = (
            #1042#1089#1077' '#1072#1075#1077#1085#1090#1099)
        end
        object DownPanel: TPanel
          Left = 0
          Top = 445
          Width = 894
          Height = 133
          Align = alBottom
          BevelOuter = bvNone
          TabOrder = 8
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
          object Label14: TLabel
            Left = 264
            Top = 80
            Width = 16
            Height = 13
            Caption = #1055#1054
          end
          object Label16: TLabel
            Left = 8
            Top = 80
            Width = 105
            Height = 13
            Caption = #1044#1072#1090#1072' '#1074#1099#1076#1072#1095#1080' '#1087#1086#1083#1080#1089#1072
          end
          object Label17: TLabel
            Left = 408
            Top = 80
            Width = 196
            Height = 13
            Caption = #1076#1083#1103' '#1077#1078#1077#1084#1077#1089#1103#1095#1085#1086#1075#1086' '#1086#1090#1095#1077#1090#1072' '#1074' '#1044#1080#1088#1077#1082#1094#1080#1102
          end
          object Apply: TButton
            Left = 300
            Top = 101
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
              #1059#1058#1045#1056#1071#1053
              #1048#1057#1055#1054#1056#1063#1045#1053
              #1059#1058#1045#1056#1071#1053' '#1040#1043#1045#1053#1058#1054#1052
              #1056#1040#1057#1058#1054#1056#1043#1053#1059#1058
              #1044#1059#1041#1051#1048#1050#1040#1058
              #1042#1047#1040#1052#1045#1053)
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
              #1070#1088' '#1083#1080#1094#1086
              #1048#1055)
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
          object dtRegTo: TDateTimePicker
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
            TabOrder = 7
          end
          object dtRegFrom: TDateTimePicker
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
            TabOrder = 8
          end
        end
      end
    end
    object StatPage: TTabSheet
      Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
      ImageIndex = 2
      object StatisticTxt: TMemo
        Left = 0
        Top = 0
        Width = 894
        Height = 557
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
      object Panel1: TPanel
        Left = 0
        Top = 557
        Width = 894
        Height = 29
        Align = alBottom
        BevelOuter = bvNone
        TabOrder = 1
        object btnRptMonth: TButton
          Left = 4
          Top = 3
          Width = 105
          Height = 25
          Caption = #1054#1090#1095#1105#1090' '#1079#1072' '#1084#1077#1089#1103#1094
          TabOrder = 0
          OnClick = btnRptMonthClick
        end
        object btnRptWeek: TButton
          Left = 112
          Top = 3
          Width = 103
          Height = 25
          Caption = #1054#1090#1095#1105#1090' '#1079#1072' '#1085#1077#1076#1077#1083#1102
          Enabled = False
          TabOrder = 1
        end
      end
    end
    object Reports: TTabSheet
      Caption = #1054#1090#1095#1105#1090#1099
      ImageIndex = 3
      object Panel3: TPanel
        Left = 0
        Top = 0
        Width = 894
        Height = 578
        Align = alClient
        BevelOuter = bvNone
        Caption = 'Panel3'
        TabOrder = 0
        DesignSize = (
          894
          578)
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
          Width = 885
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
          Width = 784
          Height = 21
          Anchors = [akLeft, akTop, akRight]
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
        object StartButton: TButton
          Left = 799
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
          Width = 885
          Height = 437
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
          Left = 802
          Top = 543
          Width = 87
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1042' Excel'
          TabOrder = 4
          OnClick = Button2Click
        end
        object IsNumbers: TCheckBox
          Left = 703
          Top = 547
          Width = 97
          Height = 17
          Anchors = [akRight, akBottom]
          Caption = #1053#1091#1084#1077#1088#1086#1074#1072#1090#1100
          TabOrder = 5
        end
        object IsUseFilter: TCheckBox
          Left = 16
          Top = 549
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
    Left = 672
    Top = 32
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
      object Reserv: TMenuItem
        Caption = #1057#1090#1088#1072#1093#1086#1074#1099#1077' '#1088#1077#1079#1077#1088#1074#1099'...'
        OnClick = ReservClick
      end
      object N9: TMenuItem
        Tag = 1
        Caption = #1056#1053#1055#1059'...'
        OnClick = ReservClick
      end
      object N111: TMenuItem
        Caption = #1046#1091#1088#1085#1072#1083' 1.1...'
        OnClick = N111Click
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
      object ImpAlvena: TMenuItem
        Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093' '#1040#1083#1100#1074#1077#1085#1099'...'
        OnClick = ImpAlvenaClick
      end
      object N15: TMenuItem
        Caption = '-'
      end
      object CurrencyRateSrc: TMenuItem
        Caption = #1048#1089#1090#1086#1095#1085#1080#1082' '#1082#1091#1088#1089#1086#1074' '#1074#1072#1083#1102#1090'...'
        OnClick = CurrencyRateSrcClick
      end
      object N5: TMenuItem
        Caption = '-'
      end
      object N2: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N2Click
      end
    end
    object N6: TMenuItem
      Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
      object N8: TMenuItem
        Caption = #1040#1074#1090#1086#1088#1099'...'
        OnClick = N8Click
      end
    end
    object N12: TMenuItem
      Caption = '|'
    end
    object Statictic: TMenuItem
      Caption = ' '
      OnClick = StaticticClick
    end
    object N13: TMenuItem
      Caption = '|'
    end
    object CompanyTitle: TMenuItem
      Caption = '.'
      Enabled = False
    end
    object DivisionInfo: TMenuItem
      Caption = '.'
      Enabled = False
    end
  end
  object MainQuery: TRxQuery
    AfterScroll = MainQueryAfterScroll
    OnCalcFields = MainQueryCalcFields
    DatabaseName = 'BelGreenDatabase'
    SQL.Strings = (
      'SELECT * FROM BELGREEN'
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
    Left = 264
    Top = 72
    object MainQuerySeria: TStringField
      DisplayLabel = #1057#1077#1088#1080#1103
      FieldName = 'Seria'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Seria'
      Size = 7
    end
    object MainQueryNumber: TFloatField
      DisplayLabel = #1053#1086#1084#1077#1088
      FieldName = 'Number'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Number'
    end
    object MainQueryPseria: TStringField
      FieldName = 'Pseria'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pseria'
      Visible = False
      Size = 7
    end
    object MainQueryPnumber: TFloatField
      FieldName = 'Pnumber'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pnumber'
      Visible = False
    end
    object MainQueryAutonmb: TStringField
      DisplayLabel = #1053#1086#1084#1077#1088' '#1072#1074#1090#1086
      FieldName = 'Autonmb'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Autonmb'
      Size = 24
    end
    object MainQueryModel: TStringField
      DisplayLabel = #1052#1086#1076#1077#1083#1100
      FieldName = 'Model'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Model'
      Size = 24
    end
    object MainQueryOwner: TStringField
      DisplayLabel = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100
      FieldName = 'Owner'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Owner'
      Size = 48
    end
    object MainQueryOwnType: TStringField
      DisplayLabel = #1058#1080#1087' '#1089#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1103
      FieldName = 'OwnType'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".OwnType'
      Visible = False
      Size = 1
    end
    object MainQueryOrg: TStringField
      FieldName = 'Org'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Org'
      Visible = False
      Size = 48
    end
    object MainQueryAddr: TStringField
      FieldName = 'Addr'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Addr'
      Visible = False
      Size = 48
    end
    object MainQueryDtfrom: TDateField
      DisplayLabel = #1057
      FieldName = 'Dtfrom'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Dtfrom'
    end
    object MainQueryPeriod: TStringField
      DisplayLabel = #1055#1077#1088#1080#1086#1076
      FieldName = 'Period'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Period'
      Size = 15
    end
    object MainQueryDtto: TDateField
      DisplayLabel = #1055#1054
      FieldName = 'Dtto'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Dtto'
    end
    object MainQueryLetter: TStringField
      DisplayLabel = #1058#1080#1087
      FieldName = 'Letter'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Letter'
      Size = 4
    end
    object MainQueryPlace: TStringField
      FieldName = 'Place'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Place'
      Visible = False
      Size = 32
    end
    object MainQueryRegdate: TDateField
      FieldName = 'Regdate'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Regdate'
      Visible = False
    end
    object MainQueryRegtime: TStringField
      FieldName = 'Regtime'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Regtime'
      Visible = False
      Size = 5
    end
    object MainQueryTarif: TFloatField
      DisplayLabel = #1058#1072#1088#1080#1092
      FieldName = 'Tarif'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Tarif'
    end
    object MainQueryPay1: TFloatField
      DisplayLabel = #1054#1087#1083#1072#1090#1072' 1'
      FieldName = 'Pay1'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pay1'
      DisplayFormat = '0.##;#;#'
    end
    object MainQueryPay1c: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1086#1087#1083#1072#1090#1099' 1'
      FieldName = 'Pay1c'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pay1c'
      Size = 3
    end
    object MainQueryPay2: TFloatField
      DisplayLabel = #1054#1087#1083#1072#1090#1072' 2'
      FieldName = 'Pay2'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pay2'
      DisplayFormat = '0.##;#;#'
    end
    object MainQueryPay2c: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1086#1087#1083#1072#1090#1099' 2'
      FieldName = 'Pay2c'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pay2c'
      Size = 3
    end
    object MainQueryPaydate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1087#1083#1072#1090#1099
      FieldName = 'Paydate'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Paydate'
    end
    object MainQueryAgcode: TStringField
      DisplayLabel = #1040#1075#1077#1085#1090
      FieldName = 'Agcode'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Agcode'
      Visible = False
      Size = 4
    end
    object MainQueryAgType: TStringField
      FieldName = 'AgType'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".AgType'
      Visible = False
      Size = 1
    end
    object MainQueryPcnt: TFloatField
      FieldName = 'Pcnt'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Pcnt'
      Visible = False
    end
    object MainQueryState: TStringField
      DisplayLabel = #1057#1086#1089#1090#1086#1103#1085#1080#1077
      FieldName = 'State'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".State'
      Visible = False
      Size = 1
    end
    object MainQueryType: TStringField
      FieldName = 'Type'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Type'
      Visible = False
      Size = 1
    end
    object MainQueryRepdate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
      FieldName = 'Repdate'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Repdate'
    end
    object MainQueryCountry: TFloatField
      FieldName = 'Country'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Country'
      Visible = False
    end
    object MainQueryText: TStringField
      DisplayLabel = #1055#1088#1080#1094#1077#1087' '#1082' '#1043#1040
      FieldName = 'Text'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".Text'
      Size = 7
    end
    object MainQueryStateText: TStringField
      DisplayLabel = #1057#1086#1089#1090#1086#1103#1085#1080#1077' '
      FieldKind = fkCalculated
      FieldName = 'StateText'
      Visible = False
      Size = 32
      Calculated = True
    end
    object MainQueryAgName: TStringField
      DisplayLabel = #1040#1075#1077#1085#1090' '
      FieldKind = fkLookup
      FieldName = 'AgName'
      LookupDataSet = AllAgents
      LookupKeyFields = 'Agent_code'
      LookupResultField = 'Name'
      KeyFields = 'Agcode'
      Size = 32
      Lookup = True
    end
    object MainQueryDUPTYPE: TStringField
      FieldName = 'DUPTYPE'
      Origin = 'BELGREENDATABASE."BELGREEN.DB".DUPTYPE'
      Size = 1
    end
  end
  object BelGreenDatabase: TDatabase
    AliasName = 'BASO'
    DatabaseName = 'BelGreenDatabase'
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
    DatabaseName = 'BelGreenDatabase'
    SQL.Strings = (
      'SELECT * FROM  AGENT WHERE BELGREEN IS NOT NULL AND BELGREEN>=0'
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
    DatabaseName = 'BelGreenDatabase'
    TableName = 'AGENT.DB'
    Left = 464
    Top = 216
  end
  object WorkSQL: TQuery
    DatabaseName = 'BelGreenDatabase'
    Left = 88
    Top = 200
  end
  object BUROTbl: TTable
    Exclusive = True
    Left = 516
    Top = 256
  end
  object BUROTbl2: TTable
    Exclusive = True
    Left = 524
    Top = 296
  end
  object LocateTbl: TTable
    DatabaseName = 'BelGreenDatabase'
    TableName = 'BELGREEN.DB'
    Left = 584
    Top = 296
  end
  object NExportTbl: TTable
    DatabaseName = 'BelGreenDatabase'
    TableName = 'BELGREENEXP.DB'
    Left = 32
    Top = 200
  end
  object REPDataSource: TDataSource
    DataSet = REPQuery
    Left = 168
    Top = 124
  end
  object REPQuery: TRxQuery
    DatabaseName = 'BelGreenDatabase'
    Macros = <>
    Left = 16
    Top = 160
  end
  object UbTable_: TTable
    DatabaseName = 'BelGreenDatabase'
    Filtered = True
    TableName = 'bgrnub.DB'
    Left = 473
    Top = 80
    object UbTable_Ser: TStringField
      FieldName = 'Ser'
      Visible = False
      Size = 5
    end
    object UbTable_Nmb: TFloatField
      FieldName = 'Nmb'
      Visible = False
    end
    object UbTable_RegDate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
      FieldName = 'RegDate'
    end
    object UbTable_RegNmb: TStringField
      DisplayLabel = #1056#1077#1075'. '#1053#1086#1084#1077#1088
      FieldName = 'RegNmb'
      Size = 10
    end
    object UbTable_ComplName: TStringField
      FieldName = 'ComplName'
      Visible = False
      Size = 16
    end
    object UbTable_ComplNmb: TStringField
      FieldName = 'ComplNmb'
      Visible = False
      Size = 10
    end
    object UbTable_ComplDate: TDateField
      FieldName = 'ComplDate'
      Visible = False
    end
    object UbTable_UbSum: TFloatField
      DisplayLabel = #1057#1091#1084#1084#1072
      FieldName = 'UbSum'
    end
    object UbTable_UbCurr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072
      FieldName = 'UbCurr'
      Size = 3
    end
    object UbTable_PayDate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1087#1083#1072#1090#1099
      FieldName = 'PayDate'
    end
    object UbTable_PaySum: TFloatField
      DisplayLabel = #1057#1091#1084#1084#1072' '#1086#1087#1083#1072#1090#1099
      FieldName = 'PaySum'
    end
    object UbTable_PayCurr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1086#1087#1083#1072#1090#1099
      FieldName = 'PayCurr'
      Size = 3
    end
    object UbTable_FreeDate: TDateField
      FieldName = 'FreeDate'
      Visible = False
    end
    object UbTable_FreeSum: TFloatField
      FieldName = 'FreeSum'
      Visible = False
    end
    object UbTable_FreeCurr: TStringField
      FieldName = 'FreeCurr'
      Visible = False
      Size = 3
    end
    object UbTable_ClDate: TDateField
      FieldName = 'ClDate'
      Visible = False
    end
    object UbTable_Text: TStringField
      DisplayLabel = #1050#1086#1084#1084#1077#1085#1090#1072#1088#1080#1080
      FieldName = 'Text'
      Size = 48
    end
    object UbTable_BrSum: TFloatField
      FieldName = 'BrSum'
      Visible = False
    end
    object UbTable_BrCurr: TStringField
      FieldName = 'BrCurr'
      Visible = False
      Size = 3
    end
    object UbTable_Country: TStringField
      DisplayLabel = #1057#1090#1088#1072#1085#1072
      FieldName = 'Country'
      Size = 24
    end
    object UbTable_State: TFloatField
      FieldName = 'State'
      Visible = False
    end
  end
  object dsUb: TDataSource
    DataSet = UbTable_
    Left = 549
    Top = 80
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
    Filter = #1060#1072#1081#1083#1099' '#1040#1083#1100#1074#1077#1085#1099' '#1041#1077#1083#1047#1050'|green*.db|'#1060#1072#1081#1083#1099' Paradox|*.db'
    Title = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093' '#1040#1083#1100#1074#1077#1085#1099
    Left = 444
    Top = 168
  end
  object FormStorage: TFormStorage
    IniFileName = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077
    IniSection = #1041#1077#1083#1047#1050
    Options = []
    StoredProps.Strings = (
      'NmbFrom.Text'
      'NmbTo.Text'
      'RepDtFrom.Date'
      'RepDtFrom.Checked'
      'RepDtTo.Date'
      'RepDtTo.Checked')
    StoredValues = <>
    Left = 512
    Top = 212
  end
  object WorkSQL2: TQuery
    DatabaseName = 'BelGreenDatabase'
    Left = 88
    Top = 268
  end
  object SaveDialog: TSaveDialog
    DefaultExt = 'unload'
    Filter = #1042#1089#1077' '#1092#1072#1081#1083#1099' (*.*)|*.*'
    InitialDir = 'C:\'
    Title = #1042#1099#1075#1088#1091#1079#1082#1072' '#1076#1072#1085#1085#1099#1093' '#1074' '#1044#1080#1088#1077#1082#1094#1080#1102
    Left = 36
    Top = 408
  end
end
