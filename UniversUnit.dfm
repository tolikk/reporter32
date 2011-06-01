object UNIVERS: TUNIVERS
  Left = 396
  Top = 205
  Width = 584
  Height = 611
  Caption = #1059#1085#1080#1074#1077#1088#1089#1072#1083#1100#1085#1099#1081' '#1087#1086#1083#1080#1089
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Label5: TLabel
    Left = 188
    Top = 381
    Width = 7
    Height = 13
    Caption = #1057
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label6: TLabel
    Left = 308
    Top = 381
    Width = 16
    Height = 13
    Caption = #1055#1054
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Panel2: TPanel
    Left = 0
    Top = 524
    Width = 576
    Height = 33
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      576
      33)
    object CountLabel: TLabel
      Left = 2
      Top = 11
      Width = 183
      Height = 13
      AutoSize = False
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Panel3: TPanel
      Left = 296
      Top = 0
      Width = 280
      Height = 33
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
      object btnClose: TButton
        Left = 200
        Top = 5
        Width = 75
        Height = 25
        Caption = #1047#1072#1082#1088#1099#1090#1100
        TabOrder = 0
        OnClick = btnCloseClick
      end
    end
    object ProgressBar: TProgressBar
      Left = 192
      Top = 9
      Width = 299
      Height = 16
      Anchors = [akLeft, akTop, akRight]
      Min = 0
      Max = 100
      TabOrder = 1
    end
  end
  object TabbedNotebook: TPageControl
    Left = 0
    Top = 0
    Width = 576
    Height = 524
    ActivePage = RepTab
    Align = alClient
    TabIndex = 0
    TabOrder = 1
    OnChange = TabbedNotebookChange
    object RepTab: TTabSheet
      Caption = #1055#1086#1083#1080#1089#1099
      object MainGrid: TRxDBGrid
        Left = 0
        Top = 0
        Width = 568
        Height = 465
        Align = alClient
        Color = 15269887
        DataSource = DataSource
        TabOrder = 0
        TitleFont.Charset = RUSSIAN_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnGetCellProps = MainGridGetCellProps
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
            FieldName = 'Vid'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PrevS'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PrevN'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FU'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RegDate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StartDate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Another'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'From'
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
            FieldName = 'Insurer'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsAddr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsOther'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsPcnt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Zastr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Vigod'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'VigodAddr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'LizoZast'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Terr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PlaceIns'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Variant'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RealSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RealCurr'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsCur'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Wait'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PremSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PremCur'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Feer'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FeeSum'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FeeCur'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FeeDate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FeeDoc'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'FeeTyp'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'SrokPay'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Agent'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AgPcnt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AgType'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'State'
            Visible = True
          end>
      end
      object Panel4: TPanel
        Left = 0
        Top = 465
        Width = 568
        Height = 31
        Align = alBottom
        BevelOuter = bvNone
        TabOrder = 1
        DesignSize = (
          568
          31)
        object btnSort: TButton
          Left = 408
          Top = 4
          Width = 85
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1057#1086#1088#1090#1080#1088#1086#1074#1082#1072'...'
          TabOrder = 0
          OnClick = btnSortClick
        end
        object btnExcel: TButton
          Left = 492
          Top = 4
          Width = 75
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1074' Excel'
          TabOrder = 1
          OnClick = btnExcelClick
        end
      end
    end
    object FilterTab: TTabSheet
      Caption = #1060#1080#1083#1100#1090#1088
      ImageIndex = 1
      object Label11: TLabel
        Left = 8
        Top = 8
        Width = 129
        Height = 57
        AutoSize = False
        Caption = #1040#1075#1077#1085#1090#1099#13#10'('#1077#1089#1083#1080' '#1085#1080#1095#1077#1075#1086' '#1085#1077' '#1087#1086#1084#1077#1095#1077#1085#1086', '#1090#1086' '#1087#1086#1082#1072#1079#1072#1090#1100' '#1074#1089#1077#1093')'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        WordWrap = True
      end
      object Label1: TLabel
        Left = 184
        Top = 189
        Width = 7
        Height = 13
        Caption = #1057
      end
      object Label2: TLabel
        Left = 304
        Top = 189
        Width = 15
        Height = 13
        Caption = #1055#1054
      end
      object Label3: TLabel
        Left = 184
        Top = 213
        Width = 7
        Height = 13
        Caption = #1057
      end
      object Label4: TLabel
        Left = 304
        Top = 213
        Width = 15
        Height = 13
        Caption = #1055#1054
      end
      object Label7: TLabel
        Left = 184
        Top = 237
        Width = 7
        Height = 13
        Caption = #1057
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object Label10: TLabel
        Left = 304
        Top = 237
        Width = 16
        Height = 13
        Caption = #1055#1054
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object SpeedButton: TSpeedButton
        Left = 156
        Top = 88
        Width = 23
        Height = 21
        Caption = '!'
        Flat = True
        OnClick = SpeedButtonClick
      end
      object CntChecked: TLabel
        Left = 8
        Top = 72
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
      object ListTemplates: TComboBox
        Left = 8
        Top = 88
        Width = 145
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 0
        OnChange = ListTemplatesChange
        Items.Strings = (
          #1042#1089#1077' '#1072#1075#1077#1085#1090#1099)
      end
      object AgentList: TCheckListBox
        Left = 200
        Top = 4
        Width = 361
        Height = 177
        Flat = False
        ItemHeight = 13
        TabOrder = 1
        OnClick = AgentListClick
      end
      object IsNumber: TCheckBox
        Left = 8
        Top = 187
        Width = 105
        Height = 17
        Caption = #1053#1086#1084#1077#1088
        TabOrder = 2
      end
      object NumberFrom: TEdit
        Left = 200
        Top = 185
        Width = 97
        Height = 21
        TabOrder = 3
      end
      object NumberTo: TEdit
        Left = 328
        Top = 185
        Width = 233
        Height = 21
        TabOrder = 4
      end
      object IsRegDate: TCheckBox
        Left = 8
        Top = 211
        Width = 121
        Height = 17
        Caption = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
        TabOrder = 5
      end
      object RegDateFrom: TDateTimePicker
        Left = 200
        Top = 209
        Width = 97
        Height = 21
        CalAlignment = dtaLeft
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Kind = dtkDate
        ParseInput = False
        TabOrder = 6
      end
      object RegDateTo: TDateTimePicker
        Left = 328
        Top = 209
        Width = 233
        Height = 21
        CalAlignment = dtaLeft
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Kind = dtkDate
        ParseInput = False
        TabOrder = 7
      end
      object IsPayDate: TCheckBox
        Left = 8
        Top = 235
        Width = 121
        Height = 17
        Caption = #1044#1072#1090#1072' '#1087#1083#1072#1090#1077#1078#1072
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 8
      end
      object PayDateFrom: TDateTimePicker
        Left = 200
        Top = 233
        Width = 97
        Height = 21
        CalAlignment = dtaLeft
        CalColors.TextColor = 10485760
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 9
      end
      object PayDateTo: TDateTimePicker
        Left = 328
        Top = 233
        Width = 233
        Height = 21
        CalAlignment = dtaLeft
        CalColors.TextColor = 10485760
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 10
      end
      object IsState: TCheckBox
        Left = 8
        Top = 309
        Width = 161
        Height = 17
        Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
      end
      object State: TComboBox
        Left = 200
        Top = 307
        Width = 361
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 12
        Items.Strings = (
          #1053#1054#1056#1052#1040#1051#1068#1053#1067#1049
          #1059#1058#1045#1056#1071#1053
          #1048#1057#1055#1054#1056#1063#1045#1053
          #1056#1040#1057#1058#1054#1056#1043#1053#1059#1058
          #1044#1059#1041#1051#1048#1050#1040#1058)
      end
      object IsRepDate: TCheckBox
        Left = 8
        Top = 259
        Width = 121
        Height = 17
        Caption = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
      end
      object DateRepStart: TDateTimePicker
        Left = 200
        Top = 257
        Width = 97
        Height = 21
        CalAlignment = dtaLeft
        CalColors.TextColor = 10485760
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 14
      end
      object DateRepEnd: TDateTimePicker
        Left = 328
        Top = 257
        Width = 233
        Height = 21
        CalAlignment = dtaLeft
        CalColors.TextColor = 10485760
        Date = 36805.6934977778
        Time = 36805.6934977778
        ShowCheckbox = True
        Checked = False
        DateFormat = dfShort
        DateMode = dmComboBox
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Kind = dtkDate
        ParseInput = False
        ParentFont = False
        TabOrder = 15
      end
      object btnApply: TButton
        Left = 486
        Top = 357
        Width = 75
        Height = 25
        Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
        TabOrder = 16
        OnClick = btnApplyClick
      end
      object listOwnType: TComboBox
        Left = 432
        Top = 283
        Width = 129
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 17
        Items.Strings = (
          #1053#1077' '#1074#1072#1078#1085#1086
          #1060#1080#1079' '#1083#1080#1094#1086
          #1070#1088' '#1083#1080#1094#1086)
      end
      object IsStopDate: TCheckBox
        Left = 97
        Top = 259
        Width = 97
        Height = 17
        Caption = '+ '#1056#1072#1089#1090#1086#1088#1078#1077#1085#1080#1077
        TabOrder = 18
      end
      object IsOwner: TCheckBox
        Left = 8
        Top = 283
        Width = 97
        Height = 17
        Caption = #1042#1083#1072#1076#1077#1083#1077#1094
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 19
      end
      object OwnerLike: TEdit
        Left = 200
        Top = 283
        Width = 229
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 20
      end
      object IsInsurType: TCheckBox
        Left = 8
        Top = 333
        Width = 161
        Height = 17
        Caption = #1042#1080#1076' '#1089#1090#1088#1072#1093#1086#1074#1072#1085#1080#1103
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 21
      end
      object InsurTypes: TComboBox
        Left = 200
        Top = 331
        Width = 361
        Height = 21
        Style = csDropDownList
        ItemHeight = 0
        TabOrder = 22
      end
    end
    object StatPage: TTabSheet
      Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
      ImageIndex = 2
      object StatisticTxt: TMemo
        Left = 0
        Top = 0
        Width = 568
        Height = 496
        Align = alClient
        Color = clBlack
        Font.Charset = RUSSIAN_CHARSET
        Font.Color = clLime
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Lines.Strings = (
          #1047#1076#1077#1089#1100' '#1073#1091#1076#1077#1090' '#1089#1090#1072#1090#1080#1089#1090#1080#1082#1072)
        ParentFont = False
        TabOrder = 0
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1054#1090#1095#1105#1090#1099
      ImageIndex = 3
      DesignSize = (
        568
        496)
      object Label8: TLabel
        Left = 4
        Top = 4
        Width = 32
        Height = 13
        Caption = #1054#1090#1095#1105#1090
      end
      object Label9: TLabel
        Left = 4
        Top = 44
        Width = 339
        Height = 13
        Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' ('#1077#1089#1083#1080' '#1085#1077#1089#1082#1086#1083#1100#1082#1086', '#1090#1086' '#1087#1077#1088#1077#1095#1080#1089#1083#1080' '#1095#1077#1088#1077#1079' '#1090#1086#1095#1082#1091' '#1089' '#1079#1072#1087#1103#1090#1086#1081')'
      end
      object Label12: TLabel
        Left = 4
        Top = 84
        Width = 53
        Height = 13
        Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090
      end
      object RepList: TComboBox
        Left = 4
        Top = 20
        Width = 561
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
        Width = 467
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
      end
      object StartButton: TButton
        Left = 476
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
        Width = 559
        Height = 373
        Anchors = [akLeft, akTop, akRight, akBottom]
        DataSource = REPDS
        TabOrder = 3
        TitleFont.Charset = RUSSIAN_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object IsNumbers: TCheckBox
        Left = 376
        Top = 477
        Width = 97
        Height = 17
        Anchors = [akRight, akBottom]
        Caption = #1053#1091#1084#1077#1088#1086#1074#1072#1090#1100
        TabOrder = 4
      end
      object Button2: TButton
        Left = 475
        Top = 475
        Width = 87
        Height = 21
        Anchors = [akRight, akBottom]
        Caption = #1042' Excel'
        TabOrder = 5
        OnClick = Button2Click
      end
      object IsUseFilter: TCheckBox
        Left = 4
        Top = 475
        Width = 133
        Height = 17
        Anchors = [akLeft, akBottom]
        Caption = #1042' '#1090#1077#1082#1091#1097#1077#1081' '#1074#1099#1073#1086#1088#1082#1077
        TabOrder = 6
      end
    end
  end
  object MainQuery: TRxQuery
    DatabaseName = 'UniversDb'
    SQL.Strings = (
      'SELECT'
      '  *'
      'FROM'
      '  UNIVERS'
      '%WHERE'
      '%ORDER')
    Macros = <
      item
        DataType = ftString
        Name = 'WHERE'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'ORDER'
        ParamType = ptInput
      end>
    Left = 388
    Top = 80
    object MainQuerySeria: TStringField
      DisplayLabel = #1057#1077#1088#1080#1103
      FieldName = 'Seria'
      Origin = 'UNIVERSDB."Univers.DB".Seria'
      Size = 3
    end
    object MainQueryNumber: TFloatField
      DisplayLabel = #1053#1086#1084#1077#1088
      FieldName = 'Number'
      Origin = 'UNIVERSDB."Univers.DB".Number'
    end
    object MainQueryVid: TStringField
      DisplayLabel = #1042#1080#1076
      FieldName = 'Vid'
      Origin = 'UNIVERSDB."Univers.DB".Vid'
      Size = 4
    end
    object MainQueryPrevS: TStringField
      FieldName = 'PrevS'
      Origin = 'UNIVERSDB."Univers.DB".PrevS'
      Size = 3
    end
    object MainQueryPrevN: TFloatField
      FieldName = 'PrevN'
      Origin = 'UNIVERSDB."Univers.DB".PrevN'
    end
    object MainQueryFU: TStringField
      FieldName = 'FU'
      Origin = 'UNIVERSDB."Univers.DB".FU'
      Size = 1
    end
    object MainQueryRegDate: TDateField
      DisplayLabel = #1042#1099#1076#1072#1085
      FieldName = 'RegDate'
      Origin = 'UNIVERSDB."Univers.DB".RegDate'
    end
    object MainQueryStartDate: TDateField
      FieldName = 'StartDate'
      Origin = 'UNIVERSDB."Univers.DB".StartDate'
    end
    object MainQueryAnother: TStringField
      FieldName = 'Another'
      Origin = 'UNIVERSDB."Univers.DB".Another'
    end
    object MainQueryFrom: TDateField
      FieldName = 'From'
      Origin = 'UNIVERSDB."Univers.DB".From'
    end
    object MainQueryTo: TDateField
      FieldName = 'To'
      Origin = 'UNIVERSDB."Univers.DB".To'
    end
    object MainQueryPeriod: TStringField
      DisplayLabel = #1055#1077#1088#1080#1086#1076
      FieldName = 'Period'
      Origin = 'UNIVERSDB."Univers.DB".Period'
      Size = 10
    end
    object MainQueryInsurer: TStringField
      DisplayLabel = #1057#1090#1088'-'#1083#1100
      FieldName = 'Insurer'
      Origin = 'UNIVERSDB."Univers.DB".Insurer'
      Size = 60
    end
    object MainQueryInsAddr: TStringField
      DisplayLabel = #1040#1076#1088#1077#1089' '#1089#1090'-'#1083#1103
      FieldName = 'InsAddr'
      Origin = 'UNIVERSDB."Univers.DB".InsAddr'
      Size = 60
    end
    object MainQueryInsOther: TStringField
      FieldName = 'InsOther'
      Origin = 'UNIVERSDB."Univers.DB".InsOther'
      Size = 30
    end
    object MainQueryInsObj: TMemoField
      FieldName = 'InsObj'
      Origin = 'UNIVERSDB."Univers.DB".InsObj'
      BlobType = ftMemo
      Size = 1
    end
    object MainQueryInsPcnt: TFloatField
      FieldName = 'InsPcnt'
      Origin = 'UNIVERSDB."Univers.DB".InsPcnt'
    end
    object MainQueryZastr: TStringField
      FieldName = 'Zastr'
      Origin = 'UNIVERSDB."Univers.DB".Zastr'
      Size = 60
    end
    object MainQueryVigod: TStringField
      FieldName = 'Vigod'
      Origin = 'UNIVERSDB."Univers.DB".Vigod'
      Size = 60
    end
    object MainQueryVigodAddr: TStringField
      FieldName = 'VigodAddr'
      Origin = 'UNIVERSDB."Univers.DB".VigodAddr'
      Size = 60
    end
    object MainQueryLizoZast: TStringField
      FieldName = 'LizoZast'
      Origin = 'UNIVERSDB."Univers.DB".LizoZast'
      Size = 60
    end
    object MainQueryPlaceIns: TStringField
      FieldName = 'PlaceIns'
      Origin = 'UNIVERSDB."Univers.DB".PlaceIns'
      Size = 30
    end
    object MainQueryInsEvents: TMemoField
      FieldName = 'InsEvents'
      Origin = 'UNIVERSDB."Univers.DB".InsEvents'
      BlobType = ftMemo
      Size = 1
    end
    object MainQueryPropertyIns: TMemoField
      FieldName = 'PropertyIns'
      Origin = 'UNIVERSDB."Univers.DB".PropertyIns'
      BlobType = ftMemo
      Size = 1
    end
    object MainQueryVariant: TStringField
      FieldName = 'Variant'
      Origin = 'UNIVERSDB."Univers.DB".Variant'
      Size = 60
    end
    object MainQueryRealSum: TFloatField
      FieldName = 'RealSum'
      Origin = 'UNIVERSDB."Univers.DB".RealSum'
    end
    object MainQueryRealCurr: TStringField
      FieldName = 'RealCurr'
      Origin = 'UNIVERSDB."Univers.DB".RealCurr'
      Size = 3
    end
    object MainQueryInsSum: TFloatField
      FieldName = 'InsSum'
      Origin = 'UNIVERSDB."Univers.DB".InsSum'
    end
    object MainQueryInsCur: TStringField
      FieldName = 'InsCur'
      Origin = 'UNIVERSDB."Univers.DB".InsCur'
      Size = 3
    end
    object MainQueryFransh: TMemoField
      FieldName = 'Fransh'
      Origin = 'UNIVERSDB."Univers.DB".Fransh'
      BlobType = ftMemo
      Size = 1
    end
    object MainQueryWait: TStringField
      FieldName = 'Wait'
      Origin = 'UNIVERSDB."Univers.DB".Wait'
      Size = 30
    end
    object MainQueryPremSum: TFloatField
      FieldName = 'PremSum'
      Origin = 'UNIVERSDB."Univers.DB".PremSum'
    end
    object MainQueryPremCur: TStringField
      FieldName = 'PremCur'
      Origin = 'UNIVERSDB."Univers.DB".PremCur'
      Size = 3
    end
    object MainQueryFeer: TStringField
      FieldName = 'Feer'
      Origin = 'UNIVERSDB."Univers.DB".Feer'
      Size = 30
    end
    object MainQueryFeeSum: TFloatField
      FieldName = 'FeeSum'
      Origin = 'UNIVERSDB."Univers.DB".FeeSum'
    end
    object MainQueryFeeCur: TStringField
      FieldName = 'FeeCur'
      Origin = 'UNIVERSDB."Univers.DB".FeeCur'
      Size = 3
    end
    object MainQueryFeeDate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1087#1083#1072#1090#1099
      FieldName = 'FeeDate'
      Origin = 'UNIVERSDB."Univers.DB".FeeDate'
    end
    object MainQueryRepDate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
      FieldName = 'RepDate'
      Origin = 'UNIVERSDB."UNIVERS.DB".RepDate'
    end
    object MainQueryFeeDoc: TStringField
      FieldName = 'FeeDoc'
      Origin = 'UNIVERSDB."Univers.DB".FeeDoc'
      Size = 10
    end
    object MainQueryFeeTyp: TStringField
      FieldName = 'FeeTyp'
      Origin = 'UNIVERSDB."Univers.DB".FeeTyp'
      Size = 1
    end
    object MainQuerySrokPay: TStringField
      FieldName = 'SrokPay'
      Origin = 'UNIVERSDB."Univers.DB".SrokPay'
      Size = 100
    end
    object MainQueryOtherCond: TMemoField
      FieldName = 'OtherCond'
      Origin = 'UNIVERSDB."Univers.DB".OtherCond'
      BlobType = ftMemo
      Size = 1
    end
    object MainQueryAgent: TStringField
      FieldName = 'Agent'
      Origin = 'UNIVERSDB."Univers.DB".Agent'
      Size = 4
    end
    object MainQueryAgPcnt: TFloatField
      FieldName = 'AgPcnt'
      Origin = 'UNIVERSDB."Univers.DB".AgPcnt'
    end
    object MainQueryAgType: TStringField
      FieldName = 'AgType'
      Origin = 'UNIVERSDB."Univers.DB".AgType'
      Size = 1
    end
    object MainQueryState: TStringField
      FieldName = 'State'
      Origin = 'UNIVERSDB."Univers.DB".State'
      Size = 1
    end
  end
  object UniversDb: TDatabase
    AliasName = 'BASO'
    Connected = True
    DatabaseName = 'UniversDb'
    SessionName = 'Default'
    Left = 60
    Top = 32
  end
  object DataSource: TDataSource
    DataSet = MainQuery
    Left = 92
    Top = 168
  end
  object AgentsTbl: TQuery
    DatabaseName = 'UniversDb'
    SQL.Strings = (
      'SELECT'
      '    AGENT_CODE, NAME'
      'FROM'
      '   AGENT'
      'WHERE'
      '   RUSS != -1 AND RUSS IS NOT NULL'
      'ORDER BY'
      '   NAME')
    Left = 128
    Top = 48
    object AgentsTblAGENT_CODE: TStringField
      FieldName = 'AGENT_CODE'
      Origin = '"AGENT.DB".Agent_code'
      Size = 4
    end
    object AgentsTblNAME: TStringField
      FieldName = 'NAME'
      Origin = '"AGENT.DB".Name'
      Size = 60
    end
  end
  object Timer: TTimer
    Enabled = False
    Interval = 350
    OnTimer = TimerTimer
    Left = 56
    Top = 144
  end
  object PopupMenu: TPopupMenu
    Left = 136
    Top = 156
    object MenuItem1: TMenuItem
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1090#1077#1082#1091#1097#1077#1077' '#1089#1086#1089#1090#1086#1103#1085#1080#1077' '#1082#1072#1082'...'
      OnClick = MenuItem1Click
    end
    object N2: TMenuItem
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1080#1079' '#1085#1072#1073#1086#1088#1072'...'
      OnClick = N2Click
    end
  end
  object REPQuery: TRxQuery
    DatabaseName = 'UniversDb'
    Macros = <>
    Left = 424
    Top = 20
  end
  object REPDS: TDataSource
    DataSet = REPQuery
    Left = 152
    Top = 404
  end
  object WorkSQL: TQuery
    DatabaseName = 'UniversDb'
    Left = 176
    Top = 128
  end
  object MainMenu: TMainMenu
    Left = 248
    Top = 16
    object N1: TMenuItem
      Caption = #1054#1090#1095#1105#1090
      object N4: TMenuItem
        Caption = #1060#1080#1083#1100#1090#1088
        OnClick = N4Click
      end
      object N5: TMenuItem
        Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
        OnClick = N5Click
      end
      object N6: TMenuItem
        Caption = #1054#1090#1095#1105#1090#1099
        OnClick = N6Click
      end
      object InputSum: TMenuItem
        Caption = #1055#1086#1083#1091#1095#1077#1085#1085#1099#1077' '#1074#1079#1085#1086#1089#1099' + '#1050#1056'...'
        OnClick = InputSumClick
      end
      object N9: TMenuItem
        Caption = #1054#1090#1095#1105#1090' '#1087#1086' '#1072#1075#1077#1085#1090#1091'...'
        OnClick = N9Click
      end
      object N7: TMenuItem
        Caption = '-'
      end
      object mnuExport: TMenuItem
        Caption = #1069#1082#1089#1087#1086#1088#1090'  '#1074' '#1089#1080#1089#1090#1077#1084#1091' '#1055#1054#1051#1048#1057'...'
        Default = True
        OnClick = mnuExportClick
      end
      object N8: TMenuItem
        Caption = '-'
      end
      object N3: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N3Click
      end
    end
  end
  object OpenDialogPx: TOpenDialog
    DefaultExt = 'db'
    Filter = 'Paradox '#1092#1072#1081#1083#1099'|*.db'
    Title = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093' '#1040#1083#1100#1074#1077#1085#1099
    Left = 300
    Top = 16
  end
  object SaveDialog: TSaveDialog
    DefaultExt = 'unload'
    Filter = #1042#1089#1077' '#1092#1072#1081#1083#1099' (*.*)|*.*'
    InitialDir = 'C:\'
    Title = #1042#1099#1075#1088#1091#1079#1082#1072' '#1076#1072#1085#1085#1099#1093' '#1074' '#1044#1080#1088#1077#1082#1094#1080#1102
    Left = 36
    Top = 408
  end
  object FormStorage: TFormStorage
    IniFileName = #1057#1090#1088#1072#1093#1086#1074#1072#1085#1080#1077
    IniSection = 'Univers'
    UseRegistry = True
    StoredProps.Strings = (
      'NumberFrom.Text'
      'NumberTo.Text'
      'RegDateFrom.Date'
      'RegDateFrom.Checked'
      'RegDateTo.Date'
      'RegDateTo.Checked')
    StoredValues = <>
    Left = 104
    Top = 408
  end
  object auxTable: TTable
    DatabaseName = 'UniversDb'
    TableName = 'MANDRUSS.DB'
    Left = 48
    Top = 240
  end
end
