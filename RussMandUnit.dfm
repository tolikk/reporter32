object RUSSMAND: TRUSSMAND
  Left = 326
  Top = 160
  Width = 600
  Height = 600
  Caption = #1054#1073#1103#1079#1072#1083#1086#1074#1082#1072' '#1056#1086#1089#1089#1080#1080
  Color = clBtnFace
  Constraints.MinHeight = 600
  Constraints.MinWidth = 600
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
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
    Top = 513
    Width = 592
    Height = 33
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      592
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
      Left = 312
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
      Width = 315
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
    Width = 592
    Height = 513
    ActivePage = FilterTab
    Align = alClient
    TabIndex = 1
    TabOrder = 1
    OnChange = TabbedNotebookChange
    object RepTab: TTabSheet
      Caption = #1055#1086#1083#1080#1089#1099
      object MainGrid: TRxDBGrid
        Left = 0
        Top = 0
        Width = 584
        Height = 454
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
            FieldName = 'RegDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsStart'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InsEnd'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Insurer'
            Width = 131
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Owner'
            Width = 123
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'OwnerT'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Marka'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AutoNmb'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Charact'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Tarif'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PayDate'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AgName'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'AgPcnt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StopDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RepDt'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PrevSN'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay1Info'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Pay2Info'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'CompanyName'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StateName'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RetSumm'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'resident'
            Visible = True
          end>
      end
      object Panel4: TPanel
        Left = 0
        Top = 454
        Width = 584
        Height = 31
        Align = alBottom
        BevelOuter = bvNone
        TabOrder = 1
        DesignSize = (
          584
          31)
        object btnSort: TButton
          Left = 424
          Top = 4
          Width = 85
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1057#1086#1088#1090#1080#1088#1086#1074#1082#1072'...'
          TabOrder = 0
          OnClick = btnSortClick
        end
        object btnExcel: TButton
          Left = 508
          Top = 4
          Width = 75
          Height = 25
          Anchors = [akRight, akBottom]
          Caption = #1074' Excel'
          TabOrder = 1
          OnClick = btnExcelClick
        end
        object btnReserveCalc: TButton
          Left = 0
          Top = 4
          Width = 117
          Height = 25
          Caption = #1057#1090#1088#1072#1093#1086#1074#1086#1081' '#1088#1077#1079#1077#1088#1074
          Font.Charset = RUSSIAN_CHARSET
          Font.Color = 8421631
          Font.Height = -11
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          TabOrder = 2
          OnClick = btnReserveCalcClick
        end
        object btnMag11: TButton
          Left = 164
          Top = 4
          Width = 75
          Height = 25
          Caption = #1046#1091#1088#1085#1072#1083' 1.1'
          TabOrder = 3
          OnClick = btnMag11Click
        end
        object btnRNPU: TButton
          Tag = 1
          Left = 116
          Top = 4
          Width = 49
          Height = 25
          Caption = #1056#1053#1055#1059
          TabOrder = 4
          OnClick = btnReserveCalcClick
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
        Top = 289
        Width = 7
        Height = 13
        Caption = #1057
      end
      object Label2: TLabel
        Left = 304
        Top = 289
        Width = 15
        Height = 13
        Caption = #1055#1054
      end
      object Label3: TLabel
        Left = 184
        Top = 313
        Width = 7
        Height = 13
        Caption = #1057
      end
      object Label4: TLabel
        Left = 304
        Top = 313
        Width = 15
        Height = 13
        Caption = #1055#1054
      end
      object Label7: TLabel
        Left = 184
        Top = 337
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
        Top = 337
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
        Height = 277
        Flat = False
        ItemHeight = 13
        TabOrder = 1
        OnClick = AgentListClick
      end
      object IsNumber: TCheckBox
        Left = 8
        Top = 287
        Width = 105
        Height = 17
        Caption = #1053#1086#1084#1077#1088
        TabOrder = 2
      end
      object NumberFrom: TEdit
        Left = 200
        Top = 285
        Width = 97
        Height = 21
        TabOrder = 3
      end
      object NumberTo: TEdit
        Left = 328
        Top = 285
        Width = 233
        Height = 21
        TabOrder = 4
      end
      object IsRegDate: TCheckBox
        Left = 8
        Top = 311
        Width = 121
        Height = 17
        Caption = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
        TabOrder = 5
      end
      object RegDateFrom: TDateTimePicker
        Left = 200
        Top = 309
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
        Top = 309
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
        Top = 335
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
        Top = 333
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
        Top = 333
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
      object IsOwner: TCheckBox
        Left = 8
        Top = 383
        Width = 97
        Height = 17
        Caption = #1042#1083#1072#1076#1077#1083#1077#1094
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 11
      end
      object OwnerLike: TEdit
        Left = 200
        Top = 383
        Width = 229
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 12
      end
      object IsState: TCheckBox
        Left = 8
        Top = 409
        Width = 161
        Height = 17
        Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 13
      end
      object State: TComboBox
        Left = 200
        Top = 407
        Width = 361
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 14
        Items.Strings = (
          #1053#1054#1056#1052#1040#1051#1068#1053#1067#1049
          #1059#1058#1045#1056#1071#1053
          #1048#1057#1055#1054#1056#1063#1045#1053
          #1056#1040#1057#1058#1054#1056#1043#1053#1059#1058
          #1044#1059#1041#1051#1048#1050#1040#1058)
      end
      object IsRepDate: TCheckBox
        Left = 8
        Top = 359
        Width = 121
        Height = 17
        Caption = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 15
        OnClick = IsRepDateClick
      end
      object DateRepStart: TDateTimePicker
        Left = 200
        Top = 357
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
        TabOrder = 16
      end
      object DateRepEnd: TDateTimePicker
        Left = 328
        Top = 357
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
        TabOrder = 17
      end
      object btnApply: TButton
        Left = 486
        Top = 457
        Width = 75
        Height = 25
        Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
        TabOrder = 18
        OnClick = btnApplyClick
      end
      object listCompanies: TComboBox
        Left = 200
        Top = 431
        Width = 361
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 19
      end
      object IsFiltCompany: TCheckBox
        Left = 8
        Top = 433
        Width = 161
        Height = 17
        Caption = #1050#1086#1084#1087#1072#1085#1080#1103
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 20
      end
      object listOwnType: TComboBox
        Left = 432
        Top = 383
        Width = 129
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 21
        Items.Strings = (
          #1053#1077' '#1074#1072#1078#1085#1086
          #1060#1080#1079' '#1083#1080#1094#1086
          #1070#1088' '#1083#1080#1094#1086)
      end
      object IsStopDate: TCheckBox
        Left = 97
        Top = 354
        Width = 97
        Height = 17
        Caption = '+ '#1056#1072#1089#1090#1086#1088#1078#1077#1085#1080#1077
        TabOrder = 22
      end
      object IsBad: TCheckBox
        Left = 97
        Top = 366
        Width = 97
        Height = 17
        Caption = '+ '#1048#1089#1087#1086#1088#1095#1077#1085
        TabOrder = 23
      end
    end
    object StatPage: TTabSheet
      Caption = #1057#1090#1072#1090#1080#1089#1090#1080#1082#1072
      ImageIndex = 2
      object StatisticTxt: TMemo
        Left = 0
        Top = 0
        Width = 584
        Height = 485
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
        584
        485)
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
        Width = 577
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
        Width = 483
        Height = 21
        Anchors = [akLeft, akTop, akRight]
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
      end
      object StartButton: TButton
        Left = 492
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
        Width = 575
        Height = 362
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
        Left = 392
        Top = 471
        Width = 97
        Height = 17
        Anchors = [akRight, akBottom]
        Caption = #1053#1091#1084#1077#1088#1086#1074#1072#1090#1100
        TabOrder = 4
      end
      object Button2: TButton
        Left = 491
        Top = 467
        Width = 87
        Height = 25
        Anchors = [akRight, akBottom]
        Caption = #1042' Excel'
        TabOrder = 5
        OnClick = Button2Click
      end
      object IsUseFilter: TCheckBox
        Left = 4
        Top = 469
        Width = 133
        Height = 17
        Anchors = [akLeft, akBottom]
        Caption = #1042' '#1090#1077#1082#1091#1097#1077#1081' '#1074#1099#1073#1086#1088#1082#1077
        TabOrder = 6
      end
    end
  end
  object MainQuery: TRxQuery
    OnCalcFields = MainQueryCalcFields
    DatabaseName = 'RussDatabase'
    SQL.Strings = (
      'SELECT'
      '  *'
      'FROM'
      '  MANDRUSS'
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
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Seria'
      Size = 4
    end
    object MainQueryNumber: TFloatField
      DisplayLabel = #1053#1086#1084#1077#1088
      FieldName = 'Number'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Number'
    end
    object MainQueryPSer: TStringField
      FieldName = 'PSer'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".PSer'
      Visible = False
      Size = 4
    end
    object MainQueryPNmb: TFloatField
      FieldName = 'PNmb'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".PNmb'
      Visible = False
    end
    object MainQueryRegDt: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088#1072#1094#1080#1080
      FieldName = 'RegDt'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".RegDt'
    end
    object MainQueryInsStart: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1085#1072#1095#1072#1083#1072
      FieldName = 'InsStart'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".InsStart'
    end
    object MainQueryTmStart: TStringField
      FieldName = 'TmStart'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".TmStart'
      Visible = False
      Size = 5
    end
    object MainQueryPeriod: TSmallintField
      DisplayLabel = #1055#1077#1088#1080#1086#1076
      FieldName = 'Period'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Period'
    end
    object MainQueryInsEnd: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1082#1086#1085#1095#1072#1085#1080#1103
      FieldName = 'InsEnd'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".InsEnd'
    end
    object MainQueryInsurer: TStringField
      DisplayLabel = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100
      FieldName = 'Insurer'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Insurer'
      Size = 48
    end
    object MainQueryInsurerT: TStringField
      FieldName = 'InsurerT'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".InsurerT'
      Visible = False
      Size = 1
    end
    object MainQueryOwner: TStringField
      DisplayLabel = #1042#1083#1072#1076#1077#1083#1077#1094
      FieldName = 'Owner'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Owner'
      Size = 48
    end
    object MainQueryOwnerT: TStringField
      DisplayLabel = #1058#1080#1087' '#1074#1083#1072#1076#1077#1083#1100#1094#1072
      FieldName = 'OwnerT'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".OwnerT'
      Visible = False
      Size = 1
    end
    object MainQueryMarka: TStringField
      DisplayLabel = #1052#1072#1088#1082#1072
      FieldName = 'Marka'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Marka'
      Size = 24
    end
    object MainQueryAutoId: TStringField
      FieldName = 'AutoId'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".AutoId'
      Visible = False
      Size = 17
    end
    object MainQueryPassSer: TStringField
      FieldName = 'PassSer'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".PassSer'
      Visible = False
      Size = 4
    end
    object MainQueryPassNmb: TFloatField
      FieldName = 'PassNmb'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".PassNmb'
      Visible = False
    end
    object MainQueryAutoNmb: TStringField
      DisplayLabel = #1053#1086#1084#1077#1088' '#1072#1074#1090#1086
      FieldName = 'AutoNmb'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".AutoNmb'
      Size = 12
    end
    object MainQuerySpecSr: TStringField
      FieldName = 'SpecSr'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".SpecSr'
      Visible = False
      Size = 2
    end
    object MainQuerySpecNmb: TFloatField
      FieldName = 'SpecNmb'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".SpecNmb'
      Visible = False
    end
    object MainQueryTarifGrp: TSmallintField
      FieldName = 'TarifGrp'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".TarifGrp'
      Visible = False
    end
    object MainQueryCharact: TSmallintField
      DisplayLabel = #1061#1072#1088#1072#1082#1090#1077#1088#1080#1089#1090#1080#1082#1072
      FieldName = 'Charact'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Charact'
    end
    object MainQueryTarif: TFloatField
      DisplayLabel = #1058#1072#1088#1080#1092
      FieldName = 'Tarif'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Tarif'
    end
    object MainQueryPayDate: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1087#1083#1072#1090#1077#1078#1072
      FieldName = 'PayDate'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".PayDate'
    end
    object MainQueryPay1: TFloatField
      DisplayLabel = '1'#1103' '#1086#1087#1083#1072#1090#1072
      FieldName = 'Pay1'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Pay1'
    end
    object MainQueryCurr1: TStringField
      DisplayLabel = '1'#1103' '#1054#1087#1083#1072#1090#1072' ('#1042#1072#1083#1102#1090#1072')'
      FieldName = 'Curr1'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Curr1'
      Size = 3
    end
    object MainQueryPay2: TFloatField
      DisplayLabel = '2'#1103' '#1054#1087#1083#1072#1090#1072
      FieldName = 'Pay2'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Pay2'
    end
    object MainQueryCurr2: TStringField
      DisplayLabel = '2'#1103' '#1054#1087#1083#1072#1090#1072' ('#1042#1072#1083#1102#1090#1072')'
      FieldName = 'Curr2'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Curr2'
      Size = 3
    end
    object MainQueryAgent: TStringField
      FieldName = 'Agent'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".Agent'
      Size = 4
    end
    object MainQueryAgType: TStringField
      FieldName = 'AgType'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".AgType'
      Visible = False
      Size = 1
    end
    object MainQueryAgPcnt: TFloatField
      DisplayLabel = '% '#1040#1075#1077#1085#1090#1072
      FieldName = 'AgPcnt'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".AgPcnt'
    end
    object MainQueryState: TSmallintField
      DisplayLabel = #1057#1086#1089#1090#1086#1103#1085#1080#1077
      FieldName = 'State'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".State'
    end
    object MainQueryStopDt: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1088#1072#1089#1090#1086#1088#1078#1077#1085#1080#1103
      FieldName = 'StopDt'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".StopDt'
    end
    object MainQueryRetSum: TFloatField
      DisplayLabel = #1057#1091#1084#1084#1072' '#1074#1086#1079#1074#1088#1072#1090#1072
      FieldName = 'RetSum'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".RetSum'
    end
    object MainQueryRetCurr: TStringField
      DisplayLabel = #1042#1072#1083#1102#1090#1072' '#1074#1086#1079#1074#1088#1072#1090#1072
      FieldName = 'RetCurr'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".RetCurr'
      Size = 3
    end
    object MainQueryInsComp: TSmallintField
      DisplayLabel = #1057#1090#1088#1072#1093#1086#1074#1072#1103' '#1082#1086#1084#1087#1072#1085#1080#1103
      FieldName = 'InsComp'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".InsComp'
    end
    object MainQueryRepDt: TDateField
      DisplayLabel = #1044#1072#1090#1072' '#1086#1090#1095#1105#1090#1072
      FieldName = 'RepDt'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".RepDt'
    end
    object MainQueryPrevSN: TStringField
      DisplayLabel = #1042#1079#1072#1084#1077#1085
      FieldKind = fkCalculated
      FieldName = 'PrevSN'
      Size = 16
      Calculated = True
    end
    object MainQueryPay1Info: TStringField
      DisplayLabel = '1'#1103' '#1054#1087#1083#1072#1090#1072
      FieldKind = fkCalculated
      FieldName = 'Pay1Info'
      Size = 16
      Calculated = True
    end
    object MainQueryPay2Info: TStringField
      DisplayLabel = '2'#1103' '#1054#1087#1083#1072#1090#1072
      FieldKind = fkCalculated
      FieldName = 'Pay2Info'
      Size = 16
      Calculated = True
    end
    object MainQueryAgName: TStringField
      DisplayLabel = #1040#1075#1077#1085#1090' '
      FieldKind = fkLookup
      FieldName = 'AgName'
      LookupDataSet = AgentsTbl
      LookupKeyFields = 'AGENT_CODE'
      LookupResultField = 'NAME'
      KeyFields = 'Agent'
      Lookup = True
    end
    object MainQueryCompanyName: TStringField
      DisplayLabel = #1050#1086#1084#1087#1072#1085#1080#1103' '
      FieldKind = fkCalculated
      FieldName = 'CompanyName'
      Size = 24
      Calculated = True
    end
    object MainQueryStateName: TStringField
      DisplayLabel = #1057#1086#1089#1090#1086#1103#1085#1080#1077' '
      FieldKind = fkCalculated
      FieldName = 'StateName'
      Size = 16
      Calculated = True
    end
    object MainQueryRetSumm: TStringField
      DisplayLabel = #1042#1086#1079#1074#1088#1072#1090
      FieldKind = fkCalculated
      FieldName = 'RetSumm'
      Size = 16
      Calculated = True
    end
    object MainQueryresident: TStringField
      DisplayLabel = #1056#1077#1079#1080#1076#1077#1085#1090
      FieldName = 'resident'
      Origin = 'RUSSDATABASE."MANDRUSS.DB".resident'
      Size = 1
    end
  end
  object RussDatabase: TDatabase
    AliasName = 'BASO'
    DatabaseName = 'RussDatabase'
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
    DatabaseName = 'RussDatabase'
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
    DatabaseName = 'RussDatabase'
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
    DatabaseName = 'RussDatabase'
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
      object ImpAlvena: TMenuItem
        Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093' '#1040#1083#1100#1074#1077#1085#1099'...'
        OnClick = ImpAlvenaClick
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
    Filter = #1060#1072#1081#1083#1099' '#1040#1083#1100#1074#1077#1085#1072' '#1087#1086' '#1056#1086#1089#1089#1080#1080'|GORU*.db|Paradox '#1092#1072#1081#1083#1099'|*.db'
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
    IniSection = #1056#1086#1089#1089#1080#1103
    UseRegistry = True
    StoredProps.Strings = (
      'NumberFrom.Text'
      'NumberTo.Text'
      'RegDateFrom.Date'
      'RegDateFrom.Checked'
      'RegDateTo.Date'
      'RegDateTo.Checked'
      'PayDateFrom.Date'
      'PayDateFrom.Checked'
      'PayDateTo.Date'
      'PayDateTo.Checked'
      'DateRepStart.Date'
      'DateRepStart.Checked'
      'DateRepEnd.Date'
      'DateRepEnd.Checked')
    StoredValues = <>
    Left = 104
    Top = 408
  end
  object auxTable: TTable
    DatabaseName = 'RussDatabase'
    TableName = 'MANDRUSS.DB'
    Left = 48
    Top = 240
  end
end
