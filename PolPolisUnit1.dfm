object PolPolisData: TPolPolisData
  Left = 225
  Top = 150
  Width = 660
  Height = 393
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSizeToolWin
  Caption = #1055#1086#1083#1080#1089
  Color = clBtnFace
  Constraints.MaxHeight = 393
  Constraints.MinHeight = 393
  Constraints.MinWidth = 660
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  KeyPreview = True
  OldCreateOrder = False
  Position = poOwnerFormCenter
  ShowHint = True
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object PanelData: TPanel
    Left = 0
    Top = 0
    Width = 652
    Height = 366
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      652
      366)
    object Label1: TLabel
      Left = 5
      Top = 319
      Width = 30
      Height = 13
      Caption = #1040#1075#1077#1085#1090
    end
    object Label2: TLabel
      Left = 267
      Top = 320
      Width = 11
      Height = 13
      Caption = '%'
    end
    object Label4: TLabel
      Left = 6
      Top = 58
      Width = 85
      Height = 13
      Caption = #1047#1072#1089#1090#1088#1072#1093#1086#1074#1072#1085#1085#1099#1081
    end
    object Label7: TLabel
      Left = 6
      Top = 82
      Width = 66
      Height = 13
      Caption = #1043#1086#1088#1086#1076'/'#1059#1083#1080#1094#1072
    end
    object Label8: TLabel
      Left = 6
      Top = 137
      Width = 88
      Height = 13
      Caption = #1052#1072#1088#1082#1072'/'#1052#1086#1076#1077#1083#1100#1058#1057
    end
    object Label9: TLabel
      Left = 6
      Top = 114
      Width = 75
      Height = 13
      Caption = #1053#1086#1084#1077#1088#1085#1086#1081' '#1079#1085#1072#1082
    end
    object Label10: TLabel
      Left = 220
      Top = 114
      Width = 110
      Height = 13
      Caption = #1053#1086#1084#1077#1088' '#1082#1091#1079#1086#1074#1072' '#1080#1083#1080' VIN'
    end
    object Label6: TLabel
      Left = 6
      Top = 217
      Width = 66
      Height = 13
      Caption = #1043#1086#1088#1086#1076'/'#1059#1083#1080#1094#1072
    end
    object Label11: TLabel
      Left = 484
      Top = 8
      Width = 64
      Height = 13
      Caption = #1056#1077#1075#1080#1089#1090#1088#1072#1094#1080#1103
    end
    object Label12: TLabel
      Left = 484
      Top = 268
      Width = 67
      Height = 13
      Caption = #1044#1072#1090#1072' '#1086#1087#1083#1072#1090#1099
    end
    object Label13: TLabel
      Left = 8
      Top = 268
      Width = 86
      Height = 13
      Caption = #1054#1087#1083#1072#1090#1072' (1 '#1080' 2 '#1095'.)'
    end
    object Label14: TLabel
      Left = 4
      Top = 8
      Width = 69
      Height = 13
      Caption = #1057#1077#1088#1080#1103', '#1053#1086#1084#1077#1088
    end
    object Label3: TLabel
      Left = 4
      Top = 32
      Width = 65
      Height = 13
      Caption = #1044#1072#1090#1072' '#1085#1072#1095#1072#1083#1072
    end
    object Label15: TLabel
      Left = 483
      Top = 32
      Width = 56
      Height = 13
      Caption = #1054#1082#1086#1085#1095#1072#1085#1080#1077
    end
    object Label16: TLabel
      Left = 176
      Top = 162
      Width = 153
      Height = 13
      Caption = #1054#1073#1098#1105#1084' '#1076#1074'./'#1043#1088#1091#1079#1086#1087#1086#1076#1098#1105#1084#1085#1086#1089#1090#1100
    end
    object Label17: TLabel
      Left = 6
      Top = 244
      Width = 32
      Height = 13
      Caption = #1058#1072#1088#1080#1092
    end
    object Label18: TLabel
      Left = 8
      Top = 162
      Width = 30
      Height = 13
      Caption = #1041#1091#1082#1074#1072
    end
    object Label19: TLabel
      Left = 4
      Top = 347
      Width = 63
      Height = 13
      Anchors = [akLeft, akBottom]
      Caption = #1057#1054#1057#1058#1054#1071#1053#1048#1045
    end
    object Label21: TLabel
      Left = 283
      Top = 32
      Width = 73
      Height = 13
      Caption = #1044#1083#1080#1090#1077#1083#1100#1085#1086#1089#1090#1100
    end
    object DBText1: TDBText
      Left = 72
      Top = 347
      Width = 133
      Height = 14
      Anchors = [akLeft, akBottom]
      Color = 16776176
      DataField = 'StateName'
      DataSource = DataSource
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label20: TLabel
      Left = 184
      Top = 244
      Width = 37
      Height = 13
      Caption = #1055#1088#1077#1084#1080#1103
    end
    object Label22: TLabel
      Left = 484
      Top = 82
      Width = 59
      Height = 13
      Caption = #1044#1086#1084'/'#1050#1074#1072#1088#1090'.'
    end
    object Label24: TLabel
      Left = 484
      Top = 58
      Width = 37
      Height = 13
      Caption = #1057#1090#1088#1072#1085#1072
    end
    object Label23: TLabel
      Left = 484
      Top = 114
      Width = 67
      Height = 13
      Caption = #8470#1076#1074#1080#1075#1072#1090#1077#1083#1103
    end
    object Label25: TLabel
      Left = 484
      Top = 138
      Width = 19
      Height = 13
      Caption = #1042#1080#1076
    end
    object Label26: TLabel
      Left = 300
      Top = 138
      Width = 18
      Height = 13
      Caption = #1058#1080#1087
    end
    object Label27: TLabel
      Left = 484
      Top = 217
      Width = 59
      Height = 13
      Caption = #1044#1086#1084'/'#1050#1074#1072#1088#1090'.'
    end
    object Label28: TLabel
      Left = 484
      Top = 193
      Width = 37
      Height = 13
      Caption = #1057#1090#1088#1072#1085#1072
    end
    object Bevel1: TBevel
      Left = 0
      Top = 51
      Width = 652
      Height = 51
      Anchors = [akLeft, akTop, akRight]
    end
    object Bevel2: TBevel
      Left = 0
      Top = 104
      Width = 652
      Height = 81
      Anchors = [akLeft, akTop, akRight]
    end
    object Bevel3: TBevel
      Left = 0
      Top = 237
      Width = 651
      Height = 52
      Anchors = [akLeft, akTop, akRight]
    end
    object Label29: TLabel
      Left = 6
      Top = 296
      Width = 58
      Height = 13
      Caption = #1056#1072#1089#1090#1086#1088#1075#1085#1091#1090
    end
    object Label30: TLabel
      Left = 232
      Top = 295
      Width = 41
      Height = 13
      Caption = #1042#1086#1079#1074#1088#1072#1090
    end
    object btnCopyInsData: TSpeedButton
      Left = 0
      Top = 188
      Width = 97
      Height = 19
      Caption = #1057#1090#1088#1072#1093#1086#1074#1072#1090#1077#1083#1100'    '
      Flat = True
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsUnderline]
      ParentFont = False
      OnClick = btnCopyInsDataClick
    end
    object dbAgentType: TDBCheckBox
      Left = 98
      Top = 318
      Width = 121
      Height = 17
      Caption = #1070#1088#1080#1076#1080#1095#1077#1089#1082#1086#1077' '#1083#1080#1094#1086
      Color = 14811135
      DataField = 'Uridich'
      DataSource = DataSource
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
      TabOrder = 41
      ValueChecked = 'True'
      ValueUnchecked = 'False'
    end
    object dbAgentPcnt: TDBEdit
      Left = 224
      Top = 316
      Width = 41
      Height = 21
      AutoSize = False
      Color = 14811135
      DataField = 'AgPercent'
      DataSource = DataSource
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 42
    end
    object dbAgent: TRxDBLookupCombo
      Left = 280
      Top = 316
      Width = 370
      Height = 21
      DropDownCount = 24
      Color = 14811135
      DataField = 'Agent'
      DataSource = DataSource
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Anchors = [akLeft, akTop, akRight]
      LookupDisplay = 'Name'
      ParentFont = False
      TabOrder = 43
    end
    object dbOwn: TDBEdit
      Left = 172
      Top = 54
      Width = 305
      Height = 21
      Color = 14811135
      DataField = 'AUTOOWNER'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 10
    end
    object dbOwnCity: TDBEdit
      Left = 100
      Top = 78
      Width = 133
      Height = 21
      Hint = #1060#1086#1088#1084#1072#1090' '#1087#1086#1083#1103': '#1043#1086#1088#1086#1076' '#1059#1083#1080#1094#1072
      Color = 14811135
      DataField = 'owncity'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 12
    end
    object dbMarka: TDBEdit
      Left = 100
      Top = 134
      Width = 185
      Height = 21
      Color = 14811135
      DataField = 'MARKA'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      PopupMenu = PopupMenuAuto
      TabOrder = 19
    end
    object dnAutoNmb: TDBEdit
      Left = 100
      Top = 110
      Width = 113
      Height = 21
      Color = 14811135
      DataField = 'AutoNumber'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 16
    end
    object dbBodyNmb: TDBEdit
      Left = 336
      Top = 110
      Width = 141
      Height = 21
      Color = 14811135
      DataField = 'BODYNO'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 17
    end
    object dbIns: TDBEdit
      Left = 172
      Top = 189
      Width = 305
      Height = 21
      Color = 14811135
      DataField = 'Name'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 25
    end
    object dbInsCity: TDBEdit
      Left = 100
      Top = 213
      Width = 129
      Height = 21
      Hint = #1060#1086#1088#1084#1072#1090' '#1087#1086#1083#1103': '#1043#1086#1088#1086#1076' '#1059#1083#1080#1094#1072
      Color = 14811135
      DataField = 'inscity'
      DataSource = DataSource
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 27
    end
    object dbReg: TDBEdit
      Left = 556
      Top = 4
      Width = 92
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      Color = 14811135
      DataField = 'RepDate'
      DataSource = DataSource
      TabOrder = 4
    end
    object dbPay: TDBEdit
      Left = 556
      Top = 264
      Width = 93
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      Color = 14811135
      DataField = 'PayDate'
      DataSource = DataSource
      TabOrder = 37
    end
    object dbPay1: TDBEdit
      Left = 100
      Top = 264
      Width = 125
      Height = 21
      Color = 14811135
      DataField = 'PremiumPay'
      DataSource = DataSource
      TabOrder = 33
    end
    object dbCurr1: TRxDBComboBox
      Left = 228
      Top = 264
      Width = 49
      Height = 21
      Color = 14811135
      DataField = 'PremiumCurr'
      DataSource = DataSource
      DropDownCount = 5
      ItemHeight = 13
      Items.Strings = (
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
      TabOrder = 34
      Values.Strings = (
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
    end
    object dbSeria: TDBEdit
      Left = 100
      Top = 4
      Width = 65
      Height = 21
      Color = 14811135
      DataField = 'Seria'
      DataSource = DataSource
      TabOrder = 0
    end
    object dbNumber: TDBEdit
      Left = 168
      Top = 4
      Width = 97
      Height = 21
      Color = 14811135
      DataField = 'Number'
      DataSource = DataSource
      TabOrder = 1
    end
    object dbPay2: TDBEdit
      Left = 280
      Top = 264
      Width = 129
      Height = 21
      Color = 14811135
      DataField = 'PremiumPay2'
      DataSource = DataSource
      TabOrder = 35
    end
    object dbCurr2: TRxDBComboBox
      Left = 410
      Top = 264
      Width = 67
      Height = 21
      Color = 14811135
      DataField = 'PremiumCurr2'
      DataSource = DataSource
      DropDownCount = 5
      ItemHeight = 13
      Items.Strings = (
        ''
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
      TabOrder = 36
      Values.Strings = (
        ''
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
    end
    object dbFrom: TDBEdit
      Left = 100
      Top = 28
      Width = 125
      Height = 21
      DataField = 'StartDate'
      DataSource = DataSource
      TabOrder = 5
    end
    object dbTo: TDBEdit
      Left = 556
      Top = 28
      Width = 92
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'EndDate'
      DataSource = DataSource
      ReadOnly = True
      TabOrder = 8
    end
    object dbFromTime: TDBEdit
      Left = 228
      Top = 28
      Width = 37
      Height = 21
      Hint = #1042#1088#1077#1084#1103' '#1085#1072#1095#1072#1083#1072
      DataField = 'StartTime'
      DataSource = DataSource
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnKeyDown = dbFromTimeKeyDown
    end
    object dbPeriod: TDBEdit
      Left = 364
      Top = 28
      Width = 113
      Height = 21
      Hint = #1047#1076#1077#1089#1100' '#1085#1091#1078#1085#1086' '#1091#1082#1072#1079#1072#1090#1100' '#1073#1099#1089#1090#1088#1099#1084' '#1089#1087#1086#1089#1086#1073#1086#1084
      DataField = 'Period'
      DataSource = DataSource
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
    end
    object dbCharact: TDBEdit
      Left = 336
      Top = 158
      Width = 313
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'CHARACT'
      DataSource = DataSource
      TabOrder = 23
    end
    object dbCountry: TDBComboBox
      Left = 556
      Top = 54
      Width = 93
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      CharCase = ecUpperCase
      DataField = 'OWNCNTRY'
      DataSource = DataSource
      DropDownCount = 15
      ItemHeight = 13
      Items.Strings = (
        'BY'
        'RU'
        'LT'
        'UA'
        'KZ'
        'AM'
        'KG'
        'AZ'
        'GE'
        'UZ'
        'TM'
        'MD')
      TabOrder = 11
    end
    object dbTarif: TDBEdit
      Left = 100
      Top = 240
      Width = 73
      Height = 21
      DataField = 'TraifCode'
      DataSource = DataSource
      ReadOnly = True
      TabOrder = 31
    end
    object Save: TButton
      Left = 500
      Top = 344
      Width = 75
      Height = 21
      Anchors = [akRight, akBottom]
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
      TabOrder = 47
      OnClick = SaveClick
    end
    object New: TButton
      Left = 232
      Top = 344
      Width = 75
      Height = 21
      Hint = #1057#1086#1079#1076#1072#1090#1100' '#1082#1086#1087#1080#1102' '#1090#1077#1082#1091#1097#1077#1075#1086' '#1087#1086#1083#1080#1089#1072
      Anchors = [akRight, akBottom]
      Caption = #1050#1086#1087#1080#1103
      TabOrder = 44
      OnClick = NewClick
    end
    object dbLetter: TDBComboBox
      Left = 100
      Top = 158
      Width = 65
      Height = 21
      CharCase = ecUpperCase
      DataField = 'Letter'
      DataSource = DataSource
      ItemHeight = 13
      Items.Strings = (
        'O'
        'A'
        'AT'
        'M'
        'P'
        'C'
        'N')
      TabOrder = 22
    end
    object btnContinue: TButton
      Left = 420
      Top = 344
      Width = 75
      Height = 21
      Anchors = [akRight, akBottom]
      Caption = #1055#1088#1086#1076#1083#1080#1090#1100
      TabOrder = 46
      OnClick = btnContinueClick
    end
    object btnDublicat: TButton
      Left = 312
      Top = 344
      Width = 107
      Height = 21
      Anchors = [akRight, akBottom]
      Caption = #1042#1099#1076#1072#1090#1100' '#1076#1091#1073#1083#1080#1082#1072#1090
      TabOrder = 45
      OnClick = btnDublicatClick
    end
    object dbPNumber: TDBEdit
      Left = 364
      Top = 4
      Width = 113
      Height = 21
      DataField = 'PNumber'
      DataSource = DataSource
      TabOrder = 3
    end
    object btnCancel: TButton
      Left = 576
      Top = 344
      Width = 75
      Height = 21
      Anchors = [akRight, akBottom]
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 48
      OnClick = btnCancelClick
    end
    object dbIsDup: TDBCheckBox
      Left = 270
      Top = 8
      Width = 92
      Height = 17
      Caption = #1044#1091#1073#1083#1080#1082#1072#1090' '#1076#1083#1103
      DataField = 'IsDup'
      DataSource = DataSource
      TabOrder = 2
      ValueChecked = 'Y'
      ValueUnchecked = 'N'
    end
    object dbPremium: TDBEdit
      Left = 228
      Top = 240
      Width = 249
      Height = 21
      DataField = 'PremiumVal'
      DataSource = DataSource
      ReadOnly = True
      TabOrder = 32
    end
    object dbOwnType: TDBCheckBox
      Left = 100
      Top = 56
      Width = 67
      Height = 17
      Caption = #1070#1088'. '#1083#1080#1094#1086
      Color = 14811135
      DataField = 'OWNTYPE'
      DataSource = DataSource
      ParentColor = False
      TabOrder = 9
      ValueChecked = 'U'
      ValueUnchecked = 'F'
    end
    object dbHouse: TDBEdit
      Left = 556
      Top = 78
      Width = 45
      Height = 21
      DataField = 'OWNHOUSE'
      DataSource = DataSource
      TabOrder = 14
    end
    object dbFlat: TDBEdit
      Left = 604
      Top = 78
      Width = 45
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'OWNFLAT'
      DataSource = DataSource
      TabOrder = 15
    end
    object dbEngine: TDBEdit
      Left = 556
      Top = 110
      Width = 93
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'EngNmb'
      DataSource = DataSource
      TabOrder = 18
    end
    object dbAutoType: TDBComboBox
      Left = 336
      Top = 134
      Width = 141
      Height = 21
      CharCase = ecUpperCase
      DataField = 'AutoType'
      DataSource = DataSource
      ItemHeight = 13
      Items.Strings = (
        'SEDAN'
        'KOMBI'
        'CABRIO')
      TabOrder = 20
    end
    object dbInsType: TDBCheckBox
      Left = 99
      Top = 190
      Width = 73
      Height = 17
      Caption = #1070#1088'. '#1083#1080#1094#1086
      Color = 14811135
      DataField = 'InsType'
      DataSource = DataSource
      ParentColor = False
      TabOrder = 24
      ValueChecked = 'U'
      ValueUnchecked = 'F'
    end
    object dbHouse2: TDBEdit
      Left = 556
      Top = 213
      Width = 45
      Height = 21
      DataField = 'InsHouse'
      DataSource = DataSource
      TabOrder = 29
    end
    object dbFlat2: TDBEdit
      Left = 603
      Top = 213
      Width = 45
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'InsFlat'
      DataSource = DataSource
      TabOrder = 30
    end
    object dbCountry2: TDBComboBox
      Left = 555
      Top = 189
      Width = 93
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      CharCase = ecUpperCase
      DataField = 'InsCntry'
      DataSource = DataSource
      DropDownCount = 15
      ItemHeight = 13
      Items.Strings = (
        'BY'
        'RU'
        'LT'
        'UA'
        'KZ'
        'AM'
        'KG'
        'AZ'
        'GE'
        'UZ'
        'TM'
        'MD')
      TabOrder = 26
    end
    object dbRetDate: TDBEdit
      Left = 100
      Top = 292
      Width = 125
      Height = 21
      DataField = 'RetDate'
      DataSource = DataSource
      TabOrder = 38
    end
    object dbRetSum: TDBEdit
      Left = 280
      Top = 292
      Width = 129
      Height = 21
      DataField = 'RetSum'
      DataSource = DataSource
      TabOrder = 39
    end
    object dbRetCurr: TRxDBComboBox
      Left = 410
      Top = 292
      Width = 67
      Height = 21
      DataField = 'RetCurr'
      DataSource = DataSource
      DropDownCount = 5
      ItemHeight = 13
      Items.Strings = (
        ''
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
      TabOrder = 40
      Values.Strings = (
        ''
        'BRB'
        'RUR'
        'USD'
        'EUR'
        'DM')
    end
    object dbStreet: TDBEdit
      Left = 235
      Top = 77
      Width = 242
      Height = 21
      Color = 14811135
      DataField = 'ownstreet'
      DataSource = DataSource
      TabOrder = 13
    end
    object dbInsStreet: TDBEdit
      Left = 232
      Top = 213
      Width = 245
      Height = 21
      Color = 14811135
      DataField = 'insstreet'
      DataSource = DataSource
      TabOrder = 28
    end
    object dbAutoVid: TDBLookupComboBox
      Left = 556
      Top = 132
      Width = 93
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      DataField = 'auto_vid'
      DataSource = DataSource
      DropDownRows = 12
      DropDownWidth = 200
      KeyField = 'Id'
      ListField = 'Name'
      ListSource = PolandPolises.DataSourceAutoVid
      TabOrder = 21
    end
  end
  object DataSource: TDataSource
    DataSet = PolandPolises.PolandQuery
    Left = 576
    Top = 288
  end
  object PopupMenuContinue: TPopupMenu
    Left = 480
    Top = 292
    object N151: TMenuItem
      Tag = 15
      Caption = #1085#1072' 15 '#1076#1085#1077#1081
      OnClick = COntinue
    end
    object N301: TMenuItem
      Tag = 30
      Caption = #1085#1072' 30 '#1076#1085#1077#1081
      OnClick = COntinue
    end
    object N12: TMenuItem
      Tag = 2
      Caption = #1085#1072' 2 '#1084#1077#1089#1103#1094#1072
      OnClick = COntinue
    end
    object N13: TMenuItem
      Tag = 3
      Caption = #1085#1072' 3 '#1084#1077#1089#1103#1094#1072
      OnClick = COntinue
    end
    object N14: TMenuItem
      Tag = 4
      Caption = #1085#1072' 4 '#1084#1077#1089#1103#1094#1072
      OnClick = COntinue
    end
    object N15: TMenuItem
      Tag = 5
      Caption = #1085#1072' 5 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N16: TMenuItem
      Tag = 6
      Caption = #1085#1072' 6 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N17: TMenuItem
      Tag = 7
      Caption = #1085#1072' 7 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N18: TMenuItem
      Tag = 8
      Caption = #1085#1072' 8 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N19: TMenuItem
      Tag = 9
      Caption = #1085#1072' 9 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N111: TMenuItem
      Tag = 10
      Caption = #1085#1072' 10 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
    object N112: TMenuItem
      Tag = 11
      Caption = #1085#1072' 11 '#1084#1077#1089#1103#1094#1077#1074
      OnClick = COntinue
    end
  end
  object Timer: TTimer
    Enabled = False
    Interval = 500
    OnTimer = TimerTimer
    Left = 20
    Top = 12
  end
  object PopupMenuAuto: TPopupMenu
    OnPopup = PopupMenuAutoPopup
    Left = 180
    Top = 140
  end
end
