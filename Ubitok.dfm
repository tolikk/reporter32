object UbitokFrm: TUbitokFrm
  Left = 261
  Top = 122
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1059#1073#1099#1090#1086#1082
  ClientHeight = 409
  ClientWidth = 435
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 12
    Top = 8
    Width = 77
    Height = 13
    Caption = #1056#1077#1075#1080#1089#1090#1088'. '#1085#1086#1084#1077#1088
  end
  object Label2: TLabel
    Left = 228
    Top = 8
    Width = 73
    Height = 13
    Caption = #1044#1072#1090#1072' '#1088#1077#1075#1080#1089#1090#1088'.'
  end
  object Label6: TLabel
    Left = 8
    Top = 104
    Width = 72
    Height = 13
    Caption = #1057#1091#1084#1084#1072' '#1091#1073#1099#1090#1082#1072
  end
  object Label11: TLabel
    Left = 12
    Top = 385
    Width = 78
    Height = 13
    Caption = #1044#1072#1090#1072' '#1079#1072#1082#1088#1099#1090#1080#1103
  end
  object Label12: TLabel
    Left = 12
    Top = 336
    Width = 64
    Height = 26
    Caption = #1044#1080#1072#1075#1085#1086#1079' '#1080#13#10#1055#1088#1080#1084#1077#1095#1072#1085#1080#1077' '
  end
  object Label13: TLabel
    Left = 12
    Top = 316
    Width = 37
    Height = 13
    Caption = #1057#1090#1088#1072#1085#1072
  end
  object Label14: TLabel
    Left = 8
    Top = 128
    Width = 65
    Height = 13
    Caption = #1076#1083#1103' '#1041#1086#1088#1076#1077#1088#1086
  end
  object Label99: TLabel
    Left = 8
    Top = 152
    Width = 61
    Height = 13
    Caption = #1057#1086#1089#1090#1086#1103#1085#1080#1077
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object RegNmb: TEdit
    Left = 100
    Top = 4
    Width = 121
    Height = 21
    TabOrder = 0
  end
  object RegDate: TDateTimePicker
    Left = 304
    Top = 4
    Width = 125
    Height = 21
    CalAlignment = dtaLeft
    Date = 37719.5657284375
    Time = 37719.5657284375
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 1
  end
  object Summa: TRxSpinEdit
    Left = 100
    Top = 100
    Width = 241
    Height = 21
    ValueType = vtFloat
    TabOrder = 2
  end
  object Currency: TComboBox
    Left = 344
    Top = 100
    Width = 85
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 3
    Items.Strings = (
      'USD'
      'RUR'
      'EUR'
      'BRB'
      'DM')
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 172
    Width = 421
    Height = 65
    Caption = ' '#1054#1055#1051#1040#1058#1040' '
    TabOrder = 4
    object Label7: TLabel
      Left = 12
      Top = 16
      Width = 26
      Height = 13
      Caption = #1044#1072#1090#1072
    end
    object Label8: TLabel
      Left = 12
      Top = 40
      Width = 31
      Height = 13
      Caption = #1057#1091#1084#1084#1072
    end
    object PaySum: TRxSpinEdit
      Left = 92
      Top = 36
      Width = 245
      Height = 21
      ValueType = vtFloat
      TabOrder = 0
    end
    object PayDate: TDateTimePicker
      Left = 92
      Top = 12
      Width = 121
      Height = 21
      CalAlignment = dtaLeft
      Date = 37719.5794046528
      Time = 37719.5794046528
      ShowCheckbox = True
      Checked = False
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 1
      OnChange = PayDateChange
    end
    object PayCurr: TComboBox
      Left = 340
      Top = 36
      Width = 73
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      Items.Strings = (
        ''
        'USD'
        'RUR'
        'EUR'
        'BRB'
        'DM')
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 240
    Width = 421
    Height = 69
    Caption = ' '#1042#1067#1057#1042#1054#1041#1054#1046#1044#1045#1053#1053#1067#1045' '#1054#1041#1071#1047#1040#1058#1045#1051#1068#1057#1058#1042#1040' '
    TabOrder = 5
    object Label9: TLabel
      Left = 12
      Top = 20
      Width = 26
      Height = 13
      Caption = #1044#1072#1090#1072
    end
    object Label10: TLabel
      Left = 12
      Top = 44
      Width = 31
      Height = 13
      Caption = #1057#1091#1084#1084#1072
    end
    object FreeSum: TRxSpinEdit
      Left = 92
      Top = 40
      Width = 245
      Height = 21
      ValueType = vtFloat
      TabOrder = 0
    end
    object FreeDate: TDateTimePicker
      Left = 92
      Top = 16
      Width = 121
      Height = 21
      CalAlignment = dtaLeft
      Date = 37719.5794046528
      Time = 37719.5794046528
      ShowCheckbox = True
      Checked = False
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 1
      OnChange = FreeDateChange
    end
    object FreeCurr: TComboBox
      Left = 340
      Top = 40
      Width = 73
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 2
      Items.Strings = (
        ''
        'USD'
        'RUR'
        'EUR'
        'BRB'
        'DM')
    end
  end
  object CloseDate: TDateTimePicker
    Left = 100
    Top = 382
    Width = 117
    Height = 21
    CalAlignment = dtaLeft
    Date = 37719.5657284375
    Time = 37719.5657284375
    DateFormat = dfShort
    DateMode = dmComboBox
    Kind = dtkDate
    ParseInput = False
    TabOrder = 6
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 28
    Width = 421
    Height = 65
    Caption = ' '#1047#1040#1071#1042#1051#1045#1053#1048#1045' '
    TabOrder = 7
    object Label5: TLabel
      Left = 216
      Top = 16
      Width = 73
      Height = 13
      Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
    end
    object Label3: TLabel
      Left = 12
      Top = 16
      Width = 26
      Height = 13
      Caption = #1044#1072#1090#1072
    end
    object Label4: TLabel
      Left = 12
      Top = 40
      Width = 31
      Height = 13
      Caption = #1053#1086#1084#1077#1088
    end
    object ComplainName: TEdit
      Left = 296
      Top = 12
      Width = 117
      Height = 21
      MaxLength = 16
      TabOrder = 0
      Text = #1047#1040#1071#1042#1051#1045#1053#1048#1045
    end
    object ComplainDate: TDateTimePicker
      Left = 92
      Top = 12
      Width = 121
      Height = 21
      CalAlignment = dtaLeft
      Date = 37719.565745081
      Time = 37719.565745081
      DateFormat = dfShort
      DateMode = dmComboBox
      Kind = dtkDate
      ParseInput = False
      TabOrder = 1
    end
    object ComplainNmb: TEdit
      Left = 92
      Top = 36
      Width = 321
      Height = 21
      MaxLength = 10
      TabOrder = 2
    end
  end
  object SaveBtn: TButton
    Left = 280
    Top = 380
    Width = 75
    Height = 25
    Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
    TabOrder = 8
    OnClick = SaveBtnClick
  end
  object Button2: TButton
    Left = 356
    Top = 380
    Width = 75
    Height = 25
    Caption = #1047#1072#1082#1088#1099#1090#1100
    TabOrder = 9
    OnClick = Button2Click
  end
  object Comments: TMemo
    Left = 100
    Top = 336
    Width = 329
    Height = 37
    MaxLength = 48
    TabOrder = 10
  end
  object Country: TEdit
    Left = 100
    Top = 312
    Width = 329
    Height = 21
    MaxLength = 24
    TabOrder = 11
  end
  object Bordero: TRxSpinEdit
    Left = 100
    Top = 124
    Width = 241
    Height = 21
    ValueType = vtFloat
    TabOrder = 12
  end
  object BorderoCurr: TComboBox
    Left = 344
    Top = 124
    Width = 85
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 13
    Items.Strings = (
      'USD'
      'RUR'
      'EUR'
      'BRB'
      'DM')
  end
  object StateCb: TComboBox
    Left = 100
    Top = 148
    Width = 329
    Height = 21
    Style = csDropDownList
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ItemHeight = 13
    ParentFont = False
    TabOrder = 14
    OnChange = StateCbChange
    Items.Strings = (
      #1047#1072#1103#1074#1083#1077#1085#1086
      #1054#1090#1082#1072#1079#1072#1085#1086
      #1054#1087#1083#1072#1095#1077#1085#1086)
  end
end
