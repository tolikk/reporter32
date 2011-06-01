object AddFldForm: TAddFldForm
  Left = 335
  Top = 254
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #1044#1086#1073#1072#1074#1083#1103#1077#1084' '#1087#1086#1083#1077' '#1074' '#1090#1072#1073#1083#1080#1094#1091
  ClientHeight = 197
  ClientWidth = 445
  Color = clBtnFace
  Constraints.MaxHeight = 231
  Constraints.MinHeight = 224
  Constraints.MinWidth = 200
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    445
    197)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 4
    Top = 6
    Width = 42
    Height = 13
    Caption = #1058#1072#1073#1083#1080#1094#1072
  end
  object lbNameFld: TLabel
    Left = 4
    Top = 44
    Width = 150
    Height = 13
    Caption = #1053#1072#1079#1074#1072#1085#1080#1077' '#1076#1086#1073#1072#1074#1083#1103#1077#1084#1086#1075#1086' '#1087#1086#1083#1103
  end
  object Label2: TLabel
    Left = 4
    Top = 84
    Width = 120
    Height = 13
    Caption = #1058#1080#1087' '#1076#1086#1073#1072#1074#1083#1103#1077#1084#1086#1075#1086' '#1087#1086#1083#1103
  end
  object lbSizeFld: TLabel
    Left = 4
    Top = 124
    Width = 62
    Height = 13
    Caption = #1056#1072#1079#1084#1077#1088' '#1087#1086#1083#1103
  end
  object TableName: TEdit
    Left = 4
    Top = 20
    Width = 436
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    ReadOnly = True
    TabOrder = 0
  end
  object NameFld: TEdit
    Left = 4
    Top = 60
    Width = 436
    Height = 21
    Anchors = [akLeft, akTop, akRight]
    MaxLength = 16
    TabOrder = 1
  end
  object FldType: TComboBox
    Left = 4
    Top = 100
    Width = 436
    Height = 21
    Style = csDropDownList
    Anchors = [akLeft, akTop, akRight]
    ItemHeight = 13
    TabOrder = 2
    Items.Strings = (
      #1057#1090#1088#1086#1082#1072
      #1062#1077#1083#1086#1077' '#1095#1080#1089#1083#1086' 2 '#1073#1072#1081#1090#1072
      #1062#1077#1083#1086#1077' '#1095#1080#1089#1083#1086' 4 '#1073#1072#1081#1090#1072
      #1044#1088#1086#1073#1085#1086#1077' '#1095#1080#1089#1083#1086
      #1044#1072#1090#1072)
  end
  object SizeFld: TRxSpinEdit
    Left = 4
    Top = 140
    Width = 436
    Height = 21
    MaxValue = 254
    MinValue = 1
    Value = 1
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 3
  end
  object Ok: TButton
    Left = 333
    Top = 172
    Width = 107
    Height = 21
    Action = FldSize
    Anchors = [akTop, akRight]
    Caption = #1044#1086#1073#1072#1074#1080#1090#1100
    TabOrder = 4
  end
  object ActionList: TActionList
    Left = 320
    Top = 16
    object FldSize: TAction
      Caption = 'FldSize'
      OnExecute = OkClick
      OnUpdate = FldSizeUpdate
    end
  end
end
