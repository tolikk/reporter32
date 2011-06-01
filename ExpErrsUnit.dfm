object ExpErrsFlag: TExpErrsFlag
  Left = 251
  Top = 106
  Width = 748
  Height = 475
  Caption = 'Ошибки экспорта (часть 2)'
  Color = clBtnFace
  Constraints.MinHeight = 300
  Constraints.MinWidth = 540
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 420
    Width = 740
    Height = 28
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object CloseBtn: TButton
      Left = 609
      Top = 2
      Width = 131
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Закрыть'
      TabOrder = 0
      OnClick = CloseBtnClick
    end
    object SaveBtn: TButton
      Left = 473
      Top = 2
      Width = 131
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Сохранить в файл...'
      TabOrder = 1
      OnClick = SaveBtnClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 740
    Height = 41
    Align = alTop
    TabOrder = 1
    object Label1: TLabel
      Left = 9
      Top = 13
      Width = 205
      Height = 13
      Caption = 'Директория с данными для БЮРО'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DirectoryEdit: TDirectoryEdit
      Left = 225
      Top = 9
      Width = 193
      Height = 21
      DialogKind = dkWin32
      DialogText = 'Директория'
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      NumGlyphs = 1
      ParentFont = False
      TabOrder = 0
    end
    object ExecBtn: TButton
      Left = 427
      Top = 9
      Width = 96
      Height = 21
      Caption = 'Выполнить'
      TabOrder = 1
      OnClick = ExecBtnClick
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 41
    Width = 740
    Height = 379
    Align = alClient
    TabOrder = 2
    object Splitter1: TSplitter
      Left = 321
      Top = 25
      Width = 8
      Height = 353
      Cursor = crHSplit
      ResizeStyle = rsUpdate
    end
    object DBGrid3: TDBGrid
      Left = 1
      Top = 25
      Width = 320
      Height = 353
      Align = alLeft
      DataSource = DataSource3
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 738
      Height = 24
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object LabelLeft3: TLabel
        Left = 2
        Top = 6
        Width = 120
        Height = 13
        Caption = 'Дубликаты владельцев'
      end
    end
    object DBGrid4: TDBGrid
      Left = 329
      Top = 25
      Width = 410
      Height = 353
      Align = alClient
      DataSource = DataSource4
      TabOrder = 2
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
  end
  object SaveDialog: TSaveDialog
    DefaultExt = 'txt'
    FileName = 'Ошибки3'
    Filter = 'Текстовые фалы|*.txt'
    Title = 'Сохранить'
    Left = 24
    Top = 376
  end
  object DataSource3: TDataSource
    DataSet = RxQuery3
    Left = 208
    Top = 272
  end
  object DataSource4: TDataSource
    DataSet = RxQuery4
    Left = 296
    Top = 272
  end
  object RxQuery3: TRxQuery
    SQL.Strings = (
      'SELECT * FROM %TABLE3 WHERE MSG_ID IN (%LIST)'
      'UNION ALL'
      'SELECT * FROM %TABLE3 WHERE MSG_ID IN (%LIST1)'
      'UNION ALL'
      'SELECT * FROM %TABLE3 WHERE MSG_ID IN (%LIST2)'
      'ORDER BY MSG_ID')
    Macros = <
      item
        DataType = ftString
        Name = 'TABLE3'
        ParamType = ptInput
        Value = #39'E:\1\3'#39
      end
      item
        DataType = ftString
        Name = 'LIST'
        ParamType = ptInput
        Value = #39'0'#39
      end
      item
        DataType = ftString
        Name = 'LIST1'
        ParamType = ptInput
        Value = '0=0'
      end
      item
        DataType = ftString
        Name = 'LIST2'
        ParamType = ptInput
        Value = '0=0'
      end>
    Left = 208
    Top = 217
  end
  object RxQuery4: TRxQuery
    Macros = <>
    Left = 304
    Top = 217
  end
  object WorkSQL: TQuery
    Left = 208
    Top = 153
  end
end
