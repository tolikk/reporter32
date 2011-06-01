object DecoderFrm: TDecoderFrm
  Left = 263
  Top = 152
  BorderStyle = bsDialog
  Caption = 'Перекодировка'
  ClientHeight = 102
  ClientWidth = 123
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Kode: TEdit
    Left = 0
    Top = 0
    Width = 121
    Height = 21
    TabOrder = 0
    OnChange = KodeChange
  end
  object Result: TEdit
    Left = 0
    Top = 24
    Width = 121
    Height = 21
    ReadOnly = True
    TabOrder = 1
  end
  object Button1: TButton
    Left = 48
    Top = 76
    Width = 75
    Height = 22
    Caption = 'Закрыть'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Kode2: TEdit
    Left = 0
    Top = 48
    Width = 121
    Height = 21
    ReadOnly = True
    TabOrder = 3
  end
end
