object MainMenu: TMainMenu
  Left = 335
  Top = 249
  BorderStyle = bsNone
  ClientHeight = 256
  ClientWidth = 236
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 0
    Top = 0
    Width = 237
    Height = 22
    Caption = 'Обязательное стразование ТС'
    Flat = True
    OnClick = SpeedButton1Click
  end
  object SpeedButton2: TSpeedButton
    Left = 0
    Top = 20
    Width = 237
    Height = 22
    Caption = 'ПОЛЬША'
    Flat = True
    OnClick = SpeedButton2Click
  end
  object SpeedButton3: TSpeedButton
    Left = 0
    Top = 40
    Width = 237
    Height = 22
    Caption = 'Обязательное стразование ТС'
    Flat = True
  end
  object SpeedButton4: TSpeedButton
    Left = 0
    Top = 60
    Width = 237
    Height = 22
    Caption = 'Обязательное стразование ТС'
    Flat = True
  end
  object SpeedButton5: TSpeedButton
    Left = 0
    Top = 80
    Width = 237
    Height = 22
    Caption = 'Обязательное стразование ТС'
    Flat = True
  end
end
