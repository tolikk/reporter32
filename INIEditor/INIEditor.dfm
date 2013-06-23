object MainForm: TMainForm
  Left = 271
  Top = 275
  Width = 601
  Height = 296
  Caption = #1056#1077#1076#1072#1082#1090#1086#1088' blank.ini'
  Color = clBtnFace
  Constraints.MinHeight = 200
  Constraints.MinWidth = 450
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object DownPanel: TPanel
    Left = 0
    Top = 233
    Width = 593
    Height = 29
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      593
      29)
    object Label1: TLabel
      Left = 5
      Top = 9
      Width = 315
      Height = 13
      Caption = 'F2 - '#1048#1079#1084#1077#1085#1080#1090#1100'   Enter - '#1047#1072#1082#1086#1085#1095#1080#1090#1100'   Esc - '#1086#1090#1084#1077#1085#1072'   Del - '#1091#1076#1072#1083#1080#1090#1100
    end
    object Button1: TButton
      Left = 462
      Top = 2
      Width = 131
      Height = 25
      Anchors = [akTop, akRight]
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 0
      OnClick = Button1Click
    end
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 593
    Height = 233
    ActivePage = TabSheet1
    Align = alClient
    TabIndex = 0
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = #1052#1072#1088#1082#1080' '#1072#1074#1090#1086
      OnResize = TabSheet1Resize
      object SplitterAutoMark: TSplitter
        Left = 377
        Top = 0
        Width = 5
        Height = 205
        Cursor = crHSplit
      end
      object AutoMarkGrid: TStringGrid
        Left = 0
        Top = 0
        Width = 377
        Height = 205
        Align = alLeft
        Color = clInfoBk
        ColCount = 2
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 1
        FixedRows = 0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goEditing]
        ParentFont = False
        TabOrder = 0
        OnGetEditText = AutoMarkGridGetEditText
        OnKeyDown = AutoMarkGridKeyDown
        OnSelectCell = AutoMarkGridSelectCell
        OnSetEditText = AutoMarkGridSetEditText
      end
      object AutoSubMarkGrid: TStringGrid
        Left = 382
        Top = 0
        Width = 203
        Height = 205
        Align = alClient
        Color = clInfoBk
        ColCount = 1
        DefaultRowHeight = 18
        FixedCols = 0
        RowCount = 1
        FixedRows = 0
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing]
        ParentFont = False
        TabOrder = 1
        OnGetEditText = AutoSubMarkGridGetEditText
        OnKeyDown = AutoSubMarkGridKeyDown
        OnSetEditText = AutoSubMarkGridSetEditText
      end
    end
  end
end
