�
 TLEBENFILTER 0?  TPF0TLebenFilterLebenFilterLeftTopbWidth�Height)BorderIconsbiSystemMenu Caption!   Стразование жизниColor	clBtnFaceConstraints.MaxWidth�Constraints.MinHeight� Constraints.MinWidth�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderPositionpoDesktopCenterOnCreate
FormCreate	OnDestroyFormDestroyOnShowFormShow
DesignSize� PixelsPerInch`
TextHeight TLabelLabel2LeftTop� Width� HeightAnchorsakBottom Caption/   Дата платежа                       C  TLabelLabel3LeftTop� WidthHeightAnchorsakBottom Caption     TLabel
CntCheckedLeftTop Width� HeightAutoSizeCaption   Выбрано 0Font.CharsetDEFAULT_CHARSET
Font.ColorclRedFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabel LeftTopWidth� HeightAnchorsakBottom Caption0   Начало страхования            C  TLabelLabel4LeftTopWidthHeightAnchorsakBottom Caption     TLabelLabel5LeftTop+Width� HeightAnchorsakBottom Caption0   Конец страхования              C  TLabelLabel6LeftTop*WidthHeightAnchorsakBottom Caption     TLabelLabel1LeftTopCWidth� HeightAnchorsakBottom Caption0   Продолжительность            С  TLabelLabel7LeftTopBWidthHeightAnchorsakBottom Caption     TSpeedButtonSpeedButtonLeft� Top8WidthHeightCaption!Flat	OnClickSpeedButtonClick  TLabelLabel8LeftTop\Width6HeightAnchorsakBottom Caption	   AA8AB0=A  TLabelLabel9LeftTop�Width� HeightAnchorsakBottom Caption/   Дата расторжения               C  TLabelLabel10LeftTop�WidthHeightAnchorsakBottom Caption     TDateTimePickerDateFromLeft� Top� WidthiHeightAnchorsakBottom CalAlignmentdtaLeftDate �Ы��@Time �Ы��@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder   TDateTimePickerDateToLeft0Top� Width� HeightAnchorsakBottom CalAlignmentdtaLeftDate ��e���@Time ��e���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TCheckListBox	AgentListLeft� TopWidthHeight� OnClickCheckAgentListClickCheckAnchorsakLeftakTopakBottom ColorclMoneyGreenFlatFont.CharsetDEFAULT_CHARSET
Font.ColorclNavyFont.Height�	Font.NameMS Sans Serif
Font.Style 
ItemHeight
ParentFontTabOrderOnClickAgentListClick  TPanelPanel1Left Top�Width�HeightAlignalBottom
BevelOuterbvNoneTabOrder TBitBtnBitBtn1LeftTopWidth� HeightCaption	   @8<5=8BLModalResultTabOrder   	TCheckBox	IsShowBadLeftTopWidth� HeightCaption+   Показывать испорченныеTabOrder   	TCheckBoxIsAgentLeftTopWidthQHeightCaption   35=BKTabOrder  TDateTimePicker	StartFromLeft� TopWidthiHeightAnchorsakBottom CalAlignmentdtaLeftDate �Ы��@Time �Ы��@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerStartToLeft0TopWidth� HeightAnchorsakBottom CalAlignmentdtaLeftDate ��e���@Time ��e���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerEndFromLeft� Top'WidthiHeightAnchorsakBottom CalAlignmentdtaLeftDate �Ы��@Time �Ы��@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerEndToLeft0Top'Width� HeightAnchorsakBottom CalAlignmentdtaLeftDate ��e���@Time ��e���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TEditDurationsFromLeft� Top?WidthhHeightAnchorsakBottom TabOrder	  TEditDurationsToLeft0Top?Width� HeightAnchorsakBottom TabOrder
  	TComboBoxListTemplatesLeftTop8WidthyHeightStylecsDropDownList
ItemHeightTabOrderOnChangeListTemplatesChangeItems.Strings   Все агенты   TCheckListBox
ListAssistLeft� TopUWidthHeight+AnchorsakLeftakBottom Color�ʦ 
ItemHeightTabOrder  	TCheckBoxIsStateLeftTop�Width� HeightAnchorsakLeftakBottom Caption	   !>AB>O=85Font.CharsetDEFAULT_CHARSET
Font.ColorclNavyFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrder  	TComboBoxStateLeft� Top�WidthHeightStylecsDropDownListAnchorsakLeftakRightakBottom 
ItemHeightTabOrderItems.Strings   НОРМАЛЬНЫЙ   ИСПОРЧЕН   РАСТОРГНУТ   	TCheckBoxIsHaveUbitokLeftTop�Width� HeightAnchorsakLeftakBottom Caption'   Есть убытки по полисуTabOrder  	TComboBoxlistInsTypeLeft� Top�WidthHeightStylecsDropDownListAnchorsakLeftakRightakBottom 
ItemHeightTabOrderItems.Strings   Физ лицо   Юр лицо   	TCheckBoxIsInsurerTypeLeftTop�Width� HeightAnchorsakLeftakBottom Caption   Тип страхователяFont.CharsetDEFAULT_CHARSET
Font.ColorclNavyFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrder  TDateTimePickerStopDateFromLeft� Top�WidthiHeightAnchorsakBottom CalAlignmentdtaLeftDate �Ы��@Time �Ы��@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePicker
StopDateToLeft0Top�Width� HeightAnchorsakBottom CalAlignmentdtaLeftDate ��e���@Time ��e���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TQuery	AgentsTblDatabaseNameLEBEN2SQL.StringsSELECT     AGENT_CODE,    NAMEFROM 	    AGENTWHERE    LEBEN2PCNT<>-1ORDER BY    NAME Left� Top TStringFieldAgentsTblAGENT_CODE	FieldName
AGENT_CODEOrigin"Agent.DB".Agent_codeSize  TStringFieldAgentsTblNAME	FieldNameNAMEOrigin"Agent.DB".NameSize<   TFormStorageFormStorageIniFileName!   Страхование жизни
IniSection   $8;LB@Options UseRegistry	StoredProps.StringsDateFrom.DateDateTo.DateIsAgent.CheckedStartFrom.DateStartTo.DateEndFrom.Date
EndTo.DateDateFrom.CheckedDateTo.CheckedEndFrom.CheckedEndTo.CheckedStartFrom.CheckedStartTo.Checked StoredValues Left� Top  
TPopupMenu	PopupMenuLeft`TopX 	TMenuItemN1Caption>   Сохранить текущее состояние как...OnClickN1Click  	TMenuItemN2Caption#   Удалить из набора...OnClickN2Click    