�
 TKORRECTDATA 0_  TPF0TKorrectDataKorrectDataLeft3TopWidthHeightXBorderIconsbiSystemMenu
biMaximize Caption5   Корректировка данных по ОСТСColor	clBtnFaceConstraints.MinHeight� Constraints.MinWidthFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	FormStylefsStayOnTopOldCreateOrderVisible	OnClose	FormCloseOnCreate
FormCreatePixelsPerInch`
TextHeight 	TSplitterSplitterLeft�TopQWidthHeight�CursorcrHSplitMinSize�   TPanelPanel2Left Top WidthHeightQAlignalTop
BevelOuterbvNoneCtl3D	ParentCtl3DParentShowHintShowHint	TabOrder  TLabelLabel5LeftTopWidthqHeightAutoSizeCaption#   Что корректировать  TLabelLabel6LeftTop8Width<HeightCaption
   !>@B8@>2:0  TLabelLabel7Left	Top WidthHeightCaption   Где?  TLabelLabel1Left�Top
WidthHeightCaption     TLabelLabel2Left� Top Width%HeightCaption   5@A8O  	TComboBoxClassFldLeftxTopWidthyHeightStylecsDropDownList
ItemHeightTabOrder OnChangeClassFldChangeItems.Strings   ФИО   Адреса!   Марки автомобилей#   Номера автомобилей   Номер кузова   Номер шасси   	TComboBox	SortComboLeftxTop3WidthyHeightStylecsDropDownList
ItemHeightTabOrderOnChangeSortComboChangeItems.Strings   Номер полиса   Возрастание   Убывание   TButtonButton2Left�TopWidthjHeightCaption   ><0=4K	PopupMenuCommandsMenuTabOrderOnClickButton1Click  	TComboBox	TableTypeLeftxTopWidthyHeightStylecsDropDownList
ItemHeightTabOrderOnChangeSortComboChangeItems.Strings
   Везде   В полисах$   В других владельцах   В выплатах   TEditNmbFromLeft@TopWidthAHeightTabOrder  TEditNmbToLeft�TopWidthIHeightTabOrder  	TComboBoxVerLeft@TopWidth� HeightStylecsDropDownList
ItemHeightTabOrderOnChangeClassFldChangeItems.Strings2   Все   	TCheckBoxIsPeriodLeft� Top5WidthAHeightCaption   5@8>4TabOrderOnClickClassFldChange  	TComboBoxMonthesLeft@Top2WidthiHeightStylecsDropDownListDropDownCount
ItemHeightTabOrderOnChange	CheckDataItems.Strings   (1) Январь   (2) Февраль   (3) Март   (4) Апрель
   (5) Май   (6) Июнь   (7) Июль   (8) Август   (9) Сентябрь   (10) Октябрь   (11) Ноябрь   (12) Декабрь   TRxSpinEditYearLeft�Top2WidthAHeightMaxValue      0�
@MinValue       �	@Value       �	@TabOrder	OnChange	CheckData  	TCheckBoxIsNumberLeft� TopWidthIHeightCaption   &Номера CTabOrder
OnClickIsNumberClick  TButtonSaveChLeft�TopWidthjHeightCaption   Сохранить...EnabledTabOrderOnClickSaveChClick  TRxSpinEdit
FindNumberLeft�Top3WidthjHeightMaxValue     ��@MinValue       ��?Value       ��?ColorclMenuTabOrderOnChangeFindNumberChange   	TRxDBGridRxDBGridLeft TopQWidth�Height�AlignalLeft
DataSourceGetStringsSrcOptions	dgEditingdgTitlesdgIndicatordgColumnResize
dgColLines
dgRowLinesdgCancelOnExit TabOrderTitleFont.CharsetDEFAULT_CHARSETTitleFont.ColorclWindowTextTitleFont.Height�TitleFont.NameMS Sans SerifTitleFont.Style OnGetCellParamsRxDBGridGetCellParamsColumnsExpanded	FieldNameSERIAReadOnly	Visible	 Expanded	FieldNameNUMBERReadOnly	Visible	 Expanded	FieldNameSTRWidth� Visible	 Expanded	FieldNameTBLReadOnly	Visible	    TMemoOutWndLeft�TopQWidthHeight�AlignalClient
ScrollBars
ssVerticalTabOrder  TQuery
GetStringsCachedUpdates	BeforeInsertGetStringsBeforeInsert
BeforeEditGetStringsBeforeEdit	AfterEditGetStringsAfterEditBeforeDeleteGetStringsBeforeDeleteOnUpdateRecordGetStringsUpdateRecordDatabaseNameDBRequestLive	SQL.Strings9SELECT SERIA,NUMBER,SERIA AS STR, '' AS TBL FROM MANDATOR Left Top�  TStringFieldGetStringsSERIADisplayLabel   !5@8O	FieldNameSERIAOrigin"MANDATOR.DB".SeriaSize  TFloatFieldGetStringsNUMBERDisplayLabel   ><5@	FieldNameNUMBEROrigin"MANDATOR.DB".Number  TStringFieldGetStringsSTRDisplayLabel<None>	FieldNameSTROrigin"MANDATOR.DB".SeriaSize2  TStringFieldGetStringsTBLDisplayLabel   "01;8F0	FieldNameTBLSize   TDataSourceGetStringsSrcDataSet
GetStringsLefthTop�   
TPopupMenuCommandsMenuOnPopupCommandsMenuPopupLeft� Top�  	TMenuItemN1Caption.   1. Всё к верхнему региструEnabledOnClickN1Click  	TMenuItemN2CaptionG   2. Удалить лишние пробелы и спецсимволыEnabledOnClickN2Click  	TMenuItem
BadManMenuCaption9   Операция 1, 2  для аварий (be carefully)EnabledVisibleOnClickBadManMenuClick  	TMenuItemN3CaptionG   3. Перевести английские буквы в русскиеEnabledOnClickN3Click  	TMenuItemN4CaptionL   4. Заменить последовательность на другую..OnClickN4Click  	TMenuItemN5Caption'   5. Переместить слово...OnClickN5Click  	TMenuItemDivider1Caption-Visible  	TMenuItem	ChAddressCaption?VisibleOnClickChAddressClick   TTableChTableDatabaseNameDBFiltered	Left Top�   
TUpdateSQL	UpdateSQLLeft Top�   TQueryWorkSQLDatabaseNameDBLeft@Top�    