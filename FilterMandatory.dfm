�
 TMANDATORYFILTERDLG 0�8  TPF0TMandatoryFilterDlgMandatoryFilterDlgLeft\TopLWidth�Height�BorderStylebsSizeToolWinCaption   $8;LB@Color	clBtnFaceConstraints.MaxHeight Constraints.MaxWidth�Constraints.MinHeight�Constraints.MinWidth�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderPositionpoDesktopCenterOnCreate
FormCreate	OnDestroyFormDestroyOnResize
FormResizeOnShowFormShowPixelsPerInch`
TextHeight TPanel	DownPanelLeft TopNWidth�HeightAlignalBottom
BevelOuterbvNoneTabOrder  TLabelLabel24LeftTopWidthUHeightCaption   Номера полисов  TPanelPanel2Left� Top Width0HeightAlignalRight
BevelOuterbvNoneTabOrder  TBitBtnBitBtn1Left� TopWidthfHeightCaption	   @8<5=8BLTabOrder KindbkOK   TEditListNumbersLeftdTopWidthHeightHint�   Перечисли номера через запятую или тире (не более нескольких сотен полисов)ParentShowHintShowHint	TabOrder   TPanel	DataPanelLeft Top� Width�Height�AlignalBottom
BevelOuterbvNoneTabOrder TLabelLabel1Left� TopJWidthHeightAutoSizeCaption   !Font.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel2LeftBTopJWidthHeightAutoSizeCaption   Font.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel4Left� Top� WidthHeightAutoSizeCaption   !  TLabelLabel5LeftBTop� WidthHeightAutoSizeCaption     TLabelLabel6Left� Top� WidthHeightAutoSizeCaption   !  TLabelLabel7LeftBTop� WidthHeightAutoSizeCaption     TBevelBevel2Left� Top WidthHeight�Shape
bsLeftLine  TLabelLabel10Left� Top5WidthHeightAutoSizeCaption   !  TLabelLabel11LeftBTop5WidthHeightAutoSizeCaption     TLabelLabel12LeftTopVWidthGHeightHint>   Данные проверяются по дате отчётаCaption   или испорчен)Font.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.StylefsUnderline 
ParentFont  TLabelLabel13Left� Top� WidthHeightAutoSizeCaption   !  TLabelLabel14LeftBTop� WidthHeightAutoSizeCaption     TLabelLabel8Left� TopdWidthHeightAutoSizeCaption   !  TLabelLabel9LeftBTopdWidthHeightAutoSizeCaption     TLabelLabel15Left� TopWidthHeightAutoSizeCaption   !  TLabelLabel16LeftBTopWidthHeightAutoSizeCaption     TLabelLabel17LeftTopLWidth� HeightCaptionP   (Может конфликтовать с Квартальным отчётом)WordWrap	  TLabelLabel18Left� Top	WidthHeightAutoSizeCaption   !  TLabelLabel19LeftBTop	WidthHeightAutoSizeCaption     TLabelLabel20Left� TopWidthHeightAutoSizeCaption   !  TLabelLabel21LeftBTopWidthHeightAutoSizeCaption     TLabelLabel22Left� Top#WidthHeightAutoSizeCaption   !  TLabelLabel23LeftBTop#WidthHeightAutoSizeCaption     TLabelLabel25LeftTop�WidthCHeightCaption   >4>BG5B=K9  TLabelLabel26LeftTop�Width!HeightCaption   'C6>9  TBevelBevel1Left Top�Width�Height	Shape	bsTopLine  	TCheckBox	IsPayDateLeftTopHWidth� HeightHint>   Данные проверяются по дате отчётаCaption.   Дата отчёта (1я или 2я датаColorclMenuFont.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.StylefsUnderline ParentColor
ParentFontTabOrder   TDateTimePicker
Pay1Start_Left� TopFWidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder  TDateTimePickerPay1End_LeftXTopFWidth� HeightCalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder  	TCheckBox	IsInsurerLeftTopzWidthaHeightCaption   !B@0E>20B5;LTabOrder  	TComboBoxInsNameModeLeft� TopuWidthaHeightStylecsDropDownList
ItemHeightTabOrderItems.Strings   Начинается С   Заканчивается НА   Содержит   TEditPartNameLeftXTopuWidth� HeightTabOrder  	TCheckBox
IsInsStartLeftTop� Width� HeightCaption#   Начало страхованияTabOrder  TDateTimePicker	InsStart1Left� Top� WidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePicker	InsStart2LeftXTop� Width� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBoxIsInsEndLeftTop� Width� HeightCaption!   Конец страхованияTabOrder	  TDateTimePickerInsEnd1Left� Top� WidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder
  TDateTimePickerInsEnd2LeftXTop� Width� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBoxIsPeriodLeftTop� Width� HeightCaption#   Период страхованияTabOrder  	TComboBoxPeriodLeft� Top� Width	HeightStylecsDropDownListDropDownCount
ItemHeightTabOrder  	TCheckBoxIs2FeeLeftTop� WidthaHeightCaption   2й взносTabOrder  	TComboBox	SecondFeeLeft� Top� Width	HeightStylecsDropDownList
ItemHeightTabOrderItems.Strings   Нет   Есть   	TCheckBoxIsPay2LeftTop� WidthaHeightCaption   2й платёжTabOrder  	TComboBox	SecondPayLeft� Top� Width	HeightStylecsDropDownList
ItemHeightTabOrderItems.Strings   Не был   Был   	TCheckBox	IsPutDateLeftTop3Width� HeightCaption   Дата выдачиTabOrder  TDateTimePickerPutDate1Left� Top1WidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerPutDate2LeftXTop1Width� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBoxIsPayNumberLeftTop� Width� HeightCaption   № 1й платёжкиTabOrder  TRxSpinEdit
PayNumber1Left� Top� WidthaHeightTabOrder  TRxSpinEdit
PayNumber2LeftXTop� Width� HeightTabOrder  	TCheckBoxIsWasActiveLeftTopdWidth� HeightCaption3   Квартальный отчёт за периодFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsUnderline 
ParentFontTabOrder  TDateTimePickerWasActiveFromLeft� Top`WidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerWasActiveToLeftXTop`Width� HeightCalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBox
IsFee2DateLeftTop
Width� HeightCaption   Дата 2го взносаTabOrder  TDateTimePickerFee2DateFromLeft� TopWidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePicker
Fee2DateToLeftYTopWidth� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TListBoxStatesLeft� Top<Width	HeightAHint�   Если вы будете выбирать состояния и держать нажатой
кнопку Ctrl то можно выбрать несколько состоянийColumns
ItemHeightItems.Strings   НОРМАЛЬНЫЙ   ИСПОРЧЕН!   ДОСРОЧНО ЗАВЕРШЁН   УТЕРЯН   ЗАКОНЧИЛСЯ   ПО НЕУПЛАТЕ   ДУБЛИКАТ MultiSelect	ParentShowHintShowHint	TabOrder  	TCheckBoxIsStateLeftTop<WidthaHeightCaption	   !>AB>O=85TabOrder  	TCheckBox	IsRepDateLeftTopWidth� HeightCaption   Дата отчёта 1TabOrder   TDateTimePickerRepDate1Left� TopWidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder!  TDateTimePickerRepDate2LeftWTopWidth� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder"  	TCheckBox
IsRepDate2LeftTopWidth� HeightCaption   Дата отчёта 2TabOrder#  TDateTimePicker	Rep2Date1Left� TopWidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder$  TDateTimePicker	Rep2Date2LeftWTopWidth� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder%  	TComboBox	InsurTypeLefthTopuWidthSHeightStylecsDropDownList
ItemHeightTabOrder&Items.Strings
   Любой
   Физич   Юридич   ИП   	TCheckBoxIsUpDTLeftTop!Width� HeightCaption(   Дата изменения данныхFont.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrder'  TDateTimePickerUpDtFromLeft� TopWidthaHeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder(  TDateTimePickerUpDtToLeftXTopWidth� HeightCalAlignmentdtaLeftDate 8����@Time 8����@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder)  	TCheckBoxIs1SULeftToplWidthaHeightCaption   1СУTabOrder*  	TCheckBox	IsStoppedLeftfTopUWidthXHeightCaption	   @5:@0I5=TabOrder+  	TComboBoxIsReportLeft� Top�Width	HeightStylecsDropDownList
ItemHeightTabOrder,Items.Strings   Не важно   Нет   Да   	TComboBoxIsOtherLeft� Top�Width	HeightStylecsDropDownList
ItemHeightTabOrder-Items.Strings   Не важно   Нет   Да   	TCheckBoxIsAdditionalPayLeftTop�Width� HeightCaption   Доплаты/ВозвратыTabOrder.   TPanelPanel1Left Top Width�Height� AlignalClient
BevelOuterbvNoneTabOrder
DesignSize��   TBevelDividerLeft� Top Width	Height� AnchorsakLeftakTopakBottom Shape
bsLeftLine  TLabelLabel3LeftTopWidth� Height9AutoSizeCaption[   Агенты
(если ничего не помечено, то показать всех)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFontWordWrap	  TLabel
CntCheckedLeftTop@Width� HeightAutoSizeCaption
CntCheckedFont.CharsetDEFAULT_CHARSET
Font.ColorclRedFont.Height�	Font.NameMS Sans Serif
Font.StylefsBoldfsUnderline 
ParentFont  TSpeedButtonSpeedButtonLeft� TopPWidthHeightCaption!Flat	OnClickSpeedButtonClick  TCheckListBox	AgentListLeft� Top Width	Height� OnClickCheckAgentListClickAlignalRightFlat
ItemHeightTabOrder OnClickAgentListClick  TPanelPanel3Left�Top WidthHeight� AlignalRight
BevelOuterbvNoneTabOrder  	TComboBoxListTemplatesLeftTopPWidth� HeightStylecsDropDownList
ItemHeightTabOrderOnChangeListTemplatesChangeItems.Strings   Все агенты    TFormStorageFormStorageFilterIniFileName/   Обязательное страхование
IniSection   $8;LB@UseRegistry	StoredProps.StringsFee2DateFrom.DateFee2DateTo.DateFee2DateFrom.CheckedFee2DateTo.CheckedInsEnd1.DateInsEnd2.DateInsEnd1.CheckedInsEnd2.CheckedInsStart1.DateInsStart2.DateInsStart1.CheckedInsStart2.CheckedRepDate1.DateRepDate2.DateRepDate1.CheckedRepDate2.CheckedRep2Date1.DateRep2Date2.DateRep2Date1.CheckedRep2Date2.CheckedPutDate1.DatePutDate2.DatePutDate1.CheckedPutDate2.Checked StoredValues Left� Top  TFormStorageFormStorageIniFileName
Reporter32
IniSectionMandatoryFilterOptions UseRegistry	StoredProps.StringsWasActiveTo.DateWasActiveFrom.DateInsEnd1.DateInsEnd2.DateInsStart1.DateInsStart2.DatePay1End_.DatePay1Start_.DateRep2Date1.DateRep2Date1.CheckedRep2Date2.DateRep2Date2.CheckedRepDate1.DateRepDate1.CheckedRepDate2.DateRepDate2.CheckedUpDtFrom.DateUpDtFrom.CheckedUpDtTo.DateUpDtTo.Checked StoredValues Left,Top  
TPopupMenu	PopupMenuLeftpTopl 	TMenuItemN1Caption>   Сохранить текущее состояние как...OnClickN1Click  	TMenuItemN2Caption#   Удалить из набора...OnClickN2Click   TQuery	AgentsTblDatabaseNameDBSQL.StringsSELECT    AGENT_CODE, NAMEFROM   AGENTWHERE   MANDPCNT<>-1ORDER BY   NAME Left� Top TStringFieldAgentsTblAGENT_CODE	FieldName
AGENT_CODEOrigin"AGENT.DB".Agent_codeSize  TStringFieldAgentsTblNAME	FieldNameNAMEOrigin"AGENT.DB".NameSize<    