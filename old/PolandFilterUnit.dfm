�
 TPOLANDFILTER 0�  TPF0TPolandFilterPolandFilterLeftPTop� Width�Height�BorderIconsbiSystemMenu Caption   $8;LB@Color	clBtnFaceConstraints.MaxWidth�Constraints.MinWidth�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderPositionpoScreenCenterOnCreate
FormCreateOnShowFormShow
DesignSize�� PixelsPerInch`
TextHeight TLabelLabel1Left� Top� WidthHeightAnchorsakLeftakBottom Caption   !  TLabelLabel2Left0Top� WidthHeightAnchorsakLeftakBottom Caption     TLabelLabel3Left� Top� WidthHeightAnchorsakLeftakBottom Caption   !  TLabelLabel4Left0Top� WidthHeightAnchorsakLeftakBottom Caption     TLabelLabel5Left� Top� WidthHeightAnchorsakLeftakBottom Caption   !  TLabelLabel6Left0Top� WidthHeightAnchorsakLeftakBottom Caption     TBevelBevel1Left ToplWidth�HeightAnchorsakLeftakBottom   TLabelLabel8Left� Top7Width	HeightAnchorsakLeftakBottom AutoSizeCaption   !  TLabelLabel9Left8Top7WidthHeightAnchorsakLeftakBottom AutoSizeCaption     TLabelLabel7Left� Top� WidthHeightAnchorsakLeftakBottom Caption   !  TLabelLabel10Left0Top� WidthHeightAnchorsakLeftakBottom Caption     TLabelLabel11LeftTopWidth� Height9AutoSizeCaption[   Агенты
(если ничего не помечено, то показать всех)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFontWordWrap	  TLabel
CntCheckedLeftTop@Width� HeightAutoSizeCaption
CntCheckedFont.CharsetDEFAULT_CHARSET
Font.ColorclRedFont.Height�	Font.NameMS Sans Serif
Font.StylefsBoldfsUnderline 
ParentFont  TSpeedButtonSpeedButtonLeft� TopPWidthHeightCaption!Flat	OnClickSpeedButtonClick  	TCheckBoxIsNumberLeftTop� WidthiHeightAnchorsakLeftakBottom Caption   ><5@TabOrder   	TCheckBox	IsRegDateLeftTop� WidthyHeightAnchorsakLeftakBottom Caption   Дата регистрацииTabOrder  TEdit
NumberFromLeft� Top� WidthaHeightAnchorsakLeftakBottom TabOrder  TEditNumberToLeftHTop� WidthaHeightAnchorsakLeftakBottom TabOrder  TDateTimePickerRegDateFromLeft� Top� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePicker	RegDateToLeftHTop� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBoxIsStartDateLeftTop� Width� HeightAnchorsakLeftakBottom Caption,   Дата начала страхованияTabOrder  TDateTimePickerStartDateFromLeft� Top� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerStartDateToLeftHTop� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TButtonButton1Left@ToptWidthkHeightAnchorsakLeftakBottom Caption   B<5=8BLModalResultTabOrder	  TButtonButton2Left� ToptWidthlHeightAnchorsakLeftakBottom Caption	   @8<5=8BLModalResultTabOrder
  	TCheckBoxIsOwnerLeftTopWidthaHeightAnchorsakLeftakBottom Caption   ;045;5FTabOrder  TEdit	OwnerLikeLeft� TopWidth� HeightAnchorsakLeftakBottom TabOrder  	TCheckBoxIsAutoNumberLeftTopWidth� HeightAnchorsakLeftakBottom Caption   Номер автомобиляTabOrder  TEditAutoNumberLikeLeft� TopWidth� HeightAnchorsakLeftakBottom TabOrder  	TCheckBoxIsWasActiveLeftTop7Width� HeightAnchorsakLeftakBottom Caption3   Квартальный отчёт за периодFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsUnderline 
ParentFontTabOrder  TDateTimePickerWasActiveFromLeft� Top3WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerWasActiveToLeftKTop3Width`HeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate 8����@Time 8����@
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  	TCheckBoxIsOformDateLeftTop� WidthyHeightAnchorsakLeftakBottom Caption   Дата платежаTabOrder  TDateTimePickerOformDateFromLeft� Top� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TDateTimePickerOformDateToLeftHTop� WidthaHeightAnchorsakLeftakBottom CalAlignmentdtaLeftDate h��ŏ@Time h��ŏ@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TCheckListBox	AgentListLeft� TopWidth� Height� OnClickCheckAgentListClickAnchorsakLeftakTopakBottom Flat
ItemHeightTabOrderOnClickAgentListClick  	TComboBoxListTemplatesLeftTopPWidth� HeightStylecsDropDownList
ItemHeightTabOrderOnChangeListTemplatesChangeItems.Strings   Все агенты   	TCheckBoxIsStateLeftTopMWidth� HeightAnchorsakLeftakBottom Caption	   !>AB>O=85Font.CharsetDEFAULT_CHARSET
Font.ColorclNavyFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrder  	TComboBoxState_Left� TopMWidth� HeightStylecsDropDownListAnchorsakLeftakBottom 
ItemHeightTabOrderItems.Strings   НОРМАЛЬНЫЙ   УТЕРЯН   ИСПОРЧЕН   РАСТОРГНУТ   
TPopupMenu	PopupMenuLeft� Top 	TMenuItemN1Caption>   Сохранить текущее состояние как...OnClickN1Click  	TMenuItemN2Caption#   Удалить из набора...OnClickN2Click   TQuery	AgentsTblDatabaseNameDBPOLANDSQL.StringsSELECT    AGENT_CODE, NAMEFROM   AGENTWHERE   POLANDPCNT IS NOT NULLOR   POLANDPCNT >= 0ORDER BY   NAME Left� Top( TStringFieldAgentsTblAGENT_CODE	FieldName
AGENT_CODEOrigin"AGENT.DB".Agent_codeSize  TStringFieldAgentsTblNAME	FieldNameNAMEOrigin"AGENT.DB".NameSize<   TFormStorageFormStorageIniFileName
Reporter32
IniSectionPOLANDUseRegistry	StoredProps.StringsOformDateFrom.DateOformDateTo.DateWasActiveFrom.DateRegDateTo.DateRegDateFrom.DateWasActiveTo.DateOformDateTo.CheckedOformDateFrom.CheckedWasActiveFrom.CheckedWasActiveTo.CheckedRegDateFrom.CheckedRegDateTo.Checked StoredValues Left(Top(   