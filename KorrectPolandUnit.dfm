�
 TKORRECTPOLAND 0�  TPF0TKorrectPolandKorrectPolandLeftTop� BorderIconsbiSystemMenu BorderStylebsDialogCaption'   Корректировка данныхClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderPositionpoScreenCenterOnCreate
FormCreateOnShowFormShowPixelsPerInch`
TextHeight TButtonButton1Left8Top� WidthKHeightCaption   0:@KBLModalResultTabOrder   TButtonButton2Left� Top� WidthKHeightCaption	   K?>;=8BLTabOrderOnClickButton2Click  	TGroupBox	GroupBox1LeftTopWidth�HeightACaption!     Изменяемые поля  TabOrder 	TCheckBoxIsSetAgPercentLeftTopWidthaHeightCaption   Процент агентуTabOrder   	TCheckBoxIsSetUrFizichLeftTop(Width� HeightCaption+   Юридическое/ФизическоеTabOrder  TRxSpinEdit	AgPercentLeft� TopWidth� HeightMaxValue       �@	ValueTypevtFloatTabOrder  	TComboBox
IsUrFizichLeft� Top%Width� HeightStylecsDropDownList
ItemHeightTabOrderItems.Strings   Физическое   Юридическое    	TGroupBox	GroupBox2LeftTopPWidth�Height� Caption:     В каких полисах изменять поля  TabOrder TLabelLabel1LeftTop*WidthHeightCaption     TLabelLabel2LeftTopBWidthHeightCaption     	TCheckBoxIsFiltSeriaLeftTopWidthaHeightCaption   !5@8OTabOrder   	TCheckBoxIsFiltNumberLeftTop(Width� HeightCaption/   Номер                                   СTabOrder  	TCheckBoxIsFiltAgentLeftTopXWidthaHeightCaption   35=BTabOrder  	TCheckBox
IsFiltNullLeftToppWidth)HeightCaptionP   Изменять только если значение неопределеноEnabledTabOrder  TEdit	NumberMinLeft� Top&WidthYHeightTabOrder  TEdit	NumberMaxLeft Top&WidthYHeightTabOrder  TRxDBLookupComboSeriesComboLeft� TopWidth� HeightDropDownCountLookupFieldSERIALookupDisplaySERIALookupSourceDataSourceSeriesTabOrder  TRxDBLookupCombo
AgentComboLeft� TopTWidth� HeightDropDownCountLookupField
Agent_codeLookupDisplayNameLookupSourceDataSourceAgTabOrder  	TCheckBox	IsRegDateLeftTop@Width� HeightCaption/   Дата выдачи                        СTabOrder  TDateTimePickerRegDateFromLeft� Top=WidthYHeightCalAlignmentdtaLeftDate 0|�$jH�@Time 0|�$jH�@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder	  TDateTimePicker	RegDateToLeft Top=WidthYHeightCalAlignmentdtaLeftDate `�P;jH�@Time `�P;jH�@ShowCheckbox	
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder
   TDataSourceDataSourceAgDataSetAgTableLeft8Top�   TDataSourceDataSourceSeriesDataSetSeriesLeft0Top@  TQueryWorkSQLLeftTop�   TQueryAgTableSQL.StringsSELECT *
FROM AGENTORDER BY NAME Left� Top�  TStringFieldAgTableAgent_code	FieldName
Agent_codeOrigin"Agent.DB".Agent_codeSize  TStringFieldAgTableName	FieldNameNameOrigin"Agent.DB".NameSize<   TRxQuerySeriesSQL.StringsSELECT  DISTINCT SERIAFROM %TABLE MacrosDataTypeftStringNameTABLE	ParamTypeptInputValue0=0  Left� TopH   