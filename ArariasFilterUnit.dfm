�
 TLISTAVARIAS 0�2  TPF0TListAvariasListAvariasLeft2Top� Width�Height�BorderIconsbiSystemMenu Caption   20@88Color	clBtnFaceConstraints.MinHeight�Constraints.MinWidth�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderOnClose	FormCloseOnCreate
FormCreate	OnDestroyFormDestroyPixelsPerInch`
TextHeight TPanelPanel1Left Top.Width�HeightGAlignalBottom
BevelOuterbvNoneTabOrder  TSpeedButtonApplyFilterBtnLeftTopWidthHeight9Flat	
Glyph.Data
z  v  BMv      v   (       @                                �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� ����������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������OnClickApplyFilterBtnClick  	TGroupBox	GroupBox1LeftTopWidthVHeight@Caption    Фильтр Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFontTabOrder  TLabelLabel2LeftTopWidthuHeightCaption&   Дата под. заявления СFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel4LeftTop(WidthZHeightCaption   Суммы за периодFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel3Left� TopWidthHeightCaption   Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel5Left� Top(WidthHeightCaption   Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TDateTimePickerRegDateFromLeftTopWidth\HeightCalAlignmentdtaLeftDate �yG���@Time �yG���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder   TDateTimePicker
PeriodFromLeftTop$Width\HeightCalAlignmentdtaLeftDate �yG���@Time �yG���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder  TDateTimePicker	RegDateToLeft� TopWidthaHeightCalAlignmentdtaLeftDate �yG���@Time �yG���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder  TDateTimePickerPeriodToLeft� Top$WidthaHeightCalAlignmentdtaLeftDate �yG���@Time �yG���@ShowCheckbox	Checked
DateFormatdfShortDateMode
dmComboBoxFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style KinddtkDate
ParseInput
ParentFontTabOrder   TPanelPanel2LeftGTop WidthmHeightGAlignalRight
BevelOuterbvNoneTabOrder TButtonButton1Left Top0WidthiHeightCaption   0:@KBLTabOrder OnClickButton1Click  TButtonButton2Left TopWidthiHeightCaption   В ExcelTabOrderOnClickButton2Click  TButtonButton3Left TopWidthiHeightCaption   Закрыть делаTabOrderOnClickButton3Click   	TComboBox	StateListLeft[Top*WidthtHeightStylecsDropDownList
ItemHeightTabOrderOnChangeStateListChangeItems.Strings    Не урегулирован   Урегулирован   	TCheckBoxIsShowPaysDetailLeftZTopWidthHHeightCaption   K?;0BKTabOrder  	TCheckBoxIsZayavLeft�TopWidthaHeightCaption   + ЗаявленныеTabOrder   	TRxDBGrid
AvariaGridLeft Top Width�Height.AlignalClientBorderStylebsNone
DataSource
DataSourceOptionsdgTitlesdgIndicatordgColumnResize
dgColLines
dgRowLinesdgCancelOnExitdgMultiSelect TabOrderTitleFont.CharsetDEFAULT_CHARSETTitleFont.ColorclWindowTextTitleFont.Height�TitleFont.NameMS Sans SerifTitleFont.Style 	FixedColsMultiSelect	OnGetCellPropsAvariaGridGetCellPropsColumnsExpanded	FieldNameSERIAWidth)Visible	 Expanded	FieldNameNUMBERWidthAVisible	 Expanded	FieldNameNWidthVisible	 Expanded	FieldName	StateNameVisible	 Expanded	FieldNameOutSummaVisible	 Expanded	FieldName	DeclSummaVisible	 Expanded	FieldNamePayInfoVisible	 Expanded	FieldNameAVDATEVisible	 Expanded	FieldNameWRDATEVisible	 Expanded	FieldNameDtClLossVisible	 Expanded	FieldNameNWORKVisible	 Expanded	FieldNameFIOWidth� Visible	 Expanded	FieldNameBADMANWidth~Visible	 Expanded	FieldName	AvPayTypeVisible	 Expanded	FieldNameDVisible	 Expanded	FieldNameSUMMAVisible	 Expanded	FieldNameCURRVisible	 Expanded	FieldNameZATRSTRVisible	 Expanded	FieldNameRISKSTRVisible	    TPanelMsgPanelLeftTopWidthPHeightTabOrder TLabelLabel1LeftTopWidthHeightCaptionI   Заполни фильтр и нажми мограющую кнопкуFont.CharsetDEFAULT_CHARSET
Font.ColorclBlueFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   TRxQueryRxAvariaQueryOnCalcFieldsRxAvariaQueryCalcFieldsDatabaseNameDBSQL.StringsSELECT 	   SERIA,
   NUMBER,   A.N,
   AVDATE,
   WRDATE,	   NWORK,   FIO,
   BADMAN,   ISTC1,OCENTC,   ISTC2,OTHTC,   ISHL1,OCENHL,   ISHL2,OTHHL,   ISIM1,OCENIM,   ISIM2,OTHIM,
   SUMPAY,   Sum1_1, Sum1_1V,   Sum2_1, Sum2_1V,   Sum3_1, Sum3_1V,   V1_Date,V2_Date,V3_Date,   What,    V1,    DtClLoss,    UpdateDate,
   %FIELDSFROM   MANDAV2 A%WHEREORDER BY	   SERIA,
   NUMBER,   A.N MacrosDataTypeftStringNameFIELDS	ParamTypeptInputValue1 DataTypeftStringNameWHERE	ParamTypeptInput  Left Topp TStringFieldRxAvariaQuerySERIA	AlignmenttaCenterDisplayLabel   !5@8O	FieldNameSERIAOrigin"MANDAV2.DB".SeriaSize  TFloatFieldRxAvariaQueryNUMBERDisplayLabel   ><5@	FieldNameNUMBEROrigin"MANDAV2.DB".Number  TFloatFieldRxAvariaQueryN	FieldNameNOrigin"MANDAV2.DB".N  
TDateFieldRxAvariaQueryAVDATE	AlignmenttaCenterDisplayLabel   Дата аварии	FieldNameAVDATEOrigin"MANDAV2.DB".AvDate  
TDateFieldRxAvariaQueryWRDATE	AlignmenttaCenterDisplayLabel   Подача заявления	FieldNameWRDATEOrigin"MANDAV2.DB".WrDate  TStringFieldRxAvariaQueryNWORKDisplayLabel   !  45;0	FieldNameNWORKOrigin"MANDAV2.DB".NWorkSize
  TStringFieldRxAvariaQueryFIODisplayLabel   >B5@?52H89	FieldNameFIOOrigin"MANDAV2.DB".FIOSize2  TStringFieldRxAvariaQueryBADMANDisplayLabel   8=>2=8:	FieldNameBADMANOrigin"MANDAV2.DB".BadManSize2  TStringFieldRxAvariaQueryISTC1	FieldNameISTC1Origin"MANDAV2.DB".IsTC1VisibleSize  TFloatFieldRxAvariaQueryOCENTCDisplayLabel   Оценка(ТС)	FieldNameOCENTCOrigin"MANDAV2.DB".OcenTCVisible  TStringFieldRxAvariaQueryISTC2	FieldNameISTC2Origin"MANDAV2.DB".IsTC2VisibleSize  TFloatFieldRxAvariaQueryOTHTCDisplayLabel   Проч. расходы(ТС)	FieldNameOTHTCOrigin"MANDAV2.DB".OthTCVisible  TStringFieldRxAvariaQueryISHL1	FieldNameISHL1Origin"MANDAV2.DB".IsHL1VisibleSize  TFloatFieldRxAvariaQueryOCENHLDisplayLabel   Оценка (Жизнь)	FieldNameOCENHLOrigin"MANDAV2.DB".OcenHLVisible  TStringFieldRxAvariaQueryISHL2	FieldNameISHL2Origin"MANDAV2.DB".IsHL2VisibleSize  TFloatFieldRxAvariaQueryOTHHLDisplayLabel$   Проч. расходы(Жизнь)	FieldNameOTHHLOrigin"MANDAV2.DB".OthHLVisible  TStringFieldRxAvariaQueryISIM1	FieldNameISIM1Origin"MANDAV2.DB".IsIM1VisibleSize  TFloatFieldRxAvariaQueryOCENIMDisplayLabel!   Оценка (Имущество)	FieldNameOCENIMOrigin"MANDAV2.DB".OcenImVisible  TStringFieldRxAvariaQueryISIM2	FieldNameISIM2Origin"MANDAV2.DB".IsIM2VisibleSize  TFloatFieldRxAvariaQueryOTHIMDisplayLabel,   Проч. расходы(Имущество)	FieldNameOTHIMOrigin"MANDAV2.DB".OthImVisible  TFloatFieldRxAvariaQuerySum1_1DisplayLabel    Возмещение (ТС), EUR	FieldNameSum1_1Origin"MANDAV2.DB".Sum1_1Visible  TFloatFieldRxAvariaQuerySum1_1VDisplayLabel#   Возмещение (ТС), РУБ	FieldNameSum1_1VOrigin"MANDAV2.DB".Sum1_1vVisible  TFloatFieldRxAvariaQuerySum2_1DisplayLabel"   Возмещение (Жзн), EUR	FieldNameSum2_1Origin"MANDAV2.DB".Sum2_1Visible  TFloatFieldRxAvariaQuerySum2_1VDisplayLabel%   Возмещение (Жзн), РУБ	FieldNameSum2_1VOrigin"MANDAV2.DB".Sum2_1VVisible  TFloatFieldRxAvariaQuerySum3_1DisplayLabel    Возмещение (Им), EUR	FieldNameSum3_1Origin"MANDAV2.DB".Sum3_1Visible  TFloatFieldRxAvariaQuerySum3_1VDisplayLabel#   Возмещение (Им), РУБ	FieldNameSum3_1VOrigin"MANDAV2.DB".Sum3_1VVisible  TFloatFieldRxAvariaQuerySUMPAYDisplayLabel	   K?;0G5=>	FieldNameSUMPAYOrigin"MANDAV2.DB".SumPayVisible  
TDateFieldRxAvariaQueryV1_Date	FieldNameV1_DateOrigin"MANDAV2.DB".V1_DateVisible  
TDateFieldRxAvariaQueryV2_Date	FieldNameV2_DateOrigin"MANDAV2.DB".V2_DateVisible  
TDateFieldRxAvariaQueryV3_Date	FieldNameV3_DateOrigin"MANDAV2.DB".V3_DateVisible  TStringFieldRxAvariaQueryPayDatesDisplayLabel   Даты выплат	FieldKindfkCalculated	FieldNamePayDatesVisibleSize(
Calculated	  TFloatFieldRxAvariaQueryOutSummaDisplayLabel   Выпаты, руб	FieldKindfkCalculated	FieldNameOutSumma
Calculated	  TFloatFieldRxAvariaQueryDeclSummaDisplayLabel   Заявлено, руб	FieldKindfkCalculated	FieldName	DeclSumma
Calculated	  TStringFieldRxAvariaQueryPayInfoDisplayLabel   Прочие платежиDisplayWidth2	FieldKindfkCalculated	FieldNamePayInfoSize� 
Calculated	  TStringFieldRxAvariaQueryStateNameDisplayLabel   Состояние дела	FieldKindfkCalculated	FieldName	StateNameSize
Calculated	  TFloatFieldRxAvariaQueryWhat	FieldNameWhatOriginDB."MANDAV2.DB".WhatVisible  TStringFieldRxAvariaQueryTYPE	FieldNameTYPEOriginDB."AVPAYS.DB".TypeVisibleSize  TStringFieldRxAvariaQueryAvPayTypeDisplayLabel   "8?	FieldKindfkCalculated	FieldName	AvPayType
Calculated	  
TDateFieldRxAvariaQueryDDisplayLabel   0B0	FieldNameDOriginDB."AVPAYS.DB".D  TStringFieldRxAvariaQueryCURRDisplayLabel   0;NB0	FieldNameCURROriginDB."AVPAYS.DB".CURRSize  TFloatFieldRxAvariaQuerySUMMADisplayLabel   !C<<0	FieldNameSUMMAOriginDB."AVPAYS.DB".SUM  TStringFieldRxAvariaQueryRISK	FieldNameRISKOriginDB."AVPAYS.DB".RISKVisibleSize  TStringFieldRxAvariaQueryV1	FieldNameV1VisibleSize  TStringFieldRxAvariaQueryRISKSTRDisplayLabel    8A:	FieldKindfkCalculated	FieldNameRISKSTR
Calculated	  TStringFieldRxAvariaQueryZATRSTRDisplayLabel   0B@0B0	FieldKindfkCalculated	FieldNameZATRSTR
Calculated	  TSmallintFieldRxAvariaQueryZATR	FieldNameZATRVisible  
TDateFieldRxAvariaQueryDtClLossDisplayLabel   Дата закрытия	FieldNameDtClLoss  
TDateFieldRxAvariaQueryUpdateDate	FieldName
UpdateDate   TDataSource
DataSourceDataSetRxAvariaQueryLeftpTopp  TRxQueryWorkSQLDatabaseNameDBMacros Left Top�   TTableAVTableDatabaseNameDB	TableName
MANDAV2.DBLeft Topp  TFormStorageFormStorageIniFileName   !B@0E>20=85
IniSection   20@88StoredProps.StringsPeriodFrom.DatePeriodFrom.CheckedPeriodTo.DatePeriodTo.Checked StoredValues LeftPTop   