�
 TAVSQLFORM 0�%  TPF0
TAvSQLForm	AvSQLFormLeft� Top� Width�Height�BorderIconsbiSystemMenu
biMaximize Caption   20@88Color	clBtnFaceFont.CharsetRUSSIAN_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameTahoma
Font.Style OldCreateOrderOnClose	FormCloseOnCreate
FormCreate	OnDestroyFormDestroyPixelsPerInch`
TextHeight 	TRxDBGridRxDBGridLeft Top Width�Height�AlignalClient
DataSourceDataSourceAVTabOrder TitleFont.CharsetRUSSIAN_CHARSETTitleFont.ColorclWindowTextTitleFont.Height�TitleFont.NameTahomaTitleFont.Style OnGetCellPropsRxDBGridGetCellPropsColumnsExpanded	FieldNameSERVisible	 Expanded	FieldNameNMBVisible	 Expanded	FieldNameNVisible	 Expanded	FieldNameAVDTVisible	 Expanded	FieldNameREGDTVisible	 Expanded	FieldNameWRKNMBVisible	 Expanded	FieldNameFIOVisible	 Expanded	FieldNameADDRVisible	 Expanded	FieldNameMARKAVisible	 Expanded	FieldNameAUTONMBVisible	 Expanded	FieldNameBADFIOVisible	 Expanded	FieldNamePAYVisible	 Expanded	FieldNameCURRVisible	    TPanelPanel1Left Top�Width�Height"AlignalBottom
BevelOuterbvNoneTabOrder TLabelLabel1LeftTopWidthiHeightAutoSizeCaption!   Количество аварий  TLabelAvCountLeftpTopWidthaHeightAutoSize  TPanelPanel2LeftXTop WidthAHeight"AlignalRight
BevelOuterbvNoneTabOrder 
DesignSizeA"  TButtonButton1Left� TopWidthtHeightAnchorsakRightakBottom Caption   0:@KBLTabOrder OnClickButton1Click  TButtonButton2Left@TopWidth{HeightCaption   В ExcelTabOrderOnClickButton2Click    TRxQueryAVQueryDatabaseNameSQLDBSQL.StringsSELECT AV.SER AS SER,AV.NMB AS NMB,
AV.N AS N,AV.AVDT AS AVDT,AV.REGDT AS REGDT,AV.PLACE_CODE AS PLACE_CODE,AV.PLACE AS PLACE,AV.WRKNMB AS WRKNMB,AV.FIO AS FIO,AV.ISFIZ AS ISFIZ,AV.ISRSD AS ISRSD,AV.DOCTYPE AS DOCTYPE,AV.DOCSN AS DOCSN,AV.VLADEN AS VLADEN,AV.ADRCODE AS ADRCODE,AV.ADDR AS ADDR,
AV.K AS K,AV.MARKA AS MARKA,AV.LTR AS LTR,AV.CHRCT AS CHRCT,AV.AUTONMB AS AUTONMB,AV.BODYSHS AS BODYSHS,AV.BADFIO AS BADFIO,AV.STAJ AS STAJ,AV.REGRESS AS REGRESS,AV.DECISION AS DECISION,AV.DAMAGECODE AS DAMAGECODE,AV.DTFREE AS DTFREE,AV.DTCLOSE AS DTCLOSE,AV.PAY AS PAY,AV.CURR AS CURR,AV.FS1E AS FS1E,AV.FS1B AS FS1B,AV.RS1E AS RS1E,AV.RS1B AS RS1B,AV.FS2E AS FS2E,AV.FS2B AS FS2B,AV.RS2E AS RS2E,AV.RS2B AS RS2B,AV.FS3E AS FS3E,AV.FS3B AS FS3B,AV.RS3E AS RS3E,AV.RS3B AS RS3B,AV.DT1 AS DT1,AV.DT2 AS DT2,AV.DT3 AS DT3,AV.OZ1 AS OZ1,AV.OZ2 AS OZ2,AV.OZ3 AS OZ3,AV.OTH1 AS OTH1,AV.OTH2 AS OTH2,AV.OTH3 AS OTH3,AV.OZS1 AS OZS1,AV.OZS2 AS OZS2,AV.OZS3 AS OZS3,AV.OTHS1 AS OTHS1,AV.OTHS2 AS OTHS2,AV.OTHS3 AS OTHS3,AV.DMG1 AS DMG1,AV.DMG2 AS DMG2,AV.DMG3 AS DMG3,AV.OWNRCODE AS OWNRCODE,AV.BASECARCODE AS BASECARCODE,AV.CARCODE AS CARCODE,AV.AVCODE AS AVCODE,AV.INSCOMP AS INSCOMP,AV.UPDDT AS UPDDTFROM    SYSADM.MANDAV AV,    SYSADM.MANDATOR MWHERE    M.SER=AV.SERAND    M.NMB=AV.NMBAND   %WHERE   ORDER BY   AV.SER, AV.NMB,AV.N MacrosDataTypeftStringNameWHERE	ParamTypeptInputValue0=0  LeftTop TStringField
AVQuerySERDisplayLabel   !5@8O	FieldNameSEROriginSQLDB.MANDAV.SERSize  TIntegerField
AVQueryNMBDisplayLabel   ><5@	FieldNameNMBOriginSQLDB.MANDAV.NMB  TSmallintFieldAVQueryN	FieldNameNOriginSQLDB.MANDAV.N  TDateTimeFieldAVQueryAVDTDisplayLabel   Дата аварии	FieldNameAVDTOriginSQLDB.MANDAV.AVDT  TDateTimeFieldAVQueryREGDTDisplayLabel   Дата регистрации	FieldNameREGDTOriginSQLDB.MANDAV.REGDT  TStringFieldAVQueryPLACE_CODE	FieldName
PLACE_CODEOriginSQLDB.MANDAV.PLACE_CODEVisibleSize  TStringFieldAVQueryPLACEDisplayLabel   Место аварии	FieldNamePLACEOriginSQLDB.MANDAV.PLACESize2  TStringFieldAVQueryWRKNMBDisplayLabel   Номер дела	FieldNameWRKNMBOriginSQLDB.MANDAV.WRKNMBSize  TStringField
AVQueryFIODisplayLabel   >B5@?52H89	FieldNameFIOOriginSQLDB.MANDAV.FIOSize2  TStringFieldAVQueryISFIZ	FieldNameISFIZOriginSQLDB.MANDAV.ISFIZVisibleSize  TStringFieldAVQueryISRSD	FieldNameISRSDOriginSQLDB.MANDAV.ISRSDVisibleSize  TSmallintFieldAVQueryDOCTYPE	FieldNameDOCTYPEOriginSQLDB.MANDAV.DOCTYPEVisible  TStringFieldAVQueryDOCSNDisplayLabel   Номер документа	FieldNameDOCSNOriginSQLDB.MANDAV.DOCSNSize  TSmallintFieldAVQueryVLADEN	FieldNameVLADENOriginSQLDB.MANDAV.VLADENVisible  TStringFieldAVQueryADRCODE	FieldNameADRCODEOriginSQLDB.MANDAV.ADRCODEVisibleSize  TStringFieldAVQueryADDRDisplayLabel   4@5A	FieldNameADDROriginSQLDB.MANDAV.ADDRSize2  TFloatFieldAVQueryK	FieldNameKOriginSQLDB.MANDAV.KVisible  TStringFieldAVQueryMARKADisplayLabel   Марка авто	FieldNameMARKAOriginSQLDB.MANDAV.MARKASize  TStringField
AVQueryLTRDisplayLabel   C:20	FieldNameLTROriginSQLDB.MANDAV.LTRSize  TFloatFieldAVQueryCHRCTDisplayLabel%   Характеристика авто	FieldNameCHRCTOriginSQLDB.MANDAV.CHRCT  TStringFieldAVQueryAUTONMBDisplayLabel   Номер авто	FieldNameAUTONMBOriginSQLDB.MANDAV.AUTONMBSize
  TStringFieldAVQueryBODYSHSDisplayLabel   Номер шасси	FieldNameBODYSHSOriginSQLDB.MANDAV.BODYSHSSize  TStringFieldAVQueryBADFIODisplayLabel   8=>2=8:	FieldNameBADFIOOriginSQLDB.MANDAV.BADFIOSize2  TSmallintFieldAVQuerySTAJDisplayLabel   !B06	FieldNameSTAJOriginSQLDB.MANDAV.STAJ  TIntegerFieldAVQueryREGRESS	FieldNameREGRESSOriginSQLDB.MANDAV.REGRESSVisible  TIntegerFieldAVQueryDECISION	FieldNameDECISIONOriginSQLDB.MANDAV.DECISIONVisible  TStringFieldAVQueryDAMAGECODE	FieldName
DAMAGECODEOriginSQLDB.MANDAV.DAMAGECODEVisibleSize  TDateTimeFieldAVQueryDTFREEDisplayLabel#   Дата освоб. резерва	FieldNameDTFREEOriginSQLDB.MANDAV.DTFREE  TDateTimeFieldAVQueryDTCLOSEDisplayLabel!   Дата закр. резерва	FieldNameDTCLOSEOriginSQLDB.MANDAV.DTCLOSE  TFloatField
AVQueryPAYDisplayLabel   ?;0G5=>	FieldNamePAYOriginSQLDB.MANDAV.PAY  TStringFieldAVQueryCURRDisplayLabel   Валюта оплаты	FieldNameCURROriginSQLDB.MANDAV.CURRSize  TFloatFieldAVQueryFS1E	FieldNameFS1EOriginSQLDB.MANDAV.FS1E  TFloatFieldAVQueryFS1B	FieldNameFS1BOriginSQLDB.MANDAV.FS1B  TFloatFieldAVQueryRS1E	FieldNameRS1EOriginSQLDB.MANDAV.RS1E  TFloatFieldAVQueryRS1B	FieldNameRS1BOriginSQLDB.MANDAV.RS1B  TFloatFieldAVQueryFS2E	FieldNameFS2EOriginSQLDB.MANDAV.FS2E  TFloatFieldAVQueryFS2B	FieldNameFS2BOriginSQLDB.MANDAV.FS2B  TFloatFieldAVQueryRS2E	FieldNameRS2EOriginSQLDB.MANDAV.RS2E  TFloatFieldAVQueryRS2B	FieldNameRS2BOriginSQLDB.MANDAV.RS2B  TFloatFieldAVQueryFS3E	FieldNameFS3EOriginSQLDB.MANDAV.FS3E  TFloatFieldAVQueryFS3B	FieldNameFS3BOriginSQLDB.MANDAV.FS3B  TFloatFieldAVQueryRS3E	FieldNameRS3EOriginSQLDB.MANDAV.RS3E  TFloatFieldAVQueryRS3B	FieldNameRS3BOriginSQLDB.MANDAV.RS3B  TDateTimeField
AVQueryDT1DisplayLabel   Дата оплаты (ТС)	FieldNameDT1OriginSQLDB.MANDAV.DT1  TDateTimeField
AVQueryDT2DisplayLabel"   Дата оплаты (Жизнь)	FieldNameDT2OriginSQLDB.MANDAV.DT2  TDateTimeField
AVQueryDT3DisplayLabel*   Дата оплаты (Имущество)	FieldNameDT3OriginSQLDB.MANDAV.DT3  TStringField
AVQueryOZ1	FieldNameOZ1OriginSQLDB.MANDAV.OZ1Size  TStringField
AVQueryOZ2	FieldNameOZ2OriginSQLDB.MANDAV.OZ2Size  TStringField
AVQueryOZ3	FieldNameOZ3OriginSQLDB.MANDAV.OZ3Size  TStringFieldAVQueryOTH1	FieldNameOTH1OriginSQLDB.MANDAV.OTH1Size  TStringFieldAVQueryOTH2	FieldNameOTH2OriginSQLDB.MANDAV.OTH2Size  TStringFieldAVQueryOTH3	FieldNameOTH3OriginSQLDB.MANDAV.OTH3Size  TFloatFieldAVQueryOZS1	FieldNameOZS1OriginSQLDB.MANDAV.OZS1  TFloatFieldAVQueryOZS2	FieldNameOZS2OriginSQLDB.MANDAV.OZS2  TFloatFieldAVQueryOZS3	FieldNameOZS3OriginSQLDB.MANDAV.OZS3  TFloatFieldAVQueryOTHS1	FieldNameOTHS1OriginSQLDB.MANDAV.OTHS1  TFloatFieldAVQueryOTHS2	FieldNameOTHS2OriginSQLDB.MANDAV.OTHS2  TFloatFieldAVQueryOTHS3	FieldNameOTHS3OriginSQLDB.MANDAV.OTHS3  TFloatFieldAVQueryDMG1DisplayLabel   Заяв. убыток (ТС)	FieldNameDMG1OriginSQLDB.MANDAV.DMG1  TFloatFieldAVQueryDMG2DisplayLabel#   Заяв. убыток (Жизнь)	FieldNameDMG2OriginSQLDB.MANDAV.DMG2  TFloatFieldAVQueryDMG3DisplayLabel!   Заяв. убыток (Имущ)	FieldNameDMG3OriginSQLDB.MANDAV.DMG3  TIntegerFieldAVQueryOWNRCODE	FieldNameOWNRCODEOriginSQLDB.MANDAV.OWNRCODEVisible  TIntegerFieldAVQueryBASECARCODE	FieldNameBASECARCODEOriginSQLDB.MANDAV.BASECARCODEVisible  TIntegerFieldAVQueryCARCODE	FieldNameCARCODEOriginSQLDB.MANDAV.CARCODEVisible  TIntegerFieldAVQueryAVCODE	FieldNameAVCODEOriginSQLDB.MANDAV.AVCODEVisible  TSmallintFieldAVQueryINSCOMP	FieldNameINSCOMPOriginSQLDB.MANDAV.INSCOMPVisible  
TDateFieldAVQueryUPDDTDisplayLabel2   Дата обновления информации	FieldNameUPDDTOriginSQLDB.MANDAV.UPDDT   TDataSourceDataSourceAVDataSetAVQueryLeftXTop   