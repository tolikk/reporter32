�
 TMANDATORRESERV 0g  TPF0TMandatorReservMandatorReservLeft�TopBorderIconsbiSystemMenu BorderStylebsDialogCaption?   Страховые резервы (Базовая премия)ClientHeight� ClientWidthuColor	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrderPositionpoScreenCenterOnCreate
FormCreateOnShowFormShowPixelsPerInch`
TextHeight TLabelLabel1LeftTopWidthqHeightAutoSizeCaption   Дата отчета  TLabelLabel2LeftTop WidthxHeightCaption%   Отчисления в фонды, %  TLabelLabel3LeftTop8Width� HeightCaption0   Налог на физическое лицо, %  TLabelLabel4LeftTopPWidth� HeightCaption2   Налог на юридическое лицо, %  TButtonButton1Left Top� WidthpHeightCaption   0:@KBLModalResultTabOrder   TDateTimePicker
ReportDateLeft� TopWidth� HeightCalAlignmentdtaLeftDate ��+6oH�@Time ��+6oH�@
DateFormatdfShortDateMode
dmComboBoxKinddtkDate
ParseInputTabOrder  TRxSpinEditFondPercentLeft� TopWidth� HeightMaxValue       �@	ValueTypevtFloatValue       �@TabOrder  TRxSpinEdit	FizichTaxLeft� Top4Width� HeightMaxValue       �@	ValueTypevtFloatValue       �@TabOrder  TRxSpinEdit
UridichTaxLeft� TopLWidth� Height	ValueTypevtFloatTabOrder  	TCheckBoxIsAgentLeftTophWidthaHeightCaption   Для агентаTabOrder  TRxDBLookupCombo
AgentComboLeft� TopdWidth� HeightDropDownCountLookupField
Agent_codeLookupDisplayNameLookupSourceDataSourceAgentsTabOrder  TButtonCalcBtnLeft� Top� WidthkHeightCaption    0AGQBTabOrderOnClickCalcBtnClick  TEdit
StatusTextLeft Top� WidthtHeightAutoSizeColorclMenuCtl3D	ParentCtl3DTabOrder  TQueryAgentsDatabaseNameDBSQL.StringsSELECT *
FROM AGENTORDER BY NAME Left� Top`  TDataSourceDataSourceAgentsDataSetAgentsLeftpTop`  TQueryWorkSQLDatabaseNameDBLeftTop�    