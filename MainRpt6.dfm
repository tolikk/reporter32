�
 TMAINREPORT6 0  TPF0TMainReport6MainReport6LeftpTop� BorderStylebsToolWindowCaption   Отчёт в БЮРОClientHeighteClientWidth5Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold OldCreateOrderPositionpoScreenCenterOnCreate
FormCreatePixelsPerInch`
TextHeight TLabelLabel1LeftTopWidthHeightAutoSizeCaption>   За какой период формировать отчёт  TLabelLabel2LeftTopXWidth� HeightCaption1   Дополнительная информация  TLabelLabel3LeftTop0Width� HeightCaption3   Куда сохранять файлы отчёта  TLabelLabel4LeftTop� Width)Height9AutoSizeCaption  Вы уверены, что у всех агентов правильно установлен код подразделения? Ошибки в кодах приведут к разрушению Базы Данных Бюро по транспортному страхованию!!!Font.CharsetDEFAULT_CHARSET
Font.ColorclRedFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFontWordWrap	  	TComboBoxMonthesLeft`TopWidth� HeightStylecsDropDownListDropDownCount
ItemHeightTabOrder OnChangeMonthesChangeItems.Strings   (1) Январь   (2) Февраль   (3) Март   (4) Апрель
   (5) Май   (6) Июнь   (7) Июль   (8) Август   (9) Сентябрь   (10) Октябрь   (11) Ноябрь   (12) Декабрь   TRxSpinEditYearLeft� TopWidthAHeightMaxValue      0�
@MinValue       �	@Value       �	@TabOrderOnChange
YearChange  TMemoMemo1LeftTophWidth)HeightIReadOnly	
ScrollBarsssBothTabOrderWordWrap  TButtonButton1LeftTopHWidth� HeightCaption   >5E0;8TabOrderOnClickButton1Click  TButtonButton2Left� TopHWidth� HeightCaption   B<5=8BLModalResultTabOrder  TDirectoryEditOutDirectoryLeftTop@Width)Height	NumGlyphsTabOrderTextA:\  	TCheckBoxUseTabLeftTop Width	HeightCaption6   Использовать режим ускоренияChecked	State	cbCheckedTabOrder  	TCheckBoxIsUsePrepareLeftTopWidth	HeightCaption@   Произвести предварительный анализChecked	State	cbCheckedTabOrder  	TCheckBoxIsTestExportLeftTop Width)HeightCaptionD   Тестовый прогон (для проверки данных)Font.CharsetRUSSIAN_CHARSET
Font.ColorclBlueFont.Height�	Font.NameArial
Font.StylefsBoldfsItalic 
ParentFontTabOrderOnClickIsTestExportClick  	TComboBox
FlagPeriodLeftTopWidthWHeightStylecsDropDownList
ItemHeightTabOrder	Items.Strings   За   До (включительно)%   После (включительно)   	TCheckBoxIsIndexLeftTop0Width!HeightCaption;   Создать индексы перед экспортомTabOrder
   