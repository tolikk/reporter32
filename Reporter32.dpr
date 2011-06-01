program Reporter32;



uses
  Forms,
  StrUtils,
  SysUtils,
  MF in 'MF.pas' {MainForm},
  FilterMandatory in 'FilterMandatory.pas' {MandatoryFilterDlg},
  Sort in 'Sort.pas' {SortForm},
  Explorer in 'Explorer.pas' {Explore},
  Korrect in 'Korrect.pas' {KorrectData},
  ChWord in 'ChWord.pas' {ChWordsDlg},
  Move_word in 'Move_word.pas' {MoveWord},
  ShowInfoUnit in 'ShowInfoUnit.pas' {ShowInfo},
  RatesUnit in 'RatesUnit.pas' {Rates},
  TmplNameUnit in 'TmplNameUnit.pas' {GetTemplName},
  CurrRatesUnit in 'CurrRatesUnit.pas' {CurrencyRates},
  MandSellCurrUnit in 'MandSellCurrUnit.pas' {MandSellCurr},
  MoveCurrPeriods in 'MoveCurrPeriods.pas' {MoveCurr},
  InpRezParams in 'InpRezParams.pas' {InitRezerv},
  DaysForSellCurrUnit in 'DaysForSellCurrUnit.pas' {DaysForSellCurr},
  AboutUnit in 'AboutUnit.pas' {About},
  WaitFormUnit in 'WaitFormUnit.pas' {WaitForm},
  AvarUnit in 'AvarUnit.pas' {AvForm},
  GurnalUbitokUnit in 'GurnalUbitokUnit.pas' {FilterForUbitok},
  ArariasFilterUnit in 'ArariasFilterUnit.pas' {ListAvarias},
  AgentUnit in 'AgentUnit.pas' {F_Agent},
  impUnit in 'impUnit.pas' {ImportData},
  ImpFormParam in 'ImpFormParam.pas' {AskImport},
  DataMod in 'DataMod.pas' {CustomerData: TDataModule},
  PolisNmbs in 'PolisNmbs.pas' {PolisNmbsComp},
  Rep2C in 'Rep2C.pas' {Form2C},
  MagazineReport in 'MagazineReport.pas' {MagazineForm},
  AddFld in 'AddFld.pas' {AddFldForm},
  TableModifier in 'TableModifier.pas' {Modifier},
  BelGreenRpt in 'BelGreenRpt.pas' {BelGreenForm},
  InpRezBG in 'InpRezBG.pas' {InitRezervBG_nouse},
  GetRepDateUnit in 'GetRepDateUnit.pas' {GetRepDate},
  GetPeriodUnit in 'GetPeriodUnit.pas' {GetPeriod},
  GetTwoDatesUnit in 'GetTwoDatesUnit.pas' {GetTwoDates},
  RussMandUnit in 'RussMandUnit.pas' {RUSSMAND},
  UnitImportPx in 'UnitImportPx.pas' {formImportPx},
  RepParamsAgntFrm in 'RepParamsAgntFrm.pas' {RepParams},
  RepParamsFrm in 'RepParamsFrm.pas' {RepParams},
  PcntFormUnit in 'PcntFormUnit.pas' {PcntForm},
  Private2005 in 'Private2005.pas' {Private2005Form},
  DM in 'DM.pas' {DMM: TDataModule},
  PayReportFilterUnit in 'PayReportFilterUnit.pas' {PayReportFilter},
  DBTables,
  ExpectedPayFilter in 'ExpectedPayFilter.pas' {ExpectedPaysForm},
  GetPeriodOnly in 'GetPeriodOnly.pas' {GetPeriodForm},
  UniversUnit in 'UniversUnit.pas' {UNIVERS};

{$R *.RES}
var
    StartParam : string;
    sr: TSearchRec;
begin
  Application.Initialize;
  Application.Title := '';

  if FindFirst(GetAppDir + '*qsq*.db', faArchive, sr) = 0 then
  begin
      repeat
        DeleteFile(GetAppDir + sr.Name);
      until FindNext(sr) <> 0;
      FindClose(sr);
  end;

  BLANK_INI := GetAliasPath('BASO') + '\blank.ini';
  if not FileExists(BLANK_INI) then
  begin
    BLANK_INI := 'blank.ini';
  end ;

  if (ParamCount > 0) then
      StartParam := ParamStr(1);
  Application.CreateForm(TDMM, DMM);

  bKomplexMode := StartParam = 'KOMPLEX';

  if (StartParam <> '') AND (StartParam <> 'KOMPLEX') AND (StartParam <> 'ALVENA') AND (StartParam <> 'IMPBASO') then begin
      //if StartParam = 'POLAND' then begin
//          Application.CreateForm(TTechCalculation, TechCalculation);
//  Application.CreateForm(TShowInfo, ShowInfo);
//  Application.CreateForm(TGetFormula, GetFormula);
//  Application.CreateForm(TCurrencyRates, CurrencyRates);
//  Application.CreateForm(TSortForm, SortForm);
//  Application.CreateForm(TInitRezerv, InitRezerv);
//  Application.CreateForm(TDaysForSellCurr, DaysForSellCurr);
//  Application.CreateForm(TAbout, About);
//  Application.CreateForm(TGetTemplName, GetTemplName);
//  Application.CreateForm(TCustomerData, CustomerData);
//  Application.CreateForm(TGetRepDate, GetRepDate);
//  Application.CreateForm(TGetPeriod, GetPeriod);
//  Application.CreateForm(TMsgOk, MsgOk);
//  end;
       if StartParam = 'RANGE' then begin
          Application.Title := 'Диапазоны номеров';
          Application.CreateForm(TPolisNmbsComp, PolisNmbsComp);
       end;
       if StartParam = 'RUSSMAND' then begin
          Application.Title := 'Обязаловка России';
          Application.CreateForm(TRUSSMAND, RUSSMAND);
  Application.CreateForm(TSortForm, SortForm);
  Application.CreateForm(TGetRepDate, GetRepDate);
  Application.CreateForm(TCurrencyRates, CurrencyRates);
  Application.CreateForm(TInitRezerv, InitRezerv);
  Application.CreateForm(TGetTemplName, GetTemplName);
  Application.CreateForm(TRepParams, RepParams);
  Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
  Application.CreateForm(TShowInfo, ShowInfo);
  Application.CreateForm(TPcntForm, PcntForm);
  //Application.CreateForm(TUbitokFrm, UbitokFrm);
  end;
       if StartParam = 'BELGREEN' then begin
          Application.Title := 'Белорусская Зелёная карта';
          Application.CreateForm(TBelGreenForm, BelGreenForm);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TInitRezerv, InitRezerv);
          Application.CreateForm(TDaysForSellCurr, DaysForSellCurr);
          Application.CreateForm(TShowInfo, ShowInfo);
          //Application.CreateForm(TUbitokFrm, UbitokFrm);
          Application.CreateForm(TGetRepDate, GetRepDate);
          Application.CreateForm(TGetTemplName, GetTemplName);
          Application.CreateForm(TRepParams, RepParams);
          Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
          Application.CreateForm(TPcntForm, PcntForm);
          Application.CreateForm(TAbout, About);
       end;
       if StartParam = 'UNIVERS' then begin
          Application.Title := 'Универсальный полис';
          Application.CreateForm(TUNIVERS, UNIVERS);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TGetTemplName, GetTemplName);
          Application.CreateForm(TRepParams, RepParams);
          Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
          Application.CreateForm(TAbout, About);
       end;
       if StartParam = 'PRIV2005' then begin
          Application.Title := 'Страхование жизни 2005';
          Application.CreateForm(TPrivate2005Form, Private2005Form);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TShowInfo, ShowInfo);
          Application.CreateForm(TRepParams, RepParams);
          Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
          Application.CreateForm(TGetTemplName, GetTemplName);
       end;
       if StartParam = 'TOOL' then begin
          Application.Title := 'Утилита';
          Application.CreateForm(TModifier, Modifier);
          Application.CreateForm(TAddFldForm, AddFldForm);
       end;
       if StartParam = '2C' then begin
          Application.Title := '2C';
          Application.CreateForm(TForm2C, Form2C);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
       end;
       if StartParam = 'IMP' then begin
          Application.Title := 'Импорт данных';
          Application.CreateForm(TImportData, ImportData);
          Application.CreateForm(TAskImport, AskImport);
       end;
       if (StartParam = 'AGENTS') OR (StartParam = 'AGENTSSQL') then begin
          Application.Title := 'Агенты';
          Application.CreateForm(TF_Agent, F_Agent);
          Application.CreateForm(TAbout, About);
       end;
{       if StartParam = 'MANDSQL' then begin
          Application.Title := 'ОБТС';
          Application.CreateForm(TMandSQL, MandSQLForm);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TPasswordDlg, PasswordDlg);
          Application.CreateForm(TAbout, About);
          Application.CreateForm(TExpBUROForm, ExpBUROForm);
          Application.CreateForm(TShowInfo, ShowInfo);
          Application.CreateForm(TChWordsDlg, ChWordsDlg);
          Application.CreateForm(TMoveWord, MoveWord);
          Application.CreateForm(TGetTemplName, GetTemplName);
          Application.CreateForm(TDecoderFrm, DecoderFrm);
          Application.CreateForm(TExpErrsFlag, ExpErrsFlag);
          Application.CreateForm(TPolisNmbsComp, PolisNmbsComp);
       end;
}      // if StartParam = 'SOVAG' then begin
       //   Application.Title := 'SOVAG';
       //   Application.CreateForm(TSOVAG, SOVAG);
       //   Application.CreateForm(TSortForm, SortForm);
       //   Application.CreateForm(TAbout, About);
       //end;
       //if StartParam = 'BULSTRAD' then begin
       //   Application.Title := 'BULSTRAD';
       //   Application.CreateForm(TBULSTRAD, BULSTRAD);
       //   Application.CreateForm(TSortForm, SortForm);
       //   Application.CreateForm(TAbout, About);
       //end;
       if StartParam = 'RATES' then begin
          Application.Title := 'Курсы валют';
          Application.CreateForm(TRates, Rates);
          Application.CreateForm(TAbout, About);
       end;
{       if StartParam = 'RATES_SQL' then begin
          Application.Title := 'Курсы валют в Centura';
          Application.CreateForm(TRatesSQL, RatesSQL);
          Application.CreateForm(TAbout, About);
       end;
      { if StartParam = 'LEBEN2' then begin
          Application.Title := 'Страхование жизни';
          Application.CreateForm(TInsLeben2, InsLeben2);
          Application.CreateForm(TLebenFilter, LebenFilter);
          Application.CreateForm(TAbout, About);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
          Application.CreateForm(TMoveCurr, MoveCurr);
          Application.CreateForm(TInitRezervLeben, InitRezervLeben);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TDaysForSellCurr, DaysForSellCurr);
          Application.CreateForm(TGetTemplName, GetTemplName);
          Application.CreateForm(TMagazineLeben, MagazineLeben);
          Application.CreateForm(TUbitokFrm, UbitokFrm);
          Application.CreateForm(TRepParams, RepParams);
          Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
          Application.CreateForm(TShowInfo, ShowInfo);
       end;
       if StartParam = 'FOREIGN' then begin
          Application.Title := 'Страхование иностранцев';
          Application.CreateForm(TInsForeign, InsForeign);
          Application.CreateForm(TForeignFilter, ForeignFilter);
          Application.CreateForm(TDaysForSellCurr, DaysForSellCurr);
          Application.CreateForm(TCurrencyRates, CurrencyRates);
          Application.CreateForm(TSortForm, SortForm);
          Application.CreateForm(TGetTemplName, GetTemplName);
          Application.CreateForm(TAbout, About);
       end }
  end
  else begin
      Application.CreateForm(TMainForm, MainForm);
      if MainForm.IsExit then exit;
      Application.CreateForm(TGetPeriodForm, GetPeriodForm);
      Application.CreateForm(TExpectedPaysForm, ExpectedPaysForm);
      Application.CreateForm(TPayReportFilter, PayReportFilter);
      Application.CreateForm(TSortForm, SortForm);
      Application.CreateForm(TChWordsDlg, ChWordsDlg);
      Application.CreateForm(TMoveWord, MoveWord);
      Application.CreateForm(TShowInfo, ShowInfo);
      Application.CreateForm(TGetTemplName, GetTemplName);
      Application.CreateForm(TFilterForUbitok, FilterForUbitok);
      Application.CreateForm(TAbout, About);
      Application.CreateForm(TMagazineForm, MagazineForm);
      Application.CreateForm(TGetTwoDates, GetTwoDates);
      Application.CreateForm(TCurrencyRates, CurrencyRates);
      Application.CreateForm(TformImportPx, formImportPx);
      Application.CreateForm(TRepParams, RepParams);
      Application.CreateForm(TRepParamsAgnt, RepParamsAgnt);
      Application.CreateForm(TPcntForm, PcntForm);
      Application.CreateForm(TInitRezerv, InitRezerv);
  end;
  Application.Run;
end.
