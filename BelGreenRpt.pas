unit BelGreenRpt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, ExtCtrls, Db, DBTables, RxQuery, Grids, DBGrids, RXDBCtrl,
  StdCtrls, ComCtrls, ActnList, CheckLst, rxstrutils, RXSplit, DataMod, InpRezBG,
  Buttons, math, Placemnt;

const
    BG_SECT : string = 'BELGREEN';

type
  TBelGreenForm = class(TForm)
    StatusPanel: TPanel;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    Reserv: TMenuItem;
    N7: TMenuItem;
    Excel1: TMenuItem;
    MainQuery: TRxQuery;
    BelGreenDatabase: TDatabase;
    DataSource: TDataSource;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryPseria: TStringField;
    MainQueryPnumber: TFloatField;
    MainQueryAutonmb: TStringField;
    MainQueryModel: TStringField;
    MainQueryOwner: TStringField;
    MainQueryOwnType: TStringField;
    MainQueryOrg: TStringField;
    MainQueryAddr: TStringField;
    MainQueryDtfrom: TDateField;
    MainQueryPeriod: TStringField;
    MainQueryDtto: TDateField;
    MainQueryLetter: TStringField;
    MainQueryPlace: TStringField;
    MainQueryRegdate: TDateField;
    MainQueryRegtime: TStringField;
    MainQueryTarif: TFloatField;
    MainQueryPay1: TFloatField;
    MainQueryPay1c: TStringField;
    MainQueryPay2: TFloatField;
    MainQueryPay2c: TStringField;
    MainQueryPaydate: TDateField;
    MainQueryAgcode: TStringField;
    MainQueryAgType: TStringField;
    MainQueryPcnt: TFloatField;
    MainQueryState: TStringField;
    MainQueryType: TStringField;
    MainQueryRepdate: TDateField;
    MainQueryCountry: TFloatField;
    MainQueryText: TStringField;
    Panel2: TPanel;
    Button1: TButton;
    PageControl: TPageControl;
    DataPage: TTabSheet;
    FilterPage: TTabSheet;
    MainGrid: TRxDBGrid;
    ActionList: TActionList;
    AgQuery: TQuery;
    AgQueryAgent_code: TStringField;
    AgQueryName: TStringField;
    MainQueryStateText: TStringField;
    Statictic: TMenuItem;
    AllAgents: TTable;
    MainQueryAgName: TStringField;
    FixSerNmb: TMenuItem;
    ProgressBar: TProgressBar;
    WorkSQL: TQuery;
    BUROTbl: TTable;
    BUROTbl2: TTable;
    LocateTbl: TTable;
    StatPage: TTabSheet;
    StatisticTxt: TMemo;
    Panel1: TPanel;
    btnRptMonth: TButton;
    btnRptWeek: TButton;
    NExportTbl: TTable;
    Reports: TTabSheet;
    REPDataSource: TDataSource;
    REPQuery: TRxQuery;
    Panel3: TPanel;
    Label7: TLabel;
    RepList: TComboBox;
    Label8: TLabel;
    ParamsText: TEdit;
    StartButton: TButton;
    Label9: TLabel;
    RxDBGrid1: TRxDBGrid;
    Button2: TButton;
    Splitter: TRxSplitter;
    UbPanel: TPanel;
    Label10: TLabel;
    RxDBGrid2: TRxDBGrid;
    Panel4: TPanel;
    DelBtn: TButton;
    AddBtn: TButton;
    EditBtn: TButton;
    UbTable_: TTable;
    UbTable_Ser: TStringField;
    UbTable_Nmb: TFloatField;
    UbTable_RegDate: TDateField;
    UbTable_RegNmb: TStringField;
    UbTable_ComplName: TStringField;
    UbTable_ComplNmb: TStringField;
    UbTable_ComplDate: TDateField;
    UbTable_UbSum: TFloatField;
    UbTable_UbCurr: TStringField;
    UbTable_PayDate: TDateField;
    UbTable_PaySum: TFloatField;
    UbTable_PayCurr: TStringField;
    UbTable_FreeDate: TDateField;
    UbTable_FreeSum: TFloatField;
    UbTable_FreeCurr: TStringField;
    UbTable_ClDate: TDateField;
    UbTable_Text: TStringField;
    UbTable_BrSum: TFloatField;
    UbTable_BrCurr: TStringField;
    UbTable_Country: TStringField;
    UbTable_State: TFloatField;
    dsUb: TDataSource;
    IsNumbers: TCheckBox;
    N9: TMenuItem;
    N111: TMenuItem;
    IsUseFilter: TCheckBox;
    ImpAlvena: TMenuItem;
    N10: TMenuItem;
    StatusBar: TLabel;
    PopupMenu: TPopupMenu;
    MenuItem1: TMenuItem;
    DelMenu: TMenuItem;
    Panel5: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label12: TLabel;
    CntChecked: TLabel;
    SpeedButton: TSpeedButton;
    NmbFrom: TEdit;
    NmbTo: TEdit;
    RepDtFrom: TDateTimePicker;
    RepDtTo: TDateTimePicker;
    IsAgents: TCheckBox;
    ListAgents: TCheckListBox;
    IsBreak: TCheckBox;
    ListTemplates: TComboBox;
    MainQueryDUPTYPE: TStringField;
    InputSum: TMenuItem;
    N11: TMenuItem;
    OpenDialogPx: TOpenDialog;
    FormStorage: TFormStorage;
    ExpInfo: TLabel;
    MenuGlobo: TMenuItem;
    DownPanel: TPanel;
    Label6: TLabel;
    Label11: TLabel;
    Label13: TLabel;
    Label15: TLabel;
    Apply: TButton;
    Insurer: TEdit;
    StateCombo: TComboBox;
    EndDateFrom: TDateTimePicker;
    EndDateTo: TDateTimePicker;
    listInsType: TComboBox;
    InsurType: TComboBox;
    WorkSQL2: TQuery;
    SaveDialog: TSaveDialog;
    dtRegTo: TDateTimePicker;
    Label14: TLabel;
    dtRegFrom: TDateTimePicker;
    Label16: TLabel;
    Label17: TLabel;
    DivisionInfo: TMenuItem;
    N6: TMenuItem;
    N8: TMenuItem;
    CompanyTitle: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    CurrencyRateSrc: TMenuItem;
    N15: TMenuItem;
    procedure N2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure ApplyClick(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Excel1Click(Sender: TObject);
    procedure FixSerNmbClick(Sender: TObject);
    procedure ReservClick(Sender: TObject);
    procedure BuroMnuClick(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure RepListChange(Sender: TObject);
    procedure StartButtonClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure btnRptMonthClick(Sender: TObject);
    procedure AddBtnClick(Sender: TObject);
    procedure EditBtnClick(Sender: TObject);
    procedure DelBtnClick(Sender: TObject);
    procedure MainQueryAfterScroll(DataSet: TDataSet);
    procedure N8Click(Sender: TObject);
    procedure N111Click(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure ListAgentsClick(Sender: TObject);
    procedure InputSumClick(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure ImpAlvenaClick(Sender: TObject);
    procedure DelMenuClick(Sender: TObject);
    procedure MenuGloboClick(Sender: TObject);
    procedure Panel5Resize(Sender: TObject);
    procedure StaticticClick(Sender: TObject);
    procedure CurrencyRateSrcClick(Sender: TObject);
  private
    { Private declarations }
    listAgCodes : TStringList;
    MSEXCEL : Variant;

    function GetFilter : string;
    function GetOrder : string;
    procedure FillDocsData(IsRNPU : boolean; var DataList : TList; RepDate : TDateTime; obj : TaxPercents; List1, ListRest : TASKOneTaxList);
    procedure ImportAlvena(DBFile : string);
  public
    { Public declarations }
    procedure CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List1, ListRest : TASKOneTaxList);
  end;

  function DecodeNBByOwner(s : string) : string;

var
  BelGreenForm: TBelGreenForm;
  sdecode : string;

implementation

uses Sort, WaitFormUnit, dateutil, Variants,
  BGBuroUnit, inifiles, ShowInfoUnit, Ubitok, 
  InpRezParams, TmplNameUnit, registry, RepParamsFrm, RepParamsAgntFrm,
  PcntFormUnit, RussMandUnit, StrUtils, mf, DM, AboutUnit, CurrRatesUnit;

{$R *.DFM}

procedure TBelGreenForm.N2Click(Sender: TObject);
begin
    Close
end;

procedure TBelGreenForm.Button1Click(Sender: TObject);
begin
    Close
end;

procedure ConnectFilter(var filter : string; Operation, Condition : string);
begin
     if Length(filter) <> 0 then
        filter := filter + ' ' + Operation + ' ';
     filter := filter + Condition;
end;

procedure TBelGreenForm.FormCreate(Sender: TObject);
var
    INI : TIniFile;
    i : integer;
    rINI : TRegIniFile;
    str : string;
begin
    CompanyTitle.Caption := COMPAMYNAME;
     INI := TIniFile.Create(BLANK_INI);
     DivisionInfo.Caption := 'ПОДРАЗДЕЛЕНИЕ: ' + INI.ReadString('DIVISION', 'ID', '?');
     INI.Free;
    DMM.HandBookSQL.DatabaseName := BelGreenDatabase.Name;
    RepDtFrom.Date := Date;
    RepDtTo.Date := Date;
    RepDtFrom.Checked := false;
    RepDtTo.Checked := false;

    listAgCodes := TStringList.Create;
    StateCombo.ItemIndex := 0;

     rINI := TRegIniFile.Create('БелЗК');
     for i := 0 to 100 do begin
         str := rINI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
         //else
         //    break;
     end;
     rINI.Free;

    try
        AgQuery.Open;
        while not AgQuery.Eof do begin
            ListAgents.Items.Add(AgQuery.FieldByName('Name').AsString);
            listAgCodes.Add(AgQueryAgent_code.AsString);
            AgQuery.Next;
        end;
        AgQuery.Close;

        PageControl.ActivePage := FilterPage;
    except
      on E : Exception do begin
        MessageDlg('Ошибка запуска программы' + E.Message, mtError, [mbOk], 0);
      end;
    end;

    //Заполним список отчётов
    INI := TIniFile.Create('blreps.ini');
    for i := 1 to 100 do begin
        if INI.ReadString(BG_SECT, 'QueryName' + IntToStr(i), '') <> '' then
            RepList.Items.Add(INI.ReadString(BG_SECT, 'QueryName' + IntToStr(i), ''))
        else
            break;
    end;
    INI.Free;
    Panel5Resize(nil)
end;

procedure TBelGreenForm.FormDestroy(Sender: TObject);
begin
    listAgCodes.Free
end;

procedure TBelGreenForm.MainQueryCalcFields(DataSet: TDataSet);
begin
    MainQueryStateText.AsString := '';
    if MainQueryState.IsNull then MainQueryStateText.AsString := '???'
    else
    if MainQueryState.AsInteger = 1 then MainQueryStateText.AsString := 'ЗАМЕНЁН'
    else
    if MainQueryState.AsInteger = 2 then MainQueryStateText.AsString := 'ИСПОРЧЕН'
    else
    if MainQueryState.AsInteger = 3 then MainQueryStateText.AsString := 'УТЕРЯН АГЕНТОМ'
    else
    if MainQueryState.AsInteger = 4 then MainQueryStateText.AsString := 'РАСТОРГНУТ'
    else
    if MainQueryState.AsInteger = 5 then MainQueryStateText.AsString := 'УТЕРЯН'
    else
    if MainQueryState.AsInteger = 0 then begin
        if (MainQueryPseria.AsString <> '') AND (MainQueryDUPTYPE.AsInteger = 1) then MainQueryStateText.AsString := 'ДУБЛИКАТ';
        if (MainQueryPseria.AsString <> '') AND (MainQueryDUPTYPE.AsInteger = 0) then MainQueryStateText.AsString := 'ВЗАМЕН';
    end
end;

procedure TBelGreenForm.ApplyClick(Sender: TObject);
begin
    MainQuery.Close;
    MainQuery.MacroByName('WHERE').AsString := GetFilter();
    MainQuery.MacroByName('ORDER').AsString := GetOrder();
    try
        Statictic.Caption := '';
        Screen.Cursor := crHourglass;
        MainQuery.Open;
        Statictic.Caption := 'Найдено записей ' + IntToStr(MainQuery.RecordCount);
        PageControl.ActivePageIndex := 0;
    except
      on E : Exception do begin
        Statictic.Caption := E.Message;
        MessageDlg('Ошибка выполнения команды' + E.Message, mtError, [mbOk], 0);
      end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TBelGreenForm.N3Click(Sender: TObject);
begin
    PageControl.ActivePageIndex := 1;
end;

function TBelGreenForm.GetFilter : string;
var
    s, s_a, s_r : string;
    i : integer;
begin
    s := '';
    if NmbFrom.Text <> '' then
        ConnectFilter(s, 'AND', 'NUMBER>=' + NmbFrom.Text);
    if NmbTo.Text <> '' then
        ConnectFilter(s, 'AND', 'NUMBER<=' + NmbTo.Text);
    if RepDtFrom.Checked then
        ConnectFilter(s, 'AND', 'REPDATE>=''' + DateToStr(RepDtFrom.Date) + '''');
    if RepDtTo.Checked then
        ConnectFilter(s, 'AND', 'REPDATE<=''' + DateToStr(RepDtTo.Date) + '''');
    if Insurer.Text <> '' then
        ConnectFilter(s, 'AND', 'OWNER LIKE ''%' + Insurer.Text + '%''');
    if InsurType.ItemIndex = 1 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''F''');
    if InsurType.ItemIndex = 2 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''U''');
    if StateCombo.ItemIndex >= 1 then begin
        if StateCombo.ItemIndex = 6 then //ДУБЛИКАТ
            ConnectFilter(s, 'AND', '(PSERIA IS NOT NULL AND DUPTYPE=''0'')')
        else
        if StateCombo.ItemIndex = 7 then //ЗАМЕНА
            ConnectFilter(s, 'AND', '(PSERIA IS NOT NULL AND DUPTYPE=''1'')')
        else begin
            ConnectFilter(s, 'AND', 'STATE LIKE ''' + IntToStr(StateCombo.ItemIndex - 1) + '''');
            if StateCombo.ItemIndex = 1 then //НОРМАЛЬНЫЙ
                ConnectFilter(s, 'AND', '(PSERIA IS NULL OR PSERIA='''')');
        end
    end;

    if EndDateFrom.Checked then
        ConnectFilter(s, 'AND', 'DTTO>=''' + DateToStr(EndDateFrom.Date) + '''');
    if EndDateTo.Checked then
        ConnectFilter(s, 'AND', 'DTTO<=''' + DateToStr(EndDateTo.Date) + '''');

    if dtRegFrom.Checked then
        ConnectFilter(s, 'AND', 'REGDATE>=''' + DateToStr(dtRegFrom.Date) + '''');
    if dtRegTo.Checked then
        ConnectFilter(s, 'AND', 'REGDATE<=''' + DateToStr(dtRegTo.Date) + '''');

    if listInsType.ItemIndex = 1 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''F''');
    if listInsType.ItemIndex = 2 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''U''');
    if listInsType.ItemIndex = 3 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''I''');

    if IsBreak.Checked AND (RepDtFrom.Checked OR RepDtTo.Checked)then begin
        if RepDtFrom.Checked then
            ConnectFilter(s_r, 'AND', 'RETDATE>=''' + DateToStr(RepDtFrom.Date) + '''');
        if RepDtTo.Checked then
            ConnectFilter(s_r, 'AND', 'RETDATE<=''' + DateToStr(RepDtTo.Date) + '''');

        if s <> '' then s := '(' + s + ')';
        ConnectFilter(s, 'OR', s_r);
    end;

    if s <> '' then s := '(' + s + ')';

    if IsAgents.Checked then begin
        s_a := '';
        for i := 0 to ListAgents.Items.Count - 1 do
            if ListAgents.Checked[i] = true then begin
                if s_a <> '' then s_a := s_a + ',';
                s_a := s_a + '''' + listAgCodes[i] + '''';
            end;
         if s_a <> '' then
            ConnectFilter(s, 'AND', 'AGCODE IN (' + s_a + ')');
    end;


    if s = '' then GetFilter := '1=1'
    else GetFilter := s;
end;

function TBelGreenForm.GetOrder : string;
begin
    if SortForm.GetOrder('') <> '' then GetOrder := SortForm.GetOrder('ORDER BY')
    else GetOrder := 'ORDER BY SERIA, NUMBER';
end;

procedure TBelGreenForm.N4Click(Sender: TObject);
begin
    SortForm.m_Query := MainQuery;
    if SortForm.ShowModal = mrOk then
        ApplyClick(nil);
end;

procedure TBelGreenForm.Excel1Click(Sender: TObject);
begin
    DatSetToExcel(MainQuery, true);
end;

procedure TBelGreenForm.FixSerNmbClick(Sender: TObject);
begin
    if MainGrid.FixedCols = 0 then MainGrid.FixedCols := 2
    else MainGrid.FixedCols := 0;
    FixSerNmb.Checked := MainGrid.FixedCols > 0;
end;

procedure TBelGreenForm.ReservClick(Sender: TObject);
var
    obj : TaxPercents;
begin
    InitRezerv.RestoreName('БелЗК');
    if InitRezerv.ShowModal = mrOk then begin
         Screen.Cursor := crHourGlass;
         obj := TaxPercents.Create;
         obj.FizPcntBefore := InitRezerv.FizichTax1.Value;
         obj.FizPcntAfter := InitRezerv.FizTax2.Value;
         obj.UrPcntBefore := InitRezerv.UridichTax1.Value;
         obj.UrPcntAfter := InitRezerv.UrTax2.Value;
         obj.UrDt := InitRezerv.UrTaxDate.Date;
         obj.FizDt := InitRezerv.FizTaxDate.Date;
         obj.CompPcntBefore := InitRezerv.CompanyPercent1.Value;
         obj.CompPcntAfter := InitRezerv.CompanyPercent2.Value;
         obj.CompPcntDt := InitRezerv.CompanyPercentDate.Date;
         CreateReport(TControl(Sender).Tag = 1,
                      InitRezerv.RepData.Date,
                      obj,
                      //InitRezerv.FizichTax.Value,
                      //InitRezerv.UridichTax.Value,
                      //InitRezerv.CompanyPercent.Value,
                      InitRezerv.IsFast.Checked,
                      InitRezerv.List,
                      InitRezerv.ListRest);
         obj.Free;
         Screen.Cursor := crDefault;
         Application.BringToFront;
    end;
end;

procedure TBelGreenForm.CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List1, ListRest : TASKOneTaxList);
var
    ListData : TList;
    i : integer;
begin
    ListData := TList.Create;
    try
         FillDocsData(IsRNPU, ListData, TRUNC(RepDate), obj, List1, ListRest);
         if not InitExcel(MSEXCEL) then begin
            MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
            exit;
         end;
         MSEXCEL.Visible := true;
         
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;

         OutTo10Form(IsRNPU, 'Белорусская зелёная карта', MSEXCEL, RepDate, ListData, IsFast, BG_SECT);

         //TransferToExecl('Белорусская зелёная карта', RepDate, BG_SECT, ListData, MSEXCEL, IsFast);
    except
        on E : Exception do begin
            Application.BringToFront;
            if WaitForm <> nil then WaitForm.Close;
            MessageDlg('Возникла ошибка при формировании отчёта'#13 + e.Message, mtInformation, [mbOk], 0);
        end;
    end;

    for i := 0 to ListData.Count - 1 do Dispose(ListData.Items[i]);
    ListData.Free;
    if WaitForm <> nil then WaitForm.Close;
    ProgressBar.Position := 0;
end;

procedure TBelGreenForm.FillDocsData(IsRNPU : boolean; var DataList : TList; RepDate : TDateTime; obj : TaxPercents; List1, ListRest : TASKOneTaxList);
var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : PGroupRestrax;
begin
     //Те кторые действуют на дату или те которые Б 30 дней и оплачены в предшеств. отчёту месяц
     //Сумма берётся вся пришедшая до отчётной даты
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT');
          Add('   Seria, Number, PayDate Repdate, dtfrom StartDate, dtto EndDate, RegDate, pay1 PremiumPay, pay2 PremiumPay2, pay1c PremiumCurr, pay2c PremiumCurr2, AgType Uridich, pcnt AgPercent');
          Add('FROM');
          Add('   BELGREEN');
          Add('WHERE');
          Add('   PAY1 is not null AND'); //RegDate - дата оплаты
          if IsRNPU then
              Add('   (paydate between :RepDateSubMonth and :RepDate)') //попал в перио
          else begin
              Add('   (paydate < :RepDate AND dtto >= :RepDate AND (RetDate is null or RetDate < :RepDate) OR'); //попал в период
              Add('   (paydate >= :RepDateSubMonth AND paydate < :RepDate AND (dtto - dtfrom + 1) <= 30))');
          end;
          Add('ORDER BY');
          Add('   Seria, Number');
          ParamByName('RepDate').AsDateTime := TRUNC(RepDate);
          if IsRNPU then
              ParamByName('RepDateSubMonth').AsDateTime := IncDate(TRUNC(RepDate), 0, -12, 0) + 1
          else
              ParamByName('RepDateSubMonth').AsDateTime := IncDate(TRUNC(RepDate), 0, -1, 0) + 1;
          try
              Open;
              if EOF then begin
                 MessageDlg('Не обнаружено полисов по указанному условию', mtInformation, [mbOk], 0);
                 exit;
              end;
          except
              on E : Exception do begin
                  MessageDlg('Возникла ошибка при выполнении запроса к Базе Данных'#13 + e.Message, mtInformation, [mbOk], 0);
                  WorkSQL.Close;
                  exit;
              end;
          end;
     end;

     while not WorkSQL.EOF do begin
         Summa := 0;
         InitStructure_3(Info);

        ProgressBar.Position := TRUNC(WorkSQL.RecNo * 100 / WorkSQL.RecordCount);
         with WorkSQL do begin
           if FieldByName('PremiumCurr').AsString <> '' then
              FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'U', obj.GetPcnt(FieldByName('RegDate').AsDateTime, true), obj.GetPcnt(FieldByName('RegDate').AsDateTime, false), obj.GetCompPcnt(FieldByName('RegDate').AsDateTime), Info, FieldByName('RegDate').AsDateTime, FieldByName('PremiumPay').AsFloat, FieldByName('PremiumCurr').AsString, FieldByName('AgPercent').AsFloat, GetOneTax(List1, FieldByName('RegDate').AsDateTime));
           if FieldByName('PremiumCurr2').AsString <> '' then
              FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'U', obj.GetPcnt(FieldByName('RegDate').AsDateTime, true), obj.GetPcnt(FieldByName('RegDate').AsDateTime, false), obj.GetCompPcnt(FieldByName('RegDate').AsDateTime), Info, FieldByName('RegDate').AsDateTime, FieldByName('PremiumPay2').AsFloat, FieldByName('PremiumCurr2').AsString, FieldByName('AgPercent').AsFloat, GetOneTax(List1, FieldByName('RegDate').AsDateTime));
         end;

         //Append Info to container
         for i := 1 to 3 do begin
             if Info[i].Curr <> '' then begin
                 New(InfoPtr);
                 {
                 InfoPtr^ := Info[i];
                 InfoPtr.Number := WorkSQL.FieldByName('Number').AsFloat;
                 InfoPtr.RegDate := WorkSQL.FieldByName('RepDate').AsDateTime;
                 InfoPtr.FromDate := WorkSQL.FieldByName('StartDate').AsDateTime;
                 InfoPtr.ToDate := WorkSQL.FieldByName('EndDate').AsDateTime;
                 DataList.Add(InfoPtr);}
                  InfoPtr.Curr := Info[i].Curr;
                  InfoPtr.BruttoPremium := Info[i].BruttoPremium;
                  InfoPtr.AgentSum := Info[i].AgOtchislenie;
                  InfoPtr.FondSum := Info[i].OneTaxOtchislenie;
                  InfoPtr.BasePremium := InfoPtr.BruttoPremium -
                                         InfoPtr.AgentSum -
                                         InfoPtr.FondSum;

                  InfoPtr.RestraxPart := 0;
                  InfoPtr.Insurer := '';
                  InfoPtr.InsObject := '';

                  InfoPtr.AllSummaCurr := 'EUR';
                  InfoPtr.AllSumma := 250000;

                  InfoPtr.Seria := WorkSQL.FieldByName('Seria').AsString;
                  InfoPtr.Number := WorkSQL.FieldByName('Number').AsFloat;
                  //InfoPtr.RegDate := WorkSQL.FieldByName('RepDate').AsDateTime;
                  InfoPtr.DateFrom := WorkSQL.FieldByName('StartDate').AsDateTime;
                  InfoPtr.DateTo := WorkSQL.FieldByName('EndDate').AsDateTime;

                  if RepDate < InfoPtr.DateFrom then
                      InfoPtr.ReserveNZPSumm := InfoPtr.BasePremium
                  else
                  if (InfoPtr.DateTo - InfoPtr.DateFrom) <= 30 then
                      InfoPtr.ReserveNZPSumm := InfoPtr.BasePremium / 2
                  else
                      InfoPtr.ReserveNZPSumm := InfoPtr.BasePremium * (InfoPtr.DateTo - RepDate + 1) / (InfoPtr.DateTo - InfoPtr.DateFrom + 1);

                  if IsRNPU then
                      InfoPtr.ReserveNZPSumm := InfoPtr.BasePremium / 10;

                  InfoPtr.RestraxPart := InfoPtr.ReserveNZPSumm * GetOneTax(ListRest, WorkSQL.FieldByName('RegDate').AsDateTime) / 100;
                  InfoPtr.StraxPart := InfoPtr.ReserveNZPSumm - InfoPtr.RestraxPart;

                  //InfoPtr.StraxPart := InfoPtr.ReserveNZPSumm;

                  DataList.Add(InfoPtr);
             end;
         end;

         WorkSQL.Next;
     end;
     ProgressBar.Position := 0;
end;

function NmbToStr(N : integer) : string;
var
    s : string;
begin
    s := IntToStr(N);
    while Length(s) < 7 do s := '0' + s;
    Result := s;
end;

procedure raise_Exception(var msg : string; s : string; agntname : string);
begin
    try
        BelGreenForm.AgQuery.Open;
        if BelGreenForm.AgQuery.Locate('Agent_code', agntname, []) then
            agntname := BelGreenForm.AgQuery.FieldByName('Name').AsString;
    except
    end;

    if Pos(s, Msg) = 0 then
        msg := msg + #13#10 + s + ' ' + agntname;
end;

function EnglishStr(s : string) : string;
begin
    s := Trim(UpperCase(s));
end;

function IsEnglish(Str : string) : boolean;
var
    i : integer;
begin
    IsEnglish := false;
    for i := 1 to Length(Str) do
        if ((Str[i] >= 'A') AND (Str[i] <= 'Z')) OR
           ((Str[i] >= 'a') AND (Str[i] <= 'z')) then begin
           IsEnglish := true;
           break;
         end;
end;

function IsRussian(Str : string; ExcludeLettersKey : string = '') : boolean;
var
    i : integer;
    s_excl : string;
    INI : TIniFile;
begin
    if ExcludeLettersKey <> '' then begin
        INI := TIniFile.Create(BLANK_INI);
        s_excl := INI.ReadString(BG_SECT, ExcludeLettersKey, '');
        INI.Free;
    end;

	for i := 1 to Length(Str) do begin
		if (Str[i] >= 'а') AND  (Str[i] <= 'я') OR (Str[i] >= 'А') AND (Str[i] <= 'Я') then begin
            if Pos(Str[i], s_excl) = 0 then begin
                IsRussian := true;
                exit;
            end;
        end;
	end;
    IsRussian := false;
end;

//Заменяет symb на newsymb и удаляет двойные пробелы
function ReplaceSymbols(str : string; symb : string = '"''[](),.-:*`_'; newsymb : string = '          ') : string;
var
    i, pos_n : integer;
begin
    for i := 1 to Length(symb) do begin
        while true do begin
            pos_n := Pos(symb[i], str);
            if pos_n = 0 then break;
            str := Copy(str, 1, pos_n - 1) + newsymb[i] + Copy(str, pos_n + 1, 1024);
        end;
    end;
    while true do begin
        pos_n := Pos('  ', str);
        if pos_n = 0 then break;
        str := Copy(str, 1, pos_n - 1) + Copy(str, pos_n + 1, 1024);
    end;
    ReplaceSymbols := Trim(str);
end;

//заменяет символы из symb на символ newsymb
function _KillSymbols(str, symb : string; newsymb : string = '') : string;
var
    i, pos_n : integer;
begin
    for i := 1 to Length(symb) do begin
        while true do begin
            pos_n := Pos(symb[i], str);
            if pos_n = 0 then break;
            str := Copy(str, 1, pos_n - 1) + newsymb + Copy(str, pos_n + 1, 1024);
        end;
    end;
    _KillSymbols := str;
end;

procedure TBelGreenForm.BuroMnuClick(Sender: TObject);
var
    errText : string;
    INI : TIniFile;
    buff1, buff2 : array[0..127] of char;
    BuroPath : string;
    N, All : integer;
    ReportStr : string;
    toFirst, toSecond : integer;
    s : string;
    NExport : integer;
    Ser : string;
    Nmb : integer;
    NStr : string;
    IsRuss, IsEngl, IsDigit, IsBad : boolean;
   // LostCount : integer;
   // LostSer : array[1..256] of string;
   // LostNmb : array[1..256] of integer;
   // Stopped : boolean;
begin
(*
    if BG_Buro.ShowModal <> mrOK then exit;
    DeleteFile(BG_Buro.Path.Text + 'Z1020001.dbf');
    DeleteFile(BG_Buro.Path.Text + 'Z1020002.dbf');
//    LostCount := 0;
//    Stopped := false;

    toFirst := 0;
    toSecond := 0;
    INI := TIniFile.Create(BLANK_INI);
    BuroPath := INI.ReadString(BG_SECT, 'BuroDirectory', '');
    if BuroPath[Length(BuroPath)] <> '\' then BuroPath := BuroPath + '\';
    if (FALSE = CopyFile(StrPCopy(buff1, BuroPath + 'Z1020001.dbf'), StrPCopy(buff2, BG_Buro.Path.Text + 'Z1020001.dbf'), FALSE)) then begin
        MessageDlg('Ошибка копирования Z1020001.dbf или Z1020002.dbf из ' + BuroPath + ' в ' + BG_Buro.Path.Text, mtInformation, [mbOk], 0);
        INI.Free;
        exit;
    end;
    if (FALSE = CopyFile(StrPCopy(buff1, BuroPath + 'Z1020002.dbf'), StrPCopy(buff2, BG_Buro.Path.Text + 'Z1020002.dbf'), FALSE)) then begin
        MessageDlg('Ошибка копирования Z1020001.dbf или Z1020002.dbf из ' + BuroPath + ' в ' + BG_Buro.Path.Text, mtInformation, [mbOk], 0);
        INI.Free;
        exit;
    end;
    INI.Free;

    BUROTbl.Close;
    BUROTbl2.Close;
    BUROTbl.TableName := BG_Buro.Path.Text + 'Z1020001.dbf';
    BUROTbl2.TableName := BG_Buro.Path.Text + 'Z1020002.dbf';

{ "Нормальный", "Земенён", "Испорчен", "Утерян" расторгнут };

    with WorkSQL, WorkSQL.SQL do begin
        Close;
        Clear;
        Add('SELECT MAX(NEXPORT) AS NE FROM BELGREENEXP');
        Open;
        NExport := 0;
        if not FieldByName('NE').IsNull then
            NExport := FieldByName('NE').AsInteger;

        Close;
        Clear;
        Add('SELECT B.* FROM BELGREEN B');
        //if BG_Buro.NExport.ItemIndex = 0 then begin
        //    Add('WHERE UPD > :YEARAGO OR UPD IS NULL');
        //    ParamByName('YEARAGO').AsDateTime := Date - 370;
        //end;
        if BG_Buro.NExport.ItemIndex = 1 then begin
            Add(',BELGREENEXP E');
            Add('WHERE E.SERIA=B.SERIA AND E.NUMBER=B.NUMBER AND NEXPORT=:NEXP');
            ParamByName('NEXP').AsInteger := NExport;
        end;
        if BG_Buro.NExport.ItemIndex = 2 then begin
            Add(',BELGREENEXP E');
            Add('WHERE E.SERIA=B.SERIA AND E.NUMBER=B.NUMBER AND NEXPORT=:NEXP');
            ParamByName('NEXP').AsInteger := NExport - 1;
        end;

        NExport := NExport + 1;

        Screen.Cursor := crHourglass;
        try
            ExpInfo.Caption := 'Подготовка данных...';
            ExpInfo.Update;

            NExportTbl.Open;
            BUROTbl.Open;
            BUROTbl2.Open;
            Open;
            All := RecordCount;
            ReportStr := 'Всего обработано ' + InttoStr(All) + ' полисов'#13;
            N := 1;
            while ((not Eof){ AND (not Stopped)) OR (LostCount > 0}) do begin
                ProgressBar.Position := round(n * 100 / All);
                Inc(N);
                NStr := IntToStr(N);

                ExpInfo.Caption := IntToStr(n) + ' из ' + IntToStr(All);
                ExpInfo.Update;

                //if N = 21503 then
                //    MessageBeep(0);

                //if (not Eof) AND (not Stopped) then begin
                    if BG_Buro.NExport.ItemIndex = 0 then begin
                        if NExportTbl.Locate('seria;number', vararrayof([FieldByName('Seria').AsString, FieldByName('Number').AsInteger]), []) then begin
                            Next;
                            continue;
                        end
                        else
                        if not BG_Buro.IsTest.Checked then begin
                            NExportTbl.Append;
                            NExportTbl.FieldByName('Seria').AsString := FieldByName('Seria').AsString;
                            NExportTbl.FieldByName('Number').AsInteger := FieldByName('Number').AsInteger;
                            NExportTbl.FieldByName('NExport').AsInteger := NExport;
                            NExportTbl.FieldByName('UPD').AsDateTime := Date;//FieldByName('UPD').AsDateTime;
                            NExportTbl.Post;
                        end;
                    end;

                if FieldByName('Number').AsInteger = 390744 then
                    MessageBeep(0);
                {end
                else begin
                    //Найти мот он уже есть всё таки
                    if not BUROTbl.Locate('POLICY_NUM', LostNmb[LostCount], []) then begin
                        Close;
                        Clear;
                        Add('SELECT B.* FROM BELGREEN B');
                        Open;
                        if not Locate('seria;number', vararrayof([LostSer[LostCount], LostNmb[LostCount]]), []) then
                            raise_Exception(errText, 'Полис ' + LostSer[LostCount] + '/' + IntToStr(LostNmb[LostCount]) + ' утерян, но не найден ', '')
                        else
                            if FieldByName('STATE').AsString <> '1' then
                                raise_Exception(errText, 'Полис ' + LostSer[LostCount] + '/' + IntToStr(LostNmb[LostCount]) + ' утерян, но не правильное состяние ', '');
                    end
                    else begin
                        LostCount := LostCount - 1;
                        continue;
                    end;
                    LostCount := LostCount - 1;
                end;
                }
                if FieldByName('State').AsString = '' then begin
                    raise_Exception(errText, 'Поле State не определено для полиса  ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    Next;
                    continue;
                end;

                if FieldByName('State').AsString[1] in ['2', '3', '4'] then begin //испорчен утерян расторгнут
                    BUROTbl2.Append;
                    if FieldByName('State').AsString[1] = '2' then
                        BUROTbl2.FieldByName('MSG_STATUS').AsString := 'I'
                    else
                    if FieldByName('State').AsString[1] = '3' then
                        BUROTbl2.FieldByName('MSG_STATUS').AsString := 'U'
                    else
                    if FieldByName('State').AsString[1] = '4' then begin
                        BUROTbl2.FieldByName('MSG_STATUS').AsString := 'Z';
                        BUROTbl2.FieldByName('DATE_TERM').AsDateTime := FieldByName('RetDate').AsDateTime;
                        if FieldByName('RetDate').IsNull then
                           raise_Exception(errText, 'Нет даты возврата ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                        BUROTbl2.FieldByName('SUM_RET1').AsFloat := FieldByName('Ret1').AsFloat;
                        if FieldByName('Ret1Cur').AsString = 'BRB' then
                            BUROTbl2.FieldByName('CURRENCYR1').AsString := 'BYR'
                        else
                            BUROTbl2.FieldByName('CURRENCYR1').AsString := FieldByName('Ret1Cur').AsString;
                    end;
                    BUROTbl2.FieldByName('COMPANY').AsString := '02';
                    BUROTbl2.FieldByName('POLICY_NUM').AsString := NmbToStr(FieldByName('NUMBER').AsInteger);
                    BUROTbl2.Post;
                    Inc(toSecond);
                end;

                if FieldByName('State').AsString[1] in ['0', '1', '4', '5'] then begin
                    //Проверка на корректность
                    if Pos('+', FieldByName('LETTER').AsString) <> 0 then begin
                        LocateTbl.Open;
                        if (FieldByName('Tarif').AsFloat >= 0.01) then begin //Основной
                            Ser := FieldByName('PRCPSER').AsString;
                            Nmb := FieldByName('PRCPNMB').AsInteger;
                            if Ser = '' then begin
                                Ser := FieldByName('SERIA').AsString;
                                Nmb := FieldByName('NUMBER').AsInteger + 1;
                            end;
                            if not LocateTbl.Locate('seria;number', vararrayof([Ser, Nmb]), []) then
                                raise_Exception(errText, 'Не найден полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' для ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                            else begin
                                if (LocateTbl.FieldByName('Tarif').AsFloat >= 0.01) OR (Pos('+', LocateTbl.FieldByName('LETTER').AsString) = 0) then
                                    raise_Exception(errText, 'Полис ' + Ser + '/' + InttoStr(Nmb) + ' не является сцепкой для полиса ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                                else begin
                                    if LocateTbl.FieldByName('PLACE').AsString <> FieldByName('PLACE').AsString then
                                        raise_Exception(errText, 'Полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' выдан не там где основной полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                                    if LocateTbl.FieldByName('REGDATE').AsString <> FieldByName('REGDATE').AsString then
                                        raise_Exception(errText, 'Полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' выдан не тогда когда основной полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                                    if LocateTbl.FieldByName('DTFROM').AsString <> FieldByName('DTFROM').AsString then
                                        raise_Exception(errText, 'Полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' даты начала страхования на совпадают ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                                    if LocateTbl.FieldByName('DTTO').AsString <> FieldByName('DTTO').AsString then
                                        raise_Exception(errText, 'Полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' даты окончания страхования на совпадают ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                                    if not LocateTbl.FieldByName('STATE').AsInteger in [0, 1] then
                                        raise_Exception(errText, 'Полис-сцепка ' + Ser + '/' + InttoStr(Nmb) + ' для ' + FieldByName('NUMBER').AsString + ' имеет неправильное состояние', FieldByName('AGCODE').AsString);
                                end;
                            end;
                        end;
                        if (FieldByName('Tarif').AsFloat < 0.01) AND (FieldByName('PSERIA').AsString = '') then begin //не дубликат
                            if (not LocateTbl.Locate('prcpser;prcpnmb', vararrayof([FieldByName('SERIA').AsString, FieldByName('NUMBER').AsInteger]), [])) AND
                               (not LocateTbl.Locate('seria;number', vararrayof([FieldByName('SERIA').AsString, FieldByName('NUMBER').AsInteger - 1]), [])) then
                                raise_Exception(errText, 'Не найден основной полис для сцепки ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                            else begin
                                if ((LocateTbl.FieldByName('Tarif').AsFloat < 0.01){ AND (LocateTbl.FieldByName('PSERIA').AsString = '')}) OR (Pos('+', LocateTbl.FieldByName('LETTER').AsString) = 0) then
                                    raise_Exception(errText, 'Полис ' + LocateTbl.FieldByName('SERIA').AsString + '/' + LocateTbl.FieldByName('NUMBER').AsString + ' не является основным для сцепки ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                                else begin
                                    if not LocateTbl.FieldByName('STATE').AsInteger in [0, 1] then
                                        raise_Exception(errText, 'Основной полис ' + LocateTbl.FieldByName('NUMBER').AsString + ' для сцепки ' + FieldByName('NUMBER').AsString + ' имеет неправильное состояние', FieldByName('AGCODE').AsString);
                                end;
                            end;
                        end
                    end;

                    BUROTbl.Append;

                    if FieldByName('PNUMBER').AsInteger > 0 then begin
                       LocateTbl.Open;
                       if not LocateTbl.Locate('seria;number', vararrayof([FieldByName('PSERIA').AsString, FieldByName('PNUMBER').AsInteger]), []) then
                           raise_Exception(errText, 'Не найден заменённый полис ' + FieldByName('PNUMBER').AsString + ' см. дубликат ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                       BUROTbl.FieldByName('MPOLICY_N').AsString := NmbToStr(FieldByName('PNUMBER').AsInteger);

                       if FieldByName('LETTER').AsString <> LocateTbl.FieldByName('LETTER').AsString then
                           raise_Exception(errText, 'Не совпадает тип ТС у дубликата ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                       //if LostCount < 256 then begin
                       //    Inc(LostCount);
                       //    LostSer[LostCount] := FieldByName('PSERIA').AsString;
                       //    LostNmb[LostCount] := FieldByName('PNUMBER').AsInteger;
                       //end;
                    end
                    else
                       BUROTbl.FieldByName('MPOLICY_N').AsString := '';
                    //

                    BUROTbl.FieldByName('COMPANY').AsString := '02';
                    BUROTbl.FieldByName('POLICY_NUM').AsString := NmbToStr(FieldByName('NUMBER').AsInteger);
                    if FieldByName('PAYDATE').AsDateTime > FieldByName('DTFROM').AsDateTime then
                        if FieldByName('PNUMBER').AsInteger = 0 then
                           raise_Exception(errText, 'Оплата после начала. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                    BUROTbl.FieldByName('DATE_BEG').AsDateTime := FieldByName('DTFROM').AsDateTime;
                    BUROTbl.FieldByName('DATE_END').AsDateTime := FieldByName('DTTO').AsDateTime;
                    if FieldByName('COUNTRY').AsInteger = 0 then
                        BUROTbl.FieldByName('TERRITIRY').AsString := '2'
                    else
                        BUROTbl.FieldByName('TERRITIRY').AsString := '1';

                    if FieldByName('LETTER').AsString = '' then
                        raise_Exception(errText, 'Не определена буква авто. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                    else
                        BUROTbl.FieldByName('VEH_TYPE').AsString := FieldByName('LETTER').AsString[1];
                    if FieldByName('TEXT').AsString <> '' then
                        BUROTbl.FieldByName('VEH_TYPE').AsString := 'F';

                    BUROTbl.FieldByName('VEH_MARKA').AsString := ReplaceSymbols(FieldByName('MODEL').AsString);
                    ExpStr(BUROTbl.FieldByName('VEH_MARKA').AsString, IsRuss, IsEngl, IsDigit, IsBad, false);
                    if IsBad or IsRuss then
                        raise_Exception(errText, 'Марка авто плохая. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    if IsRussian(FieldByName('AUTONMB').AsString, 'AutoNmb_AllowChars') then
                        raise_Exception(errText, 'В номере авто ' + FieldByName('AUTONMB').AsString + ' русские буквы. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    BUROTbl.FieldByName('VEH_SIGN').AsString := ReplaceStr(ReplaceSymbols(FieldByName('AUTONMB').AsString), ' ', '');
                    ExpStr(BUROTbl.FieldByName('VEH_SIGN').AsString, IsRuss, IsEngl, IsDigit, IsBad, false);
                    if IsBad then
                        raise_Exception(errText, 'Номер ' + BUROTbl.FieldByName('VEH_SIGN').AsString + ' плохой. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    BUROTbl.FieldByName('VEH_BODY').AsString := ReplaceSymbols(FieldByName('BODY').AsString);
                    BUROTbl.FieldByName('VEH_MOTOR').AsString := ReplaceSymbols(FieldByName('MOTOR').AsString);
                    BUROTbl.FieldByName('NAME').AsString := ReplaceSymbols(FieldByName('OWNER').AsString);
                    ExpStr(BUROTbl.FieldByName('NAME').AsString, IsRuss, IsEngl, IsDigit, IsBad, false);
                    if IsBad or IsRuss then
                        raise_Exception(errText, 'Имя страхователя плохое. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    BUROTbl.FieldByName('ADDRESS').AsString := ReplaceSymbols(FieldByName('ADDR').AsString);
                    ExpStr(BUROTbl.FieldByName('ADDRESS').AsString, IsRuss, IsEngl, IsDigit, IsBad, false);
                    if IsBad or IsRuss then
                        raise_Exception(errText, 'Адрес страхователя плохой. Полис ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    if FieldByName('RESIDENT').AsString = 'Y' then
                        BUROTbl.FieldByName('R_N').AsString := 'R'
                    else
                        BUROTbl.FieldByName('R_N').AsString := 'N';

                    if FieldByName('OWNTYPE').AsString = '' then
                    begin
                        raise Exception.Create('Тип страхователя неопределён');
                    end;

                    if FieldByName('OWNTYPE').AsString[1] in [ 'F' , 'I' ] then
                        BUROTbl.FieldByName('F_L').AsString := 'F'
                    else
                        BUROTbl.FieldByName('F_L').AsString := 'L';

                    BUROTbl.FieldByName('PREM_SUM1').AsFloat := FieldByName('PAY1').AsFloat;
                    if FieldByName('PAY1').AsFloat > 0 then begin
                        if FieldByName('PAY1C').AsString = '' then
                            raise_Exception(errText, 'Не определена валюта 1й оплаты в полисе ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                        if FieldByName('PAY1C').AsString = 'BRB' then
                            BUROTbl.FieldByName('CURRENCY1').AsString := 'BYR'
                        else
                            BUROTbl.FieldByName('CURRENCY1').AsString := FieldByName('PAY1C').AsString;
                    end
                    else
                    if (FieldByName('PNUMBER').AsInteger = 0) AND (FieldByName('PAY2C').AsString <> '') then
                        raise_Exception(errText, 'Зачем то определена валюта 1й оплаты в полисе ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                    else
                    if (FieldByName('Text').Value = '') then
                        raise_Exception(errText, 'Полис ' + FieldByName('NUMBER').AsString + ' не определена сумма оплаты 1, но есть валюта.', FieldByName('AGCODE').AsString);

                    BUROTbl.FieldByName('PREM_SUM2').AsFloat := FieldByName('PAY2').AsFloat;
                    if FieldByName('PAY2').AsFloat > 0 then begin
                        if FieldByName('PAY2C').AsString = '' then
                            raise_Exception(errText, 'Не определена валюта 2й оплаты в полисе ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                        if FieldByName('PAY2C').AsString = 'BRB' then
                            BUROTbl.FieldByName('CURRENCY2').AsString := 'BYR'
                        else
                            BUROTbl.FieldByName('CURRENCY2').AsString := FieldByName('PAY2C').AsString;
                    end
                    else
                    if (FieldByName('PNUMBER').AsInteger = 0) AND (FieldByName('PAY2C').AsString <> '') then
                           raise_Exception(errText, 'Зачем то определена валюта 2й оплаты в полисе ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                    //AgQuery.Open;
                    //if AgQuery.Locate('Agent_code', FieldByName('AgCode').AsString, []) = false then
                    //    raise_Exception(errText, 'Не найден агент ' + FieldByName('AgCode').AsString);
                    BUROTbl.FieldByName('INSURER').AsString := ReplaceSymbols(AnsiUpperCase(FieldByName('AGNAME').AsString)); 
                    if IsEnglish(BUROTbl.FieldByName('INSURER').AsString) then
                        raise_Exception(errText, 'Агент страховщика имеет английские буквы ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    if Length(BUROTbl.FieldByName('INSURER').AsString) = 0 then
                        raise_Exception(errText, 'Не указан агент страховщика (длина AGNAME) ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    if FieldByName('PLACE').AsString = '' then
                        raise_Exception(errText, 'Не указано место продажи полиса ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                    if IsEnglish(FieldByName('PLACE').AsString) then
                        raise_Exception(errText, 'Место продажи полиса с англ. буквами' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                    BUROTbl.FieldByName('SALE_PLACE').AsString := UpperCase(FieldByName('PLACE').AsString);
                    s := ReplaceStr(FieldByName('REGTIME').AsString, '-', ':');
                    while(Length(s) < 5) do s := '0' + s;
                    if Pos(';', s) <> 0 then s[Pos(';', s)] := ':';
                    if Pos('.', s) <> 0 then s[Pos('.', s)] := ':';
                    if Pos(':', s) <> 3 then
                        raise_Exception(errText, 'Не правильный формат времени (' + FieldByName('REGTIME').AsString + ') продажи полиса ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString)
                    else
                    if(StrToInt(Copy(s, 1, Pos(':', s) - 1)) > 23) then
                        raise_Exception(errText, 'Время ' + FieldByName('REGTIME').AsString + ' продажи полиса ' + FieldByName('NUMBER').AsString + ' не правильное', FieldByName('AGCODE').AsString);

                    if(StrToInt(Copy(s, Pos(':', s) + 1, 3)) > 59) then
                        raise_Exception(errText, 'Время ' + FieldByName('REGTIME').AsString + ' продажи полиса ' + FieldByName('NUMBER').AsString + ' не правильное', FieldByName('AGCODE').AsString);

                    if FieldByName('PAYDATE').AsDateTime < EncodeDate(2003, 3, 1) then
                        raise_Exception(errText, 'Дата продажи не правильная ' + FieldByName('PAYDATE').AsString + ' ' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);

                    BUROTbl.FieldByName('SALE_DATE').AsDateTime := FieldByName('REGDATE').AsDateTime;
                    BUROTbl.FieldByName('SALE_TIME').AsString := s;
                    //Полис с прицепом
                    if (Pos('+', FieldByName('LETTER').AsString) <> 0) AND (FieldByName('Tarif').AsFloat >= 0.01) then
                        BUROTbl.FieldByName('TPOLICY_N').AsString := NmbToStr(Nmb)
                    else
                        BUROTbl.FieldByName('TPOLICY_N').AsString := '';

                    if FieldByName('STATE').AsString = '1' then begin //Замена
                       LocateTbl.Open;
                       if LocateTbl.Locate('pseria;pnumber', vararrayof([FieldByName('SERIA').AsString, FieldByName('NUMBER').AsInteger]), []) then begin
                           //BUROTbl.FieldByName('MPOLICY_N').AsString := NmbToStr(LocateTbl.FieldByName('NUMBER').AsInteger);
                           if FieldByName('PAYDATE').AsDateTime > LocateTbl.FieldByName('PAYDATE').AsDateTime then
                               raise_Exception(errText, 'Дата продажи заменённого ' + FieldByName('NUMBER').AsString + ' позже чем у дубликата ' + LocateTbl.FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                           if LocateTbl.FieldByName('LETTER').AsString <> FieldByName('LETTER').AsString then
                               raise_Exception(errText, 'тип ТС дубликата должны совпадать ' + LocateTbl.FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                           //if LocateTbl.FieldByName('PAYDATE').AsDateTime <> LocateTbl.FieldByName('REGDATE').AsDateTime then
                           //    raise_Exception(errText, 'Дата выдачи и оплаты дубликата должны совпадать ' + LocateTbl.FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                           //end
                       end
                       else
                           raise_Exception(errText, 'Не найден дубликат полиса ' + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString, FieldByName('AGCODE').AsString);
                    end;
                    //else
                    //    BUROTbl.FieldByName('MPOLICY_N').AsString := '';

                    BUROTbl.Post;
                    Inc(toFirst);
                end;
                //else
                if not (FieldByName('State').AsString[1] in ['0', '1', '2', '3', '4', '5']) then
                    raise_Exception(errText, 'Поле State для полиса  ' + FieldByName('NUMBER').AsString + ' имеет не правильное состояние', FieldByName('AGCODE').AsString);
                Next;

                //if not Stopped then
                //    Stopped := Eof;
            end;
        except
          on E : Exception do begin
            raise_Exception(errText, 'Ошибка экспорта данных в БЮРО (' + FieldByName('NUMBER').AsString + ')'#13#10 + E.Message, FieldByName('AGCODE').AsString);
          end;
        end;

        Close;
    end;
    AgQuery.Close;
    ExpInfo.Caption := '';
    
    if BG_Buro.IsTest.Checked or ((errText <> '') AND ((GetAsyncKeyState(VK_SHIFT) AND $8000) = 0)) then begin
        BUROTbl.EmptyTable;
        BUROTbl2.EmptyTable;
    end;
    BUROTbl.Close;
    BUROTbl2.Close;
    NExportTbl.Close;
    ProgressBar.Position := 0;
    Screen.Cursor := crDefault;

    if errText <> '' then begin
        ShowInfo.Caption := 'Ошибки экспорта в БЮРО';
        ShowInfo.InfoPanel.Text := errText;
        ShowInfo.ShowModal;
    end
    else begin
        ReportStr := ReportStr + 'В первую таблицу попало ' + InttoStr(toFirst) + ' полисов'#13;
        ReportStr := ReportStr + 'В вторую таблицу попало ' + InttoStr(toSecond) + ' полисов'#13;
        MessageDlg(ReportStr, mtInformation, [mbOk], 0);
    end;
    *)
end;

procedure TBelGreenForm.PageControlChange(Sender: TObject);
var
    s : string;
begin
    StatisticTxt.Text := '';
    if PageControl.ActivePage = StatPage then begin
        try
            Screen.Cursor := crHourglass;
            StatisticTxt.Lines.Add('Отчёт создаётся на основании фильтра');
            with WorkSQL, WorkSQL.SQL do begin
                //if RepDtFrom.Checked OR RepDtTo.Checked then begin
                    s := 'ОТЧЁТНЫЙ ПЕРИОД ';
                    if RepDtFrom.Checked then
                        s := s + ' C ' + DateToStr(RepDtFrom.Date);
                    if RepDtTo.Checked then
                        s := s + ' ПО ' + DateToStr(RepDtTo.Date);
                    if RepDtFrom.Checked and RepDtto.Checked then
                        s := s + #13#10' за ' + IntToStr(Round(RepDtTo.Date-RepDtFrom.Date+1)) + ' дней';

                    if not (RepDtFrom.Checked or RepDtto.Checked) then
                        s := s + 'НЕ УКАЗАН';

                    StatisticTxt.Lines.Add(s);
                    StatisticTxt.Lines.Add('');
                //end;

                Close;
                Clear;
                Add('SELECT SUM(TARIF) AS S');
                Add('FROM BELGREEN');
                Add('WHERE TARIF>0 AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('Тариф');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString);
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(PAY1) AS S, PAY1C AS CUR');
                Add('FROM BELGREEN');
                Add('WHERE PAY1C<>'''' AND ' + GetFilter());
                Add('GROUP BY PAY1C');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Первая часть');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(PAY2) AS S, PAY2C AS CUR');
                Add('FROM BELGREEN');
                Add('WHERE PAY2C<>'''' AND ' + GetFilter());
                Add('GROUP BY PAY2C');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Вторая часть');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;
                Close;
                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM BELGREEN');
                Add('WHERE STATE=''2'' AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Испорчено');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('C').AsString);
                    Next;
                end;
                Close;
                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM BELGREEN');
                Add('WHERE STATE=''3'' AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Утеряных полисов');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('C').AsString);
                    Next;
                end;
                Close;
                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM BELGREEN');
//                Add('WHERE STATE=''1'' AND ' + GetFilter());
                Add('WHERE PNUMBER>0 AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Дубликатов');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('C').AsString);
                    Next;
                end;
                Close;
                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM BELGREEN');
                Add('WHERE TEXT<>'''' AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Сцепок');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('C').AsString);
                    Next;
                end;

                if RepDtFrom.Checked and RepDtto.Checked then begin
                    Close;
                    Clear;
                    Add('SELECT COUNT(*), SUM(Ret1), Ret1Cur');
                    Add('FROM BELGREEN');
                    Add('WHERE RETDATE >= ''' + DateToStr(RepDtFrom.Date) + ''' ');
                    Add('AND RETDATE <= ''' + DateToStr(RepDtTo.Date) + ''' ');
                    Add('GROUP BY ret1cur');
                    Open;
                    StatisticTxt.Lines.Add('');
                    StatisticTxt.Lines.Add('Расторгнуто');
                    while not Eof do begin
                        StatisticTxt.Lines.Add(Fields[0].AsString + ' шт.  ' + Fields[1].AsString  + '  ' + Fields[2].AsString);
                        Next;
                    end;
                end
            end;
        except
          on E : Exception do begin
            StatisticTxt.Text := E.Message;
          end;
        end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TBelGreenForm.RepListChange(Sender: TObject);
var
    INI : TIniFile;
begin
    REPQuery.Close;
    INI := TIniFile.Create('blreps.ini');
    ParamsText.Hint := INI.ReadString(BG_SECT, 'QueryParams' + IntToStr(RepList.ItemIndex + 1), '');
    ParamsText.Text := INI.ReadString(BG_SECT, 'LastParams' + IntToStr(RepList.ItemIndex + 1), '');
    if ParamsText.Text = '' then
        ParamsText.Text := ParamsText.Hint; 
    INI.Free;
end;

procedure TBelGreenForm.StartButtonClick(Sender: TObject);
{var
    INI : TIniFile;
    i : integer;
    s, param : string;}
var
    s : string;
begin
    s := '';
    if IsUseFilter.Checked then s := GetFilter();
    ExecLittleReport('blreps.ini', BG_SECT, ParamsText.Text, RepList.ItemIndex + 1, REPQuery, GetFilter());
{
    REPQuery.Close;
    REPQuery.SQL.Clear;
    INI := TIniFile.Create('blreps.ini');
    for i := 1 to 100 do begin
        s := INI.ReadString(BG_SECT, 'QueryText' + IntToStr(RepList.ItemIndex + 1) + '_' + IntToStr(i), '');
        if s <> '' then
            REPQuery.SQL.Add(s)
        else
            break;
    end;
    //1.1.2003;2.2.1003
    for i := 0 to REPQuery.MacroCount - 1 do begin
        s := INI.ReadString(BG_SECT, REPQuery.Macros[i].Name, '');
        param := ExtractWord(i + 1, ParamsText.Text, [';']);
        if s = 'D' then begin
            while Pos('/', param) <> 0 do
                param[Pos('/', param)] := '.';
            while Pos('-', param) <> 0 do
                param[Pos('-', param)] := '.';
            try
                if (StrToInt(ExtractWord(1, param, ['.'])) > 31) OR
                   (StrToInt(ExtractWord(2, param, ['.'])) > 12) OR
                   (StrToInt(ExtractWord(3, param, ['.'])) < 1900) then
                    raise Exception.Create('');
            except
                MessageDlg('Дата ' + ExtractWord(i + 1, ParamsText.Text, [';']) + ' введена не правильно', mtInformation, [mbOk], 0);
                exit;
            end;
            param := '''' + param + '''';
        end;
        if s = 'S' then param := '''' + param + '''';
        REPQuery.MacroByName(REPQuery.Macros[i].Name).AsString := param;
    end;
    INI.WriteString(BG_SECT, 'LastParams' + IntToStr(RepList.ItemIndex + 1), ParamsText.Text);
    INI.Free;

    try
        Screen.Cursor := crHourGlass;
        REPQuery.Open;
    except
        on E : Exception do
            MessageDlg('Ошибка выполнения запроса'#13 + e.Message, mtInformation, [mbOk], 0);
    end;
    Screen.Cursor := crDefault; }
end;

procedure TBelGreenForm.Button2Click(Sender: TObject);
begin
    if REPQuery.Active then
        DatSetToExcel(REPQuery, IsNumbers.Checked);
end;

procedure TBelGreenForm.btnRptMonthClick(Sender: TObject);
var
    D, M, Y : WORD;
begin
    DecodeDate(Date, Y, M, D);
    RepDtFrom.Date := EncodeDate(Y, M, 1);
    RepDtTo.Date := Date;
    PageControlChange(nil);
    ApplyClick(nil);
    PageControl.ActivePageIndex := 2;
end;

procedure TBelGreenForm.AddBtnClick(Sender: TObject);
begin
{    if (not UbTable.Active) then exit;
    UbitokFrm.DataQuery := nil;
    UbitokFrm.SaveQuery := UbTable;
    UbTable.Append;
    MainGrid.Options := MainGrid.Options + [dgRowSelect];
    UbitokFrm.MinDate := MainQueryDtfrom.AsDateTime;
    UbitokFrm.MaxDate := MainQueryDtTo.AsDateTime;
    if UbitokFrm.ShowModal = mrOk then begin
        try
            UbTable.FieldByName('Ser').AsString := MainQuerySeria.AsString;
            UbTable.FieldByName('Nmb').AsInteger := MainQueryNumber.AsInteger;
            UbTable.Post;
        except
          on E : Exception do begin
            MessageDlg('Ошибка записи убытка'#13 + E.Message, mtError, [mbOk], 0);
            UbTable.Cancel;
          end;
        end;
        UbTable.Close;
        MainQueryAfterScroll(nil);
    end
    else
        UbTable.Cancel;
    MainGrid.Options := MainGrid.Options - [dgRowSelect];   }
end;

procedure TBelGreenForm.EditBtnClick(Sender: TObject);
begin
{    if (not UbTable.Active) OR (UbTable.RecordCount = 0) then exit;
    UbitokFrm.DataQuery := UbTable;
    UbTable.Edit;
    UbitokFrm.SaveQuery := UbTable;
    UbitokFrm.MinDate := MainQueryDtfrom.AsDateTime;
    UbitokFrm.MaxDate := MainQueryDtTo.AsDateTime;
    if UbitokFrm.ShowModal = mrOk then begin
        try
            UbTable.Post;
        except
          on E : Exception do begin
            MessageDlg('Ошибка записи убытка'#13 + E.Message, mtError, [mbOk], 0);
            UbTable.Cancel;
          end;
        end;
        UbTable.Close;
        MainQueryAfterScroll(nil);
    end
    else
        UbTable.Cancel; }
end;

procedure TBelGreenForm.DelBtnClick(Sender: TObject);
begin
{    if (not UbTable.Active) OR (UbTable.RecordCount = 0) then exit;
    if MessageDlg('Удалить убыток?', mtInformation, [mbYes, mbNo], 0) = mrYes then
      UbTable.Delete;      }
end;

procedure TBelGreenForm.MainQueryAfterScroll(DataSet: TDataSet);
begin
//    UbTable.Filter := 'SER=''' + MainQuerySeria.AsString + ''' AND NMB=' + MainQueryNumber.AsString;
//    UbTable.Open;
end;

procedure TBelGreenForm.N8Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по белорусской зелёной карте';
    About.ShowModal
end;

procedure TBelGreenForm.N111Click(Sender: TObject);
begin
    CreateSimpleMagazine(WorkSQL, 'SELECT SERIA,NUMBER,DTFROM INSSTART,DTTO INSEND,OWNER INSURER, MODEL INSURER2, 1 ALLSUMMA, '''' ALLSUMMACURr, ' +
                                  'pay1 sum1, pay1c sum1curr, pay2 sum2, pay2c sum2curr, paydate PAYMENTDATE, pcnt AgPercent, AgType AgUr FROM belgreen T WHERE paydate is not null and pay1 >= 0.01 and paydate BETWEEN :D1 AND :D2', 'Белорусская зелёная карта');
    Application.BringToFront;                              
end;

procedure TBelGreenForm.ImportAlvena(DBFile : string);
var
    TableMain, TablePay2 : TTable;
    BelGreen : TTable;
    i, n : integer;
    Seria, Number : string;
    NSeria, NNumber : string;
    dv : integer;
    xType : boolean;
    Field : TField;
    IsNewRecord : boolean;
    PolisId, RateCurrency : string;
    TableTarif : integer;
    NAll, NPos : integer;
    errStr : string;
    SuccsessCount : integer;
    UpdateCount : integer;
    UpdateFlag : boolean;
    IsResident : boolean;
    IsRussian, IsEnglish, IsDigit, IsBad : boolean;
    fld : integer;
begin
    errStr := '';
    TableMain := TTable.Create(self);
    TablePay2 := TTable.Create(self);
    BelGreen := TTable.Create(self);
    BelGreen.TableName := 'belgreen.db';
    BelGreen.DatabaseName := 'BelGreenDatabase';

    for i := Length(DBFile) downto 1 do
        if DBFile[i] = '\' then break;

    if PcntForm.ShowModal <> mrOk then exit;
//    if MessageDlg('7% Альвене?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
  //      exit;

    UpdateCount := 0;

    try
        TableMain.TableName := DBFile;
        n := Pos('.DB', AnsiUpperCase(DBFile)) - 1;
        TablePay2.TableName := Copy(DBFile, 1, n) + '$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '_$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '__$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '___$.DB';
        TableMain.Open;
        TablePay2.Open;
        BelGreen.Open;

        Screen.Cursor := crHourGlass;
        NAll := TableMain.RecordCount;
        NPos := 1;
        SuccsessCount := 0;
        while not TableMain.Eof do begin
            StatusBar.Caption := InttoStr(NPos) + ' из ' + IntToStr(NAll);
            StatusBar.Update;
            NPos := NPos + 1;

            try
                if not TableMain.FieldByName('Состояние').AsInteger in [1, 2, 3] then
                    raise Exception.Create('Полис имеет непонятное состояние.');
                if (TableMain.FieldByName('Состояние').AsInteger = 2) AND
                   (TableMain.FieldByName('Расторжение действия').IsNull) then
                    raise Exception.Create('Полис расторгнут но не указана дата расторжения');
                //if (not TableMain.FieldByName('Расторжение действия').IsNull) AND
                //   (TableMain.FieldByName('Состояние').AsInteger <> 2) then
                //    raise Exception.Create('Полис имеет дату расторжения но по состоянию ДЕЙСТВУЕТ');
                //if TableMain.FieldByName('Адрес_Страна').AsString <> 'Беларусь' then
                //    raise Exception.Create('Полис зарегистрирован не на резидента.');
                //if TableMain.FieldByName('Скидка тарифа').AsFloat <> 0 then
                //    raise Exception.Create('Импорт полиса  со скидкой пока не поддержимается.');
                Field := TableMain.FieldByName('Номер полиса');
                PolisId := TableMain.FieldByName('Номер полиса').AsString;

                if PolisId = 'BY02/0411816' then
                    messagebeep(0);

                dv := Pos('/', Field.AsString);
                if dv = 0 then
                    raise Exception.Create('Не правильно указана номер/серия ' + Field.AsString);
                Seria := Copy(Field.AsString, 1, dv - 1);
                Number := Copy(Field.AsString, dv + 1, 1024);

                IsNewRecord := false;
                UpdateFlag := false;
                if not BelGreen.Locate('Seria;Number', vararrayof([Seria, StrToInt(Number)]), []) then begin
                    BelGreen.Append;
                    BelGreen.ClearFields;
                    IsNewRecord := true;
                    BelGreen.FieldByName('Seria').AsString := Seria;
                    BelGreen.FieldByName('Number').AsInteger := StrToInt(Number);
                end
                else begin
                    UpdateFlag := true;
                    //Сравнить агента!!!
                    if BelGreen.FieldByName('AgCode').AsString <> 'АЛВН' then
                        raise Exception.Create('Полис ' + Field.AsString + ' уже зарегистрирован не на АЛЬВЕНУ');
                    //!!if BelGreen.FieldByName('State').AsInteger = 2 then begin
                    //!!    if TableMain.FieldByName('Состояние').AsInteger <> 2 then 444
                    //!!        raise Exception.Create('Полис в БД ' + Field.AsString + ' досрочно завершён или расторгнут.');
                    //!!end;
                    //!!if (TableMain.FieldByName('Состояние').AsInteger = 3) AND (Mandator.FieldByName('State').AsInteger <> 1) OR
                    //!!    (TableMain.FieldByName('Состояние').AsInteger <> 3) AND (Mandator.FieldByName('State').AsInteger = 1) then
                    //!!    raise Exception.Create('Полис ' + Field.AsString + ' досрочно конфликт состояний.');
                    Belgreen.Edit;
                end;

//                if Number = '4848300' then
//                 MessageBeep(0);

                BelGreen.FieldByName('AgCode').AsString := 'АЛВН';
                BelGreen.FieldByName('AGNAME').Value := 'АЛЬВЕНА';
                BelGreen.FieldByName('UPD').AsDateTime := Date;
                BelGreen.FieldByName('RegDate').AsDateTime := TableMain.FieldByName('Дата выдачи').AsDateTime;
                BelGreen.FieldByName('RegTime').AsString := CalcAlvenaTime(TableMain.FieldByName('Время выдачи полиса').AsInteger);

                if TableMain.FieldByName('Состояние').AsInteger = 3 then begin
                    BelGreen.FieldByName('Owner').AsString := 'ИСПОРЧЕН';
                    BelGreen.FieldByName('State').AsInteger := 2;
                    BelGreen.Post;
                    TableMain.Next;
                    SuccsessCount := SuccsessCount + 1;
                    if UpdateFlag then Inc(UpdateCount);
                    continue;
                end;

                BelGreen.FieldByName('Autonmb').Value := TableMain.FieldByName('Номер машины').Value;
                BelGreen.FieldByName('Model').Value := TableMain.FieldByName('Марка машины').Value;
                //BelGreen.FieldByName('Body').Value := TableMain.FieldByName('Марка машины').Value;
                BelGreen.FieldByName('Motor').Value := TableMain.FieldByName('Номер шасси').Value;
                BelGreen.FieldByName('Owner').Value := TableMain.FieldByName('Страхователь').Value;
                //BelGreen.FieldByName('Org').Value := TableMain.FieldByName('Страхователь').Value;
                BelGreen.FieldByName('Addr').Value := PrepareStr(
                                                            TableMain.FieldByName('Адрес_Город').AsString,
                                                            TableMain.FieldByName('Адрес_Улица').AsString,
                                                            TableMain.FieldByName('Адрес_Дом').AsString
                                                            //TableMain.FieldByName('Адрес_Телефон').AsString
                                                            );
                BelGreen.FieldByName('Dtfrom').Value := TableMain.FieldByName('Начало действия').Value;
                BelGreen.FieldByName('Dtto').Value := TableMain.FieldByName('Окончание действия').Value;
                BelGreen.FieldByName('Period').Value := CalcPeriod(BelGreen.FieldByName('Dtfrom').AsDateTime,
                                                                   BelGreen.FieldByName('Dtto').AsDateTime);
                if BelGreen.FieldByName('Period').Value = '15 суток' then
                    BelGreen.FieldByName('Period').Value := '15 дней';

                BelGreen.FieldByName('Letter').Value := UpperCase(TableMain.FieldByName('Категория машины').Value);
                if (TableMain.FieldByName('Парный полис').AsString <> '') and ((TableMain.FieldByName('Категория машины').Value = 'C') or (TableMain.FieldByName('Категория машины').Value = 'F')) then BelGreen.FieldByName('Letter').Value := 'C+F2'
                else
                if TableMain.FieldByName('Категория машины').Value = 'f' then BelGreen.FieldByName('Letter').Value := 'F1'
                else
                if TableMain.FieldByName('Категория машины').Value = 'F' then BelGreen.FieldByName('Letter').Value := 'F2';

                case TableMain.FieldByName('Филиал').AsInteger of
                    1 : BelGreen.FieldByName('Place').Value := 'МИНСК';
                    2 : BelGreen.FieldByName('Place').Value := 'БРЕСТ';
                    3 : BelGreen.FieldByName('Place').Value := 'ВИТЕБСК';
                    4 : BelGreen.FieldByName('Place').Value := 'ГОМЕЛЬ';
                    5 : BelGreen.FieldByName('Place').Value := 'ГРОДНО';
                    6 : BelGreen.FieldByName('Place').Value := 'СОЛИГОРСК';
                    7 : BelGreen.FieldByName('Place').Value := 'ЖОДИНО';
                    8 : BelGreen.FieldByName('Place').Value := 'ПРЕДСТАВИТЕЛЬСТВО';
                    9 : BelGreen.FieldByName('Place').Value := 'ПИНСК';
                    10 : BelGreen.FieldByName('Place').Value := 'МОГИЛЕВ';
                    11 : BelGreen.FieldByName('Place').Value := 'БЕРЕСТОВИЦА';
                    12 : BelGreen.FieldByName('Place').Value := 'МОЗЫРЬ';
                    13 : BelGreen.FieldByName('Place').Value := 'БАРАНОВИЧИ';
                    else raise Exception.Create('Полис в БД ' + Field.AsString + ' не определено место выдачи.');
                end;

                if TableMain.FieldByName('Валюта с_тарифа').AsString <> 'EUR' then
                    raise Exception.Create('Полис ' + Field.AsString + ' имеет тариф не EUR');
                BelGreen.FieldByName('Tarif').Value := TableMain.FieldByName('Страховой тариф').Value;
                BelGreen.FieldByName('Pay1').Value := TableMain.FieldByName('Страховой платеж').Value;
                BelGreen.FieldByName('Pay1c').Value := TableMain.FieldByName('Валюта с_платежа').Value;
                if TableMain.FieldByName('Валюта с_платежа').AsString = 'BYB' then
                    BelGreen.FieldByName('Pay1c').AsString := 'BRB';
                BelGreen.FieldByName('Paydate').Value := TableMain.FieldByName('Дата с_платежа').Value;

                if not (TableMain.FieldByName('Форма с_платежа').AsString[1] in ['*', 'Б']) then begin

                    TablePay2.Close;
                    TablePay2.Filter := '[Номер полиса]=''' + Field.AsString + ''' AND [Вид платежа]=1';
                    TablePay2.Filtered := True;
                    TablePay2.Open;

                    BelGreen.FieldByName('Pay1c').AsString := '';

                    while not TablePay2.Eof do begin
                          if (TablePay2.RecordCount = 0) or (TablePay2.RecordCount > 2) then
                              raise Exception.Create('Платежи по полису не найдены или их слишком много ' + IntToStr(TablePay2.RecordCount));
                          //if (TablePay2.RecordCount = 1) and (TablePay2.FieldByName('Дата платежа').AsDateTime <> BelGreen.FieldByName('Paydate').AsDateTime) then
                          //    raise Exception.Create('Не найден первый платёж');

                          if BelGreen.FieldByName('Pay1c').AsString = '' then begin
                              BelGreen.FieldByName('Pay1').Value := TablePay2.FieldByName('Сумма платежа').Value;
                              BelGreen.FieldByName('Pay1c').Value := TablePay2.FieldByName('Валюта платежа').Value;
                              if TablePay2.FieldByName('Валюта платежа').AsString = 'BYB' then
                                  BelGreen.FieldByName('Pay1c').AsString := 'BRB';
                          end
                          else begin
                              BelGreen.FieldByName('Pay2').Value := TablePay2.FieldByName('Сумма платежа').Value;
                              BelGreen.FieldByName('Pay2c').Value := TablePay2.FieldByName('Валюта платежа').Value;
                              if TablePay2.FieldByName('Валюта платежа').AsString = 'BYB' then
                                  BelGreen.FieldByName('Pay12').AsString := 'BRB';
                          end;
                          TablePay2.Next
                    end
                end;

                //if (BelGreen.FieldByName('Pay1c').Value <> '') and (BelGreen.FieldByName('Pay1').IsNull) then
                //    if FieldByName('PNUMBER').AsInteger aa then BelGreen.FieldByName('Pay1c').Value = '';
                    //raise Exception.Create('Полис ' + Field.AsString + ' не определена сумма оплаты 1, но есть валюта.');
                //if (BelGreen.FieldByName('Pay2c').Value <> '') and (BelGreen.FieldByName('Pay2').IsNull) then
                //    raise Exception.Create('Полис ' + Field.AsString + ' не определена сумма оплаты 2, но есть валюта.');

                BelGreen.FieldByName('Pcnt').Value := PcntForm.PcntBefore.Value;
                if BelGreen.FieldByName('Paydate').Value >= PcntForm.Date.Date then BelGreen.FieldByName('Pcnt').Value := PcntForm.PcntAfter.Value;

                BelGreen.FieldByName('State').Value := 0;
                //if TableMain.FieldByName('Дубликат полиса').AsString <> '' then begin
                //    BelGreen.FieldByName('State').AsInteger := 4;
                //end;
                //!!
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'A' then
                    BelGreen.FieldByName('Type').Value := 0;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'B' then
                    BelGreen.FieldByName('Type').Value := 4;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'C' then
                    BelGreen.FieldByName('Type').Value := 2;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'C+F2' then
                    BelGreen.FieldByName('Type').Value := 6;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'E' then
                    BelGreen.FieldByName('Type').Value := 5;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'F1' then
                    BelGreen.FieldByName('Type').Value := 1;
                if UpperCase(Belgreen.FieldByName('Letter').AsString) = 'F2' then
                    BelGreen.FieldByName('Type').Value := 3;

                BelGreen.FieldByName('Repdate').Value := BelGreen.FieldByName('Paydate').Value;
                BelGreen.FieldByName('Country').Value := 0;
                if TableMain.FieldByName('Страны действия').AsString = 'A' then
                    BelGreen.FieldByName('Country').Value := 1;
                if (TableMain.FieldByName('Парный полис').AsString <> '') AND (UpperCase(TableMain.FieldByName('Категория машины').Value) = 'F') then
                    BelGreen.FieldByName('Text').Value := 'Сцепка';
                BelGreen.FieldByName('AgType').Value := 'U';
                BelGreen.FieldByName('OwnType').Value := 'U';
                if (TableMain.FieldByName('Характеристика клиента').AsInteger AND $0001) = 0 then
                    BelGreen.FieldByName('OwnType').Value := 'F';
                BelGreen.FieldByName('Resident').Value := 'Y';
                if (TableMain.FieldByName('Характеристика клиента').AsInteger and $0002) = $0002 then
                    BelGreen.FieldByName('Resident').Value := 'N';

                if not TableMain.FieldByName('Расторжение действия').IsNull then begin
                    BelGreen.FieldByName('RetDate').Value := TableMain.FieldByName('Расторжение действия').Value;
                    //BelGreen.FieldByName('Ret1').Value :=
                    //BelGreen.FieldByName('Ret1Cur').Value :=
                    BelGreen.FieldByName('State').Value := 4;
                end;

                xType := (BelGreen.FieldByName('Type').AsString = '6') OR
                         (BelGreen.FieldByName('Type').AsString = 'O') OR //I + 6
                         (BelGreen.FieldByName('Type').AsString = 'G'); //A +6
                if (xType) AND (UpperCase(TableMain.FieldByName('Категория машины').AsString) = 'C') then begin
                    dv := Pos('/', TableMain.FieldByName('Парный полис').Value);
                    if dv = 0 then
                        raise Exception.Create('Не правильно указана номер/серия парного полиса' + Field.AsString);
                    BelGreen.FieldByName('PrcpSer').Value := Copy(TableMain.FieldByName('Парный полис').Value, 1, dv - 1);
                    BelGreen.FieldByName('PrcpNmb').Value := Copy(TableMain.FieldByName('Парный полис').Value, dv + 1, 1024);
                end;

                for fld := 0 to BelGreen.FieldCount - 1 do
                    if BelGreen.Fields[i].DataType = ftDate then
                        if not BelGreen.Fields[i].IsNull then
                            if (BelGreen.Fields[i].AsDateTime < (Date - 400)) OR (BelGreen.Fields[i].AsDateTime > (Date + 400)) then
                                raise Exception.Create('Полис ' + Field.AsString + ' дата не в диапазоне ' + BelGreen.Fields[i].FieldName + '=' + BelGreen.Fields[i].AsString);

                BelGreen.Post;
                SuccsessCount := SuccsessCount + 1;
                if UpdateFlag then Inc(UpdateCount);

                {
                if TableMain.FieldByName('Дубликат полиса').AsString <> '' then begin
                    Field := TableMain.FieldByName('Дубликат полиса');
                    dv := Pos('/', Field.AsString);
                    if dv = 0 then
                        raise Exception.Create('Не правильно указана номер/серия дубликата' + Field.AsString);
                    NSeria := Copy(Field.AsString, 1, dv);
                    NNumber := Copy(Field.AsString, dv + 1, 1024);
                    if not Belgreen.Locate('Seria;Number', vararrayof([NSeria, StrToInt(NNumber)]), []) then
                        errStr := errStr + PolisId + ' Не найден полис дубликат ' + Field.AsString + #13#10
                    else begin
                        Belgreen.Edit;
                        Belgreen.FieldByName('PSeria').AsString := NSeria;
                        Belgreen.FieldByName('PNumber').AsInteger := StrToInt(NNumber);
                        Belgreen.FieldByName('Duptype').AsInteger := 0;
                        Belgreen.Post;
                    end;
                end;
                }
            except
                on  E : Exception do begin
                    if IsNewRecord then
                        Belgreen.Delete
                    else
                    if Belgreen.State in [dsEdit] then
                        BelGreen.Cancel;

                    errStr := errStr + PolisId + ' ' + e.Message + #13#10;
                end;
            end;

            TableMain.Next;
        end;
    except
        on  E : Exception do begin
            MessageDlg('Ошибка импорта. ' + e.Message, mtInformation, [mbOk], 0);
        end
    end;

    Screen.Cursor := crDefault;
    BelGreen.Free;
    TableMain.Free;
    TablePay2.Free;
//    TableOwn.Free;


    StatusBar.Caption := '';

    MessageDlg(Format('Импортировано %d из %d. Из них обновлено %d', [SuccsessCount, NAll, UpdateCount]), mtInformation, [mbOk], 0);

    if errStr <> '' then begin
        ShowInfo.Caption := 'Результат импорта данных Альвены';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end;
end;


procedure TBelGreenForm.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top + SpeedButton.Height ;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TBelGreenForm.MenuItem1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('БелЗК');
     if GetTemplName.ShowModal = mrOk then begin
         for i := 0 to 100 do
           if INI.ReadString('AgTemplates', IntToStr(i), '') = '' then break;
         index := i;
         str := '';
         for i := 0 to ListAgents.Items.Count - 1 do
             if ListAgents.Checked[i] then
                 str := str + listAgCodes[i] + ',';
         INI.WriteString('AgTemplates', IntToStr(index), str);
         INI.WriteString('AgTemplates', 'Name' + IntToStr(index), GetTemplName.Name.Text);
         ListTemplates.ItemIndex := ListTemplates.Items.Add(GetTemplName.Name.Text);
     end;
     INI.Free;
end;

procedure TBelGreenForm.ListTemplatesChange(Sender: TObject);
var
     i, fnd_idx, fnd_token : integer;
     INI : TRegIniFile;
     str : string;
     Code : string;
begin
     //if ListTemplates.ItemIndex = 0 then begin
       for i := 0 to ListAgents.Items.Count - 1 do
            ListAgents.Checked[i] := false;
       //exit;
     //end;

     Code := '';
     INI := TRegIniFile.Create('БелЗК');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str = ListTemplates.Text then begin
             str := INI.ReadString('AgTemplates', IntToStr(i), '');

             for fnd_token := 1 to Length(str) do begin
                 if str[fnd_token] = ',' then begin
                     for fnd_idx := 0 to ListAgents.Items.Count - 1 do
                         if listAgCodes[fnd_idx] = Code then
                             ListAgents.Checked[fnd_idx] := true;

                     Code := '';
                 end
                 else
                     Code := Code + str[fnd_token];
             end;
         end;
     end;
     INI.Free;

     ListAgentsClick(nil);
end;

procedure TBelGreenForm.ListAgentsClick(Sender: TObject);
var
     Count, i : integer;
begin
     Count := 0;
     for i := 0 to ListAgents.Items.Count - 1 do
         if ListAgents.Checked[i] then
             Inc(Count);
     CntChecked.Caption := 'Помечено ' + IntToStr(Count);

     if Sender <> nil then
         ListTemplates.ItemIndex := -1;
end;

procedure TBelGreenForm.InputSumClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT PAY1,PAY1C,PAYDATE,NUMBER,REPDATE FROM BELGREEN WHERE PNUMBER=0 AND PAYDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT PAY2,PAY2C,PAYDATE,NUMBER,REPDATE FROM BELGREEN WHERE PNUMBER=0 AND PAYDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkSQL);
end;

procedure TBelGreenForm.N11Click(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    //RepParamsAgnt.filtParam := 'BELGREEN > 0';
    if RepParamsAgnt.ShowModal <> mrOk then exit;
    if RepParamsAgnt.listAgnt.ItemIndex = -1 then
    begin
        MessageDlg('Выделите одного агента', mtInformation, [mbOk], 0);
        exit;
    end;

    DtFrom := min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    DtTo := max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    sqlText := 'SELECT PAY1,PAY1C,REPDATE,NUMBER,PCNT FROM BELGREEN WHERE AGCODE=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PNUMBER=0 AND PAYDATE IS NOT NULL AND REPDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT PAY2,PAY2C,REPDATE,NUMBER,PCNT FROM BELGREEN WHERE AGCODE=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PNUMBER=0 AND PAYDATE IS NOT NULL AND REPDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Items[RepParamsAgnt.listAgnt.ItemIndex], WorkSQL);
end;

procedure TBelGreenForm.ImpAlvenaClick(Sender: TObject);
begin
    if not OpenDialogPx.Execute then exit;
    ImportAlvena(OpenDialogPx.FileName);
end;

function DecodeR(s : string) : string;
begin
    DecodeR := '?';
    if s = 'Y' then DecodeR := '1';
    if s = 'N' then DecodeR := '2';
end;

//нал безнал по типу лица
function DecodeNBByOwner(s : string) : string;
begin
    DecodeNBByOwner := '?';
    if s = 'U' then DecodeNBByOwner := '2';
    if (s = 'F') OR (s = 'I') then DecodeNBByOwner := '1';
end;

//нал безнал указан в типе
function DecodeNBByOwner2(s : string; s2 : string) : string;
begin
    DecodeNBByOwner2 := '?';
    if s2 < 'A' then //old code system
    begin
        DecodeNBByOwner2 := DecodeNBByOwner(s);
    end
    else
    begin //new code system
        if s2 >= 'I' then DecodeNBByOwner2 := '1'
        else
        if s2 >= 'A' then DecodeNBByOwner2 := '2'
    end
end;

function DecodeBodyMotor(body, motor : string) : string;
begin
    if (body = '') then DecodeBodyMotor :=  motor + '|' + body
    else DecodeBodyMotor := body + '|' + motor;
end;

procedure TBelGreenForm.DelMenuClick(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('БелЗк');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

function DecodeLetter(s, pricep : string) : string;
begin
    DecodeLetter := s;
    if (s = 'C+F2') or (s = 'C+F') then begin
        if pricep <> '' then
            DecodeLetter := 'F2'
        else
            DecodeLetter := 'C+F';
    end;
end;

function DecodeBGVariant(s : string) : string;
var
    Ini : TIniFile;
begin
    DecodeBGVariant := '?';
    if s = '0' then DecodeBGVariant := '2'
    else if s = '1' then DecodeBGVariant := '1'
    else begin
        INI := TIniFile.Create(BLANK_INI);
        DecodeBGVariant := INI.ReadString(BG_SECT, 'InsVariant' + s, '?');
        INI.Free;
    end;
end;

function DecodeAgentName(s : string) : string;
begin
    if lastAgCode = s then begin
        DecodeAgentName := lastAgName;
        exit;
    end;
    lastAgCode := s;
    BelGreenForm.WorkSQL2.Close;
    BelGreenForm.WorkSQL2.SQL.Clear;
    BelGreenForm.WorkSQL2.SQL.Add('SELECT NAME FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    BelGreenForm.WorkSQL2.Open;
    if BelGreenForm.WorkSQL2.Eof then lastAgName := s
    else lastAgName := ReplaceSymbols(BelGreenForm.WorkSQL2.FieldByName('NAME').AsString);
    DecodeAgentName := lastAgName;
end;

function SConv(s : string) : string;
var
    INI : TIniFile;
begin
    s := UpperCase(s);
    if sdecode = '' then begin
        INI := TIniFile.Create(BLANK_INI);
        sdecode := INI.ReadString(BG_SECT, 'SeriaDecode', 'BY02/');
        INI.Free;
    end;
    SConv := s;
    if s = 'BY02' then SConv := sdecode;
end;

function IfEmpty(IsEmpty : boolean; Value : string) : string;
begin
    IfEmpty := Value;
    if IsEmpty then IfEmpty := '';
end;

procedure TBelGreenForm.MenuGloboClick(Sender: TObject);
var
    F1 : TextFile;
    X : TextFile;
    F3 : TextFile;
    F9 : TextFile;
    F8, A : TextFile;
    F10 : TextFile;
    F11 : TextFile;
    F12 : TextFile;
    F13 : TextFile;
    All, Curr : integer;
    IsError : boolean;
    UnloadN, WarnStr : string;
    UnloadDt : string;
    Division, sDivision, path : string;
    DD, MM, YY : WORD;
    errStr : string;
    INI : TIniFile;
    sPricepNmb : string;
    Tarif : double;
    Hour, Min, Sec, MSec: Word;
    Warn : TStringList;
    i : integer;
    PrevKod, PrevSeria, PrevNumber : string;
    InsSummPolis, InsSummEvent : integer;
    lInsSummPolis, lInsSummEvent : integer;
    InsSummCurr : string;
const
    InsTypeCode : string = '31';
begin
    errStr := '';
    UnloadN := '00';

    INI := TIniFile.Create(BLANK_INI);
    Division := INI.ReadString('DIVISION', 'ID', '');
    SaveDialog.InitialDir := INI.ReadString('DIVISION', 'OutDir', 'C:\');
    InsSummPolis := INI.ReadInteger(BG_SECT, 'InsSummPolis', 150000);
    InsSummEvent := INI.ReadInteger(BG_SECT, 'InsSummEvent', 150000);
    InsSummCurr := UpperCase(INI.ReadString(BG_SECT, 'InsSummCurr', 'EUR'));
    INI.Free;

    if DecodeCurrency(InsSummCurr) = '?' then begin
        MessageDlg('Неправильная валюта в файле ' + BLANK_INI + ' [' + BG_SECT + '] InsSummCurr=' + InsSummCurr, mtInformation, [mbOk], 0);
        exit;
    end;

    if Length(Division) < 1 then begin
        MessageDlg('Не указан код вашего подразделения', mtInformation, [mbOk], 0);
        exit;
    end;

    if Length(GetFilter) <> 0 then
    begin
        if not IsBreak.Checked then
        begin
            if MessageDlg('Вы не указали выгрузку досрочно расторгнутых (В Фильтре не установлена пометка Расторгнутые). Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
        end;
    end;

    SaveDialog.FileName := UnloadN + '.unload';
    if not SaveDialog.Execute then exit;
    path := Copy(SaveDialog.FileName, 1, Length(SaveDialog.FileName) - Pos('\', ReverseString(SaveDialog.FileName)) + 1);

    INI := TIniFile.Create(BLANK_INI);
    INI.WriteString('DIVISION', 'OutDir', path);

    sDivision := Division;
    if Length(sDivision) < 2 then sDivision := '0' + sDivision;

    Warn := TStringList.Create;

    DecodeDate(Date, YY, MM, DD);

    UnloadDt := IntToStr(MM);
    if MM < 10 then UnloadDt := '0' + UnloadDt;
    if DD < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(DD);

    DecodeTime(Now, Hour, Min, Sec, MSec);
//    UnloadN := IntToStr(Hour);
//    if Hour < 10 then UnloadN := '0' + UnloadN;
    UnloadN := Format('%02d', [Hour]) + Format('%02d', [Min]) + Format('%02d', [Sec]);

    IsError := false;
    AssignFile(F1, path + sDivision + 'D' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F1);
    AssignFile(X, path + sDivision + 'X' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(X);
    AssignFile(F3, path + sDivision + 'T' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F3);
    AssignFile(F9, path + sDivision + 'S' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F9);
    AssignFile(F8, path + sDivision + 'P' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F8);
    AssignFile(A, path + sDivision + 'A' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(A);
    AssignFile(F10, path + sDivision + 'H' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F10);
    AssignFile(F11, path + sDivision + 'Z' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F11);
    AssignFile(F12, path + sDivision + 'G' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F12);
    AssignFile(F13, path + sDivision + 'B' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F13);


    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    if (GetAsyncKeyState(vk_Shift) AND $8000) = 0 then
        WorkSQL.SQL.Add('SELECT ''' + Division + ''' AS DIVISION,T.* FROM BELGREEN T LEFT OUTER JOIN AGENT A ON (T.AGCODE=A.AGENT_CODE)')
    else
        WorkSQL.SQL.Add('SELECT A.DIVISIONCODE AS DIVISION,T.* FROM BELGREEN T LEFT OUTER JOIN AGENT A ON (T.AGCODE=A.AGENT_CODE)');
    if Length(GetFilter) <> 0 then
    begin
        WorkSQL.SQL.Add(' WHERE ' + GetFilter);
    end;

    Screen.Cursor := crHourGlass;
    WorkSQL.Open;
    if MainQuery.Active then
        All := MainQuery.RecordCount
    else
        All := WorkSQL.RecordCount;
    Curr := 1;
    with WorkSQL do begin
        while not Eof do begin
          try
            Curr := Curr + 1;
            ProgressBar.Position := round(Curr * 100 / All);
            if FieldByName('DIVISION').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' нет кода подразделения ' + FieldByName('AGCODE').AsString + ' ' + FieldByName('AGNAME').AsString);
            if FieldByName('STATE').AsInteger in [2] then begin //испорчен
                WriteLn(F13, FieldByName('DIVISION').AsString + '|8|' +
                            DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                            SConv(FieldByName('SERIA').AsString) + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                )
            end
            else if FieldByName('STATE').AsInteger in [3] then begin //утерян агентом
                WriteLn(F13, FieldByName('DIVISION').AsString + '|24|' +
                            DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                            SConv(FieldByName('SERIA').AsString) + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDATE').AsDateTime));
            end
            else begin
                    if Trim(FieldByName('OWNER').AsString) = '' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('OWNER').AsString + ' имя страхователя не указано');
                    if (FieldByName('AGTYPE').AsString = '') OR not (FieldByName('AGTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип агента не определён "' + FieldByName('AGTYPE').AsString + '" ' + FieldByName('AGNAME').AsString);
                    if (FieldByName('OWNTYPE').AsString = '') OR not (FieldByName('OWNTYPE').AsString[1] in ['F', 'U', 'I']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип страхователя не определён "' + FieldByName('OWNTYPE').AsString + '" ' + FieldByName('AGNAME').AsString);
                    if DecodePeriod(FieldByName('PERIOD').AsString, 0, 0, false) = '?' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' период не определён "' + FieldByName('PERIOD').AsString + '"');
                    if Length(FieldByName('AUTONMB').AsString) > 10 then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' номер авто больше 10 символов ' + FieldByName('AUTONMB').AsString);

                    CheckAutoNmb(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString, FieldByName('AUTONMB').AsString, Warn);

                    Tarif := FieldByName('TARIF').AsFloat;
                    if FieldByName('PSERIA').AsString <> '' then begin
                        WorkSQL2.Close;
                        WorkSQL2.SQL.Clear;
                        WorkSQL2.SQL.Add('SELECT TARIF FROM BELGREEN WHERE SERIA=''' + FieldByName('PSERIA').AsString + ''' AND NUMBER=' + FieldByName('PNUMBER').AsString);
                        WorkSQL2.Open;
                        if WorkSQL2.Eof then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' не найден утерянный №' + FieldByName('NUMBER').AsString);
                        Tarif := WorkSQL2.FieldByName('TARIF').AsFloat;
                        WorkSQL2.Close;
                    end;

                    lInsSummPolis := INI.ReadInteger(BG_SECT, 'InsSummPolis' + FieldByName('Country').AsString, InsSummPolis);
                    lInsSummEvent := INI.ReadInteger(BG_SECT, 'InsSummEvent' + FieldByName('Country').AsString, InsSummEvent);

                    WriteLn(F1, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('PCNT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('REGTIME').AsString) + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('DTFROM').AsDateTime, FieldByName('DTTO').AsDateTime, false) + '|' +
                                DecodeGlDate(FieldByName('DTFROM').AsDateTime) + '|' +
                                DecodeTimeX('00:00') + '|' +
                                DecodeGlDate(FieldByName('DTTO').AsDateTime) + '|' +
                                DecodeBGVariant(FieldByName('Country').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('OWNER').AsString + '|' +
                                DecodeFU12(FieldByName('OWNTYPE').AsString, '9') + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                '' + '|' + //PLACE_CODE
                                FieldByName('ADDR').AsString + '|' +
                                '' + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodeNBByOwner2(FieldByName('OWNTYPE').AsString, FieldByName('TYPE').AsString) + '|' + //Форма оплаты
                                ///DecodePayType('N') + '|' +!!!!!!!
                                DecodeAgentName(FieldByName('AGCODE').AsString) + '|' +
                                IfEmpty(FieldByName('TEXT').AsString <> '', IntToStr(lInsSummPolis)) + '|' + //Страховая сумма по полису (на весь период)
                                IfEmpty(FieldByName('TEXT').AsString <> '', IntToStr(lInsSummEvent)) + '|' + //Страховая сумма на один случай по полису
                                IfEmpty(FieldByName('TEXT').AsString <> '', DecodeCurrency(InsSummCurr)) + '|' + //Код валюты (страховой суммы)
                                IfEmpty(FieldByName('TEXT').AsString <> '', Format('%f', [Tarif])) + '|' + //Страховая премия
                                IfEmpty(FieldByName('TEXT').AsString <> '', Format('%f', [Tarif])) + '|' + //Страховой тариф (сумма)
                                IfEmpty(FieldByName('TEXT').AsString <> '', DecodeCurrency('EUR')) + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' //Для заметок
                    );
                    PrevKod := '';
                    PrevSeria := '';
                    PrevNumber := '';
                    if FieldByName('PNUMBER').AsInteger > 0 then begin
                        PrevKod := DecodeBlank(BG_SECT, FieldByName('PSERIA').AsString);
                        PrevSeria := FieldByName('PSERIA').AsString;
                        PrevNumber := IntToStr(FieldByName('PNUMBER').AsInteger);
                    end;
                    WriteLn(X, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('PCNT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('REGTIME').AsString) + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('DTFROM').AsDateTime, FieldByName('DTTO').AsDateTime, false) + '|' +
                                DecodeGlDate(FieldByName('DTFROM').AsDateTime) + '|' +
                                DecodeTimeX('00:00') + '|' +
                                DecodeGlDate(FieldByName('DTTO').AsDateTime) + '|' +
                                DecodeBGVariant(FieldByName('Country').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('OWNER').AsString + '|' +
                                DecodeFU12(FieldByName('OWNTYPE').AsString, '9') + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                '' + '|' + //PLACE_CODE
                                FieldByName('ADDR').AsString + '|' +
                                '' + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodeNBByOwner2(FieldByName('OWNTYPE').AsString, FieldByName('TYPE').AsString) + '|' + //Форма оплаты
                                ///DecodePayType('N') + '|' +!!!!!!!
                                DecodeAgentName(FieldByName('AGCODE').AsString) + '|' +
                                IntToStr(lInsSummPolis) + '|' + //Страховая сумма по полису (на весь период)
                                IntToStr(lInsSummEvent) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency(InsSummCurr) + '|' + //Код валюты (страховой суммы)
                                Format('%f', [Tarif]) + '|' + //Страховая премия
                                Format('%f', [Tarif]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('EUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' + '|' + //Для заметок
                                '' + '|' + //MOBTEL
                                '' + '|' + //HOMETEL
                                DecodeAgentID(FieldByName('AGCODE').AsString) + '|' + //Agent Doverennost
                                '' + '|' + //SMS
                                '' + '|' + //CONTACT
                                PrevKod + '|' +
                                PrevSeria + '|' +
                                PrevNumber
                    );



                    WriteLn(F3, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeLetter(FieldByName('LETTER').AsString, FieldByName('TEXT').AsString) + '|' + //Тип транспортного средства
                                FieldByName('MODEL').AsString + '|' +
                                FieldByName('AUTONMB').AsString + '|' +
                                '' + '|' + //Год выпуска
                                DecodeBodyMotor(FieldByName('BODY').AsString, FieldByName('MOTOR').AsString) + '|' +
                                '' + '|' + //Серия техпаспорта
                                '' + '|' + //Номер техпаспорта
                                '' + '|' + //Код характеристики
                                '' + '|' + //Значение характеристики
                                '' + '|' + //Стоимость
                                '' + '|' + //Код валюты (стоимости и страховой суммы)
                                '' + '|' + //Страховая сумма по объекту
                                'S' + '|' +
                                'BY' //Символ страны прибытия
                    );
                    if FieldByName('TEXT').AsString = '' then begin // НЕ Прицеп
                    if FieldByName('PAY1').AsFloat > 0 then
                    WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNBByOwner2(FieldByName('OWNTYPE').AsString, FieldByName('TYPE').AsString) + '|' + //Форма оплаты
                                '' + '|' +
                                DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('REPDATE').AsDateTime) + '|' + //Дата  проведено по бухгалтерии

                                FieldByName('PAY1').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('PAY1C').AsString) + '|' + //Код валюты
                                FieldByName('PCNT').AsString + '|' + //Комиссия
                                DecodeAgentName(FieldByName('AGCODE').AsString) + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('PCNT').AsFloat),
                                DecodeAgentID(FieldByName('AGCODE').AsString)
                    );
                    if FieldByName('PAY2').AsFloat > 0 then
                    WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNBByOwner2(FieldByName('OWNTYPE').AsString, FieldByName('TYPE').AsString) + '|' + //Форма оплаты
                                '' + '|' +
                                DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('REPDATE').AsDateTime) + '|' + //Дата  проведено по бухгалтерии
                                FieldByName('PAY2').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('PAY2C').AsString) + '|' + //Код валюты
                                FieldByName('PCNT').AsString + '|' + //Комиссия
                                DecodeAgentName(FieldByName('AGCODE').AsString) + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('PCNT').AsFloat),
                                DecodeAgentID(FieldByName('AGCODE').AsString)
                    );
                    end;
                    if FieldByName('STATE').AsString = '4' then
                    if FieldByName('RET1').AsFloat > 0 then
                    WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '51' + '|' + //Код события
                                DecodeNBByOwner2(FieldByName('OWNTYPE').AsString, FieldByName('TYPE').AsString) + '|' + //Форма оплаты
                                '|' +
                                DecodeGlDate(FieldByName('RETDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('RETDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('RETDATE').AsDateTime) + '|' +
                                FieldByName('RET1').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('RET1CUR').AsString) + '|' + //Код валюты
                                '|' + //Комиссия
                                '|' +
                                DecodeFU16('', 0),
                                ''
                    );
                    if FieldByName('STATE').AsString = '4' then  //РАСТОРГНУТ
                    WriteLn(F11, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '|' +
                                DecodeGlDate(FieldByName('RETDATE').AsDateTime) + '|' +
                                '2' + '|' +
                                '2' + '|' +
                                '41'
                    );
                    if FieldByName('Text').AsString <> '' then begin//Сцепка
                    WorkSQl2.SQL.Clear;
                    WorkSQl2.SQL.Add('SELECT * FROM BELGREEN WHERE PRCPSER=''' + FieldByName('SERIA').AsString + ''' AND PRCPNMB=' + FieldByName('NUMBER').AsString);
                    WorkSQl2.Open;
                    if WorkSQl2.Eof then
                        sPricepNmb := IntToStr(FieldByName('NUMBER').AsInteger - 1)
                    else
                        sPricepNmb := WorkSQl2.FieldByName('NUMBER').AsString;
                    WorkSQl2.Close;
                    WriteLn(F9, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                SConv(FieldByName('SERIA').AsString) + '|' +
                                sPricepNmb + '|' + //Main Polis
                                FieldByName('NUMBER').AsString         //Pricep
                    );
                    end;
                //end;
                //else begin //Дубликаты
                if FieldByName('PSERIA').AsString <> '' then begin
                    if FieldByName('DUPTYPE').AsString = '0' then
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('PSERIA').AsString) + '|' +
                                FieldByName('PSERIA').AsString + '|' +
                                FieldByName('PNUMBER').AsString + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '12' + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                    );
                    if FieldByName('DUPTYPE').AsString = '1' then
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(BG_SECT, FieldByName('PSERIA').AsString) + '|' +
                                FieldByName('PSERIA').AsString + '|' +
                                FieldByName('PNUMBER').AsString + '|' +
                                DecodeBlank(BG_SECT, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '9' + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                    );
                end;
            end; //if
            Next;
          except
            on E : Exception do begin
                errStr := errStr + #13#10 + e.Message + ' ' + FieldByName('NUMBER').AsString;
                IsError := true;
                Next;
            end
        end; //while
    end; //width
  end;
    INI.Free;
    CloseFile(F1);
    CloseFile(X);
    CloseFile(F3);
    CloseFile(F9);
    CloseFile(F8);
    CloseFile(A);
    CloseFile(F10);
    CloseFile(F11);
    CloseFile(F12);
    CloseFile(F13);
    ProgressBar.Position := 0;

    Screen.Cursor := crDefault;

    for i := 0 to Warn.Count - 1 do begin
        WarnStr := WarnStr + Warn[i] + #13;
    end;
    if Warn <> nil then Warn.Free;

    if not IsError then begin
        MessageDlg('Экспорт выполнен успешно', mtInformation, [mbOk], 0);
        ShowInfo.Caption := 'ПРЕДУПРЕЖДЕНИЯ!!! Возможно ошибки';
        ShowInfo.InfoPanel.Text := WarnStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        if WarnStr <> '' then  ShowInfo.ShowModal;
    end
    else begin
        ShowInfo.Caption := 'Ошибки в данных';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
end;

procedure TBelGreenForm.Panel5Resize(Sender: TObject);
begin
    ListAgents.Height := DownPanel.Top - ListAgents.Top 
end;

procedure TBelGreenForm.StaticticClick(Sender: TObject);
begin
    MessageDlg(COMPAMYNAME + ' ' + DivisionInfo.Caption, mtInformation, [mbOk], 0);
end;

procedure TBelGreenForm.CurrencyRateSrcClick(Sender: TObject);
begin
    CurrencyRates.ShowModal
end;

end.
