unit Private2005;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, ExtCtrls, Db, DBTables, RxQuery, Grids, DBGrids, RXDBCtrl,
  StdCtrls, ComCtrls, ActnList, CheckLst, rxstrutils, RXSplit, DataMod, InpRezBG,
  Buttons, math, Placemnt, strutils, shellapi;

type
  TPrivate2005Form = class(TForm)
    StatusPanel: TPanel;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N7: TMenuItem;
    Excel1: TMenuItem;
    MainQuery: TRxQuery;
    Priv2005Database: TDatabase;
    DataSource: TDataSource;
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
    Statictic: TMenuItem;
    AllAgents: TTable;
    FixSerNmb: TMenuItem;
    ProgressBar: TProgressBar;
    WorkSQL: TQuery;
    LocateTbl: TTable;
    StatPage: TTabSheet;
    StatisticTxt: TMemo;
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
    IsNumbers: TCheckBox;
    IsUseFilter: TCheckBox;
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
    ListTemplates: TComboBox;
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
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryRegDt: TDateField;
    MainQueryInsurer: TStringField;
    MainQueryInsCnt: TSmallintField;
    MainQueryAddress: TStringField;
    MainQueryBirthdate: TStringField;
    MainQueryPassport: TStringField;
    MainQueryInsurer2: TStringField;
    MainQueryFrom: TDateField;
    MainQueryFromTm: TStringField;
    MainQueryTo: TDateField;
    MainQueryPeriod: TSmallintField;
    MainQueryDuration: TSmallintField;
    MainQueryInsSum: TFloatField;
    MainQueryUridSum: TFloatField;
    MainQueryTarif: TFloatField;
    MainQueryInsSumCurr: TStringField;
    MainQueryPay: TFloatField;
    MainQueryPayDt: TDateField;
    MainQueryRepDt: TDateField;
    MainQuerySum1: TFloatField;
    MainQuerySum1Curr: TStringField;
    MainQuerySum2: TFloatField;
    MainQuerySum2Curr: TStringField;
    MainQuerystopdate: TDateField;
    MainQueryretsum: TFloatField;
    MainQueryretcur: TStringField;
    MainQueryStateText: TStringField;
    MainQueryState: TStringField;
    WorkSQL2: TQuery;
    SaveDialog: TSaveDialog;
    MainQueryPSeria: TStringField;
    MainQueryPNumber: TFloatField;
    Label10: TLabel;
    FromDateFrom: TDateTimePicker;
    Label14: TLabel;
    FromDateTo: TDateTimePicker;
    MainQueryName: TStringField;
    To1C: TMenuItem;
    SaveDialog1C: TSaveDialog;
    Table1C: TTable;
    IsStopped: TCheckBox;
    Label16: TLabel;
    Label17: TLabel;
    UpdateDtFrom: TDateTimePicker;
    Label18: TLabel;
    UpdateDtTo: TDateTimePicker;
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
    procedure PageControlChange(Sender: TObject);
    procedure RepListChange(Sender: TObject);
    procedure StartButtonClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure ListAgentsClick(Sender: TObject);
    procedure InputSumClick(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure DelMenuClick(Sender: TObject);
    procedure MenuGloboClick(Sender: TObject);
    procedure Panel5Resize(Sender: TObject);
    procedure To1CClick(Sender: TObject);
  private
    { Private declarations }
    listAgCodes : TStringList;
    MSEXCEL : Variant;
    Section: string;

    function GetFilter : string;
    function GetOrder : string;
  public
    { Public declarations }
  end;

var
  Private2005Form: TPrivate2005Form;
  lastAgCode2, lastAgName2 : string;

const
  Section : string = 'PRIV2014';

procedure Create1CTable(Table1C: TTable; Name : string);

implementation

uses Sort, WaitFormUnit, dateutil, Variants,
  BGBuroUnit, inifiles, ShowInfoUnit, Ubitok,
  InpRezParams, TmplNameUnit, registry, RepParamsFrm, RepParamsAgntFrm,
  PcntFormUnit, mf, DM;

{$R *.DFM}

procedure TPrivate2005Form.N2Click(Sender: TObject);
begin
    Close
end;

procedure TPrivate2005Form.Button1Click(Sender: TObject);
begin
    Close
end;

procedure ConnectFilter(var filter : string; Operation, Condition : string);
begin
     if Length(filter) <> 0 then
        filter := filter + ' ' + Operation + ' ';
     filter := filter + Condition;
end;

procedure TPrivate2005Form.FormCreate(Sender: TObject);
var
    INI : TIniFile;
    i : integer;
    rINI : TRegIniFile;
    str : string;
begin
    DMM.HandBookSQL.DatabaseName := Priv2005Database.Name;
    RepDtFrom.Date := Date;
    RepDtTo.Date := Date;
    RepDtFrom.Checked := false;
    RepDtTo.Checked := false;
    Section := 'PRIV2014';

    listAgCodes := TStringList.Create;
    StateCombo.ItemIndex := 0;

     rINI := TRegIniFile.Create('Жизнь 2005');
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
        if INI.ReadString(Section, 'QueryName' + IntToStr(i), '') <> '' then
            RepList.Items.Add(INI.ReadString(Section, 'QueryName' + IntToStr(i), ''))
        else
            break;
    end;
    INI.Free;
    Panel5Resize(nil)
end;

procedure TPrivate2005Form.FormDestroy(Sender: TObject);
begin
    listAgCodes.Free
end;

procedure TPrivate2005Form.MainQueryCalcFields(DataSet: TDataSet);
begin
    MainQueryStateText.AsString := '';
    if MainQueryState.IsNull then MainQueryStateText.AsString := '???'
    else
    if MainQueryState.AsInteger = 1 then MainQueryStateText.AsString := 'ИСПОРЧЕН'
    else
    if MainQueryState.AsInteger = 2 then MainQueryStateText.AsString := 'РАСТОРГНУТ'
    else
    if MainQueryState.AsInteger = 3 then MainQueryStateText.AsString := 'УТЕРЯН'
    else
    if MainQueryPSeria.AsString <> '' then MainQueryStateText.AsString := 'ДУБЛИКАТ ДЛЯ ' + MainQueryPSeria.AsString + '/' + MainQueryPNumber.AsString
end;

procedure TPrivate2005Form.ApplyClick(Sender: TObject);
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

procedure TPrivate2005Form.N3Click(Sender: TObject);
begin
    PageControl.ActivePageIndex := 1;
end;

function TPrivate2005Form.GetFilter : string;
var
    s, s_a, s_r : string;
    i : integer;
begin
    s := '';
    s_a := '';


    if RepDtFrom.Checked then
        ConnectFilter(s, 'AND', 'REPDT>=''' + DateToStr(RepDtFrom.Date) + '''');
    if RepDtTo.Checked then
        ConnectFilter(s, 'AND', 'REPDT<=''' + DateToStr(RepDtTo.Date) + '''');

    if UpdateDtFrom.Checked then
        ConnectFilter(s, 'AND', 'UPDT>=''' + DateToStr(UpdateDtFrom.Date) + '''');
    if UpdateDtTo.Checked then
        ConnectFilter(s, 'AND', 'UPDT<=''' + DateToStr(UpdateDtTo.Date) + '''');

    if IsStopped.Checked then begin
        if RepDtFrom.Checked then
            ConnectFilter(s_a, 'AND', 'STOPDATE>=''' + DateToStr(RepDtFrom.Date) + '''');
        if RepDtTo.Checked then
            ConnectFilter(s_a, 'AND', 'STOPDATE<=''' + DateToStr(RepDtTo.Date) + '''');
    end;

    if s_a <> '' then begin
        if s = '' then s := s_a
        else s := '('+s+') or ('+s_a+')';
    end;


    if NmbFrom.Text <> '' then
        ConnectFilter(s, 'AND', 'NUMBER>=' + NmbFrom.Text);
    if NmbTo.Text <> '' then
        ConnectFilter(s, 'AND', 'NUMBER<=' + NmbTo.Text);
    if Insurer.Text <> '' then
        ConnectFilter(s, 'AND', 'INSURER2 LIKE ''%' + Insurer.Text + '%''');
    if InsurType.ItemIndex = 1 then
        ConnectFilter(s, 'AND', 'INSTYPE=''F''');
    if InsurType.ItemIndex = 2 then
        ConnectFilter(s, 'AND', 'INSTYPE=''U''');
    if StateCombo.ItemIndex >= 1 then begin
        ConnectFilter(s, 'AND', '(STATE=''' + IntToStr(StateCombo.ItemIndex - 1) + ''')')
    end;

    if EndDateFrom.Checked then
        ConnectFilter(s, 'AND', 'T."TO">=''' + DateToStr(EndDateFrom.Date) + '''');
    if EndDateTo.Checked then
        ConnectFilter(s, 'AND', 'T."TO"<=''' + DateToStr(EndDateTo.Date) + '''');

    if FromDateFrom.Checked then
        ConnectFilter(s, 'AND', 'T."FROM">=''' + DateToStr(FromDateFrom.Date) + '''');
    if FromDateTo.Checked then
        ConnectFilter(s, 'AND', 'T."FROM"<=''' + DateToStr(FromDateTo.Date) + '''');

    if listInsType.ItemIndex = 1 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''F''');
    if listInsType.ItemIndex = 2 then
        ConnectFilter(s, 'AND', 'OWNTYPE=''U''');

    if s <> '' then s := '(' + s + ')';

    if IsAgents.Checked then begin
        s_a := '';
        for i := 0 to ListAgents.Items.Count - 1 do
            if ListAgents.Checked[i] = true then begin
                if s_a <> '' then s_a := s_a + ',';
                s_a := s_a + '''' + listAgCodes[i] + '''';
            end;
         if s_a <> '' then
            ConnectFilter(s, 'AND', 'AGENT IN (' + s_a + ')');
    end;

    if s = '' then GetFilter := '1=1'
    else GetFilter := s;
end;

function TPrivate2005Form.GetOrder : string;
begin
    if SortForm.GetOrder('') <> '' then GetOrder := SortForm.GetOrder('ORDER BY')
    else GetOrder := 'ORDER BY SERIA, NUMBER';
end;

procedure TPrivate2005Form.N4Click(Sender: TObject);
begin
    SortForm.m_Query := MainQuery;
    if SortForm.ShowModal = mrOk then
        ApplyClick(nil);
end;

procedure TPrivate2005Form.Excel1Click(Sender: TObject);
begin
    DatSetToExcel(MainQuery, true);
end;

procedure TPrivate2005Form.FixSerNmbClick(Sender: TObject);
begin
    if MainGrid.FixedCols = 0 then MainGrid.FixedCols := 2
    else MainGrid.FixedCols := 0;
    FixSerNmb.Checked := MainGrid.FixedCols > 0;
end;

function NmbToStr(N : integer) : string;
var
    s : string;
begin
    s := IntToStr(N);
    while Length(s) < 7 do s := '0' + s;
    Result := s;
end;

procedure TPrivate2005Form.PageControlChange(Sender: TObject);
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
                Add('FROM PRIV2005 T');
                Add('WHERE TARIF>0 AND ' + GetFilter());
                Open;
                StatisticTxt.Lines.Add('Тариф');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString);
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(SUM1) AS S, SUM1CURR AS CUR');
                Add('FROM PRIV2005 T');
                Add('WHERE SUM1CURR<>'''' AND ' + GetFilter());
                Add('GROUP BY SUM1CURR');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Первая часть');
                if FieldByName('S').AsString = '' then
                    StatisticTxt.Lines.Add('<Нет>');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(SUM2) AS S, SUM2CURR AS CUR');
                Add('FROM PRIV2005 T');
                Add('WHERE SUM2 IS NOT NULL AND SUM2CURR<>'''' AND ' + GetFilter());
                Add('GROUP BY SUM2CURR');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Вторая часть');
                if FieldByName('S').AsString = '' then
                    StatisticTxt.Lines.Add('<Нет>');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;
                Close;
                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM PRIV2005 T');
                Add('WHERE STATE=''1'' AND ' + GetFilter());
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
                Add('FROM PRIV2005 T');
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
                Add('FROM PRIV2005 T');
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

                if RepDtFrom.Checked and RepDtto.Checked then begin
                    Close;
                    Clear;
                    Add('SELECT COUNT(*), SUM(Retsum), RetCur');
                    Add('FROM PRIV2005 T');
                    Add('WHERE STOPDATE >= ''' + DateToStr(RepDtFrom.Date) + ''' ');
                    Add('AND STOPDATE <= ''' + DateToStr(RepDtTo.Date) + ''' ');
                    //Add('AND ' + GetFilter());
                    Add('GROUP BY RetCur');
                    Open;
                    StatisticTxt.Lines.Add('');
                    StatisticTxt.Lines.Add('Расторгнуто (по всей Базе)');
                    while not Eof do begin
                        StatisticTxt.Lines.Add(Fields[0].AsString + ' шт.  ' + Fields[1].AsString  + '  ' + Fields[2].AsString);
                        Next;
                    end;
                end;

                if EndDateTo.Checked then begin
                    Close;
                    Clear;
                    Add('SELECT COUNT(*)');
                    Add('FROM PRIV2005 T');
                    Add('WHERE State=''0''');
                    Add('AND T."TO">=''' + DateToStr(EndDateTo.Date) + ''' ');
                    Add('AND T."FROM"<''' + DateToStr(EndDateTo.Date) + ''' ');
                    //Add('AND ' + GetFilter());
                    Open;
                    StatisticTxt.Lines.Add('');
                    StatisticTxt.Lines.Add('Действует на ' + DateToStr(EndDateTo.Date) + ' (по всей Базе)');
                    while not Eof do begin
                        StatisticTxt.Lines.Add(Fields[0].AsString + ' шт.  ');
                        Next;
                    end;
                end;
            end;
        except
          on E : Exception do begin
            StatisticTxt.Text := E.Message;
          end;
        end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TPrivate2005Form.RepListChange(Sender: TObject);
var
    INI : TIniFile;
begin
    REPQuery.Close;
    INI := TIniFile.Create('blreps.ini');
    ParamsText.Hint := INI.ReadString(Section, 'QueryParams' + IntToStr(RepList.ItemIndex + 1), '');
    ParamsText.Text := INI.ReadString(Section, 'LastParams' + IntToStr(RepList.ItemIndex + 1), '');
    if ParamsText.Text = '' then
        ParamsText.Text := ParamsText.Hint; 
    INI.Free;
end;

procedure TPrivate2005Form.StartButtonClick(Sender: TObject);
var
    s : string;
begin
    s := '';
    if IsUseFilter.Checked then s := GetFilter();
    ExecLittleReport('blreps.ini', Section, ParamsText.Text, RepList.ItemIndex + 1, REPQuery, GetFilter());
end;

procedure TPrivate2005Form.Button2Click(Sender: TObject);
begin
    if REPQuery.Active then
        DatSetToExcel(REPQuery, IsNumbers.Checked);
end;

procedure TPrivate2005Form.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top + SpeedButton.Height ;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TPrivate2005Form.MenuItem1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Жизнь 2005');
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

procedure TPrivate2005Form.ListTemplatesChange(Sender: TObject);
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
     INI := TRegIniFile.Create('Жизнь 2005');
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

procedure TPrivate2005Form.ListAgentsClick(Sender: TObject);
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

procedure TPrivate2005Form.InputSumClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT SUM1,SUM1CURR,PAYDT,NUMBER,REPDT FROM PRIV2005 WHERE REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT SUM2,SUM2CURR,PAYDT,NUMBER,REPDT FROM PRIV2005 WHERE REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkSQL);
end;

procedure TPrivate2005Form.N11Click(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    //RepParamsAgnt.filtParam := 'LEBEN2PCNT > 0';
    if RepParamsAgnt.ShowModal <> mrOk then exit;
    if RepParamsAgnt.listAgnt.ItemIndex = -1 then
    begin
        MessageDlg('Выделите одного агента', mtInformation, [mbOk], 0);
        exit;
    end;

    DtFrom := min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    DtTo := max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    sqlText := 'SELECT SUM1,SUM1CURR,REPDT,NUMBER,AGPERCENT FROM PRIV2005 WHERE AGENT=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PAYDT IS NOT NULL AND REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT SUM2,SUM2CURR,REPDT,NUMBER,AGPERCENT FROM PRIV2005 WHERE AGENT=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PAYDT IS NOT NULL AND REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Items[RepParamsAgnt.listAgnt.ItemIndex], WorkSQL);
end;

procedure TPrivate2005Form.DelMenuClick(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Жизнь 2005');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

function DecodeFUFU(s : string; pcnt : double) : string;
begin
    DecodeFUFU := '?';
    if pcnt = 0 then DecodeFUFU := '3'
    else
    if s = 'F' then DecodeFUFU := '1'
    else
    if s = 'U' then DecodeFUFU := '6';
end;
function DecodePassport(s : string) : string;
begin
    DecodePassport := '';
    if Length(s) > 3 then DecodePassport := s;
end;

function DecodeTxtDate(s : string) : string;
var
    d, m, y : string;
begin
    DecodeTxtDate := '';

    try
        d := ExtractWord(1, s, ['/', '.', '-']);
        m := ExtractWord(2, s, ['/', '.', '-']);
        y := ExtractWord(3, s, ['/', '.', '-']);
        if(StrToInt(y) < 10) then
            y := '200' + IntToStr(StrToInt(y))
        else
        if(StrToInt(y) <= 99) then
            y := '19' + IntToStr(StrToInt(y));

        DecodeTxtDate := DecodeGlDate(EncodeDate(StrToInt(y), StrToInt(m), StrToInt(d)));
    except
    end;
end;

function DecodeOwnType(s : string) : string;
begin
    DecodeOwnType := '?';
    if s = 'F' then DecodeOwnType := '1';
    if s = 'U' then DecodeOwnType := '3';
end;

function DecodeNB(s : string) : string;
begin
    DecodeNB := '?';
    if s = 'N' then DecodeNB := '1';
    if s = 'B' then DecodeNB := '2';
end;

function DecodeAGNameL(s : string) : string;
begin
    if lastAgCode2 = s then begin
        DecodeAGNameL := lastAgName2;
        exit;
    end;
    Private2005Form.WorkSQL2.Close;
    Private2005Form.WorkSQL2.SQL.Clear;
    Private2005Form.WorkSQL2.SQL.Add('SELECT NAME FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    Private2005Form.WorkSQL2.Open;
    if Private2005Form.WorkSQL2.Eof then lastAgName2 := s
    else lastAgName2 := Private2005Form.WorkSQL2.FieldByName('NAME').AsString;
    DecodeAGNameL := lastAgName2;
    lastAgCode2 := s;
end;

procedure DecodeKS(ks : string; var List : TStringList);
{const
   decoder1_src : array[1..19] of string = ('0101', '0201', '0301', '0401', '0501',
   '0601', '0602', '0630', '0604', '0605', '0606', '0607', '0608',
   '0701', '0801', '0901', '1001', '1101', '1201');
   decoder1_dest : array[1..19] of string = ('0101', '0201', '0202', '0300', '0400',
   '0501', '0502', '0503', '', '0600', '0700', '0800', '0900',
   '1000', '1200', '1100', '', '1300', ''
   ); }
   //decoder2 : array[1..3] of string = ('12', '06', '13');
var
   INI : TIniFile;
   key : string;
   i, x : integer;
   s, fs : string;
   f : boolean;
begin
   List.Clear;

   i := 1;
   while true do begin
       s := Copy(ks, i, 4);
       Inc(i, 4);
       if s = '' then break;
       key := Format('GLOBO%d_%d', [StrToInt(Copy(s, 1, 2)), StrToInt(Copy(s, 3, 2))]);
       //fs := s;
       //key := '?';
       //if t[1] in ['Q', 'Y', 'S'] then begin
       // List.Add(decoder2[StrToInt(Copy(s, 1, 2))] + '00');
       // key := Format('K_YSQ_%d_%d', [StrToInt(Copy(s, 1, 2)), StrToInt(Copy(s, 3, 2))]);
       // f := true;
       //end
       //else begin
        {f := false;
        for x := 1 to 19 do
            if s = decoder1_src[x] then begin
                if decoder1_dest[x] = '' then begin
                    raise Exception.Create('Не найден коэффициент ' + fs + ' в ГЛОБО');
                    exit;
                end;
                List.Add(decoder1_dest[x]);
                key := Format('K%d_%d', [StrToInt(Copy(s, 1, 2)), StrToInt(Copy(s, 3, 2))]);
                f := true;
                break;
            end;
       //end;
       }
       INI := TIniFile.Create(BLANK_INI);
       s := INI.ReadString(Section, key, '');
       if s = '' then begin
            INI.Free;
            //if f then List.Delete(List.Count - 1);
            raise Exception.Create('Не найден коэффициент ' + key + ' в ГЛОБО');
            exit;
       end;
       if Pos(',', s) = 0 then
            raise Exception.Create('Не правильный коэффициент ' + key + ' в ГЛОБО ' + s);

       List.Add(Trim(Copy(s, Pos(',', s) + 1, Length(s))));
       List.Add(Trim(Copy(s, 1, Pos(',', s) - 1)));
   end;

   INI.Free;
end;

function DecodeAssist(s : string) : string;
var
     INI : TIniFile;
begin
     INI := TIniFile.Create(BLANK_INI);
     if s = '' then s := '0';
     DecodeAssist := INI.ReadString(Section, 'Assist' + s, '');
     INI.Free;
end;

function DecodeAddr(s1, s2, s : string) : string;
begin
    DecodeAddr := '';
    if AnsiUpperCase(ReplaceStr(Trim(s1), ' ', '')) = AnsiUpperCase(ReplaceStr(Trim(s2), ' ', '')) then DecodeAddr := s;
end;

function AddOne(s : string) : string;
begin
    s := IntToStr(StrToInt(s) + 1);
    while(Length(s) < 2) do s := '0' + s;
    AddOne := s;
end;

function DecodeInsTypeVid(s : string; sum, cur : string) : string;
var
     INI : TIniFile;
     kod : string;
begin
    INI := TIniFile.Create(BLANK_INI);
    kod := INI.ReadString(Section, 'Country' + s + '_' + sum + '_' + cur, '?');
    INI.Free;
    if kod = '?' then
        raise Exception.Create('Не указан код вида ' + s + '_' + sum + '_' + cur + ' в ГЛОБО ');
    DecodeInsTypeVid := kod;
end;

function DecodeBlank2005(s : string) : string;
var
    INI : TIniFile;
    kod : string;
begin
    INI := TIniFile.Create(BLANK_INI);
    kod := INI.ReadString(Section, 'Seria_' + s + '_KOD', '?');
    INI.Free;
    if kod = '?' then
        raise Exception.Create('Не указан код серии бланка ' + kod + ' в ГЛОБО ');
    DecodeBlank2005 := kod;
end;

procedure TPrivate2005Form.MenuGloboClick(Sender: TObject);
var
    F1, A : TextFile;
    L : TextFile;
    X : TextFile;
    //F3 : TextFile;
    F6 : TextFile;
//    F9 : TextFile;
    F8 : TextFile;
    F10 : TextFile;
    F11 : TextFile;
    F12 : TextFile;
    F13 : TextFile;
    F15 : TextFile;
    All, Curr : integer;
    IsError : boolean;
    UnloadN : string;
    UnloadDt : string;
    Division, sDivision : string;
    DD, MM, YY : WORD;
    errStr : string;
    List : TStringList;
    i : integer;
    INI : TIniFile;
    path : string;
    Ltr : char;
    pbuff : array[0..128] of char;
    Hour, Min, Sec, MSec: Word;
    PrevKod, PrevSeria, PrevNumber : string;
const
    InsTypeCode : string = '08';
begin
    UnloadN := '00';
    if FloatToStr(5.5)[2] <> '.' then begin
        MessageDlg('Установлен не правильный разделитель целой и дробной части. Нужна точка.', mtInformation, [mbOk], 0);
        exit;
    end;

    INI := TIniFile.Create(BLANK_INI);
    Division := INI.ReadString('DIVISION', 'ID', '');
    SaveDialog.InitialDir := INI.ReadString('DIVISION', 'OutDir', 'C:\');
    INI.Free;
    if Length(Division) = 0 then begin
        MessageDlg('Не указан код вашего подразделения', mtInformation, [mbOk], 0);
        exit;
    end;

    sDivision := Division;
    if Length(sDivision) < 2 then sDivision := '0' + sDivision;

    SaveDialog.FileName := UnloadN + '.unload';
    if not SaveDialog.Execute then exit;
    path := Copy(SaveDialog.FileName, 1, Length(SaveDialog.FileName) - Pos('\', ReverseString(SaveDialog.FileName)) + 1);

    INI := TIniFile.Create(BLANK_INI);
    INI.ReadString('DIVISION', 'OutDir', path);
    INI.Free;

    List := TStringList.Create;
    errStr := '';

    DecodeTime(Now, Hour, Min, Sec, MSec);
    //UnloadN := IntToStr(Hour);
    //if Hour < 10 then UnloadN := '0' + UnloadN;
    UnloadN := Format('%02d', [Hour]) + Format('%02d', [Min]) + Format('%02d', [Sec]);


    DecodeDate(Date, YY, MM, DD);
    UnloadDt := IntToStr(DD);
    if DD < 10 then UnloadDt := '0' + UnloadDt;
    if MM < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(MM);
    IsError := false;
    AssignFile(F1, path + sDivision + 'D' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F1);
    AssignFile(L, path + sDivision + 'L' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(L);
//    AssignFile(F3, path + sDivision + 'T' + UnloadDt + UnloadN + '.0' + InsTypeCode);
//    ReWrite(F3);
    AssignFile(F6, path + sDivision + 'C' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F6);
//    AssignFile(F9, path + sDivision + 'S' + UnloadDt + UnloadN + '.0' + InsTypeCode);
//    ReWrite(F9);
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
    AssignFile(F15, path + sDivision + 'K' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F15);
    AssignFile(X, path + sDivision + 'X' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(X);

    Screen.Cursor := crHourGlass;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT ''' + Division + ''' AS DIVISION,T.* FROM PRIV2005 T LEFT OUTER JOIN AGENT A ON (T.AGENT=A.AGENT_CODE)');
    if Length(GetFilter) <> 0 then
        WorkSQL.SQL.Add(' WHERE ' + GetFilter);
    WorkSQL.Open;
    //if MainQuery.Active then
    //    All := MainQuery.RecordCount
    //else
        All := WorkSQL.RecordCount;
    Curr := 1;
    //StatusPanel.Visible := true;
    //StatusMsg.Visible := false;
    lastAgCode := '';
    with WorkSQL do begin
        while not Eof do begin
          try
            Curr := Curr + 1;

            //if FieldByName('NUMBER').AsString = '21426' then
            //messagebeep(0);

            ProgressBar.Position := round(Curr * 100 / All);
            if FieldByName('DIVISION').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' нет кода подразделения ' + FieldByName('AGENT_CODE').AsString);
            if FieldByName('STATE').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' состояние не определено ' + FieldByName('ST').AsString);
            if FieldByName('STATE').AsInteger in [1] then begin //испорчен
                WriteLn(F13, FieldByName('DIVISION').AsString + '|8|' +
                            DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                            FieldByName('SERIA').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDT').AsDateTime)
                )
            end
            {else
            if FieldByName('STATE').AsInteger in [3] then begin //утерян
                WriteLn(F13, FieldByName('DIVISION').AsString + '|24|' +
                            DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                            FieldByName('SERIA').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                )
            end }
            else begin
                //if FieldByName('PSERIA').AsString = '' then begin
                    if Trim(FieldByName('INSURER').AsString) = '' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('INSURER').AsString + ' имя страхователя не указано');
                    if (FieldByName('AGTYPE').AsString = '') OR not (FieldByName('AGTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип агента не определён "' + FieldByName('AGTYPE').AsString + '" ');
                    //if (FieldByName('PAYMENTDATE').IsNull) then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата перечисления не определена');
                    if (FieldByName('REGDT').AsDateTime > FieldByName('FROM').AsDateTime) AND (FieldByName('PSERIA').AsString = '') then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата выписки после начала');
                    //if ((FieldByName('REGDT').AsDateTime + 5) < FieldByName('FROM').AsDateTime) AND (FieldByName('PSERIA').AsString <> '') then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата дубликата до начала');
                    if (FieldByName('PAYDT').AsDateTime > FieldByName('REGDT').AsDateTime) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата оплаты сдишком поздно');
                    //if (FieldByName('REGDT').AsDateTime > FieldByName('REGDT').AsDateTime) then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата оплаты сдишком поздно');
                    //if ((FieldByName('WRITEDATE').AsDateTime - 45) > FieldByName('PAYMENTDATE').AsDateTime) then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата выписки позже даты оплаты на много');
                    if (FieldByName('INSTYPE').AsString = '') OR not (FieldByName('INSTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип страхователя не определён "' + FieldByName('INSTYPE').AsString + '" ');
                    //if DecodePeriod(FieldByName('PERIOD').AsString) = '?' then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' период не определён "' + FieldByName('PERIOD').AsString + '"');

                    WriteLn(F1, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFUFU(FieldByName('AGTYPE').AsString, -1) + '|' +
                                DecodeGlDate(FieldByName('REGDT').AsDateTime) + '|' +
                                '' + '|' + //time
                                FieldByName('DURATION').AsString + '|' + //DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('FROM').AsDateTime, FieldByName('TO').AsDateTime, true) + '|' +
                                DecodeGlDate(FieldByName('FROM').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('FROMTM').AsString) + '|' +
                                DecodeGlDate(FieldByName('TO').AsDateTime) + '|' +
                                DecodeInsTypeVid(FieldByName('COUNTRY').AsString, FieldByName('INSSUM').AsString, FieldByName('INSSUMCURR').AsString) + '|' + //Код варианта страхования
                                IntToStr(max(FieldByName('INSCNT').AsInteger, 0)) + '|' + //Количество застрахованных
                                FieldByName('INSURER2').AsString + '|' +
                                DecodeFU12(FieldByName('INSTYPE').AsString, '9') + '|' +
                                '1' + {DecodeR(FieldByName('RESIDENT').AsString) +} '|' +
                                '' + '|' + //PLACE_CODE
                                DecodeAddr(FieldByName('INSURER').AsString, FieldByName('INSURER2').AsString, FieldByName('ADDRESS').AsString) + '|' +
                                DecodeOwnType(FieldByName('INSTYPE').AsString) + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodePayType('N') + '|' +
                                DecodeAgNameL(FieldByName('AGENT').AsString) + '|' +
                                IntToStr(FieldByName('INSSUM').AsInteger * max(FieldByName('INSCNT').AsInteger, 1)) + '|' + //Страховая сумма по полису (на весь период)
                                IntToStr(FieldByName('INSSUM').AsInteger * max(FieldByName('INSCNT').AsInteger, 1)) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency(FieldByName('INSSUMCURR').AsString) + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('PAY').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TARIF').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency(FieldByName('INSSUMCURR').AsString) + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' //Для заметок
                    );

                    PrevKod := '';
                    PrevSeria := '';
                    PrevNumber := '';
                    if FieldByName('PNUMBER').AsInteger > 0 then begin
                        PrevKod := DecodeBlank2005(FieldByName('PSERIA').AsString);
                        PrevSeria := FieldByName('PSERIA').AsString;
                        PrevNumber := IntToStr(FieldByName('PNUMBER').AsInteger);
                    end;

                    WriteLn(X, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFUFU(FieldByName('AGTYPE').AsString, -1) + '|' +
                                DecodeGlDate(FieldByName('REGDT').AsDateTime) + '|' +
                                '' + '|' + //time
                                FieldByName('DURATION').AsString + '|' + //DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('FROM').AsDateTime, FieldByName('TO').AsDateTime, true) + '|' +
                                DecodeGlDate(FieldByName('FROM').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('FROMTM').AsString) + '|' +
                                DecodeGlDate(FieldByName('TO').AsDateTime) + '|' +
                                DecodeInsTypeVid(FieldByName('COUNTRY').AsString, FieldByName('INSSUM').AsString, FieldByName('INSSUMCURR').AsString) + '|' + //Код варианта страхования
                                IntToStr(max(FieldByName('INSCNT').AsInteger, 0)) + '|' + //Количество застрахованных
                                FieldByName('INSURER2').AsString + '|' +
                                DecodeFU12(FieldByName('INSTYPE').AsString, '9') + '|' +
                                '1' + {DecodeR(FieldByName('RESIDENT').AsString) +} '|' +
                                '' + '|' + //PLACE_CODE
                                DecodeAddr(FieldByName('INSURER').AsString, FieldByName('INSURER2').AsString, FieldByName('ADDRESS').AsString) + '|' +
                                DecodeOwnType(FieldByName('INSTYPE').AsString) + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodePayType('N') + '|' +
                                DecodeAgNameL(FieldByName('AGENT').AsString) + '|' +
                                IntToStr(FieldByName('INSSUM').AsInteger * max(FieldByName('INSCNT').AsInteger, 1)) + '|' + //Страховая сумма по полису (на весь период)
                                IntToStr(FieldByName('INSSUM').AsInteger * max(FieldByName('INSCNT').AsInteger, 1)) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency(FieldByName('INSSUMCURR').AsString) + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('PAY').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TARIF').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency(FieldByName('INSSUMCURR').AsString) + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' + '|' + //Для заметок
                                '' + '|' + //MOBTEL
                                '' + '|' + //HOMETEL
                                DecodeAgentID(FieldByName('AGENT').AsString) + '|' + //Agent Doverennost
                                '' + '|' + //SMS
                                '' + '|' + //CONTACT
                                PrevKod + '|' +
                                PrevSeria + '|' +
                                PrevNumber
                    );

                    if FieldByName('INSCNT').AsInteger <= 1 then
                    begin
                        WriteLn(L, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                    DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                    FieldByName('SERIA').AsString + '|' +
                                    FieldByName('NUMBER').AsString + '|' +
                                    '1' + '|' + //Номер застрахованного
                                    DecodeCurrency(FieldByName('INSSUMCURR').AsString) + '|' + //Код валюты
                                    IntToStr(FieldByName('INSSUM').AsInteger * max(FieldByName('INSCNT').AsInteger, 1)) + '|' + //Страховая сумма
                                    FieldByName('INSURER').AsString + '|' +
                                    DecodeTxtDate(FieldByName('BIRTHDATE').AsString) + '|' + //Дата рождения
                                    DecodePassport(FieldByName('PASSPORT').AsString) + '|' + //Паспорт
                                    FieldByName('ADDRESS').AsString
                        );
                    end;
                    WriteLn(F15, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeAssist(FieldByName('ASSISTANT').AsString)
                    );
                    {WriteLn(F3, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                FieldByName('LETTER').AsString + '|' + //Тип транспортного средства
                                FieldByName('MODEL').AsString + '|' +
                                FieldByName('AUTONMB').AsString + '|' +
                                '' + '|' + //Год выпуска
                                DecodeBodyMotor(FieldByName('BODY').AsString, FieldByName('MOTOR').AsString) + '|' +
                                //FieldByName('BODY').AsString + '|' +
                                //FieldByName('MOTOR').AsString + '|' +
                                '' + '|' + //Серия техпаспорта
                                '' + '|' + //Номер техпаспорта
                                '' + '|' + //Charact
                                '' + '|' + //Стоимость
                                '' + '|' + //Код валюты (стоимости и страховой суммы)
                                '' + '|' + //Страховая сумма по объекту
                                {?} {DecodeVladen2(FieldByName('VLADEN').AsString) + '|' +
                                '' + '|'//Символ страны прибытия
                    );}

                    if FieldByName('KS').AsString <> '' then begin
                        DecodeKS(FieldByName('KS').AsString, List);
                        for i := 0 to (List.Count DIV 2) - 1 do
                        WriteLn(F6, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                    DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                    FieldByName('SERIA').AsString + '|' +
                                    FieldByName('NUMBER').AsString + '|' +
                                    '|' + //Номер застрахованного не указан
                                    List[i * 2] + '|' + //Код коэффициента
                                    '4|' + //Единица измерения Коэффициент
                                    List[i * 2 + 1] + '|' //Значение коэффициента
                        );
                    end;

                    if FieldByName('PSERIA').AsString = '' then begin
                        WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                    DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                    FieldByName('SERIA').AsString + '|' +
                                    FieldByName('NUMBER').AsString + '|' +
                                    '5' + '|' + //Код события
                                    DecodeNB(FieldByName('CASH').AsString) + '|' + //Форма оплаты
                                    '' + '|' +
                                    DecodeGlDate(FieldByName('PAYDT').AsDateTime) + '|' + //Дата платежного документа
                                    DecodeGlDate(FieldByName('PAYDT').AsDateTime) + '|' +
                                    DecodeGlDate(FieldByName('REPDT').AsDateTime) + '|' + //Дата  проведено по бухгалтерии

                                    FieldByName('SUM1').AsString + '|' + //Сумма
                                    DecodeCurrency(FieldByName('SUM1CURR').AsString) + '|' + //Код валюты
                                    FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                    DecodeAgNameL(FieldByName('AGENT').AsString) + '|' +
                                    DecodeFUFU(FieldByName('AGTYPE').AsString, FieldByName('AGPERCENT').AsFloat),
                                    DecodeAgentID(FieldByName('AGENT').AsString)
                        );

                        if DecodeCurrency(FieldByName('SUM1CURR').AsString) = '?' then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' неправильно указана валюта 1 "' + FieldByName('SUM1CURR').AsString + '"');
                    end;

                    if FieldByName('SUM2').AsFloat > 0 then begin
                    WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('CASH').AsString) + '|' + //Форма оплаты
                                '' + '|' +
                                DecodeGlDate(FieldByName('PAYDT').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAYDT').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('REPDT').AsDateTime) + '|' + //Дата  проведено по бухгалтерии
                                FieldByName('SUM2').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('SUM2CURR').AsString) + '|' + //Код валюты
                                FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                DecodeAGNameL(FieldByName('AGENT').AsString) + '|' +
                                DecodeFUFU(FieldByName('AGTYPE').AsString, FieldByName('AGPERCENT').AsFloat),
                                DecodeAgentID(FieldByName('AGENT').AsString)
                    );
                    if DecodeCurrency(FieldByName('SUM2CURR').AsString) = '?' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' неправильно указана валюта 2');
                    end;

                    if (FieldByName('STATE').AsString <> '2') AND (FieldByName('RETSUM').AsFloat >= 0.01) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' не правильное состояние полиса с возвратом');

                    if (FieldByName('STATE').AsString = '2') AND (FieldByName('RETSUM').AsFloat >= 0.01) then
                    WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '51' + '|' + //Код события
                                DecodeNB(FieldByName('CASH').AsString) + '|' + //Форма оплаты
                                '|' +
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                                FieldByName('RETSUM').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('RETCUR').AsString) + '|' + //Код валюты
                                FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                DecodeAGNameL(FieldByName('AGENT').AsString) + '|' +
                                DecodeFUFU(FieldByName('AGTYPE').AsString, FieldByName('AGPERCENT').AsFloat),
                                ''
                    );
                    if FieldByName('STATE').AsString = '2' then
                    WriteLn(F11, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '|' +
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                                '2' + '|' +
                                '2' + '|' + 
                                '41'
                    );
                    {if FieldByName('PRCPNMB').AsString <> '' then //Сцепка
                    WriteLn(F9, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                FieldByName('PRCPNMB').AsString
                    );}
                //end;
                //else begin //Дубликаты
                if FieldByName('PSERIA').AsString <> '' then begin
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank2005(FieldByName('PSERIA').AsString) + '|' +
                                FieldByName('PSERIA').AsString + '|' +
                                FieldByName('PNUMBER').AsString + '|' +
                                DecodeBlank2005(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '9' + '|' +
                                DecodeGlDate(FieldByName('REGDT').AsDateTime)
                    );
                end;
            end; //if
            Next;
          except
            on E : Exception do begin

                //if Pos('Access', e.Message) <> 0 then
                //exit;

                errStr := errStr + #13#10;
                if Pos('/', e.Message) = 0 then errStr := errStr + FieldByName('NUMBER').AsString + ' ';
                errStr := errStr + e.Message + ' ' + FieldByName('AGENT').AsString;
                IsError := true;
                Next;
                //break;
            end
        end; //while
    end; //width
  end;
    //StatusPanel.Visible := false;
    //StatusMsg.Visible := true;
    CloseFile(F1);
    CloseFile(L);
//    CloseFile(F3);
    CloseFile(F6);
//    CloseFile(F9);
    CloseFile(F8);
    CloseFile(A);
    CloseFile(F10);
    CloseFile(F11);
    CloseFile(F12);
    CloseFile(F13);
    CloseFile(F15);
    CloseFile(X);
    ProgressBar.Position := 0;

    Screen.Cursor := crDefault;
    List.Free;

    if not IsError then begin
        INI := TIniFile.Create(BLANK_INI);
        INI.WriteString(Section, 'UnloadNumber', UnloadN);
        INI.Free;
        MessageDlg('Экспорт ' + UnloadN + ' выполнен успешно'#13'Отправь все файлы с расширением ' + InsTypeCode + ' в каталоге ' + path + ' в дирекцию', mtInformation, [mbOk], 0);
        ShellExecute(0, 'open', StrPCopy(pbuff, path), 0, 0, SW_SHOWNORMAL);
    end
    else begin
        for Ltr := 'A' to 'Z' do
            DeleteFile(path + Division + Ltr + UnloadDt + UnloadN + '.0' + InsTypeCode);
        ShowInfo.Caption := 'Ошибки в данных';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
end;

procedure TPrivate2005Form.Panel5Resize(Sender: TObject);
begin
    ListAgents.Height := DownPanel.Top - ListAgents.Top
end;

procedure Create1CTable(Table1C: TTable; Name : string);
begin
    Table1C.TableName := Name;

    with Table1C.FieldDefs do begin
      Clear;
      with AddFieldDef do begin
        Name := 'Вид страхования';
        DataType := ftString;
        Size := 5;
      end;
      with AddFieldDef do begin
        Name := 'Дата';
        DataType := ftDate;
      end;
      with AddFieldDef do begin
        Name := 'Номер полиса';
        DataType := ftString;
        Size := 16;
      end;
      with AddFieldDef do begin
        Name := 'ФИО Клиента';
        DataType := ftString;
        Size := 60;
      end;
      with AddFieldDef do begin
        Name := 'Сумма';
        DataType := ftString;
        Size := 25;
      end;
      with AddFieldDef do begin
        Name := 'Валюта';
        DataType := ftString;
        Size := 5;
      end;
      with AddFieldDef do begin
        Name := 'Начальная дата';
        DataType := ftDate;
      end;
      with AddFieldDef do begin
        Name := 'Конечная дата';
        DataType := ftDate;
      end;
      with AddFieldDef do begin
        Name := 'Кассир';
        DataType := ftString;
        Size := 60;
      end;
    end;
    Table1C.CreateTable;
end;

procedure TPrivate2005Form.To1CClick(Sender: TObject);
label
    NextLabel;
const
    InsTypeCode : string = '08';
var
    PolCount, PayCount : integer;
    F : TextFile;
begin
    if not SaveDialog1C.Execute then exit;
    //Create1CTable(Table1C, SaveDialog1C.FileName);
    PolCount := 0;
    PayCount := 0;

    try
        DeleteFile(SaveDialog1C.FileName);
        AssignFile(F, SaveDialog1C.FileName);
        Rewrite(F);
        //Table1C.Open;
        MainQuery.First;

        Screen.Cursor := crHourGlass;
        MainGrid.DisableScroll;

        while not MainQuery.Eof do begin
            if MainQueryPSeria.AsString <> '' then goto NextLabel;
            if (MainQueryState.AsString = '1') then goto NextLabel;

            Inc(PolCount);
            Inc(PayCount);
            {
            Table1C.Append;
            Table1C.Fields[0].AsString := InsTypeCode;
            Table1C.Fields[1].AsDateTime := MainQueryRepDt.AsDateTime;
            Table1C.Fields[2].AsString := MainQueryNumber.AsString;
            Table1C.Fields[3].AsString := MainQueryInsurer.AsString;
            Table1C.Fields[4].AsFloat := MainQuerySum1.AsFloat;
            Table1C.Fields[5].AsString := MainQuerySum1Curr.AsString;
            Table1C.Fields[6].AsDateTime := MainQueryFrom.AsDateTime;
            Table1C.Fields[7].AsDateTime := MainQueryTo.AsDateTime;
            Table1C.Fields[8].AsString := MainQueryName.AsString;
            Table1C.Post;
            }

            WriteLn(F, InsTypeCode + ';' +
                       MainQueryRepDt.AsString + ';' +
                       MainQueryNumber.AsString + ';' +
                       MainQueryInsurer.AsString + ';' +
                       MainQuerySum1.AsString + ';' +
                       MainQuerySum1Curr.AsString + ';' +
                       MainQueryFrom.AsString + ';' +
                       MainQueryTo.AsString + ';' +
                       MainQueryName.AsString);

            if MainQuerySum2Curr.AsString <> '' then begin
                Inc(PayCount);
                {
                Table1C.Append;
                Table1C.Fields[0].AsString := InsTypeCode;
                Table1C.Fields[1].AsDateTime := MainQueryRepDt.AsDateTime;
                Table1C.Fields[2].AsString := MainQueryNumber.AsString;
                Table1C.Fields[3].AsString := MainQueryInsurer.AsString;
                Table1C.Fields[4].AsFloat := MainQuerySum2.AsFloat;
                Table1C.Fields[5].AsString := MainQuerySum2Curr.AsString;
                Table1C.Fields[6].AsDateTime := MainQueryFrom.AsDateTime;
                Table1C.Fields[7].AsDateTime := MainQueryTo.AsDateTime;
                Table1C.Fields[8].AsString := MainQueryName.AsString;
                Table1C.Post;
                }
                WriteLn(F, InsTypeCode + ';' +
                           MainQueryRepDt.AsString + ';' +
                           MainQueryNumber.AsString + ';' +
                           MainQueryInsurer.AsString + ';' +
                           MainQuerySum2.AsString + ';' +
                           MainQuerySum2Curr.AsString + ';' +
                           MainQueryFrom.AsString + ';' +
                           MainQueryTo.AsString + ';' +
                           MainQueryName.AsString);
            end;

            NextLabel:;
            MainQuery.Next;
        end;
        MessageDlg('Выгружено ' + IntToStr(PolCount) + ' полисов, ' + IntToStr(PayCount) + ' платежей', mtInformation, [mbOk], 0);
    except
        on e : Exception do
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
    end;

    MainGrid.EnableScroll;
    //MainGrid.DataSource := MainQuery;
    Screen.Cursor := crDefault;
    //Table1C.Close;
    CloseFile(F);
end;

end.
