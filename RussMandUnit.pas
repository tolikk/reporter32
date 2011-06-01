unit RussMandUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, DBGrids, RXDBCtrl, RxQuery, StdCtrls,
  ExtCtrls, ComCtrls, CheckLst, Buttons, Menus, datamod, Math, Placemnt;

type
  TRUSSMAND = class(TForm)
    MainQuery: TRxQuery;
    RussDatabase: TDatabase;
    DataSource: TDataSource;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryPSer: TStringField;
    MainQueryPNmb: TFloatField;
    MainQueryRegDt: TDateField;
    MainQueryInsStart: TDateField;
    MainQueryTmStart: TStringField;
    MainQueryPeriod: TSmallintField;
    MainQueryInsEnd: TDateField;
    MainQueryInsurer: TStringField;
    MainQueryInsurerT: TStringField;
    MainQueryOwner: TStringField;
    MainQueryOwnerT: TStringField;
    MainQueryMarka: TStringField;
    MainQueryAutoId: TStringField;
    MainQueryPassSer: TStringField;
    MainQueryPassNmb: TFloatField;
    MainQueryAutoNmb: TStringField;
    MainQuerySpecSr: TStringField;
    MainQuerySpecNmb: TFloatField;
    MainQueryTarifGrp: TSmallintField;
    MainQueryCharact: TSmallintField;
    MainQueryTarif: TFloatField;
    MainQueryPayDate: TDateField;
    MainQueryPay1: TFloatField;
    MainQueryCurr1: TStringField;
    MainQueryPay2: TFloatField;
    MainQueryCurr2: TStringField;
    MainQueryAgent: TStringField;
    MainQueryAgType: TStringField;
    MainQueryAgPcnt: TFloatField;
    MainQueryState: TSmallintField;
    MainQueryStopDt: TDateField;
    MainQueryRetSum: TFloatField;
    MainQueryRetCurr: TStringField;
    MainQueryInsComp: TSmallintField;
    MainQueryRepDt: TDateField;
    Panel2: TPanel;
    Panel3: TPanel;
    btnClose: TButton;
    TabbedNotebook: TPageControl;
    RepTab: TTabSheet;
    FilterTab: TTabSheet;
    MainGrid: TRxDBGrid;
    Label11: TLabel;
    ListTemplates: TComboBox;
    AgentList: TCheckListBox;
    IsNumber: TCheckBox;
    Label1: TLabel;
    NumberFrom: TEdit;
    Label2: TLabel;
    NumberTo: TEdit;
    IsRegDate: TCheckBox;
    Label3: TLabel;
    RegDateFrom: TDateTimePicker;
    Label4: TLabel;
    RegDateTo: TDateTimePicker;
    IsPayDate: TCheckBox;
    Label7: TLabel;
    PayDateFrom: TDateTimePicker;
    Label10: TLabel;
    PayDateTo: TDateTimePicker;
    IsOwner: TCheckBox;
    OwnerLike: TEdit;
    IsState: TCheckBox;
    State: TComboBox;
    Label5: TLabel;
    Label6: TLabel;
    AgentsTbl: TQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    IsRepDate: TCheckBox;
    DateRepStart: TDateTimePicker;
    DateRepEnd: TDateTimePicker;
    Panel4: TPanel;
    btnSort: TButton;
    btnExcel: TButton;
    SpeedButton: TSpeedButton;
    Timer: TTimer;
    CountLabel: TLabel;
    btnApply: TButton;
    PopupMenu: TPopupMenu;
    MenuItem1: TMenuItem;
    N2: TMenuItem;
    MainQueryPrevSN: TStringField;
    MainQueryPay1Info: TStringField;
    MainQueryPay2Info: TStringField;
    MainQueryAgName: TStringField;
    MainQueryCompanyName: TStringField;
    MainQueryStateName: TStringField;
    MainQueryRetSumm: TStringField;
    btnReserveCalc: TButton;
    StatPage: TTabSheet;
    TabSheet2: TTabSheet;
    StatisticTxt: TMemo;
    Label8: TLabel;
    RepList: TComboBox;
    Label9: TLabel;
    ParamsText: TEdit;
    StartButton: TButton;
    Label12: TLabel;
    RxDBGrid1: TRxDBGrid;
    IsNumbers: TCheckBox;
    Button2: TButton;
    REPQuery: TRxQuery;
    REPDS: TDataSource;
    btnMag11: TButton;
    WorkSQL: TQuery;
    ProgressBar: TProgressBar;
    btnRNPU: TButton;
    IsUseFilter: TCheckBox;
    listCompanies: TComboBox;
    IsFiltCompany: TCheckBox;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    mnuExport: TMenuItem;
    MainQueryresident: TStringField;
    listOwnType: TComboBox;
    CntChecked: TLabel;
    InputSum: TMenuItem;
    N9: TMenuItem;
    ImpAlvena: TMenuItem;
    OpenDialogPx: TOpenDialog;
    SaveDialog: TSaveDialog;
    FormStorage: TFormStorage;
    IsStopDate: TCheckBox;
    auxTable: TTable;
    IsBad: TCheckBox;
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure btnSortClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure btnExcelClick(Sender: TObject);
    procedure MainGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure RepListChange(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure StartButtonClick(Sender: TObject);
    procedure TabbedNotebookChange(Sender: TObject);
    procedure btnMag11Click(Sender: TObject);
    procedure btnReserveCalcClick(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure mnuExportClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure AgentListClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure InputSumClick(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure ImpAlvenaClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure IsRepDateClick(Sender: TObject);
  private
     { Private declarations }
     sqlOrder : string;
     Responsibility : string;
     AgCodes : TStringList;
     W : HWND;
     Company : array[0..20] of string[48];
     IniName : string;
     MSEXCEL : VARIANT;
     function GetFilter(sCond : string) : string;
     function GetOrder : string;
     procedure ReOpenQuery(Sender: TObject);
     procedure CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List : TASKOneTaxList);
     procedure ImportAlvena(DBFile : string);
  public
    { Public declarations }
  end;

RussGlobo = record
   Company : integer;
   Ser: string;
   Id: String;        
end;

procedure CreateSimpleMagazine(WorkSQL : TQuery; SQL : string; InsTypeName : string);

var
  RUSSMAND: TRUSSMAND;
  lastAgCode : string;
  lastAgName : string;
  ShowUndefCharact : integer = -1;
  BlankKodes : array[0..100] of RussGlobo;

implementation

uses inifiles, registry, Sort, InpRezParams,
  WaitFormUnit, TmplNameUnit, RepParamsFrm, RepParamsAgntFrm, ShowInfoUnit,
  PcntFormUnit, StrUtils, BelGreenRpt, RxStrUtils, GetRepDateUnit, mf, DM;

{$R *.dfm}

procedure TRUSSMAND.btnCloseClick(Sender: TObject);
begin
    Close
end;

procedure TRUSSMAND.FormCreate(Sender: TObject);
var
     BLINI : TIniFile;
     rINI : TRegIniFile;
     i : integer;
     str : string;
     R : TRect;
begin
     State.ItemIndex := 0;
     DMM.HandBookSQL.DatabaseName := RussDatabase.Name;

     rINI := TRegIniFile.Create('Россия');
     for i := 0 to 100 do begin
         str := rINI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
         //else
         //    break;
     end;
     rINI.Free;

     TabbedNotebook.ActivePageIndex := 0;
     sqlOrder := 'SERIA,NUMBER';
     
     BLINI := TIniFile.Create(BLANK_INI);
     IniName := BLINI.ReadString('RUSSIAN', 'INI', BLANK_INI);
     BLINI.Free;
     BLINI := TIniFile.Create(IniName);
     Responsibility := BLINI.ReadString('RUSSIAN', 'Responsibility', '400000');
     for i := 0 to 100 do begin
         str := BLINI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
     end;
     ListTemplates.ItemIndex := 0;

     for i := 0 to 100 do begin
         str := BLINI.ReadString('RUSSIAN', 'Company' + IntToStr(i), '');
         if str <> '' then begin
             Company[i] := str;
             listCompanies.Items.Add(str);
             listCompanies.Items.Objects[listCompanies.Items.Count - 1] := Pointer(i)
         end
     end;

    //Заполним список отчётов
    for i := 1 to 100 do begin
        if BLINI.ReadString('RUSSIAN', 'QueryName' + IntToStr(i), '') <> '' then
            RepList.Items.Add(BLINI.ReadString('RUSSIAN', 'QueryName' + IntToStr(i), ''))
        else
            break;
    end;

     {for i := 0 to 100 do begin
         str := BLINI.ReadString('RUSSMAND', 'AutoType' + IntToStr(i), '');
         if str <> '' then
             AutoTypes[i] := str;
     end;}
     BLINI.Free;

     AgCodes := TStringList.Create;
     AgentsTbl.Open;
     while not AgentsTbl.EOF do begin
           AgentList.Items.Add(AgentsTblName.AsString);
           AgCodes.Add(AgentsTblAgent_code.AsString);
           AgentsTbl.Next;
     end;
     AgentsTbl.Close;

     MainQuery.MacroByName('WHERE').AsString := GetFilter(' WHERE ');
     MainQuery.MacroByName('ORDER').AsString := GetOrder();
     MainQuery.Open;
     CountLabel.Caption := 'Количество записей ' + IntToStr(MainQuery.RecordCount);

     W := FindWindow('OWLWindow31', nil);
     if (W <> 0) then begin
       GetWindowRect(W, R);
       Left := R.Left;
       Top := R.Top;
       Width := R.Right - R.Left;
       Height := R.Bottom - R.Top;
       EnableWindow(W, false);
       BorderStyle := bsSingle;
       DeleteMenu(GetSystemMenu(Handle, false), SC_MOVE, MF_BYCOMMAND);
       DeleteMenu(GetSystemMenu(Handle, false), SC_MINIMIZE, MF_BYCOMMAND);
       if IsZoomed(W) then begin
        Left := 0;
        Top := 0;
        Width := GetSystemMetrics(SM_CXSCREEN);
        Height := GetSystemMetrics(SM_CYSCREEN);
        WindowState := wsMaximized;
       end;
       ShowWindow(Handle, SW_SHOW);
       ShowWindow(W, SW_HIDE);
     end
     else begin
        WindowState := wsMaximized;
     end;
end;

procedure Connect(var s: string; cond : string; op : string = 'AND');
begin
     if s <> '' then
         s := s + ' ' + op + ' ';
     s := s + cond;
end;

function TRUSSMAND.GetFilter(sCond: string): string;
var
     s, str, str2, str3 : string;
     i : integer;
begin
     if IsNumber.Checked then begin
        if Trim(NumberFrom.Text) <> '' then
            Connect(s, 'NUMBER>=' + NumberFrom.Text);
        if Trim(NumberTo.Text) <> '' then
            Connect(s, 'NUMBER<=' + NumberTo.Text);
     end;

     if IsRegDate.Checked then begin
        if RegDateFrom.Checked then
            Connect(s, 'REGDT>=''' + DateToStr(RegDateFrom.Date) + '''');
        if RegDateTo.Checked then
            Connect(s, 'REGDT<=''' + DateToStr(RegDateTo.Date) + '''');
     end;

     if IsPayDate.Checked then begin
        if PayDateFrom.Checked then
            Connect(s, 'PAYDATE>=''' + DateToStr(PayDateFrom.Date) + '''');
        if PayDateTo.Checked then
            Connect(s, 'PAYDATE<=''' + DateToStr(PayDateTo.Date) + '''');
     end;

     if IsRepDate.Checked then begin
        if DateRepStart.Checked then
            Connect(str, 'REPDT>=''' + DateToStr(DateRepStart.Date) + '''');
        if DateRepEnd.Checked then
            Connect(str, 'REPDT<=''' + DateToStr(DateRepEnd.Date) + '''');

        if IsStopDate.Checked then begin
           if DateRepStart.Checked then
              Connect(str2, 'STOPDT>=''' + DateToStr(DateRepStart.Date) + '''');
           if DateRepEnd.Checked then
              Connect(str2, 'STOPDT<=''' + DateToStr(DateRepEnd.Date) + '''');
        end;

        if IsBad.Checked then begin
           if DateRepStart.Checked then
              Connect(str3, 'REGDT>=''' + DateToStr(DateRepStart.Date) + '''');
           if DateRepEnd.Checked then
              Connect(str3, 'REGDT<=''' + DateToStr(DateRepEnd.Date) + '''');
        end;

        if (str <> '') then begin
            if str2 <> '' then
                Connect(str, str2, 'OR');
            if str3 <> '' then
                Connect(str, str3, 'OR');
            Connect(s, '(' + str + ')');
        end;
     end;

     //Агенты
     str := '';
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if Length(str) > 0 then
          Connect(s, '(AGENT IN (' + Copy(str, 1, Length(str) - 1) + '))');

     if IsOwner.Checked  then begin
        if (OwnerLike.Text <> '') then
            Connect(s, '(OWNER LIKE ''%' + OwnerLike.Text + '%'' OR OWNER LIKE ''%' + AnsiUpperCase(OwnerLike.Text) + '%'' OR OWNER LIKE ''%' + AnsiLowerCase(OwnerLike.Text) + '%'')');

         if listOwnType.ItemIndex = 1 then
             Connect(s, 'ownert=''F''');
         if listOwnType.ItemIndex = 2 then
             Connect(s, 'ownert=''U''');

         {if listResidType.ItemIndex = 1 then
             Connect(s, 'resident=''Y''');
         if listResidType.ItemIndex = 2 then
             Connect(s, 'resident=''N''');}
     end;

     if IsState.Checked then begin
         if State.ItemIndex = 0 then
             Connect(s, '(State = ' + IntToStr(State.ItemIndex) + ' AND (PSER='''' OR PSER IS NULL))')
         else
         if State.ItemIndex = 4 then
             Connect(s, '(PSER<>'''')')
         else
             Connect(s, '(State = ' + IntToStr(State.ItemIndex) + ')');
     end;

     if IsFiltCompany.Checked AND (listCompanies.ItemIndex >= 0) then begin
        Connect(s, '(INSCOMP=' + InttoStr(Integer(listCompanies.Items.Objects[listCompanies.ItemIndex])) + ')')
     end;

     if s <> '' then s := sCond + s;
     GetFilter := s;
end;

function TRUSSMAND.GetOrder: string;
begin
    if sqlOrder = '' then GetOrder := ''
    else GetOrder := 'ORDER BY ' + sqlOrder;
end;

procedure TRUSSMAND.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top + SpeedButton.Height;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TRUSSMAND.btnSortClick(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then begin
         sqlOrder := SortForm.GetOrder('');
         ReOpenQuery(nil);
         if TabbedNotebook.ActivePageIndex = 2 then
           Timer.Enabled := true
     end;
end;

procedure TRUSSMAND.ReOpenQuery(Sender: TObject);
begin
     MainQuery.Close;
     MainQuery.MacroByName('WHERE').AsString := GetFilter(' WHERE ');
     MainQuery.MacroByName('ORDER').AsString := GetOrder();
     CountLabel.Caption := '';
     try
        Screen.Cursor := crHourGlass;
        MainQuery.Open;
        CountLabel.Caption := 'Количество записей ' + IntToStr(MainQuery.RecordCount);
     except
        on E : Exception do
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

procedure TRUSSMAND.FormDestroy(Sender: TObject);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TRUSSMAND.TimerTimer(Sender: TObject);
begin
     Timer.Enabled := false;
end;

procedure TRUSSMAND.ListTemplatesChange(Sender: TObject);
var
     i, fnd_idx, fnd_token : integer;
     INI : TRegIniFile;
     str : string;
     Code : string;
begin
     //if ListTemplates.ItemIndex = 0 then begin
       for i := 0 to AgentList.Items.Count - 1 do
            AgentList.Checked[i] := false;
       //exit;
     //end;

     Code := '';
     INI := TRegIniFile.Create('Россия');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str = ListTemplates.Text then begin
             str := INI.ReadString('AgTemplates', IntToStr(i), '');

             for fnd_token := 1 to Length(str) do begin
                 if str[fnd_token] = ',' then begin
                     for fnd_idx := 0 to AgentList.Items.Count - 1 do
                         if AgCodes[fnd_idx] = Code then
                             AgentList.Checked[fnd_idx] := true;

                     Code := '';
                 end
                 else
                     Code := Code + str[fnd_token];
             end;
         end;
     end;
     INI.Free;

     AgentListClick(nil);
end;

procedure TRUSSMAND.btnApplyClick(Sender: TObject);
begin
     MainQuery.Close;
     MainQuery.MacroByName('WHERE').AsString := GetFilter(' WHERE ');
     MainQuery.MacroByName('ORDER').AsString := GetOrder();
     CountLabel.Caption := '';
     try
        Screen.Cursor := crHourGlass;
        MainQuery.Open;
        CountLabel.Caption := 'Количество записей ' + IntToStr(MainQuery.RecordCount);
        TabbedNotebook.ActivePageIndex := 0;
     except
        on E : Exception do
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

procedure TRUSSMAND.MainQueryCalcFields(DataSet: TDataSet);
begin
    MainQueryPrevSN.AsString := '';
    MainQueryPay1Info.AsString := '';
    MainQueryPay2Info.AsString := '';
    MainQueryCompanyName.AsString := '';
    MainQueryStateName.AsString := '';
    if MainQueryPSer.AsString <> '' then MainQueryPrevSN.AsString := MainQueryPSer.AsString + '/' + MainQueryPNmb.AsString;
    if MainQueryPay1.AsFloat > 0.001 then MainQueryPay1Info.AsString := MainQueryPay1.AsString + ' ' + MainQueryCurr1.AsString;
    if MainQueryPay2.AsFloat > 0.001 then MainQueryPay2Info.AsString := MainQueryPay2.AsString + ' ' + MainQueryCurr2.AsString;
    MainQueryCompanyName.AsString := Company[MainQueryInsComp.AsInteger];
    MainQueryStateName.AsString := State.Items[MainQueryState.AsInteger];
    if MainQueryPay1.AsFloat > 0.001 then MainQueryPay1Info.AsString := MainQueryPay1.AsString + ' ' + MainQueryCurr1.AsString;
    if MainQueryRetSum.AsFloat > 0.001 then MainQueryRetSumm.AsString := MainQueryRetSum.AsString + ' ' + MainQueryRetCurr.AsString;

end;

procedure TRUSSMAND.btnExcelClick(Sender: TObject);
begin
    DatSetToExcel(MainQuery);
end;

procedure TRUSSMAND.MainGridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
    if MainQueryState.AsInteger = 2 then
        Background := $008080FF;

    if MainQueryState.AsInteger = 3 then
        Background := clRed;

    if MainQueryPSer.AsString <> '' then
        Background := $007A7ABC;

    if Field = MainQueryCompanyName then
        Background := clSkyBlue;
    if Field = MainQueryPay1Info then
        Background := $00FFB3B3;
    if Field = MainQueryPay2Info then
        Background := $00FFB3B3;
end;

procedure TRUSSMAND.RepListChange(Sender: TObject);
var
    INI : TIniFile;
begin
    REPQuery.Close;
    INI := TIniFile.Create(IniName);
    ParamsText.Hint := INI.ReadString('RUSSIAN', 'QueryParams' + IntToStr(RepList.ItemIndex + 1), '');
    ParamsText.Text := INI.ReadString('RUSSIAN', 'LastParams' + IntToStr(RepList.ItemIndex + 1), '');
    if ParamsText.Text = '' then
        ParamsText.Text := ParamsText.Hint; 
    INI.Free;
end;

procedure TRUSSMAND.Button2Click(Sender: TObject);
begin
    if REPQuery.Active then
        DatSetToExcel(REPQuery, IsNumbers.Checked);
end;

procedure TRUSSMAND.StartButtonClick(Sender: TObject);
var
    s : string;
begin
    s := '';
    if IsUseFilter.Checked then s := GetFilter('');
    ExecLittleReport(IniName, 'RUSSIAN', ParamsText.Text, RepList.ItemIndex + 1, REPQuery, s);
end;

procedure TRUSSMAND.TabbedNotebookChange(Sender: TObject);
var
    s : string;
begin
    CountLabel.Visible := TabbedNotebook.ActivePageIndex < 2;

    StatisticTxt.Text := '';

    if TabbedNotebook.ActivePage = StatPage then begin
        try
            Screen.Cursor := crHourglass;
            StatisticTxt.Lines.Add('Отчёт создаётся на основании фильтра');
            with WorkSQL, WorkSQL.SQL do begin
                //if RepDtFrom.Checked OR RepDtTo.Checked then begin
                    s := 'ОТЧЁТНЫЙ ПЕРИОД ';
                    if DateRepStart.Checked then
                        s := s + ' C ' + DateToStr(DateRepStart.Date);
                    if DateRepEnd.Checked then
                        s := s + ' ПО ' + DateToStr(DateRepEnd.Date);
                    if DateRepStart.Checked and DateRepEnd.Checked then
                        s := s + #13#10' за ' + IntToStr(Round(DateRepEnd.Date-DateRepStart.Date+1)) + ' дней'
                    else
                        s := s + 'НЕ УКАЗАН';

                    StatisticTxt.Lines.Add(s);
                    StatisticTxt.Lines.Add('');
                //end;

                Close;
                Clear;
                Add('SELECT SUM(TARIF) AS S');
                Add('FROM MANDRUSS');
                Add('WHERE TARIF>0 ' + GetFilter(' AND '));
                Open;
                StatisticTxt.Lines.Add('Тариф');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FloatToStr(RoundTo(FieldByName('S').AsFloat, -2)));
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(PAY1) AS S, CURR1 AS CUR');
                Add('FROM MANDRUSS');
                Add('WHERE CURR1<>'''' ' + GetFilter(' AND '));
                Add('GROUP BY CURR1');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Первая часть');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FloatToStr(RoundTo(FieldByName('S').AsFloat, -2)) + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;

                Close;
                Clear;
                Add('SELECT SUM(PAY2) AS S, CURR2 AS CUR');
                Add('FROM MANDRUSS');
                Add('WHERE CURR2<>'''' ' + GetFilter(' AND '));
                Add('GROUP BY CURR2');
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Вторая часть');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FloatToStr(RoundTo(FieldByName('S').AsFloat, -2)) + ' ' + FieldByName('CUR').AsString);
                    Next;
                end;
                Close;

                Clear;
                Add('SELECT COUNT(*) AS C');
                Add('FROM MANDRUSS');
                Add('WHERE STATE=''2'' ' + GetFilter(' AND '));
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
                Add('FROM MANDRUSS');
                Add('WHERE STATE=''1'' ' + GetFilter(' AND '));
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
                Add('FROM MANDRUSS');
                Add('WHERE PNMB>0 ' + GetFilter(' AND '));
                Open;
                StatisticTxt.Lines.Add('');
                StatisticTxt.Lines.Add('Дубликатов');
                while not Eof do begin
                    StatisticTxt.Lines.Add(FieldByName('C').AsString);
                    Next;
                end;
                Close;
            end;
        except
          on E : Exception do begin
            StatisticTxt.Text := E.Message;
          end;
        end;
    end;
    Screen.Cursor := crDefault;
end;

procedure CreateSimpleMagazine(WorkSQL : TQuery; SQL : string; InsTypeName : string);
var
   MagazineLeben : integer;
   StartDate, EndDate : TDateTime;
   MonthName : string;
   ASH, ArrayData, MSEXCEL : VARIANT;
   Y, M, D : WORD;
   AgKf : double;
   Line, LastLine : integer;
   LineStr, TempStr : string;
   OutFileName : string;
begin
   if GetRepDate.ShowModal <> mrOk then exit;

     if not InitExcel(MSEXCEL) then exit;
     MSEXCEL.Visible := true;
     MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\ЖУРНАЛ УЧЁТА ЗАКЛЮЧЁННЫХ ДОГОВОРОВ.XLS');
     ASH := MSEXCEL.ActiveSheet;
     ASH.Range['C4','C4'].Value := InsTypeName;
     try
           DecodeDate(GetRepDate.RepDate.Date, Y, M, D);
           D := 1;
           if M = 1 then
               IncAMonth(Y, M, D, -1)
           else
              M := 1;

           StartDate := EncodeDate(Y, M, D);
           MonthName := FormatDateTime('MMMM YYYY', StartDate);
           IncAMonth(Y, M, D, 1);
           EndDate := EncodeDate(Y, M, D) - 1;

           Line := 22;
                            
           while(StartDate < TRUNC(GetRepDate.RepDate.Date)) do begin
               LineStr := InttoStr(Line);
               ASH.Range['A' + LineStr, 'A' + LineStr].Value := MonthName;
               Inc(Line);
               WorkSQL.SQL.Clear;
               WorkSQL.SQL.Add(SQL);
               WorkSQL.ParamByName('D1').AsdateTime := StartDate;
               WorkSQL.ParamByName('D2').AsdateTime := EndDate;
               WorkSQL.Open;


           ArrayData := VarArrayCreate([1, 1, 1, 27], varVariant);
           while not WorkSQL.EOF do begin
                 //Не испорчен и не дубликат
                 with WorkSQL do begin
                       ArrayData[1, 1] := FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                       ArrayData[1, 2] := FieldByName('INSSTART').AsString;
                       ArrayData[1, 3] := FieldByName('INSURER2').AsString;
                       ArrayData[1, 4] := FieldByName('INSURER').AsString;
                       ArrayData[1, 5] := FieldByName('AllSumma').AsFloat;
                       ArrayData[1, 6] := FieldByName('AllSummaCurr').AsString;
                       ArrayData[1, 7] := FieldByName('INSSTART').AsString;
                       ArrayData[1, 8] := FieldByName('INSEND').AsString;

                       TempStr := FieldByName('PAYMENTDATE').AsString;

                       if (FieldByName('PAYMENTDATE').ASdateTime>=StartDate) AND
                          (FieldByName('PAYMENTDATE').ASdateTime<=EndDate) then begin
                               ArrayData[1, 9] := TempStr;
                               ArrayData[1, 10] := FieldByName('Sum1Curr').AsString;
                               ArrayData[1, 11] := FieldByName('Sum1').AsFloat;
                               if FieldByName('Sum1Curr').AsString = 'BRB' then
                                   ArrayData[1, 12] := FieldByName('Sum1').AsFloat
                               else
                                   ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Sum1Curr').AsString, FieldByName('PAYMENTDATE').AsDateTime) * FieldByName('Sum1').AsFloat;

                               ArrayData[1, 13] := ArrayData[1, 11];
                               ArrayData[1, 14] := ArrayData[1, 12];
                               if (FieldByName('AgUr').AsString = 'True') OR (FieldByName('AgUr').AsString = 'Y') OR (FieldByName('AgUr').AsString = 'U') then 
                                   AgKf := 1 + GetRepDate.UrTax.Value / 100
                               else
                                   AgKf := 1 + GetRepDate.FizTax.Value / 100;
                               ArrayData[1, 15] := double(ArrayData[1, 11]) * FieldByName('AgPercent').AsFloat / 100 * AgKf;
                               ArrayData[1, 16] := double(ArrayData[1, 12]) * FieldByName('AgPercent').AsFloat / 100 * AgKf;
                               ArrayData[1, 17] := double(ArrayData[1, 13]) * FieldByName('AgPercent').AsFloat / 100 * AgKf;
                               ArrayData[1, 18] := double(ArrayData[1, 14]) * FieldByName('AgPercent').AsFloat / 100 * AgKf;
                       end
                       else begin
                               ArrayData[1, 9] := '';
                               ArrayData[1, 10] := '';
                               ArrayData[1, 11] := '';
                               ArrayData[1, 12] := '';
                               ArrayData[1, 13] := '';
                               ArrayData[1, 14] := '';
                               ArrayData[1, 15] := '';
                               ArrayData[1, 16] := '';
                               ArrayData[1, 17] := '';
                               ArrayData[1, 18] := '';
                       end;

                       //IsNotReinsurPolis := true;
                       //OutPercent := 0;
                       //if MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1] <> '' then begin
                       //    OutPercent := StrToInt(MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1]);
                       //    OutKomis := StrtoInt(MagazineLeben.StringGrid.Cells[2, FieldByName('Restrax').AsInteger + 1]);
                       //end;

                       //if OutPercent > 0 then begin
                       //      IsNotReinsurPolis := (WorkSQL.FieldByName('End').AsDateTime - WorkSQL.FieldByName('Begin').AsDateTime + 1) <= 30;
                       //      if not IsNotReinsurPolis then begin
                       //          //Проверить, отдали в перестрахование или нет
                       //          IsNotReinsurPolis := not DecideIsReinsur(MagazineLeben.RepData.Date, WorkSQL.FieldByName('PaymentDate').AsDateTime, SendDate, RestDok);
                       //      end;
                       //end;

                       ArrayData[1, 19] := '';
                       ArrayData[1, 20] := '';
                       ArrayData[1, 21] := '';
                       ArrayData[1, 22] := '';
                       ArrayData[1, 23] := '';
                       ArrayData[1, 24] := '';
                       ArrayData[1, 25] := '';
                       ArrayData[1, 26] := '';
                       ArrayData[1, 27] := '';
                       {if not IsNotReinsurPolis then begin
                           ArrayData[1, 19] := RestDok;
                           ArrayData[1, 20] := MagazineLeben.StringGrid.Cells[0, FieldByName('Restrax').AsInteger + 1];

                           ArrayData[1, 21] := FieldByName('Sum1Curr').AsString;
                           ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                           ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                           ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                           ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;

                           ArrayData[1, 26] := FieldByName('AllSumma').AsFloat * OutPercent / 100;
                           ArrayData[1, 27] := OutPercent;
                       end;}

                       LastLine := Line;
                 end;
                 ASH.Range['A' + InttoStr(Line), 'AA' + InttoStr(Line)] := ArrayData;
                 Inc(Line);

                 with WorkSQL do
                 if FieldByName('Sum2').AsFloat > 0.01 then begin
                     if (FieldByName('PAYMENTDATE').ASdateTime>=StartDate) AND
                        (FieldByName('PAYMENTDATE').ASdateTime<=EndDate) then begin
                         ArrayData[1, 10] := FieldByName('Sum2Curr').AsString;
                         ArrayData[1, 11] := FieldByName('Sum2').AsFloat;
                         if FieldByName('Sum2Curr').AsString = 'BRB' then
                             ArrayData[1, 12] := FieldByName('Sum2').AsFloat
                         else
                             ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Sum2Curr').AsString, FieldByName('PAYMENTDATE').AsDateTime) * FieldByName('Sum2').AsFloat;
                         ArrayData[1, 13] := ArrayData[1, 11];
                         ArrayData[1, 14] := ArrayData[1, 12];
                         ArrayData[1, 15] := double(ArrayData[1, 11]) * FieldByName('AgPercent').AsFloat / 100;
                         ArrayData[1, 16] := double(ArrayData[1, 12]) * FieldByName('AgPercent').AsFloat / 100;
                         ArrayData[1, 17] := double(ArrayData[1, 13]) * FieldByName('AgPercent').AsFloat / 100;
                         ArrayData[1, 18] := double(ArrayData[1, 14]) * FieldByName('AgPercent').AsFloat / 100;
                     end;

                     {if not IsNotReinsurPolis then begin
                         ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                         ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                         ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                         ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;
                     end;}
                     ASH.Range['A' + InttoStr(Line), 'AA' + InttoStr(Line)] := ArrayData;
                     Inc(Line);
                 end;

                 WorkSQL.Next;
           end;

           Inc(Line);

           VarClear(ArrayData);

               StartDate := EncodeDate(Y, M, D);
               MonthName := FormatDateTime('MMMM YYYY', StartDate);
               IncAMonth(Y, M, D, 1);
               EndDate := EncodeDate(Y, M, D) - 1;
           end;

           OutFileName := GetAppDir + 'DOCS\_ЖУРНАЛ УЗД ' + InsTypeName + '.XLS';
           DeleteFile(OutFileName);
           MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
     except
           on E : Exception do
              MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
end;

procedure TRUSSMAND.btnMag11Click(Sender: TObject);
begin
    CreateSimpleMagazine(WorkSQL, 'SELECT SERIA,NUMBER, INSSTART, INSEND,OWNER INSURER, MARKA INSURER2, 1 ALLSUMMA, '''' ALLSUMMACURr, ' +
                                  'pay1 sum1, Curr1 sum1curr, pay2 sum2, Curr2 sum2curr, paydate PAYMENTDATE, AgPcnt AgPercent, AgType AgUr FROM mandruss T WHERE paydate is not null and pay1 >= 0.01 and paydate BETWEEN :D1 AND :D2', 'Российская обязаловка');
    Application.BringToFront;                              
end;

procedure TRUSSMAND.btnReserveCalcClick(Sender: TObject);
var
    obj : TaxPercents;
begin
     InitRezerv.RestoreName('Россия');
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
                      InitRezerv.List);
         obj.Free;
         Screen.Cursor := crDefault;
         Application.BringToFront;
     end;
end;

procedure TRUSSMAND.CreateReport(IsRNPU: boolean; RepDate: TDateTime;
  obj : TaxPercents; IsFast: boolean;
  List : TASKOneTaxList);
var
    ListData : TList;
    i : integer;
begin
    ListData := TList.Create;
    try
         FillPolisesForReserve(IsRNPU, ListData, TRUNC(RepDate), obj, List,
          'SELECT  Seria, Number, Repdt repdate, Insstart StartDate, insend EndDate, PayDate, pay1 PremiumPay, pay2 PremiumPay2, curr1 PremiumCurr, curr2 PremiumCurr2, agtype Uridich, agpcnt AgPercent ' +
          'FROM MANDRUSS WHERE PayDate is not null AND',
                              WorkSQL, ProgressBar, 'insstart', 'insend', 'RUR');

         //FillDocsData(IsRNPU, ListData, TRUNC(RepDate), FizichTax, UridichTax, CompanyPercent, OneTax);
         if not InitExcel(MSEXCEL) then begin
            MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
            exit;
         end;
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;
         MSEXCEL.Visible := true;

         OutTo10Form(IsRNPU, 'Российская обязаловка', MSEXCEL, RepDate, ListData, IsFast, 'RUSSIA');
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

procedure TRUSSMAND.N3Click(Sender: TObject);
begin
    Close
end;

procedure TRUSSMAND.N4Click(Sender: TObject);
begin
    TabbedNotebook.ActivePageIndex := 1
end;

procedure TRUSSMAND.N5Click(Sender: TObject);
begin
    TabbedNotebook.ActivePageIndex := 2
end;

procedure TRUSSMAND.N6Click(Sender: TObject);
begin
    TabbedNotebook.ActivePageIndex := 3
end;

procedure AddStrValue(var s : string; s2 : string);
begin
    if s <> '' then s := s + ';';
    s := s + s2
end;

function DecodeRUSSVariant(it, rsdn : string) : string;
begin
    DecodeRUSSVariant := '?';
    if (it = 'U') AND (rsdn = 'N') then DecodeRUSSVariant := '1';
    if ((it = 'F') OR (it = 'I')) AND (rsdn = 'N') then DecodeRUSSVariant := '2';
    if (it = 'U') AND (rsdn = 'Y') then DecodeRUSSVariant := '3';
    if ((it = 'F') OR (it = 'I')) AND (rsdn = 'Y') then DecodeRUSSVariant := '4';
end;

function DecodeCharact2(Letter : string; var tgrp : string) : string;
var
    charact : string;
begin
    charact := '0';
    tgrp := '-1';

    if Letter = 'B' then begin
        tgrp := '5';
    end;
    if Letter = 'V' then begin
        tgrp := '4';
    end;
    if Letter = 'F1' then begin
        tgrp := '6';
    end;
    if Letter = 'F2' then begin
        tgrp := '7';
    end;
    if Letter = 'F3' then begin
        tgrp := '8';
    end;
    if Letter = 'A1' then begin
        tgrp := '0';
        charact := '50';
    end;
    if Letter = 'A2' then begin
        tgrp := '0';
        charact := '70';
    end;
    if Letter = 'A3' then begin
        tgrp := '0';
        charact := '95';
    end;
    if Letter = 'A4' then begin
        tgrp := '0';
        charact := '120';
    end;
    if Letter = 'A5' then begin
        tgrp := '0';
        charact := '160';
    end;
    if Letter = 'A6' then begin
        tgrp := '0';
        charact := '200';
    end;
    if Letter = 'A7' then begin
        tgrp := '0';
        charact := '210';
    end;

    if Letter = 'C1' then begin
        tgrp := '2';
        charact := '10';
    end;
    if Letter = 'C2' then begin
        tgrp := '2';
        charact := '20';
    end;

    if Letter = 'E1' then begin
        tgrp := '3';
        charact := '20';
    end;
    if Letter = 'E2' then begin
        tgrp := '3';
        charact := '25';
    end;

    if Letter = 'T1' then begin
        tgrp := '1';
        charact := '50';
    end;
    if Letter = 'T2' then begin
        tgrp := '1';
        charact := '70';
    end;
    if Letter = 'T3' then begin
        tgrp := '1';
        charact := '95';
    end;
    if Letter = 'T4' then begin
        tgrp := '1';
        charact := '120';
    end;
    if Letter = 'T5' then begin
        tgrp := '1';
        charact := '160';
    end;
    if Letter = 'T6' then begin
        tgrp := '1';
        charact := '200';
    end;
    if Letter = 'T7' then begin
        tgrp := '1';
        charact := '210';
    end;

    DecodeCharact2 := charact;
end;

function DecodeLetter(tgrp, charact : string) : string;
begin
    DecodeLetter := '?';
    if tgrp = '5' then begin
        DecodeLetter := 'B';
    end;
    if tgrp = '4' then begin
        DecodeLetter := 'V';
    end;
    if tgrp = '6' then begin
        DecodeLetter := 'F1';
    end;
    if tgrp = '7' then begin
        DecodeLetter := 'F2';
    end;
    if tgrp = '8' then begin
        DecodeLetter := 'F3';
    end;
    if tgrp = '0' then begin
        try
            StrtoInt(charact);
        except
            if ShowUndefCharact = 0 then exit;
            if ShowUndefCharact = -1 then begin
            if MessageDlg('Не определена характеристика ТС. Показать такие ошибки [Yes] или пропустить все [No]?', mtInformation, [mbYes, mbNo], 0) = mrNo then begin
                ShowUndefCharact := 0;
                exit;
            end;
            end;
            ShowUndefCharact := 1;
            raise Exception.Create('Не определена характеристика ТС');
        end;

        if StrtoInt(charact) <= 50 then DecodeLetter := 'A1'
        else if StrtoInt(charact) <= 70 then DecodeLetter := 'A2'
        else if StrtoInt(charact) <= 95 then DecodeLetter := 'A3'
        else if StrtoInt(charact) <= 120 then DecodeLetter := 'A4'
        else if StrtoInt(charact) <= 150 then DecodeLetter := 'A5'
        else if StrtoInt(charact) <= 200 then DecodeLetter := 'A6'
        else if StrtoInt(charact) > 200 then DecodeLetter := 'A7';
    end;
    if tgrp = '2' then begin
        try
            StrtoInt(charact);
        except
            if ShowUndefCharact = 0 then exit;
            raise Exception.Create('Не определена характеристика ТС');
        end;
        if StrtoInt(charact) <= 10 then DecodeLetter := 'C1'
        else if StrtoInt(charact) > 10 then DecodeLetter := 'C2';
    end;
    if tgrp = '3' then begin
        try
            StrtoInt(charact);
        except
            if ShowUndefCharact = 0 then exit;
            raise Exception.Create('Не определена характеристика ТС');
        end;
        if StrtoInt(charact) <= 20 then DecodeLetter := 'E1'
        else if StrtoInt(charact) > 20 then DecodeLetter := 'E2';
    end;
    if tgrp = '1' then begin
        try
            StrtoInt(charact);
        except
            if ShowUndefCharact = 0 then exit;
            raise Exception.Create('Не определена характеристика ТС');
        end;
        if StrtoInt(charact) <= 50 then DecodeLetter := 'T1'
        else if StrtoInt(charact) <= 70 then DecodeLetter := 'T2'
        else if StrtoInt(charact) <= 95 then DecodeLetter := 'T3'
        else if StrtoInt(charact) <= 120 then DecodeLetter := 'T4'
        else if StrtoInt(charact) <= 160 then DecodeLetter := 'T5'
        else if StrtoInt(charact) <= 200 then DecodeLetter := 'T6'
        else if StrtoInt(charact) > 200 then DecodeLetter := 'T7';
    end;
end;

function DecodeAgentName(s : string) : string;
begin
    if lastAgCode = s then begin
        DecodeAgentName := lastAgName;
        exit;
    end;
    lastAgCode := s;
    RUSSMAND.RepQuery.Close;
    RUSSMAND.RepQuery.SQL.Clear;
    RUSSMAND.RepQuery.SQL.Add('SELECT NAME FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    RUSSMAND.RepQuery.Open;
    if RUSSMAND.RepQuery.Eof then lastAgName := s
    else lastAgName := RUSSMAND.RepQuery.FieldByName('NAME').AsString;
    DecodeAgentName := lastAgName;
end;

function DecodeBlank(Company : integer; Ser: string) : string;
var
    INI : TIniFile;
    i : integer;
begin
    DecodeBlank := '?';
    //if Company = 0 then DecodeBlank := '31';
    //if Company = 1 then DecodeBlank := '32';
    //if Company = 2 then DecodeBlank := '55';

    for i := 1 to 100 do
    begin
        if BlankKodes[i].Ser = '' then
        begin
            INI := TIniFile.Create(RUSSMAND.IniName);
            BlankKodes[i].Company := Company;
            BlankKodes[i].Ser := Ser;
            BlankKodes[i].Id := INI.ReadString('RUSSIAN', 'BlankCode' + IntToStr(Company) + '_' + Ser, '?');
            INI.Free;
        end;

        if (BlankKodes[i].Company = Company) AND (BlankKodes[i].Ser = Ser) then
        begin
            if BlankKodes[i].Id <> '?' then DecodeBlank := BlankKodes[i].Id;
            exit;
        end
    end;
end;

function GetInsurType(Company : integer) : string;
var
    INI : TIniFile;
begin
    INI := TIniFile.Create(RUSSMAND.IniName);
    GetInsurType := INI.ReadString('RUSSIAN', 'InsurType' + IntToStr(Company), '?');
    INI.Free;
end;

function DecodePeriod(s : string) : string;
begin
    DecodePeriod := s;
    if s = '12' then DecodePeriod := '1Г'
    else
    if s <> '15' then DecodePeriod := s + 'М'
    else
    if StrToInt(s) > 99 then DecodePeriod := '6М';
end;

function DecodeR(s : string) : string;
begin
    DecodeR := '?';
    if s = 'Y' then DecodeR := '1';
    if s = 'N' then DecodeR := '2';
end;

function DecodeSpecZnak(s1, s2 : string) : string;
begin
    DecodeSpecZnak := '|';
    if s1 <> '' then DecodeSpecZnak := s1 + '|' + s2;
end;

function DecodeCharact(s : string) : string;
begin
    DecodeCharact := s;
    if s = '0' then DecodeCharact := '';
end;

function DecodeIdNumber(val, typ : string) : string;
begin
    if (typ = '6') or (typ = '7') or (typ = '8') then DecodeIdNumber := val
    else DecodeIdNumber := '';
end;

function DecodeIdNumberPric(val, typ : string) : string;
begin
    if (typ = '6') or (typ = '7') or (typ = '8') then DecodeIdNumberPric := ''
    else DecodeIdNumberPric := val;
end;

function DecodeCharactType(s : string) : string;
begin
    if (s = '0') or (s = '1') then DecodeCharactType := '2';
    if (s = '2') then DecodeCharactType := '3';
    if (s = '3') then DecodeCharactType := '4';
end;

function DecodeTarifGrp(Tarif : double; PersType, Resid : string; Charact : double) : integer;
var
    INI : TIniFile;
    mask, prefix, s, ts : string;
    i, j, k : integer;
    alternativeGrp : integer;
label
    Stop;
begin
    prefix := 'Tarif';
    if Resid = 'Y' then prefix := prefix + 'R';
    prefix := prefix + '_' + PersType;
    DecodeTarifGrp := -1;
    alternativeGrp := -1;
    INI := TIniFile.Create(RUSSMAND.IniName);
    for i := 0 to 100 do begin
        s := INI.ReadString('RUSSIAN', 'Vid' + IntToStr(i), '');
        if s = '' then break;
        if Pos(',', s) = 0 then s := s + ',0';
        for j := 2 to WordCount(s, [',']) do begin //charact
            mask := prefix + '_' + IntToStr(i) + '_' + ExtractWord(j, s, [',']);
            //if Charact < StrToInt(ExtractWord(j, s, [','])) then
            //    continue;
            ts := INI.ReadString('RUSSIAN', mask, '');
            if ts <> '' then begin
                for k := 1 to WordCount(s, [',']) do begin
                    if StrToFloat(ExtractWord(k, ts, [','])) = Tarif then begin
                        if Charact < StrToInt(ExtractWord(j, s, [','])) then begin
                            DecodeTarifGrp := i;
                            goto Stop;
                        end
                        else
                            alternativeGrp := i;
                    end;
                end;
            end;
        end;
    end;
    Stop:;

    if Result = -1 then
        DecodeTarifGrp := alternativeGrp;

    INI.Free;
end;

procedure TRUSSMAND.mnuExportClick(Sender: TObject);
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
    UnloadN : string;
    UnloadDt : string;
    Division, sDivision, path : string;
    DD, MM, YY : WORD;
    errStr : string;
    INI : TIniFile;
    sPricepNmb : string;
    Tarif : double;
    InsTypeCode : string;
    UndefinedGrp : integer;
    Hour, Min, Sec, MSec: Word;
    DS : TDataSet;
    PrevKod, PrevSeria, PrevNumber : string;
begin
    UndefinedGrp := 0;
    errStr := '';
    UnloadN := '00';
    InsTypeCode := '32';

    if not IsStopDate.Checked then
        if MessageDlg('Вы не указали выгрузку досрочно расторгнутых. Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
    MessageDlg('Полисы с неопределенными характеристиками будут пропущены.', mtInformation, [mbOk], 0);

    if Responsibility = '' then begin
        MessageDlg('Не указана ответственность по полису', mtInformation, [mbOk], 0);
        exit;
    end;

    INI := TIniFile.Create(BLANK_INI);
    Division := INI.ReadString('DIVISION', 'ID', '');
    SaveDialog.InitialDir := INI.ReadString('DIVISION', 'OutDir', 'C:\');
    INI.Free;
    if Length(Division) < 1 then begin
        MessageDlg('Не указан код вашего подразделения', mtInformation, [mbOk], 0);
        exit;
    end;


    SaveDialog.FileName := UnloadN + '.unload';
    if not SaveDialog.Execute then exit;

    path := Copy(SaveDialog.FileName, 1, Length(SaveDialog.FileName) - Pos('\', ReverseString(SaveDialog.FileName)) + 1);

    INI := TIniFile.Create(BLANK_INI);
    INI.WriteString('DIVISION', 'OutDir', path);
    INI.Free;

    sDivision := Division;
    if Length(sDivision) < 2 then sDivision := '0' + sDivision;

    MessageDlg('Ответственность равна ' + Responsibility, mtInformation, [mbOk], 0);

    DecodeDate(Date, YY, MM, DD);

    UnloadDt := IntToStr(MM);
    if MM < 10 then UnloadDt := '0' + UnloadDt;
    if DD < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(DD);

    DecodeTime(Now, Hour, Min, Sec, MSec);
    //UnloadN := IntToStr(Hour);
    //if Hour < 10 then UnloadN := '0' + UnloadN;
    UnloadN := Format('%02d', [Hour]) + Format('%02d', [Min]) + Format('%02d', [Sec]);

    IsError := false;

    Screen.Cursor := crHourGlass;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    if (GetAsyncKeyState(vk_Shift) AND $8000) = 0 then
        WorkSQL.SQL.Add('SELECT ''' + Division + ''' AS DIVISION,T.* FROM MANDRUSS T LEFT OUTER JOIN AGENT A ON (T.AGENT=A.AGENT_CODE)')
    else
        WorkSQL.SQL.Add('SELECT A.DIVISIONCODE AS DIVISION,T.* FROM MANDRUSS T LEFT OUTER JOIN AGENT A ON (T.AGENT=A.AGENT_CODE)');
    if Length(GetFilter(' WHERE ')) <> 0 then
        WorkSQL.SQL.Add(GetFilter(' WHERE '));
    WorkSQL.Open;
    if MainQuery.Active then
        All := MainQuery.RecordCount
    else
        All := WorkSQL.RecordCount;
    Curr := 1;
    with WorkSQL do begin
        while not Eof do begin
          try

            InsTypeCode := GetInsurType(FieldByName('INSCOMP').AsInteger);
            if InsTypeCode = '?' then
            begin
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' для компании не указан код вида страхования. Ключ ' + 'InsurType' + FieldByName('INSCOMP').AsString);
            end;
            if DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) = '?' then
            begin
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' для компании не указан код бланка в настройках BlankCode' + FieldByName('INSCOMP').AsString + '_' + FieldByName('SERIA').AsString);
            end;


            ////////////////////
            if Curr = 1 then
            begin
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
                AssignFile(A,  path + sDivision + 'A' + UnloadDt + UnloadN + '.0' + InsTypeCode);
                ReWrite(A);
                AssignFile(F10, path + sDivision + 'H' + UnloadDt + UnloadN + '.0' + InsTypeCode);
                ReWrite(F10);
                AssignFile(F11, path + sDivision + 'Z' + UnloadDt + UnloadN + '.0' + InsTypeCode);
                ReWrite(F11);
                AssignFile(F12, path + sDivision + 'G' + UnloadDt + UnloadN + '.0' + InsTypeCode);
                ReWrite(F12);
                AssignFile(F13, path + sDivision + 'B' + UnloadDt + UnloadN + '.0' + InsTypeCode);
                ReWrite(F13);
            end;
            ////////////////////

            Curr := Curr + 1;
            ProgressBar.Position := round(Curr * 100 / All);
            if FieldByName('DIVISION').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' нет кода подразделения ' + FieldByName('AGENT').AsString);
            if FieldByName('STATE').AsInteger in [2] then begin //испорчен
                WriteLn(F13, FieldByName('DIVISION').AsString + '|8|' +
                            DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                            FieldByName('SERIA').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDT').AsDateTime)
                )
            end
            else begin
                  {if FieldByName('STATE').AsInteger in [1] then begin //утерян
                      WriteLn(F13, FieldByName('DIVISION').AsString + '|24|' +
                                  DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
                                  FieldByName('SERIA').AsString + '|' +
                                  FieldByName('NUMBER').AsString + '|' +
                                  FieldByName('NUMBER').AsString + '|' +
                                  DecodeGlDate(FieldByName('REGDT').AsDateTime)
                      )
                  end;}
                //if FieldByName('PSERIA').AsString = '' then begin
                    if Trim(FieldByName('INSURER').AsString) = '' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('INSURER').AsString + ' имя страхователя не указано');
                    if (FieldByName('AGTYPE').AsString = '') OR not (FieldByName('AGTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип агента не определён "' + FieldByName('AGENT').AsString);
                    if (FieldByName('INSURERT').AsString = '') OR not (FieldByName('INSURERT').AsString[1] in ['F', 'U', 'I']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип страхователя не определён "' + FieldByName('INSURERT').AsString + '" ' + FieldByName('AGENT').AsString);
                    if DecodePeriod(FieldByName('PERIOD').AsString) = '?' then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' период не определён "' + FieldByName('PERIOD').AsString + '"');

                    if DecodeLetter(FieldByName('TARIFGRP').AsString, FieldByName('CHARACT').AsString) = '?' then begin
                        //raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' характеристика не определена');
                        Next;
                        Inc(UndefinedGrp);
                        continue;
                    end;
                    //if Length(FieldByName('AUTONMB').AsString) > 10 then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' номер авто больше 10 символов ' + FieldByName('AUTONMB').AsString);

                    DS := WorkSQL;

                    if not auxTable.active then auxTable.open;

                    if FieldByName('PNMB').AsInteger > 0 then begin
                        if not auxTable.Locate('seria;number', VarArrayOf([FieldByName('PSER').AsString, FieldByName('PNMB').AsInteger]), []) then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' не найден утерянный полис ' + FieldByName('PNMB').AsString);
                        DS := auxTable;
                    end;

                    Tarif := DS.FieldByName('TARIF').AsFloat;

                    WriteLn(F1, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('AGPCNT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDT').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString) + '|' +
                                DecodeGlDate(FieldByName('INSSTART').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('TMSTART').AsString) + '|' +
                                DecodeGlDate(FieldByName('INSEND').AsDateTime) + '|' +
                                DecodeRUSSVariant(FieldByName('INSURERT').AsString, FieldByName('RESIDENT').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('INSURER').AsString + '|' +
                                DecodeFU12(FieldByName('INSURERT').AsString, '9') + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                '' + '|' + //PLACE_CODE
                                {FieldByName('ADDR').AsString +} '|' +
                                '' + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodePayType(FieldByName('INSURERT').AsString) + '|' +
                                DecodeAgentName(FieldByName('AGENT').AsString) + '|' +
                                Responsibility + '|' + //Страховая сумма по полису (на весь период)
                                Responsibility + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('RUR') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [Tarif]) + '|' + //Страховая премия
                                Format('%f', [Tarif]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('RUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                DecodeSpecZnak(FieldByName('SPECSR').AsString, FieldByName('SPECNMB').AsString) + '|' + //Серия спецзнака (карточки)
                                '' //Для заметок
                    );

                    PrevKod := '';
                    PrevSeria := '';
                    PrevNumber := '';
                    if FieldByName('PNMB').AsInteger > 0 then begin
                        PrevKod := DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString);
                        PrevSeria := FieldByName('PSER').AsString;
                        PrevNumber := IntToStr(FieldByName('PNMB').AsInteger);
                    end;
                    WriteLn(X, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('AGPCNT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDT').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString) + '|' +
                                DecodeGlDate(FieldByName('INSSTART').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('TMSTART').AsString) + '|' +
                                DecodeGlDate(FieldByName('INSEND').AsDateTime) + '|' +
                                DecodeRUSSVariant(FieldByName('INSURERT').AsString, FieldByName('RESIDENT').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('INSURER').AsString + '|' +
                                DecodeFU12(FieldByName('INSURERT').AsString, '9') + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                '' + '|' + //PLACE_CODE
                                 '|' +
                                '' + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodePayType(FieldByName('INSURERT').AsString) + '|' +
                                DecodeAgentName(FieldByName('AGENT').AsString) + '|' +
                                Responsibility + '|' + //Страховая сумма по полису (на весь период)
                                Responsibility + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('RUR') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [Tarif]) + '|' + //Страховая премия
                                Format('%f', [Tarif]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('RUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                DecodeSpecZnak(FieldByName('SPECSR').AsString, FieldByName('SPECNMB').AsString) + '|' + //Серия спецзнака (карточки)
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



{                    if FieldByName('INSURER').AsString <> '' then
                    WriteLn(F2, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeCurrency('EUR') + '|' + //Код валюты
                                '0' + '|' + //Страховая сумма
                                FieldByName('INSURER').AsString + '|' +
                                '' + '|' + //Дата рождения
                                '' + '|' + //Паспорт
                                ''
                    );}
                    WriteLn(F3, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeLetter(FieldByName('TARIFGRP').AsString, FieldByName('CHARACT').AsString) + '|' + //Тип транспортного средства
                                FieldByName('MARKA').AsString + '|' +
                                FieldByName('AUTONMB').AsString + '|' +
                                '' + '|' + //Год выпуска
                                DecodeIdNumber(FieldByName('AUTOID').AsString,  FieldByName('TARIFGRP').AsString) + '|' +
                                DecodeIdNumberPric(FieldByName('AUTOID').AsString,  FieldByName('TARIFGRP').AsString) + '|' +
                                FieldByName('PASSSER').AsString + '|' + //Серия техпаспорта
                                FieldByName('PASSNMB').AsString + '|' + //Номер техпаспорта
                                DecodeCharactType(FieldByName('TARIFGRP').AsString) + '|' + //Код характеристики
                                DecodeCharact(FieldByName('CHARACT').AsString) + '|' + //Значение характеристики
                                '' + '|' + //Стоимость
                                '' + '|' + //Код валюты (стоимости и страховой суммы)
                                '' + '|' + //Страховая сумма по объекту
                                'S' + '|' +
                                'BY' //Символ страны прибытия
                    );


                    //Не дубликат
                    if FieldByName('PNMB').AsInteger = 0 then
                    begin
                        if Trim(FieldByName('PSER').AsString) <> '' then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' некорректные данные по утерянному полису. Не указан номер');

                        if DS.FieldByName('PAY1').AsFloat > 0 then
                            WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                        DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                        FieldByName('SERIA').AsString + '|' +
                                        FieldByName('NUMBER').AsString + '|' +
                                        '5' + '|' + //Код события
                                        DecodeNBByOwner(DS.FieldByName('INSURERT').AsString) + '|' + //Форма оплаты
                                        '' + '|' +
                                        DecodeGlDate(DS.FieldByName('PAYDATE').AsDateTime) + '|' + //Дата платежного документа
                                        DecodeGlDate(DS.FieldByName('PAYDATE').AsDateTime) + '|' +
                                        DecodeGlDate(DS.FieldByName('REPDT').AsDateTime) + '|' + //Дата  проведено по бухгалтерии

                                        DS.FieldByName('PAY1').AsString + '|' + //Сумма
                                        DecodeCurrency(DS.FieldByName('CURR1').AsString) + '|' + //Код валюты
                                        DS.FieldByName('AGPCNT').AsString + '|' + //Комиссия
                                        DecodeAgentName(DS.FieldByName('AGENT').AsString) + '|' +
                                        DecodeFU16(DS.FieldByName('AGTYPE').AsString, DS.FieldByName('AGPCNT').AsFloat),
                                        DecodeAgentID(FieldByName('AGENT').AsString)
                            );
                    end;

                    if Trim(DS.FieldByName('PSER').AsString) <> '' then
                    if DS.FieldByName('PNMB').AsInteger = 0 then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' некорректные данные по утерянному полису. Не указан номер');

                    if Trim(DS.FieldByName('PSER').AsString) = '' then
                    if DS.FieldByName('PNMB').AsInteger > 0 then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' некорректные данные по полису. Не указана серия утерянного');

                    //Не дубликат
                    if FieldByName('PNMB').AsInteger = 0 then
                    begin
                        if Trim(DS.FieldByName('PSER').AsString) <> '' then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' некорректные данные по утерянному полису. Не указан номер');

                        if DS.FieldByName('PAY2').AsFloat > 0 then
                            WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                        DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                        FieldByName('SERIA').AsString + '|' +
                                        FieldByName('NUMBER').AsString + '|' +
                                        '5' + '|' + //Код события
                                        DecodeNBByOwner(DS.FieldByName('INSURERT').AsString) + '|' + //Форма оплаты
                                        '' + '|' +
                                        DecodeGlDate(DS.FieldByName('PAYDATE').AsDateTime) + '|' + //Дата платежного документа
                                        DecodeGlDate(DS.FieldByName('PAYDATE').AsDateTime) + '|' +
                                        DecodeGlDate(DS.FieldByName('REPDT').AsDateTime) + '|' + //Дата  проведено по бухгалтерии
                                        DS.FieldByName('PAY2').AsString + '|' + //Сумма
                                        DecodeCurrency(DS.FieldByName('CURR2').AsString) + '|' + //Код валюты
                                        DS.FieldByName('AGPCNT').AsString + '|' + //Комиссия
                                        DecodeAgentName(DS.FieldByName('AGENT').AsString) + '|' +
                                        DecodeFU16(DS.FieldByName('AGTYPE').AsString, DS.FieldByName('AGPCNT').AsFloat),
                                        DecodeAgentID(FieldByName('AGENT').AsString)
                            );
                    end;
                    if FieldByName('STATE').AsString = '3' then
                    if FieldByName('RETSUM').AsFloat > 0 then begin
                        WriteLn2(F8, A, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                    DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                    FieldByName('SERIA').AsString + '|' +
                                    FieldByName('NUMBER').AsString + '|' +
                                    '51' + '|' + //Код события
                                    DecodeNBByOwner(FieldByName('INSURERT').AsString) + '|' + //Форма оплаты
                                    '|' +
                                    DecodeGlDate(FieldByName('STOPDT').AsDateTime) + '|' + //Дата платежного документа
                                    DecodeGlDate(FieldByName('STOPDT').AsDateTime) + '|' +
                                    DecodeGlDate(FieldByName('STOPDT').AsDateTime) + '|' +
                                    FieldByName('RETSUM').AsString + '|' + //Сумма
                                    DecodeCurrency(FieldByName('RETCURR').AsString) + '|' + //Код валюты
                                    '|' + //Комиссия
                                    '|' +
                                    DecodeFU16('', 0),
                                    ''
                        );
                        if FieldByName('RETCURR').AsString = '' then
                            raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' валюта возврата не определена');
                    end;
                    if FieldByName('STATE').AsString = '3' then
                    WriteLn(F11, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '|' +
                                DecodeGlDate(FieldByName('StopDt').AsDateTime) + '|' +
                                '2' + '|' +
                                '2' + '|' +
                                '42'
                    );
                //end;
                //else begin //Дубликаты
                if FieldByName('PSER').AsString <> '' then begin
                    //if FieldByName('DUPTYPE').AsString = '0' then
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('PSER').AsString + '|' +
                                FieldByName('PNMB').AsString + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger, FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '9' + '|' + //Дубликат
                                DecodeGlDate(FieldByName('REGDT').AsDateTime)
                    );
                end;
            end; //if
            Next;
          except
            on E : Exception do
            begin
                errStr := errStr + #13#10 + e.Message + ' ' + FieldByName('NUMBER').AsString;
                IsError := true;
                Next;
            end
          end;  
        end; //while
    end; //width

    try
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
    except
    end;
    ProgressBar.Position := 0;

    Screen.Cursor := crDefault;

    if not IsError then
        MessageDlg('Экспорт выполнен успешно'#13'Пропущено ' + IntToStr(UndefinedGrp) + ' полисов из-за неопределенной характеристика ТС', mtInformation, [mbOk], 0)
    else begin
        ShowInfo.Caption := 'Ошибки в данных';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
end;

procedure TRUSSMAND.MenuItem1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Россия');
     if GetTemplName.ShowModal = mrOk then begin
         for i := 0 to 100 do
           if INI.ReadString('AgTemplates', IntToStr(i), '') = '' then break;
         index := i;
         str := '';
         for i := 0 to AgentList.Items.Count - 1 do
             if AgentList.Checked[i] then
                 str := str + AgCodes[i] + ',';
         INI.WriteString('AgTemplates', IntToStr(index), str);
         INI.WriteString('AgTemplates', 'Name' + IntToStr(index), GetTemplName.Name.Text);
         ListTemplates.ItemIndex := ListTemplates.Items.Add(GetTemplName.Name.Text);
     end;
     INI.Free;
end;

procedure TRUSSMAND.AgentListClick(Sender: TObject);
var
     Count, i : integer;
begin
     Count := 0;
     for i := 0 to AgentList.Items.Count - 1 do
         if AgentList.Checked[i] then
             Inc(Count);
     CntChecked.Caption := 'Помечено ' + IntToStr(Count);

     if Sender <> nil then
         ListTemplates.ItemIndex := -1;
end;

procedure TRUSSMAND.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Россия');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

procedure TRUSSMAND.InputSumClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT PAY1,CURR1,PAYDATE,NUMBER,REPDT FROM MANDRUSS WHERE PNMB=0 AND PAYDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT PAY2,CURR2,PAYDATE,NUMBER,REPDT FROM MANDRUSS WHERE PNMB=0 AND PAYDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkSQL);
end;

procedure TRUSSMAND.N9Click(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    //RepParamsAgnt.filtParam := 'RUSS > 0';

    if RepParamsAgnt.ShowModal <> mrOk then exit;
    if RepParamsAgnt.listAgnt.ItemIndex = -1 then
    begin
        MessageDlg('Выделите одного агента', mtInformation, [mbOk], 0);
        exit;
    end;

    DtFrom := min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    DtTo := max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    sqlText := 'SELECT PAY1,CURR1,REPDT,NUMBER,AGPCNT FROM MANDRUSS WHERE AGENT=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PNMB=0 AND PAYDATE IS NOT NULL AND REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT PAY2,CURR2,REPDT,NUMBER,AGPCNT FROM MANDRUSS WHERE AGENT=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PNMB=0 AND PAYDATE IS NOT NULL AND REPDT BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Items[RepParamsAgnt.listAgnt.ItemIndex], WorkSQL);
end;

procedure TRUSSMAND.ImpAlvenaClick(Sender: TObject);
begin
    if not OpenDialogPx.Execute then exit;
    ImportAlvena(OpenDialogPx.FileName);
end;

procedure TRUSSMAND.ImportAlvena(DBFile : string);
var
    TableMain, TablePay2 : TTable;
    Russ : TTable;
    i, n : integer;
    Seria, Number : string;
    NSeria, NNumber : string;
    dv : integer;
    Field : TField;
    IsNewRecord : boolean;
    PolisId, RateCurrency : string;
    TableTarif : integer;
    NAll, NPos : integer;
    errStr, s : string;
    SuccsessCount : integer;
    UpdateCount : integer;
    UpdateFlag : boolean;
    IsResident : boolean;
    IsRussian, IsEnglish, IsDigit, IsBad : boolean;
    fld : integer;
    INI : TINIFile;
    CompanyId : integer;
    Step : integer;
    tgrp : string;
begin
    Step := -1;
    errStr := '';
    TableMain := TTable.Create(self);
    TablePay2 := TTable.Create(self);
    Russ := TTable.Create(self);
    Russ.TableName := 'mandruss.db';
    Russ.DatabaseName := 'RussDatabase';

    INI := TINIFile.Create(BLANK_INI);
    s := INI.ReadString('RUSSIAN', 'INI', '');
    if s <> '' then begin
        INI.Free;
        INI := TINIFile.Create(s);
    end;
    CompanyId := INI.ReadInteger('RUSSIAN', 'ALVENARUSS', -1);
    INI.Free;

    if CompanyId < 0 then begin
        MessageDlg('Не определён код Российской компании для Альвены', mtConfirmation, [mbOk], 0);
        exit;
    end;

    for i := Length(DBFile) downto 1 do
        if DBFile[i] = '\' then break;

    //MessageDlg('Импорт данных не доделан', mtConfirmation, [mbOk], 0);
    //exit;

    //if MessageDlg('7% Альвене?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
    //    exit;
    if PcntForm.ShowModal <> mrOk then exit;

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
        Russ.Open;

        Screen.Cursor := crHourGlass;
        NAll := TableMain.RecordCount;
        NPos := 1;
        SuccsessCount := 0;
        while not TableMain.Eof do begin
            CountLabel.Caption := InttoStr(NPos) + ' из ' + IntToStr(NAll);
            CountLabel.Update;
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




                //if PolisId = 'ВВВ/0473866792' then
                //    messagebeep(0);

                dv := Pos('/', Field.AsString);
                if dv = 0 then
                    raise Exception.Create('Не правильно указана номер/серия ' + Field.AsString);
                Seria := Copy(Field.AsString, 1, dv - 1);
                Number := Copy(Field.AsString, dv + 1, 1024);

                Step := 1;

                IsNewRecord := false;
                UpdateFlag := false;
                if not Russ.Locate('Seria;Number', vararrayof([Seria, StrToInt(Number)]), []) then begin
                    Russ.Append;
                    Russ.ClearFields;
                    IsNewRecord := true;
                    Russ.FieldByName('Seria').AsString := Seria;
                    Russ.FieldByName('Number').AsInteger := StrToInt(Number);
                end
                else begin
                    UpdateFlag := true;
                    Step := 2;
                    //Сравнить агента!!!
                    if Russ.FieldByName('Agent').AsString <> 'АЛВН' then
                        raise Exception.Create('Полис ' + Field.AsString + ' уже зарегистрирован не на АЛЬВЕНУ');
                    //!!if Russ.FieldByName('State').AsInteger = 2 then begin
                    //!!    if TableMain.FieldByName('Состояние').AsInteger <> 2 then 444
                    //!!        raise Exception.Create('Полис в БД ' + Field.AsString + ' досрочно завершён или расторгнут.');
                    //!!end;
                    //!!if (TableMain.FieldByName('Состояние').AsInteger = 3) AND (Mandator.FieldByName('State').AsInteger <> 1) OR
                    //!!    (TableMain.FieldByName('Состояние').AsInteger <> 3) AND (Mandator.FieldByName('State').AsInteger = 1) then
                    //!!    raise Exception.Create('Полис ' + Field.AsString + ' досрочно конфликт состояний.');
                    Russ.Edit;
                end;

//                if Number = '4848300' then
//                 MessageBeep(0);

                Russ.FieldByName('Agent').AsString := 'АЛВН';
                Russ.FieldByName('AgType').Value := 'U';
                //Russ.FieldByName('UPD').AsDateTime := Date;
                Russ.FieldByName('RegDt').AsDateTime := TableMain.FieldByName('Дата выдачи').AsDateTime;
                //Russ.FieldByName('RegTime').AsString := CalcAlvenaTime(TableMain.FieldByName('Время выдачи полиса').AsInteger);

                Step := 3;
                if TableMain.FieldByName('Состояние').AsInteger = 3 then begin
                    Russ.FieldByName('Insurer').AsString := 'ИСПОРЧЕН';
                    Russ.FieldByName('State').AsInteger := 2;
                    Russ.FieldByName('TarifGrp').AsInteger := -1;
                    Russ.FieldByName('InsComp').Value := CompanyId;
                    Russ.Post;
                    TableMain.Next;
                    SuccsessCount := SuccsessCount + 1;
                    if UpdateFlag then Inc(UpdateCount);
                    continue;
                end;

                Russ.FieldByName('Autonmb').Value := TableMain.FieldByName('Номер машины').Value;
                Russ.FieldByName('Marka').Value := TableMain.FieldByName('Марка машины').Value;

                Step := 4;
                dv := Pos('/', TableMain.FieldByName('Номер стикера').AsString);
                if dv = 0 then
                    errStr := errStr + PolisId + ' ' + 'Номер стикера не правильный ' + TableMain.FieldByName('Номер стикера').AsString + #13#10
                else begin
                    Russ.FieldByName('SpecSr').Value := Copy(TableMain.FieldByName('Номер стикера').AsString, 1, dv - 1);
                    if Copy(TableMain.FieldByName('Номер стикера').AsString, dv + 1, 1024) = '' then
                        errStr := errStr + PolisId + ' ' + 'Номер стикера не правильный ' + TableMain.FieldByName('Номер стикера').AsString + #13#10
                    else
                        Russ.FieldByName('SpecNmb').Value := Copy(TableMain.FieldByName('Номер стикера').AsString, dv + 1, 1024);
                end;

                Russ.FieldByName('Insurer').Value := TableMain.FieldByName('Страхователь').Value;
                Russ.FieldByName('Owner').Value := TableMain.FieldByName('Владелец').Value;
                Russ.FieldByName('AutoId').Value := TableMain.FieldByName('Номер шасси').Value;
                Russ.FieldByName('PassSer').Value := TableMain.FieldByName('ТехПаспорт_Серия').Value;
                Russ.FieldByName('PassNmb').Value := TableMain.FieldByName('ТехПаспорт_Номер').Value;
                //Russ.FieldByName('Org').Value := TableMain.FieldByName('Страхователь').Value;
                //Russ.FieldByName('Addr').Value := PrepareStr(
                //                                            TableMain.FieldByName('Адрес_Город').AsString,
                //                                            TableMain.FieldByName('Адрес_Улица').AsString,
                //                                            TableMain.FieldByName('Адрес_Дом').AsString,
                //                                            TableMain.FieldByName('Адрес_Телефон').AsString
                //                                            );
                Step := 5;
                Russ.FieldByName('InsStart').Value := TableMain.FieldByName('Начало действия').Value;
                Russ.FieldByName('TmStart').Value := CalcAlvenaTime(TableMain.FieldByName('Время начала действия').Value);
                Russ.FieldByName('InsEnd').Value := TableMain.FieldByName('Окончание действия').Value;
                Russ.FieldByName('Period').Value := CalcPeriod2(Russ.FieldByName('InsStart').AsDateTime,
                                                                   Russ.FieldByName('InsEnd').AsDateTime);
                //if Russ.FieldByName('Period').Value = '15 суток' then
                //    Russ.FieldByName('Period').Value := '15 дней';
                //Russ.FieldByName('Letter').Value := UpperCase(TableMain.FieldByName('Категория машины').Value);
                //if TableMain.FieldByName('Парный полис').AsString <> '' then Russ.FieldByName('Letter').Value := 'C+F2';
                //if TableMain.FieldByName('Категория машины').Value = 'f' then Russ.FieldByName('Letter').Value := 'F1';
                //if TableMain.FieldByName('Категория машины').Value = 'F' then Russ.FieldByName('Letter').Value := 'F2';

                {case TableMain.FieldByName('Филиал').AsInteger of
                    1 : Russ.FieldByName('Place').Value := 'МИНСК';
                    2 : Russ.FieldByName('Place').Value := 'БРЕСТ';
                    3 : Russ.FieldByName('Place').Value := 'ВИТЕБСК';
                    4 : Russ.FieldByName('Place').Value := 'ГОМЕЛЬ';
                    5 : Russ.FieldByName('Place').Value := 'ГРОДНО';
                    6 : Russ.FieldByName('Place').Value := 'СОЛИГОРСК';
                    7 : Russ.FieldByName('Place').Value := 'ЖОДИНО';
                    8 : Russ.FieldByName('Place').Value := 'ПРЕДСТАВИТЕЛЬСТВО';
                    9 : Russ.FieldByName('Place').Value := 'ПИНСК';
                    10 : Russ.FieldByName('Place').Value := 'МОГИЛЕВ';
                    11 : Russ.FieldByName('Place').Value := 'БЕРЕСТОВИЦА';
                    12 : Russ.FieldByName('Place').Value := 'МОЗЫРЬ';
                    13 : Russ.FieldByName('Place').Value := 'БАРАНОВИЧИ';
                    else raise Exception.Create('Полис в БД ' + Field.AsString + ' не определено место выдачи.');
                end;
                }
                //if TableMain.FieldByName('Валюта с_тарифа').AsString <> 'EUR' then
                //    raise Exception.Create('Полис ' + Field.AsString + ' имеет тариф не EUR');
                //Russ.FieldByName('Tarif').Value := TableMain.FieldByName('Страховой тариф').Value;

                Russ.FieldByName('Tarif').Value := TableMain.FieldByName('Страховой тариф').Value;

                Step := 6;
                Russ.FieldByName('Pay1').Value := TableMain.FieldByName('Страховой платеж').Value;
                Russ.FieldByName('Curr1').Value := TableMain.FieldByName('Валюта с_платежа').Value;
                if TableMain.FieldByName('Валюта с_платежа').AsString = 'BYB' then
                    Russ.FieldByName('Curr1').AsString := 'BRB';
                Russ.FieldByName('Paydate').Value := TableMain.FieldByName('Дата с_платежа').Value;

                if not (TableMain.FieldByName('Форма с_платежа').AsString[1] in ['*', 'Б']) then begin

                    TablePay2.Close;
                    TablePay2.Filter := '[Номер полиса]=''' + Field.AsString + ''' AND [Вид платежа]=1';
                    TablePay2.Filtered := True;
                    TablePay2.Open;

                    Russ.FieldByName('Curr1').Value := '';

                    Step := 7;
                    while not TablePay2.Eof do begin
                          if (TablePay2.RecordCount = 0) or (TablePay2.RecordCount > 2) then
                              raise Exception.Create('Платежи по полису не найдены или их слишком много ' + IntToStr(TablePay2.RecordCount));
                          //if (TablePay2.RecordCount = 1) and (TablePay2.FieldByName('Дата платежа').AsDateTime <> Russ.FieldByName('Paydate').AsDateTime) then
                          //    raise Exception.Create('Не найден первый платёж');

                          if Russ.FieldByName('Curr1').AsString = '' then begin
                              Russ.FieldByName('Pay1').Value := TablePay2.FieldByName('Сумма платежа').Value;
                              Russ.FieldByName('Curr1').Value := TablePay2.FieldByName('Валюта платежа').Value;
                              if TablePay2.FieldByName('Валюта платежа').AsString = 'BYB' then
                                  Russ.FieldByName('Curr1').AsString := 'BRB';
                          end
                          else begin
                              Russ.FieldByName('Pay2').Value := TablePay2.FieldByName('Сумма платежа').Value;
                              Russ.FieldByName('Curr2').Value := TablePay2.FieldByName('Валюта платежа').Value;
                              if TablePay2.FieldByName('Валюта платежа').AsString = 'BYB' then
                                  Russ.FieldByName('Curr2').AsString := 'BRB';
                          end;

                          TablePay2.Next
                    end;
                end;
                Step := 8;
                Russ.FieldByName('AgPcnt').Value := PcntForm.PcntBefore.Value;
                if Russ.FieldByName('Paydate').Value >= PcntForm.Date.Date then Russ.FieldByName('AgPcnt').Value := PcntForm.PcntAfter.Value;

                Russ.FieldByName('InsComp').Value := CompanyId;
                Russ.FieldByName('State').Value := 0;

                //if TableMain.FieldByName('Дубликат полиса').AsString <> '' then begin
                //    Russ.FieldByName('State').AsInteger := 4;
                //end;
                //!!
                {if UpperCase(Russ.FieldByName('Letter').AsString) = 'A' then
                    Russ.FieldByName('Type').Value := 0;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'B' then
                    Russ.FieldByName('Type').Value := 4;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'C' then
                    Russ.FieldByName('Type').Value := 2;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'C+F2' then
                    Russ.FieldByName('Type').Value := 6;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'E' then
                    Russ.FieldByName('Type').Value := 5;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'F1' then
                    Russ.FieldByName('Type').Value := 1;
                if UpperCase(Russ.FieldByName('Letter').AsString) = 'F2' then
                    Russ.FieldByName('Type').Value := 3;
                }
                Russ.FieldByName('RepDt').Value := Russ.FieldByName('Paydate').Value;
                //Russ.FieldByName('Country').Value := 0;
                //if TableMain.FieldByName('Страны действия').AsString = 'A' then
                //    Russ.FieldByName('Country').Value := 1;
                //if (TableMain.FieldByName('Парный полис').Value <> '') AND (UpperCase(TableMain.FieldByName('Категория машины').Value) = 'F') then
                //    Russ.FieldByName('Text').Value := 'Сцепка';

                Russ.FieldByName('AgType').Value := 'U';

                Step := 9;
                Russ.FieldByName('InsurerT').Value := 'U';
                if (TableMain.FieldByName('Характеристика клиента').AsInteger AND $0001) = 0 then
                    Russ.FieldByName('InsurerT').Value := 'F';

                Russ.FieldByName('OwnerT').Value := 'U';
                if (TableMain.FieldByName('Владелец_Характеристика').AsInteger AND $0001) = 0 then
                    Russ.FieldByName('OwnerT').Value := 'F';

                Russ.FieldByName('Resident').Value := 'Y';
                if (TableMain.FieldByName('Характеристика клиента').AsInteger and $0002) = $0002 then
                    Russ.FieldByName('Resident').Value := 'N';

                Russ.FieldByName('Charact').Value := DecodeCharact2(TableMain.FieldByName('Категория машины').AsString, tgrp);
                Russ.FieldByName('TarifGrp').Value := tgrp;//DecodeTarifGrp(WorkSQL.FieldByName('Tarif').Value, WorkSQL.FieldByName('OwnerT').AsString, WorkSQL.FieldByName('Resident').AsString, 0);

                Step := 10;
                if not TableMain.FieldByName('Расторжение действия').IsNull then begin
                    Step := 11;
                    Russ.FieldByName('StopDt').Value := TableMain.FieldByName('Расторжение действия').Value;
                    Russ.FieldByName('State').Value := 3;
                end;

                Step := 12;
                //if Russ.FieldByName('Type').Value = 6 {(TableMain.FieldByName('Парный полис').Value <> '') AND (UpperCase(TableMain.FieldByName('Категория машины').AsString) = 'C') }then begin
                //    dv := Pos('/', TableMain.FieldByName('Парный полис').Value);
                //    if dv = 0 then
                //        raise Exception.Create('Не правильно указана номер/серия парного полиса' + Field.AsString);
                //    Russ.FieldByName('PrcpSer').Value := Copy(TableMain.FieldByName('Парный полис').Value, 1, dv - 1);
                //    Russ.FieldByName('PrcpNmb').Value := Copy(TableMain.FieldByName('Парный полис').Value, dv + 1, 1024);
                //end;

                for fld := 0 to Russ.FieldCount - 3 do begin
                    Step := 10 * fld + 1000;
                    if Russ.Fields[i].DataType = ftDate then begin
                        if not Russ.Fields[i].IsNull then
                            if (Russ.Fields[i].AsDateTime < (Date - 400)) OR (Russ.Fields[i].AsDateTime > (Date + 400)) then
                                raise Exception.Create('Полис ' + Field.AsString + ' дата не в диапазоне ' + Russ.Fields[i].FieldName + '=' + Russ.Fields[i].AsString);
                    end;
                end;

                Russ.Post;
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
                    if not Russ.Locate('Seria;Number', vararrayof([NSeria, StrToInt(NNumber)]), []) then
                        errStr := errStr + PolisId + ' Не найден полис дубликат ' + Field.AsString + #13#10
                    else begin
                        Russ.Edit;
                        Russ.FieldByName('PSeria').AsString := NSeria;
                        Russ.FieldByName('PNumber').AsInteger := StrToInt(NNumber);
                        Russ.FieldByName('Duptype').AsInteger := 0;
                        Russ.Post;
                    end;
                end;
                }
            except
                on  E : Exception do begin
                    if IsNewRecord then
                        Russ.Delete
                    else
                    if Russ.State in [dsEdit] then
                        Russ.Cancel;

                    errStr := errStr + PolisId + ' ' + IntToStr(Step) + ' ' + e.Message + #13#10;
                end;
            end;

            TableMain.Next;
        end;
    except
        on  E : Exception do begin
            MessageDlg('Ошибка импорта. ' + IntToStr(Step) + ' ' + e.Message, mtInformation, [mbOk], 0);
        end
    end;

    Screen.Cursor := crDefault;
    Russ.Free;
    TableMain.Free;
    TablePay2.Free;
//    TableOwn.Free;


    CountLabel.Caption := '';

    MessageDlg(Format('Импортировано %d из %d. Из них обновлено %d', [SuccsessCount, NAll, UpdateCount]), mtInformation, [mbOk], 0);

    if errStr <> '' then begin
        ShowInfo.Caption := 'Результат импорта данных Альвены';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end;
end;

procedure TRUSSMAND.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
    Found, NotFound : integer;
begin
    if (GetAsyncKeyState(VK_SHIFT) AND $8000) = 0 then exit;
    Screen.Cursor := crHourGlass;
    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT * FROM MANDRUSS WHERE TARIFGRP=-1 AND TARIF IS NOT NULL');
    WorkSQL.Open;
    Found := 0;
    NotFound := 0;
    try
    while not WorkSQL.Eof do begin
        //if WorkSQL.RecNo = 1143 then
        //    MessageBeep(0);
        if -1 = DecodeTarifGrp(WorkSQL.FieldByName('Tarif').Value, WorkSQL.FieldByName('OwnerT').AsString, WorkSQL.FieldByName('Resident').AsString, WorkSQL.FieldByName('Charact').AsFloat) then
            NotFound := NotFound + 1
        else
            Found := Found + 1;
        WorkSQL.Next;
    end;
    except
        on E : Exception do
            MessageDlg(e.Message + ', ' + IntToStr(WorkSQL.RecNo), mtInformation, [mbOk], 0);
    end;
    Screen.Cursor := crDefault;

    MessageDlg('В Вашей БД существуют полисы без указания таблицы тарифов. ' + IntToStr(Found) + 'шт можно восстановить из других данных, ' + IntToStr(NotFound) + ' не определено', mtInformation, [mbOk], 0);
end;

procedure TRUSSMAND.IsRepDateClick(Sender: TObject);
begin
       IsState.Enabled := not IsRepDate.Checked;
       State.Enabled := not IsRepDate.Checked;
       if not IsState.Enabled then IsState.Checked := false;
end;

end.
