unit UniversUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, DBGrids, RXDBCtrl, RxQuery, StdCtrls,
  ExtCtrls, ComCtrls, CheckLst, Buttons, Menus, datamod, Math, Placemnt;

type
  TUNIVERS = class(TForm)
    MainQuery: TRxQuery;
    UniversDb: TDatabase;
    DataSource: TDataSource;
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
    WorkSQL: TQuery;
    ProgressBar: TProgressBar;
    IsUseFilter: TCheckBox;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    mnuExport: TMenuItem;
    listOwnType: TComboBox;
    CntChecked: TLabel;
    InputSum: TMenuItem;
    N9: TMenuItem;
    OpenDialogPx: TOpenDialog;
    SaveDialog: TSaveDialog;
    FormStorage: TFormStorage;
    IsStopDate: TCheckBox;
    auxTable: TTable;
    IsOwner: TCheckBox;
    OwnerLike: TEdit;
    IsInsurType: TCheckBox;
    InsurTypes: TComboBox;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryVid: TStringField;
    MainQueryPrevS: TStringField;
    MainQueryPrevN: TFloatField;
    MainQueryFU: TStringField;
    MainQueryRegDate: TDateField;
    MainQueryStartDate: TDateField;
    MainQueryAnother: TStringField;
    MainQueryFrom: TDateField;
    MainQueryTo: TDateField;
    MainQueryPeriod: TStringField;
    MainQueryInsurer: TStringField;
    MainQueryInsAddr: TStringField;
    MainQueryInsOther: TStringField;
    MainQueryInsObj: TMemoField;
    MainQueryInsPcnt: TFloatField;
    MainQueryZastr: TStringField;
    MainQueryVigod: TStringField;
    MainQueryVigodAddr: TStringField;
    MainQueryLizoZast: TStringField;
    MainQueryPlaceIns: TStringField;
    MainQueryInsEvents: TMemoField;
    MainQueryPropertyIns: TMemoField;
    MainQueryVariant: TStringField;
    MainQueryRealSum: TFloatField;
    MainQueryRealCurr: TStringField;
    MainQueryInsSum: TFloatField;
    MainQueryInsCur: TStringField;
    MainQueryFransh: TMemoField;
    MainQueryWait: TStringField;
    MainQueryPremSum: TFloatField;
    MainQueryPremCur: TStringField;
    MainQueryFeer: TStringField;
    MainQueryFeeSum: TFloatField;
    MainQueryFeeCur: TStringField;
    MainQueryFeeDate: TDateField;
    MainQueryFeeDoc: TStringField;
    MainQueryFeeTyp: TStringField;
    MainQuerySrokPay: TStringField;
    MainQueryOtherCond: TMemoField;
    MainQueryAgent: TStringField;
    MainQueryAgPcnt: TFloatField;
    MainQueryAgType: TStringField;
    MainQueryState: TStringField;
    MainQueryRepDate: TDateField;
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure btnSortClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure btnApplyClick(Sender: TObject);
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
  private
     { Private declarations }
     sqlOrder : string;
     AgCodes : TStringList;
     W : HWND;
     InsTypes : array[0..20] of string[48];
     IniName : string;
     MSEXCEL : VARIANT;
     function GetFilter(sCond : string) : string;
     function GetOrder : string;
     procedure ReOpenQuery(Sender: TObject);
     procedure CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List : TASKOneTaxList);
  public
    { Public declarations }
  end;

var
  UNIVERS: TUNIVERS;
  lastAgCode : string;
  lastAgName : string;
  ShowUndefCharact : integer = -1;
  BlankKodes : array[0..30] of String;
  
implementation

uses inifiles, registry, Sort, InpRezParams,
  WaitFormUnit, TmplNameUnit, RepParamsFrm, RepParamsAgntFrm, ShowInfoUnit,
  PcntFormUnit, StrUtils, BelGreenRpt, RxStrUtils, GetRepDateUnit, mf, DM;

{$R *.dfm}

procedure TUNIVERS.btnCloseClick(Sender: TObject);
begin
    Close
end;

procedure TUNIVERS.FormCreate(Sender: TObject);
var
     BLINI : TIniFile;
     rINI : TRegIniFile;
     i, divpos : integer;
     str : string;
     R : TRect;
begin
     State.ItemIndex := 0;
     DMM.HandBookSQL.DatabaseName := UniversDb.Name;

     rINI := TRegIniFile.Create('Универсальный');
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
     IniName := BLINI.ReadString('UNIVERSAL', 'INI', BLANK_INI);
     BLINI.Free;
     BLINI := TIniFile.Create(IniName);
     for i := 0 to 100 do begin
         str := BLINI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
     end;
     ListTemplates.ItemIndex := 0;

     for i := 0 to 100 do begin
         str := BLINI.ReadString('UNIVERSAL', 'Vid' + IntToStr(i), '');
         if str <> '' then begin
             divpos := Pos(',', str);
             InsTypes[i] := Copy(str, 1, divpos);
             InsurTypes.Items.Add(Copy(str, divpos+1, 100));
             InsurTypes.Items.Objects[InsurTypes.Items.Count - 1] := Pointer(i)
         end
     end;

    //Заполним список отчётов
    for i := 1 to 100 do begin
        if BLINI.ReadString('UNIVERSAL', 'QueryName' + IntToStr(i), '') <> '' then
            RepList.Items.Add(BLINI.ReadString('UNIVERSAL', 'QueryName' + IntToStr(i), ''))
        else
            break;
    end;

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

function TUNIVERS.GetFilter(sCond: string): string;
var
     s, str, str2 : string;
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
            Connect(s, 'FEEDATE>=''' + DateToStr(PayDateFrom.Date) + '''');
        if PayDateTo.Checked then
            Connect(s, 'FEEDATE<=''' + DateToStr(PayDateTo.Date) + '''');
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

        if (str <> '') then begin
            if str2 <> '' then
                Connect(str, str2, 'OR');
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
            Connect(s, '(INSURER LIKE ''%' + OwnerLike.Text + '%'' OR INSURER LIKE ''%' + AnsiUpperCase(OwnerLike.Text) + '%'' OR INSURER LIKE ''%' + AnsiLowerCase(OwnerLike.Text) + '%'')');

         if listOwnType.ItemIndex = 1 then
             Connect(s, 'FU=''F''');
         if listOwnType.ItemIndex = 2 then
             Connect(s, 'FU=''U''');
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

     if IsInsurType.Checked AND (InsurTypes.ItemIndex >= 0) then begin
        Connect(s, '(VID=''' + InsTypes[ Integer(InsurTypes.Items.Objects[InsurTypes.ItemIndex]) ] + ''')')
     end;

     if s <> '' then s := sCond + s;
     GetFilter := s;
end;

function TUNIVERS.GetOrder: string;
begin
    if sqlOrder = '' then GetOrder := ''
    else GetOrder := 'ORDER BY ' + sqlOrder;
end;

procedure TUNIVERS.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top + SpeedButton.Height;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TUNIVERS.btnSortClick(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then begin
         sqlOrder := SortForm.GetOrder('');
         ReOpenQuery(nil);
         if TabbedNotebook.ActivePageIndex = 2 then
           Timer.Enabled := true
     end;
end;

procedure TUNIVERS.ReOpenQuery(Sender: TObject);
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

procedure TUNIVERS.FormDestroy(Sender: TObject);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TUNIVERS.TimerTimer(Sender: TObject);
begin
     Timer.Enabled := false;
end;

procedure TUNIVERS.ListTemplatesChange(Sender: TObject);
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
     INI := TRegIniFile.Create('Универсальный');
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

procedure TUNIVERS.btnApplyClick(Sender: TObject);
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

procedure TUNIVERS.btnExcelClick(Sender: TObject);
begin
    DatSetToExcel(MainQuery);
end;

procedure TUNIVERS.MainGridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
    if MainQueryState.AsInteger = 1 then
        Background := $008080FF;
end;

procedure TUNIVERS.RepListChange(Sender: TObject);
var
    INI : TIniFile;
begin
    REPQuery.Close;
    INI := TIniFile.Create(IniName);
    ParamsText.Hint := INI.ReadString('UNIVERSAL', 'QueryParams' + IntToStr(RepList.ItemIndex + 1), '');
    ParamsText.Text := INI.ReadString('UNIVERSAL', 'LastParams' + IntToStr(RepList.ItemIndex + 1), '');
    if ParamsText.Text = '' then
        ParamsText.Text := ParamsText.Hint;
    INI.Free;
end;

procedure TUNIVERS.Button2Click(Sender: TObject);
begin
    if REPQuery.Active then
        DatSetToExcel(REPQuery, IsNumbers.Checked);
end;

procedure TUNIVERS.StartButtonClick(Sender: TObject);
var
    s : string;
begin
    s := '';
    if IsUseFilter.Checked then s := GetFilter('');
    ExecLittleReport(IniName, 'UNIVERSAL', ParamsText.Text, RepList.ItemIndex + 1, REPQuery, s);
end;

procedure TUNIVERS.TabbedNotebookChange(Sender: TObject);
var
    s : string;
begin
    exit;
    
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

procedure TUNIVERS.btnMag11Click(Sender: TObject);
begin
    CreateSimpleMagazine(WorkSQL, 'SELECT SERIA,NUMBER, INSSTART, INSEND,OWNER INSURER, MARKA INSURER2, 1 ALLSUMMA, '''' ALLSUMMACURr, ' +
                                  'pay1 sum1, Curr1 sum1curr, pay2 sum2, Curr2 sum2curr, paydate PAYMENTDATE, AgPcnt AgPercent, AgType AgUr FROM mandruss T WHERE paydate is not null and pay1 >= 0.01 and paydate BETWEEN :D1 AND :D2', 'Российская обязаловка');
    Application.BringToFront;                              
end;

procedure TUNIVERS.btnReserveCalcClick(Sender: TObject);
var
    obj : TaxPercents;
begin
     InitRezerv.RestoreName('Универсальный');
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

procedure TUNIVERS.CreateReport(IsRNPU: boolean; RepDate: TDateTime;
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

procedure TUNIVERS.N3Click(Sender: TObject);
begin
    Close
end;

procedure TUNIVERS.N4Click(Sender: TObject);
begin
    TabbedNotebook.ActivePageIndex := 1
end;

procedure TUNIVERS.N5Click(Sender: TObject);
begin
    TabbedNotebook.ActivePageIndex := 2
end;

procedure TUNIVERS.N6Click(Sender: TObject);
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

{
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

function DecodeBlank(Company : integer) : string;
var
    INI : TIniFile;
begin
    DecodeBlank := '?';
    if Company = 0 then DecodeBlank := '31';
    if Company = 1 then DecodeBlank := '32';
    if Company = 2 then DecodeBlank := '55';
    if (Company >= 2) AND (Company < 30) AND (BlankKodes[Company] = '') then
    begin
        INI := TIniFile.Create(RUSSMAND.IniName);
        BlankKodes[Company] := INI.ReadString('UNIVERSAL', 'BlankCode' + IntToStr(Company), '?');
        INI.Free;
    end;

    if BlankKodes[Company] <> '' then
    begin
        DecodeBlank := BlankKodes[Company];
    end
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
        s := INI.ReadString('UNIVERSAL', 'Vid' + IntToStr(i), '');
        if s = '' then break;
        if Pos(',', s) = 0 then s := s + ',0';
        for j := 2 to WordCount(s, [',']) do begin //charact
            mask := prefix + '_' + IntToStr(i) + '_' + ExtractWord(j, s, [',']);
            //if Charact < StrToInt(ExtractWord(j, s, [','])) then
            //    continue;
            ts := INI.ReadString('UNIVERSAL', mask, '');
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
}
procedure TUNIVERS.mnuExportClick(Sender: TObject);
var
    F1 : TextFile;
    //F2 : TextFile;
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
begin
    UndefinedGrp := 0;
    errStr := '';
    UnloadN := '00';
    InsTypeCode := '32';

    exit;

    if not IsStopDate.Checked then
        if MessageDlg('Вы не указали выгрузку досрочно расторгнутых. Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
    MessageDlg('Полисы с неопределенными характеристиками будут пропущены.', mtInformation, [mbOk], 0);

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

    DecodeDate(Date, YY, MM, DD);

    UnloadDt := IntToStr(MM);
    if MM < 10 then UnloadDt := '0' + UnloadDt;
    if DD < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(DD);

    DecodeTime(Now, Hour, Min, Sec, MSec);
    UnloadN := IntToStr(Hour);
    if Hour < 10 then UnloadN := '0' + UnloadN;

    IsError := false;
    AssignFile(F1, path + sDivision + 'D' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F1);
    //AssignFile(F2, path + Division + 'L' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    //ReWrite(F2);
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
{
    with WorkSQL do begin
        while not Eof do begin
          try

            InsTypeCode := '32';
            if FieldByName('INSCOMP').AsInteger = 1 then InsTypeCode := '33';  //Ресо гарантия
            if FieldByName('INSCOMP').AsInteger = 2 then InsTypeCode := '35';

            Curr := Curr + 1;
            ProgressBar.Position := round(Curr * 100 / All);
            if FieldByName('DIVISION').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' нет кода подразделения ' + FieldByName('AGENT').AsString);
            if FieldByName('STATE').AsInteger in [2] then begin //испорчен
                WriteLn(F13, FieldByName('DIVISION').AsString + '|8|' +
                            DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
                            FieldByName('SERIA').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDT').AsDateTime)
                )
            end
            else begin
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
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                '' //Для заметок
                    );
                    WriteLn(F3, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                        DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                        DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                    DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
                                FieldByName('PSER').AsString + '|' +
                                FieldByName('PNMB').AsString + '|' +
                                DecodeBlank(FieldByName('INSCOMP').AsInteger) + '|' +
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
    }
    CloseFile(F1);
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

procedure TUNIVERS.MenuItem1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Универсальный');
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

procedure TUNIVERS.AgentListClick(Sender: TObject);
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

procedure TUNIVERS.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Универсальный');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

procedure TUNIVERS.InputSumClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT FEESUM,FEECUR,FEEDATE,NUMBER,REPDATE FROM UNIVERS WHERE (PREVN=0 OR PREVN IS NULL) AND FEEDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkSQL);
end;

procedure TUNIVERS.N9Click(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParamsAgnt.ShowModal <> mrOk then exit;
    if RepParamsAgnt.listAgnt.ItemIndex = -1 then
    begin
        MessageDlg('Выделите одного агента', mtInformation, [mbOk], 0);
        exit;
    end;

    DtFrom := min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    DtTo := max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    sqlText := 'SELECT FEESUM,FEECUR,REPDATE,NUMBER,AGPCNT FROM UNIVERS WHERE AGENT=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND (PREVN=0 OR PREVN IS NULL) AND FEEDATE IS NOT NULL AND REPDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Items[RepParamsAgnt.listAgnt.ItemIndex], WorkSQL);
end;

end.
