{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}

{$MINSTACKSIZE $00004000}

{$MAXSTACKSIZE $00100000}

{$IMAGEBASE $00400000}

{$APPTYPE GUI}

unit BulStradForm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Menus, Grids, DBGrids, RXDBCtrl, Db, RxQuery, ComCtrls,
  Tabnotbk, StdCtrls, CheckLst, Buttons, registry, FR_DSet, FR_DBSet,
  FR_Class, ExtCtrls, FR_View, FR_E_TXT, FR_E_CSV, inifiles;

type
  TBULSTRAD = class(TForm)
    BULSTRADDB: TDatabase;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N3: TMenuItem;
    TabbedNotebook: TTabbedNotebook;
    MainQuery: TRxQuery;
    MainDataSource: TDataSource;
    MainGrid: TRxDBGrid;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryPSeria: TStringField;
    MainQueryPNumber: TFloatField;
    MainQueryDtFrom: TDateField;
    MainQueryDtTO: TDateField;
    MainQueryPeriod: TStringField;
    MainQueryKatg: TFloatField;
    MainQueryAutoNmb: TStringField;
    MainQueryShassiNmb: TStringField;
    MainQueryYearMade: TFloatField;
    MainQueryTerr: TFloatField;
    MainQueryOwner: TStringField;
    MainQueryAddress: TStringField;
    MainQueryTarif: TFloatField;
    MainQueryTarifC: TStringField;
    MainQueryPay1: TFloatField;
    MainQueryPay1C: TStringField;
    MainQueryPay2: TFloatField;
    MainQueryPay2C: TStringField;
    MainQueryPayDate: TDateField;
    MainQueryRegDate: TDateField;
    MainQueryRegTime: TStringField;
    MainQueryAgCode: TStringField;
    MainQueryPcnt: TFloatField;
    MainQueryState: TFloatField;
    MainQueryPrevSN: TStringField;
    MainQueryInsPeriod: TStringField;
    MainQueryTC: TStringField;
    MainQueryPlace: TStringField;
    MainQueryOwnAddr: TStringField;
    MainQueryTarifVis: TStringField;
    MainQueryPayment: TStringField;
    MainQueryStateStr: TStringField;
    MainQueryNAME: TStringField;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    CntChecked: TLabel;
    SpeedButton: TSpeedButton;
    IsNumber: TCheckBox;
    IsRegDate: TCheckBox;
    NumberFrom: TEdit;
    NumberTo: TEdit;
    RegDateFrom: TDateTimePicker;
    RegDateTo: TDateTimePicker;
    IsStartDate: TCheckBox;
    StartDateFrom: TDateTimePicker;
    StartDateTo: TDateTimePicker;
    IsOwner: TCheckBox;
    OwnerLike: TEdit;
    IsAutoNumber: TCheckBox;
    AutoNumberLike: TEdit;
    IsPayDate: TCheckBox;
    PayDateFrom: TDateTimePicker;
    PayDateTo: TDateTimePicker;
    AgentList: TCheckListBox;
    ListTemplates: TComboBox;
    PopupMenu: TPopupMenu;
    MenuItem1: TMenuItem;
    N2: TMenuItem;
    AgentsTbl: TQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    Button1: TButton;
    Memo: TMemo;
    frReportFull: TfrReport;
    frDBDataSetFull: TfrDBDataSet;
    frPreviewFull: TfrPreview;
    frCSVExport: TfrCSVExport;
    DownPanel: TPanel;
    Timer: TTimer;
    Label8: TLabel;
    MainQueryAutoTypeLetter: TStringField;
    Panel2: TPanel;
    frPreviewShort: TfrPreview;
    frReportShort: TfrReport;
    frDBDataSetShort: TfrDBDataSet;
    RxQueryShort: TRxQuery;
    RxQueryShortPERIOD: TStringField;
    RxQueryShortKATG: TFloatField;
    RxQueryShortTERR: TFloatField;
    RxQueryShortS: TFloatField;
    RxQueryShortCNT: TIntegerField;
    RxQueryShortCountryName: TStringField;
    RxQueryShortAutoLetter: TStringField;
    N4: TMenuItem;
    WorkSQL: TRxQuery;
    RxQueryShortAL: TFloatField;
    MainQueryRealTarif: TFloatField;
    RxQueryShortTARIFC: TStringField;
    MainQueryDS: TStringField;
    MainQueryDN: TFloatField;
    MainQueryDTRF: TFloatField;
    MainQueryPolis2: TStringField;
    PopupMenuExcel: TPopupMenu;
    mnuToExcel: TMenuItem;
    mnuToExcelNew: TMenuItem;
    MainQueryPTRF: TFloatField;
    MainQueryMarka: TStringField;
    MainQueryModel: TStringField;
    MainQueryCountry: TStringField;
    Panel1: TPanel;
    sbPrint: TSpeedButton;
    sbExcel: TSpeedButton;
    Panel3: TPanel;
    btnPrintShort: TButton;
    Button3: TButton;
    IsState: TCheckBox;
    State: TComboBox;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ReOpenQuery(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure TabbedNotebookChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
    procedure ToExcelClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure PrintClick(Sender: TObject);
    procedure btnPrintShortClick(Sender: TObject);
    procedure RxQueryShortCalcFields(DataSet: TDataSet);
    procedure AgentListClick(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure sbExcelClick(Sender: TObject);
    procedure mnuToExcelNewClick(Sender: TObject);
  private
    { Private declarations }
     AgCodes : TStringList;
     BadCountStr, LostCountStr, ChCountStr : string;
     W : HWND;
     sqlOrder : string;
     Countries, AutoTypes : array[0..100] of string[48];
     _badTarif : double;
     function GetFilter(sCond : string) : string;
     function GetOrder : string;
     function GetSpecialFilter : string;
  public
    { Public declarations }
  end;

var
  BULSTRAD: TBULSTRAD;

implementation

uses Sort, datamod, MF, WaitFormUnit, Variants;

{$R *.DFM}

procedure TBULSTRAD.N3Click(Sender: TObject);
begin
     Close;
end;

procedure Connect(var s: string; cond : string);
begin
     if s <> '' then
         s := s + ' AND ';
     s := s + cond;
end;

function TBULSTRAD.GetOrder : string;
begin
    if sqlOrder = '' then GetOrder := ''
    else GetOrder := 'ORDER BY ' + sqlOrder;
end;

function TBULSTRAD.GetSpecialFilter : string;
var
     s : string;
begin
     Result := GetFilter(' AND ');
end;

function TBULSTRAD.GetFilter(sCond : string) : string;
var
     s, str : string;
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
            Connect(s, 'REGDATE>=''' + DateToStr(RegDateFrom.Date) + '''');
        if RegDateTo.Checked then
            Connect(s, 'REGDATE<=''' + DateToStr(RegDateTo.Date) + '''');
     end;

     if IsStartDate.Checked then begin
        if StartDateFrom.Checked then
            Connect(s, 'DTFROM>=''' + DateToStr(StartDateFrom.Date) + '''');
        if StartDateTo.Checked then
            Connect(s, 'DTTO<=''' + DateToStr(StartDateTo.Date) + '''');
     end;

     if IsPayDate.Checked then begin
        if PayDateFrom.Checked then
            Connect(s, 'PAYDATE>=''' + DateToStr(PayDateFrom.Date) + '''');
        if PayDateTo.Checked then
            Connect(s, 'PAYDATE<=''' + DateToStr(PayDateTo.Date) + '''');
     end;

     //Агенты
     str := '';
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if Length(str) > 0 then
          Connect(s, '(AgCode IN (' + Copy(str, 1, Length(str) - 1) + '))');

     if IsOwner.Checked AND (OwnerLike.Text <> '') then begin
         Connect(s, '(NAME LIKE ''%' + OwnerLike.Text + '%'' OR NAME LIKE ''%' + AnsiUpperCase(OwnerLike.Text) + '%'' OR NAME LIKE ''%' + AnsiLowerCase(OwnerLike.Text) + '%'')');
     end;

     if IsAutoNumber.Checked AND (AutoNumberLike.Text <> '') then begin
         Connect(s, '(AutoNmb LIKE ''%' + AutoNumberLike.Text + '%'' OR AutoNmb LIKE ''%' + AnsiUpperCase(AutoNumberLike.Text) + '%'' OR AutoNmb LIKE ''%' + AnsiLowerCase(AutoNumberLike.Text) + '%'')');
     end;

     if IsState.Checked then begin
         Connect(s, '(State = ' + IntToStr(State.ItemIndex) + ')');
     end;

     if s <> '' then s := sCond + s;
     GetFilter := s;
end;

procedure TBULSTRAD.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     BLINI : TIniFile;
     i : integer;
     str : string;
     R : TRect;
begin
     _badTarif := 1;

     State.ItemIndex := 0;

     TabbedNotebook.PageIndex := 0;
     sqlOrder := 'SERIA,NUMBER';
     
     INI := TRegIniFile.Create('BULSTRAD');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
     end;
     INI.Free;
     ListTemplates.ItemIndex := 0;

     BLINI := TIniFile.Create('blank.ini');
     for i := 0 to 100 do begin
         str := BLINI.ReadString('BULSTRAD', 'Country' + IntToStr(i), '');
         if str <> '' then
             Countries[i] := str;
     end;
     for i := 0 to 100 do begin
         str := BLINI.ReadString('BULSTRAD', 'AutoType' + IntToStr(i), '');
         if str <> '' then
             AutoTypes[i] := str;
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
     Memo.Lines.Add('Количество записей ' + IntToStr(MainQuery.RecordCount));

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

procedure TBULSTRAD.ReOpenQuery(Sender: TObject);
begin
     MainQuery.Close;
     MainQuery.MacroByName('WHERE').AsString := GetFilter(' WHERE ');
     MainQuery.MacroByName('ORDER').AsString := GetOrder();
     Memo.Lines.Clear;
     try
        Screen.Cursor := crHourGlass;
        MainQuery.Open;
        Memo.Lines.Add('Количество записей ' + IntToStr(MainQuery.RecordCount));
     except
        on E : Exception do
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

procedure TBULSTRAD.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TBULSTRAD.MainQueryCalcFields(DataSet: TDataSet);
begin
     if MainQueryDS.AsString <> '' then
         MainQueryPolis2.AsString := MainQueryDS.AsString + '/' + MainQueryDN.AsString;
     if MainQueryPSeria.AsString <> '' then
          MainQueryPrevSN.AsString := MainQueryPSeria.AsString + '/' + MainQueryPNumber.AsString;
     if MainQueryDtFrom.AsString <> '' then
        MainQueryInsPeriod.AsString := MainQueryPeriod.AsString + '; ' + MainQueryDtFrom.AsString + ' - ' + MainQueryDtTo.AsString;
     if MainQueryAddress.AsString <> '' then
         MainQueryOwnAddr.AsString := MainQueryOwner.AsString + ', ' + MainQueryAddress.AsString;
     MainQueryTarifVis.AsString := MainQueryTarif.AsString + ' ' + MainQueryTarifC.AsString;
     MainQueryPayment.AsString := MainQueryPay1.AsString + ' ' + MainQueryPay1C.AsString;
     if MainQueryPay2C.AsString <> '' then
         MainQueryPayment.AsString := MainQueryPayment.AsString + '; ' + MainQueryPay2.AsString + ' ' + MainQueryPay2C.AsString;
     MainQueryStateStr.AsString := 'НОРМАЛЬНЫЙ';
     if MainQueryState.AsInteger = 1 then MainQueryStateStr.AsString := 'УТЕРЯН';
     if MainQueryState.AsInteger = 2 then MainQueryStateStr.AsString := 'ИСПОРЧЕН';

     if MainQueryState.AsInteger <> 2 then begin
         MainQueryPlace.AsString := Countries[MainQueryTerr.AsInteger];
         MainQueryTC.AsString := AutoTypes[MainQueryKatg.AsInteger];
         if Pos(',', MainQueryTC.AsString) = 0 then
             MainQueryAutoTypeLetter.AsString := MainQueryTC.AsString
         else
             MainQueryAutoTypeLetter.AsString := Copy(MainQueryTC.AsString, 1, Pos(',', MainQueryTC.AsString) - 1);
     end;

     MainQueryRealTarif.AsFloat := MainQueryTarif.AsFloat;
     if MainQueryPSeria.AsString <> '' then
         MainQueryRealTarif.AsFloat := _badTarif;
end;

procedure TBULSTRAD.TabbedNotebookChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
begin
     if NewTab = 0 then MainGrid.DataSource := MainDataSource
     else MainGrid.DataSource := nil;

     if NewTab >= 2 then
         Timer.Enabled := true
     else begin
         if TabbedNotebook.PageIndex = 2 then begin
             frReportFull.EMFPages.Clear;
             frPreviewFull.Clear;
         end;
         if TabbedNotebook.PageIndex = 3 then begin
             frReportShort.EMFPages.Clear;
             frPreviewShort.Clear;
             RxQueryShort.Close;
         end;
     end;
end;

procedure TBULSTRAD.ToExcelClick(Sender: TObject);
var
     MSEXCEL, ASH : VARIANT;
     i, fi : integer;
     strI, s : string;
     AllRecords, CurrRecord : integer;
begin
     try
         if not InitExcel(MSEXCEL) then exit;
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;
         WaitForm.ProgressBar.Position := 0;
         MSEXCEL.WorkBooks.Add;
         ASH := MSEXCEL.ActiveSheet;

         MainQuery.First;
         AllRecords := MainQuery.RecordCount;
         CurrRecord := 0;
         i := 5;
         while not MainQuery.EOF do begin
           WaitForm.ProgressBar.Position := ROUND(CurrRecord * 100 / AllRecords);
           CurrRecord := CurrRecord + 1;
           strI := IntToStr(i);
           //for fi := 0 to MainQuery.Fields.Count - 1 do begin
               ASH.Range['A' + strI, 'A'+ strI].Value := MainQuery.FieldByName('SERIA').AsString;
               ASH.Range['B' + strI, 'B'+ strI].Value := MainQuery.FieldByName('NUMBER').AsString;
               ASH.Range['C' + strI, 'C'+ strI].Value := MainQuery.FieldByName('DtFrom').AsString;
               ASH.Range['D' + strI, 'D'+ strI].Value := MainQuery.FieldByName('DtTO').AsString;
               ASH.Range['E' + strI, 'E'+ strI].Value := MainQuery.FieldByName('Period').AsString;
               ASH.Range['F' + strI, 'F'+ strI].Value := MainQuery.FieldByName('AutoTypeLetter').AsString;
               ASH.Range['G' + strI, 'G'+ strI].Value := MainQuery.FieldByName('AutoNmb').AsString;

               ASH.Range['H' + strI, 'H'+ strI].Value := MainQuery.FieldByName('Place').AsString;
               ASH.Range['I' + strI, 'I'+ strI].Value := MainQuery.FieldByName('Tarif').AsString;
               s := '';
               if MainQuery.FieldByName('Pay1C').AsString = 'USD' then s := MainQuery.FieldByName('Pay1').AsString;
               if MainQuery.FieldByName('Pay2C').AsString = 'USD' then s := MainQuery.FieldByName('Pay2').AsString;
               ASH.Range['J' + strI, 'J'+ strI].Value := s;
               s := '';
               if MainQuery.FieldByName('Pay1C').AsString = 'DM' then s := MainQuery.FieldByName('Pay1').AsString;
               if MainQuery.FieldByName('Pay2C').AsString = 'DM' then s := MainQuery.FieldByName('Pay2').AsString;
               ASH.Range['K' + strI, 'K'+ strI].Value := s;
               s := '';
               if MainQuery.FieldByName('Pay1C').AsString = 'EUR' then s := MainQuery.FieldByName('Pay1').AsString;
               if MainQuery.FieldByName('Pay2C').AsString = 'EUR' then s := MainQuery.FieldByName('Pay2').AsString;
               ASH.Range['L' + strI, 'L'+ strI].Value := s;
               s := '';
               if MainQuery.FieldByName('Pay1C').AsString = 'RUR' then s := MainQuery.FieldByName('Pay1').AsString;
               if MainQuery.FieldByName('Pay2C').AsString = 'RUR' then s := MainQuery.FieldByName('Pay2').AsString;
               ASH.Range['M' + strI, 'M'+ strI].Value := s;
               s := '';
               if MainQuery.FieldByName('Pay1C').AsString = 'BRB' then s := MainQuery.FieldByName('Pay1').AsString;
               if MainQuery.FieldByName('Pay2C').AsString = 'BRB' then s := MainQuery.FieldByName('Pay2').AsString;
               ASH.Range['N' + strI, 'N'+ strI].Value := s;
               ASH.Range['O' + strI, 'O'+ strI].Value := MainQuery.FieldByName('Name').AsString;
           //end;
           Inc(i);
           MainQuery.Next;
         end;
     except
       on E : Exception do begin
         if WaitForm <> nil then WaitForm.Close;
         MessageDlg(E.Message, mtInformation, [mbOk], 0);
       end
     end;

     if WaitForm <> nil then WaitForm.Close;
     MSEXCEL.Visible := true;
end;

procedure TBULSTRAD.TimerTimer(Sender: TObject);
var
     PeriodName : string;
     //IsShowBad : boolean;
     s : string;
begin
     if IsPayDate.Checked then begin
        if PayDateFrom.Checked then PeriodName := 'C ' + DateToStr(PayDateFrom.Date);
        if PayDateTo.Checked then begin
            PeriodName := PeriodName + '  ПО ' + DateToStr(PayDateTo.Date);
        end;
     end;

     Timer.Enabled := false;
     if TabbedNotebook.PageIndex = 2 then begin
         MainQuery.Close;
         MainQuery.MacroByName('ORDER').AsString := 'ORDER BY ORDERFLD DESC, SERIA, NUMBER';
         MainQuery.Open;
         MainQuery.First;
         frReportFull.Dictionary.Variables['PERIODNAME'] := '''' + PeriodName + '''';
         frReportFull.ShowReport;
     end;
     if TabbedNotebook.PageIndex = 3 then begin
         //IsShowBad := IsBad.Checked;
         //IsBad.Checked := false;
         s := GetFilter(' WHERE ');
         if s <> '' then
             s := s + ' AND STATE IN (0, 1)'
         else
             s := 'WHERE STATE IN (0, 1)';

         s := s + ' AND (PSERIA = '''' OR PSERIA IS NULL)'; //Не дубликаты

         RxQueryShort.MacroByName('WHERE').AsString := s;

         with WorkSQL do begin
              SQL.Clear;
              SQL.Add('SELECT COUNT(*) AS CNT FROM BULSTRAD WHERE STATE=2');//BAD
              SQL.Add(GetSpecialFilter);
              WorkSQL.Open;
              frReportShort.Dictionary.Variables['BADCOUNT'] := '''' + WorkSQL.FieldByName('CNT').AsString + '''';
              frReportShort.Dictionary.Variables['BADSUM'] := '''' + WorkSQL.FieldByName('CNT').AsString + 'исп.''';
              BadCountStr := WorkSQL.FieldByName('CNT').AsString;
              WorkSQL.Close;

              SQL.Clear;
              SQL.Add('SELECT COUNT(*) AS CNT FROM BULSTRAD %WHERE');//LOST STATE=1
              s := GetFilter(' WHERE ');
              if s <> '' then s := s + ' AND '
              else s := s + 'WHERE ';
              s := s + 'STATE = 1';
              MacroByName('WHERE').AsString := s;
              WorkSQL.Open;
              frReportShort.Dictionary.Variables['LOSTCOUNT'] := '''' + WorkSQL.FieldByName('CNT').AsString + '''';
              frReportShort.Dictionary.Variables['LOSTSUM'] := '''' + WorkSQL.FieldByName('CNT').AsString + 'утер.''';
              LostCountStr := WorkSQL.FieldByName('CNT').AsString;
              WorkSQL.Close;

              SQL.Clear;
              SQL.Add('SELECT COUNT(*) AS CNT FROM BULSTRAD %WHERE');//ЗАМЕНЫ !!!!! двойные
              s := GetFilter(' WHERE ');
              if s <> '' then s := s + ' AND '
              else s := s + 'WHERE ';
              s := s + '(PSERIA IS NOT NULL OR PSERIA <> '''')';
              MacroByName('WHERE').AsString := s;
              WorkSQL.Open;
              frReportShort.Dictionary.Variables['CHCOUNT'] := '''' + WorkSQL.FieldByName('CNT').AsString + '''';
              frReportShort.Dictionary.Variables['CHSUM'] := '''' + WorkSQL.FieldByName('CNT').AsString + 'дубл.''';
              ChCountStr := WorkSQL.FieldByName('CNT').AsString;
              WorkSQL.Close;

              SQL.Clear;
              SQL.Add('SELECT COUNT(*) AS CNT FROM BULSTRAD %WHERE');//Доп.полисы
              s := GetFilter(' WHERE ');
              if s <> '' then s := s + ' AND '
              else s := s + 'WHERE ';
              s := s + '(DS IS NOT NULL OR DS <> '''')';
              MacroByName('WHERE').AsString := s;
              WorkSQL.Open;
              frReportShort.Dictionary.Variables['DCOUNT'] := '''' + WorkSQL.FieldByName('CNT').AsString + '''';

              SQL.Clear;
              SQL.Add('SELECT SUM(DTRF) AS SM FROM BULSTRAD %WHERE');//Доп.полисы
              s := GetFilter(' WHERE ');
              if s <> '' then s := s + ' AND '
              else s := s + 'WHERE ';
              s := s + '(DS IS NOT NULL OR DS <> '''')';
              MacroByName('WHERE').AsString := s;
              WorkSQL.Open;
              frReportShort.Dictionary.Variables['DSUM'] := '''' + WorkSQL.FieldByName('SM').AsString + '''';
              WorkSQL.Close;
         end;

         //IsBad.Checked := IsShowBad;
         frReportShort.Dictionary.Variables['PERIODNAME'] := '''' + PeriodName + '''';
         frReportShort.ShowReport;
     end;
end;

procedure TBULSTRAD.PrintClick(Sender: TObject);
begin
     frPreviewFull.Print
end;

procedure TBULSTRAD.btnPrintShortClick(Sender: TObject);
begin
     frPreviewShort.Print
end;

procedure TBULSTRAD.RxQueryShortCalcFields(DataSet: TDataSet);
var
    s : string;
begin
    RxQueryShortCountryName.AsString := Countries[RxQueryShortTERR.AsInteger];
    s := AutoTypes[RxQueryShortKATG.AsInteger];
    if Pos(',', s) = 0 then
        RxQueryShortAutoLetter.AsString := MainQueryTC.AsString
    else
        RxQueryShortAutoLetter.AsString := Copy(s, 1, Pos(',', s) - 1);

    RxQueryShortAL.Clear;
    if not RxQueryShortS.IsNull then
        RxQueryShortAL.AsFloat := RxQueryShortS.AsFloat * RxQueryShortCnt.AsFloat;
end;

procedure TBULSTRAD.AgentListClick(Sender: TObject);
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

procedure TBULSTRAD.N4Click(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then begin
         sqlOrder := SortForm.GetOrder('');
         ReOpenQuery(nil);
         if TabbedNotebook.PageIndex = 2 then
           Timer.Enabled := true
     end;
end;

procedure TBULSTRAD.Button3Click(Sender: TObject);
var
     MSEXCEL, ASH : VARIANT;
     i, fi : integer;
     strI, s : string;
begin
     try
         if not InitExcel(MSEXCEL) then exit;
         MSEXCEL.Visible := true;
         MSEXCEL.WorkBooks.Add;
         ASH := MSEXCEL.ActiveSheet;

         RxQueryShort.First;
         i := 5;
         while not RxQueryShort.EOF do begin
           strI := IntToStr(i);
           ASH.Range['A' + strI, 'A'+ strI].Value := RxQueryShort.FieldByName('Period').AsString;
           ASH.Range['B' + strI, 'B'+ strI].Value := RxQueryShort.FieldByName('AutoLetter').AsString;
           ASH.Range['C' + strI, 'C'+ strI].Value := RxQueryShort.FieldByName('CountryName').AsString;
           ASH.Range['D' + strI, 'D'+ strI].Value := RxQueryShort.FieldByName('S').AsString;
           ASH.Range['E' + strI, 'E'+ strI].Value := RxQueryShort.FieldByName('CNT').AsString;
           ASH.Range['F' + strI, 'F'+ strI].Value := RxQueryShort.FieldByName('AL').AsString;
           Inc(i);
           RxQueryShort.Next;
         end;
         strI := IntToStr(i);
         ASH.Range['A' + strI, 'A'+ strI].Value := 'ЗАМЕНА';
         ASH.Range['E' + strI, 'E'+ strI].Value := ChCountStr;
         Inc(i);
         strI := IntToStr(i);
         ASH.Range['A' + strI, 'A'+ strI].Value := 'ИСПОРЧЕНО';
         ASH.Range['E' + strI, 'E'+ strI].Value := BadCountStr;
         Inc(i);
         strI := IntToStr(i);
         ASH.Range['A' + strI, 'A'+ strI].Value := 'УТЕРЯНО';
         ASH.Range['E' + strI, 'E'+ strI].Value := LostCountStr;
     except
       on E : Exception do begin
         MessageDlg(E.Message, mtInformation, [mbOk], 0);
       end;
     end

end;

procedure TBULSTRAD.FormDestroy(Sender: TObject);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TBULSTRAD.ListTemplatesChange(Sender: TObject);
var
    i : integer;
begin
     //if ListTemplates.ItemIndex = 0 then begin
       for i := 0 to AgentList.Items.Count - 1 do
            AgentList.Checked[i] := false;
       //exit;
     //end;


end;

procedure TBULSTRAD.sbExcelClick(Sender: TObject);
var
    Pnt : TPoint;
begin
    GetCursorPos(Pnt);
    PopupMenuExcel.Popup(Pnt.x, Pnt.y)
end;
{
Country0=Все страны                       
Country1=Не используется (Ранее: Страны, входящие в EEA)
Country2=Страны, не входящие в EEA
Country3=Польша и Украина
Country4=Эстония и Латвия
Эдик, ТАСК (2:05 PM) : 
0 это 1.1.
2 это 1.2
4 это  1.3
3 это 1.4
 и новая есть тарифная сетка Финляндия, Швеция и Норвегия  это 1.5
}
procedure TBULSTRAD.mnuToExcelNewClick(Sender: TObject);
var
     MSEXCEL, ASH_POLIS, ASH_POLIS_ADD, ArrayData : VARIANT;
     i, j, fi : integer;
     strI, s : string;
     AllRecords, CurrRecord : integer;
     Contract : string;
     INI : TIniFile;
begin
     INI := TIniFile.Create('blank.ini');
     Contract := INI.ReadString('BULSTRAD', 'Contract', '');

     try
         if not InitExcel(MSEXCEL) then exit;
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;
         WaitForm.ProgressBar.Position := 0;
         MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\BULSTRAD.XLS');
         ASH_POLIS := MSEXCEL.ActiveWorkbook.WorkSheets[2];
         ASH_POLIS_ADD := MSEXCEL.ActiveWorkbook.WorkSheets[1];

         MainQuery.First;
         AllRecords := MainQuery.RecordCount;
         CurrRecord := 0;
         i := 2;
         j := 2;
         ArrayData := VarArrayCreate([1, 1, 1, 14], varVariant);
         while not MainQuery.EOF do begin
           WaitForm.ProgressBar.Position := ROUND(CurrRecord * 100 / AllRecords);
           CurrRecord := CurrRecord + 1;
           strI := IntToStr(i);
           ArrayData[1, 1] := MainQuery.FieldByName('SERIA').AsString + '/' + MainQuery.FieldByName('NUMBER').AsString;
           ArrayData[1, 2] := Contract;
           ArrayData[1, 3] := MainQuery.FieldByName('DtFrom').AsString;
           ArrayData[1, 4] := MainQuery.FieldByName('DtTo').AsString;
           ArrayData[1, 5] := MainQuery.FieldByName('AutoNmb').AsString;
           ArrayData[1, 6] := MainQuery.FieldByName('ShassiNmb').AsString;
           s := INI.ReadString('BULSTRAD', 'AutoType' + MainQuery.FieldByName('Katg').AsString, '');
           s := Copy(s, 1, Pos(',', s) - 1);
           ArrayData[1, 7] := s;
           ArrayData[1, 8] := MainQuery.FieldByName('YearMade').AsString;
           ArrayData[1, 9] := MainQuery.FieldByName('Marka').AsString;
           ArrayData[1, 10] := MainQuery.FieldByName('Model').AsString;
           if MainQuery.FieldByName('Country').AsString = '' then
               ArrayData[1, 11] := 'BEL'
           else
               ArrayData[1, 11] := MainQuery.FieldByName('Country').AsString;
           ArrayData[1, 12] := MainQuery.FieldByName('RegDate').AsString;
           s := INI.ReadString('BULSTRAD', 'TarifN' + MainQuery.FieldByName('Terr').AsString, '-');
           ArrayData[1, 13] := ' ' + s;
           if MainQuery.FieldByName('TarifC').AsString = 'EUR' then
               ArrayData[1, 14] := MainQuery.FieldByName('PTRF').AsFloat
           else
               ArrayData[1, 14] := MainQuery.FieldByName('PTRF').AsString + ' ' + MainQuery.FieldByName('TarifC').AsString;

           ASH_POLIS.Range['A' + strI, 'N'+ strI] := ArrayData;

           if MainQuery.FieldByName('DS').AsString <> '' then begin
                 strI := InttoStr(j);
                 ArrayData[1, 1] := MainQuery.FieldByName('DS').AsString + '/' + MainQuery.FieldByName('DN').AsString;
                 {ArrayData[1, 2] := 'DOC';
                 ArrayData[1, 3] := MainQuery.FieldByName('DtFrom').AsString;
                 ArrayData[1, 4] := MainQuery.FieldByName('DtTo').AsString;
                 ArrayData[1, 5] := MainQuery.FieldByName('AutoNmb').AsString;
                 ArrayData[1, 6] := MainQuery.FieldByName('ShassiNmb').AsString;
                 ArrayData[1, 7] := '';
                 ArrayData[1, 8] := MainQuery.FieldByName('YearMade').AsString;
                 ArrayData[1, 9] := '';
                 ArrayData[1, 10] := '';
                 ArrayData[1, 11] := 'COUNTRY';
                 ArrayData[1, 12] := MainQuery.FieldByName('RegDate').AsString;}
                 if MainQuery.FieldByName('TarifC').AsString = 'EUR' then
                     ArrayData[1, 13] := MainQuery.FieldByName('DTRF').AsFloat
                 else
                     ArrayData[1, 13] := MainQuery.FieldByName('DTRF').AsString + ' ' + MainQuery.FieldByName('TarifC').AsString;

                 ASH_POLIS_ADD.Range['A' + strI, 'M'+ strI] := ArrayData;
                 j := j + 1;
           end;

           Inc(i);
           MainQuery.Next;
         end;
     except
       on E : Exception do begin
         if WaitForm <> nil then WaitForm.Close;
         MessageDlg(E.Message, mtInformation, [mbOk], 0);
       end
     end;

     VarClear(ArrayData);
     INI.Free;
     if WaitForm <> nil then WaitForm.Close;
     MSEXCEL.Visible := true;
     MSEXCEL.ActiveWorkbook.SaveAs(GetAppDir + 'DOCS\BULSTRAD_' + DateToStr(Date) + '.XLS');
end;

end.
