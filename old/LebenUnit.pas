unit LebenUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Db, RxQuery, Grids, DBGrids, RXDBCtrl, Menus, StdCtrls,
  ExtCtrls, FR_Class, FR_View, FR_DSet, FR_DBSet, ComCtrls, InpRezParamsFor,
  Buttons, datamod, Clipbrd;

type
  TInsLeben2 = class(TForm)
    MainMenu: TMainMenu;
    RepMenu: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    MenuCompany: TMenuItem;
    MainQuery: TRxQuery;
    Leben2Database: TDatabase;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryInsurer: TStringField;
    MainQueryAddress: TStringField;
    MainQueryBirthdate: TStringField;
    MainQueryPassport: TStringField;
    MainQueryState: TStringField;
    MainQueryConditions: TStringField;
    MainQueryInsurer2: TStringField;
    MainQueryBegin: TDateField;
    MainQueryEnd: TDateField;
    MainQueryDuration: TFloatField;
    MainQueryAllSumma: TFloatField;
    MainQueryTarif: TFloatField;
    MainQueryPayment: TFloatField;
    MainQueryPaymentDate: TDateField;
    MainQueryGroup: TFloatField;
    MainQueryWriteDate: TDateField;
    MainQueryAgent_code: TStringField;
    MainQueryCash: TFloatField;
    MainQueryInsType: TStringField;
    MainQueryPeriod1: TStringField;
    MainQueryPeriod2: TStringField;
    MainQueryKs: TStringField;
    MainQuerySum1: TFloatField;
    MainQuerySum1Curr: TStringField;
    MainQuerySum2: TFloatField;
    MainQuerySum2Curr: TStringField;
    MainQueryPeriod: TStringField;
    MainQuerySummStr: TStringField;
    DataSourceLeben: TDataSource;
    Panel1: TPanel;
    Memo: TMemo;
    MainQueryNAME: TStringField;
    CalcSQL: TRxQuery;
    InitTimer: TTimer;
    TELEFAX1: TMenuItem;
    STATISTICS1: TMenuItem;
    N6: TMenuItem;
    WorkQuery: TQuery;
    PreviewRpt: TMenuItem;
    N9: TMenuItem;
    N5: TMenuItem;
    N7: TMenuItem;
    Reserv: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    WorkSQL: TQuery;
    IsReinsurSQL: TQuery;
    StatusPanel: TPanel;
    ProgressBar: TProgressBar;
    StatusMsg: TLabel;
    N16: TMenuItem;
    MainQueryAgPercent: TFloatField;
    MainQueryUridich: TStringField;
    MainQueryRestrax: TFloatField;
    MainQueryShowRestrax: TStringField;
    Coris1: TMenuItem;
    N17: TMenuItem;
    CorisExcel: TMenuItem;
    PolCounter: TMenuItem;
    Magazine: TMenuItem;
    UbPanel: TPanel;
    Splitter: TSplitter;
    RxDBGrid1: TRxDBGrid;
    Label1: TLabel;
    Panel2: TPanel;
    DelBtn: TButton;
    AddBtn: TButton;
    UbTable: TTable;
    UbMnu: TMenuItem;
    UbTableSer: TStringField;
    UbTableNmb: TFloatField;
    UbTableRegDate: TDateField;
    UbTableRegNmb: TStringField;
    UbTableComplName: TStringField;
    UbTableComplNmb: TStringField;
    UbTableComplDate: TDateField;
    UbTableUbSum: TFloatField;
    UbTableUbCurr: TStringField;
    UbTablePayDate: TDateField;
    UbTablePaySum: TFloatField;
    UbTablePayCurr: TStringField;
    UbTableFreeDate: TDateField;
    UbTableFreeSum: TFloatField;
    UbTableFreeCurr: TStringField;
    UbTableClDate: TDateField;
    UbTableText: TStringField;
    EditBtn: TButton;
    dsUb: TDataSource;
    UbTableBrSum: TFloatField;
    UbTableBrCurr: TStringField;
    UbTableCountry: TStringField;
    Panel3: TPanel;
    Panel4: TPanel;
    Label2: TLabel;
    Seria: TEdit;
    Number: TEdit;
    Leben2Grid: TRxDBGrid;
    SpeedButton1: TSpeedButton;
    UbTableState: TFloatField;
    N18: TMenuItem;
    MainQueryInsurtype: TStringField;
    N1: TMenuItem;
    MainQueryValuta: TStringField;
    MainQuerystopdate: TDateField;
    MainQueryretsum: TFloatField;
    MainQueryretcur: TStringField;
    MainQueryst: TStringField;
    MainQueryPolisState: TStringField;
    InputSum: TMenuItem;
    N8: TMenuItem;
    Globo: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure N2Click(Sender: TObject);
    procedure InitTimerTimer(Sender: TObject);
    procedure WorkFunction(Sender: TObject);
    procedure PreviewRptClick(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure ReservClick(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure MainQueryUridichGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure Coris1Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure CorisExcelClick(Sender: TObject);
    procedure Leben2GridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure MagazineClick(Sender: TObject);
    procedure UbMnuClick(Sender: TObject);
    procedure MainQueryAfterScroll(DataSet: TDataSet);
    procedure AddBtnClick(Sender: TObject);
    procedure DelBtnClick(Sender: TObject);
    procedure EditBtnClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure InputSumClick(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure GloboClick(Sender: TObject);
  private
    { Private declarations }
    W : HWnd;
    R : TRECT;
    Total : array[1..50] of double;
    MSWORD : Variant;
    MSEXCEL : Variant;

    function FillPlanIndividualSum(Plan, s : string; StartIndex : integer) : double;
    function GetPersons2(s, Plan, Period : string) : integer;
    function FillPlanIndividual(Plan, s : string; StartIndex : integer) : integer;
    function GetIndividualSum(Plan : string; Period, filter : string) : double;
    function GetIndividual(Plan : string; Period, filter : string) : integer;
    function FillPlanGroup(Plan, s : string; StartIndex : integer) : integer;
    function GetGroupSum(Plan : string; Period, filter : string) : double;
    function GetGroup(Plan : string; Period, filter : string) : integer;
    function FillPlanGroupSum(Plan, s : string; StartIndex : integer) : double;
    function GetFilter() : string;
    procedure FillInfo;

//    procedure FillDocsDataALBUR(var DataList : TList; RepDate : TDateTime; FizichTax, UridichTax : double);
    procedure FillDocsData(IsRNPU : boolean; var {DataList, }ReinsurList : TList; RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxComis : double; Company : integer; List : TASKOneTaxList);
  public
    IsFirstStart : boolean;
    AssistNames : array[0..10] of string;
    { Public declarations }
    function GetOrder : string;
    procedure CreateReport(IsRNPU : boolean; RepDate : TDateTime; FizichTax, UridichTax : double; S : RestPercentsDataArray; SSize : integer; IsFast : boolean; List : TASKOneTaxList);
  end;

var
  InsLeben2: TInsLeben2;
  DocN : integer;
  FastDateFrom, FastDateTo : TDateTime;

implementation

uses LebenFilterUnit, ComObj, mf, AboutUnit, CurrRatesUnit, Variants,
  DaysForSellCurrUnit, MandSellCurrUnit, Leben2RepUnit,
  MoveCurrPeriods, ForeignUnit, polandUnit, DateUtil, KorrectPolandUnit,
  WaitFormUnit, inifiles, Sort, MagazineReport, MagazLeben, Ubitok,
  RepParamsFrm, RepParamsAgntFrm, Math, ShowInfoUnit, mandsql, rxstrutils;

{$R *.DFM}

procedure TInsLeben2.FormCreate(Sender: TObject);
var
     i : integer;
     INI : TIniFile;
begin
     IsFirstStart := true;

     INI := TIniFile.Create('blank.ini');
     for i := 0 to 10 do
         AssistNames[i] := INI.ReadString('PRIVATE2', 'Assist' + IntToStr(i), '');
     INI.Free;

     MenuCompany.Caption := COMPAMYNAME;
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

procedure TInsLeben2.N4Click(Sender: TObject);
begin
     Close
end;

procedure TInsLeben2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
     Action := caFree
end;

procedure TInsLeben2.MainQueryCalcFields(DataSet: TDataSet);
begin
     if MainQueryBegin.AsString <> '' then
         MainQueryPeriod.AsString := MainQueryBegin.AsString + ' - ' + MainQueryEnd.AsString;

     if MainQuerySum1.AsFloat > 0.01 then
         MainQuerySummStr.AsString := MainQuerySum1.AsString + ' ' + MainQuerySum1Curr.AsString;

     if MainQuerySum2.AsFloat > 0.01 then begin
         MainQuerySummStr.AsString := MainQuerySummStr.AsString + '; ';
         MainQuerySummStr.AsString := MainQuerySummStr.AsString + MainQuerySum2.AsString + ' ' + MainQuerySum2Curr.AsString;
     end;

     MainQueryShowRestrax.AsString := 'НЕ ЗНАЮ';
     if (MainQueryRestrax.AsInteger >= 0) AND (MainQueryRestrax.AsInteger <= 10) then
         MainQueryShowRestrax.AsString := AssistNames[MainQueryRestrax.AsInteger];

     MainQueryPolisState.AsString := 'НОРМАЛЬНЫЙ';
     if MainQueryst.AsString = '1' then MainQueryPolisState.AsString := 'ИСПОРЧЕН';
     if MainQueryst.AsString= '2' then MainQueryPolisState.AsString := 'РАСТОРГНУТ';
end;

procedure TInsLeben2.FillInfo;
var
     s : string;
begin
     Screen.Cursor := crHourGlass;
     try
         Memo.Clear;
         CalcSQL.Close;
         CalcSQL.MacroByName('WHERE').AsString := MainQuery.MacroByName('WHERE').AsString;
         CalcSQL.Open;
         while not CalcSQL.EOF do begin
            if CalcSQL.FieldByName('S').AsFloat > 0.01 then
               s := s + CalcSQL.FieldByName('S').AsString + ' ' + CalcSQL.FieldByName('CURR').AsString + ';   ';
               CalcSQL.Next;
         end;
         Memo.Lines.Add(s);
         PolCounter.Caption := 'Полисов ' + IntToStr(MainQuery.RecordCount);
         //Memo.Lines.Add('Полисов ' + IntToStr(MainQuery.RecordCount));
         CalcSQL.Close;
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

function TInsLeben2.GetOrder : string;
var
     s : string;
begin
     s := SortForm.GetOrder('');
     if s = '' then s := 'SERIA,NUMBER';
     Result := s;
end;

procedure TInsLeben2.N2Click(Sender: TObject);
var
     prevFilt : string;
begin
     prevFilt := MainQuery.MacroByName('WHERE').AsString;//LebenFilter.getFilterStr;
     LebenFilter.ShowModal;
     if prevFilt <> LebenFilter.getFilterStr then begin
        MainQuery.Close;
        MainQuery.MacroByName('WHERE').AsString := LebenFilter.getFilterStr;
        MainQuery.MacroByName('ORDER').AsString := GetOrder();
        try
            PolCounter.Caption := '';
            Screen.Cursor := crHourGlass;
            MainQuery.Open;
            Screen.Cursor := crDefault; 
            FillInfo;
        except
          on E : Exception do begin
            MessageDlg(E.Message, mtError, [mbOk], 0);
          end;
        end
     end;
end;

procedure TInsLeben2.InitTimerTimer(Sender: TObject);
begin
     try
       Screen.Cursor := crHourGlass;
       MainQuery.Close;
       InitTimer.Enabled := false;
       if IsFirstStart then
           MainQuery.MacroByName('WHERE').AsString := 'SERIA=''TL'''
       else    
           MainQuery.MacroByName('WHERE').AsString := LebenFilter.getFilterStr;
       MainQuery.MacroByName('ORDER').AsString := GetOrder;
       IsFirstStart := false;
       MainQuery.Open;
       Screen.Cursor := crDefault;
       FillInfo;
     except
       on E : Exception do
           MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end
end;

function FTS(Val : double) : string;
var
     iVal : Int64;
begin
     iVal := ROUND(Val * 100 + 0.5);
     FTS := FloatToStr(iVal / 100);
end;

procedure TInsLeben2.WorkFunction(Sender: TObject);
var
     tmpInt : integer;
     tmpFloat : double;
     s : string;
     Path : string;
     AllSum : integer;
     SuperAllSum : integer;
     SuperAllSumFloat : double;
     IsTelefax : boolean;
//     ClipRect : TRect;
begin
     IsTelefax := Sender = TELEFAX1;

     AllSum := 0;

     Path := Application.ExeName;
     for tmpInt := Length(Path) downto 1 do
        if Path[tmpInt] = '\' then break;

     Path := Copy(Path, 1, tmpInt);

     //if LebenFilter.Restrax.ItemIndex = 0 then begin
     //   MessageDlg('Выбери ассистанта.', mtInformation, [mbOk], 0);
     //   exit;
     //end;
     //LebenFilter.Restrax.ItemIndex := 0;
     s := LebenFilter.GetFilterStr;

     try
        Screen.Cursor := crHourGlass;

        if Length(s) > 0 then s := '(' + s + ') AND ';
        s := s + ' Payment > 0.009';

        MSWORD := CreateOLEObject('WORD.BASIC');
        MSWORD.AppMaximize(1);
        MSWORD.AppShow;

        if IsTelefax then begin
              MSWORD.FileOpen(Path + 'TEMPLATES\Telefax2.doc');
              MSWORD.LineDown(30);
              MSWORD.TableSelectTable;
              MSWORD.NextCell;
              MSWORD.NextCell;
              MSWORD.LineDown;

              // 1 - 7
              AllSum := GetPersons2(s, 'InsType in (''A1'', ''A2'', ''B'', ''C'')', 'Duration between 1 and 7');
              // 8 - 14
              AllSum := AllSum + GetPersons2(s, 'InsType in (''A1'', ''A2'', ''B'', ''C'')', 'Duration between 8 and 14');
              // 15 - 31
              AllSum := AllSum + GetPersons2(s, 'InsType in (''A1'', ''A2'', ''B'', ''C'')', 'Duration between 15 and 31');
              // 32 - 92
              AllSum := AllSum + GetPersons2(s, 'InsType in (''A1'', ''A2'', ''B'', ''C'')', 'Duration between 32 and 92');
              // 93 - 365
              AllSum := AllSum + GetPersons2(s, 'InsType in (''A1'', ''A2'', ''B'', ''C'')', 'Duration between 93 and 366');
              // Y
              AllSum := AllSum + GetPersons2(s, 'InsType = ''Y''', '');
              // S
              AllSum := AllSum + GetPersons2(s, 'InsType = ''S''', '');
              // Q
              AllSum := AllSum + GetPersons2(s, 'InsType = ''Q''', '');

              MSWORD.Insert(IntToStr(AllSum));

              // Возврат
              MSWORD.LineUp(8);
              MSWORD.NextCell;
              AllSum := 0;

              // 1 - 7
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType in (''A1'', ''A2'', ''B'', ''C'') AND ');
              WorkQuery.SQL.Add('(Duration between 1 and 7) AND');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // 8 - 14
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType in (''A1'', ''A2'', ''B'', ''C'') AND ');
              WorkQuery.SQL.Add('(Duration between 8 and 14) AND');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // 15 - 31
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType in (''A1'', ''A2'', ''B'', ''C'') AND ');
              WorkQuery.SQL.Add('(Duration between 15 and 31) AND');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // 32 - 92
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType in (''A1'', ''A2'', ''B'', ''C'') AND ');
              WorkQuery.SQL.Add('(Duration between 32 and 92) AND');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // 93 - 365
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType in (''A1'', ''A2'', ''B'', ''C'') AND ');
              WorkQuery.SQL.Add('(Duration between 93 and 366) AND');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // Y
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType = ''Y'' AND ');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // S
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType = ''S'' AND ');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              // Q
              WorkQuery.SQL.Clear;
              WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
              WorkQuery.SQL.Add('InsType = ''Q'' AND ');
              WorkQuery.SQL.Add(s);
              WorkQuery.Open;
              tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
              AllSum := AllSum + tmpInt;
              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(tmpInt));
              WorkQuery.Close;

              MSWORD.LineDown;
              MSWORD.Insert(IntToStr(AllSum));

              MSWORD.FileSaveAs(Path + 'DOCS\Document' + IntToStr(DocN), 0);
              DocN := DocN + 1;
        end //Telefax
        else begin
              MSWORD.FileOpen(Path + 'TEMPLATES\Statistics2.doc');
              MSWORD.LineDown(10);
              MSWORD.TableSelectTable;
              MSWORD.NextCell;
              MSWORD.LineDown;
              MSWORD.NextCell;
              MSWORD.NextCell;

              for tmpInt := 1 to 50 do
                 Total[tmpInt] := 0;

              SuperAllSum := FillPlanIndividual('A1', s, 0);
              SuperAllSum := SuperAllSum + FillPlanIndividual('A2', s, 5);
              SuperAllSum := SuperAllSum + FillPlanIndividual('B', s, 10);
              SuperAllSum := SuperAllSum + FillPlanIndividual('C', s, 15);

              // Y individual
              tmpInt := GetIndividual('Y', '', s);
              //tmpInt := tmpInt + GetIndividual('G', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[16 + 5] := tmpInt;

              // S individual
              tmpInt := GetIndividual('S', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[17 + 5] := tmpInt;

              // Q individual
              tmpInt := GetIndividual('Q', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[18 + 5] := tmpInt;
              Total[19 + 5] := SuperAllSum;

              MSWORD.Insert(IntToStr(SuperAllSum));
              MSWORD.LineUp(18 + 5);
              MSWORD.NextCell;

              //Заполняеи вторую строку
              SuperAllSum := FillPlanGroup('A1', s, 0);
              SuperAllSum := SuperAllSum + FillPlanGroup('A2', s, 5);
              SuperAllSum := SuperAllSum + FillPlanGroup('B', s, 10);
              SuperAllSum := SuperAllSum + FillPlanGroup('C', s, 15);

              // Y individual
              tmpInt := GetGroup('Y', '', s);
              //tmpInt := tmpInt + GetGroup('G', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[16 + 5] := Total[16 + 5] + tmpInt;

              // S individual
              tmpInt := GetGroup('S', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[17 + 5] := Total[17 + 5] + tmpInt;

              // Q individual
              tmpInt := GetGroup('Q', '', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              SuperAllSum := SuperAllSum + tmpInt;
              Total[18 + 5] := Total[18 + 5] + tmpInt;
              Total[19 + 5] := Total[19 + 5] + SuperAllSum;

              MSWORD.Insert(IntToStr(SuperAllSum));
              MSWORD.LineUp(18 + 5);
              MSWORD.NextCell;

              for tmpInt := 1 to 19 + 5 do begin
                  MSWORD.Insert(FTS(Total[tmpInt]));
                  if tmpInt <> (19 + 5) then MSWORD.LineDown;
              end;

              MSWORD.LineUp(18 + 5);
              MSWORD.NextCell;

              for tmpInt := 1 to 50 do
                 Total[tmpInt] := 0;

              //////////////////////////////////////////////////////////////////
              // Премии
              // Индивидуально

              tmpInt := -1000000;

              SuperAllSumFloat := FillPlanIndividualSum('A1', s, 0);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanIndividualSum('A2', s, 5);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanIndividualSum('B', s, 10);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanIndividualSum('C', s, 15);

              // Y individual
              tmpFloat := GetIndividualSum('Y', '', s);
              //tmpFloat := tmpFloat + GetIndividualSum('G', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[16 + 5] := tmpFloat;

              // S individual
              tmpFloat := GetIndividualSum('S', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[17 + 5] := tmpFloat;

              // Q individual
              tmpFloat := GetIndividualSum('Q', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[18 + 5] := tmpFloat;
              Total[19 + 5] := SuperAllSumFloat;

              MSWORD.Insert(FTS(SuperAllSumFloat));
              MSWORD.LineUp(18 + 5);
              MSWORD.NextCell;

              //Заполняеи вторую строку
              SuperAllSumFloat := FillPlanGroupSum('A1', s, 0);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanGroupSum('A2', s, 5);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanGroupSum('B', s, 10);
              SuperAllSumFloat := SuperAllSumFloat + FillPlanGroupSum('C', s, 15);

              // Y individual
              tmpFloat := GetGroupSum('Y', '', s);
              //tmpFloat := tmpFloat + GetGroupSum('G', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[16 + 5] := Total[16 + 5] + tmpFloat;

              // S individual
              tmpFloat := GetGroupSum('S', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[17 + 5] := Total[17 + 5] + tmpFloat;

              // Q individual
              tmpFloat := GetGroupSum('Q', '', s);
              MSWORD.Insert(FTS(tmpFloat));
              MSWORD.LineDown;
              SuperAllSumFloat := SuperAllSumFloat + tmpFloat;
              Total[18 + 5] := Total[18 + 5] + tmpFloat;
              Total[19 + 5] := Total[19 + 5] + SuperAllSumFloat;

              MSWORD.Insert(FTS(SuperAllSumFloat));
              MSWORD.LineUp(18 + 5);
              MSWORD.NextCell;

              for tmpInt := 1 to 19 + 5 do begin
                  MSWORD.Insert(FTS(Total[tmpInt]));
                  MSWORD.LineDown;
              end;

              MSWORD.FileSaveAs(Path + 'DOCS\Document' + IntToStr(DocN), 0);
              DocN := DocN + 1;
         end;
         ClipCursor(nil);
     except on E : Exception do begin
            Screen.Cursor := crDefault;
            ClipCursor(nil);

            try
               MSWORD.AppClose;
            except
            end;
            MessageBeep(0);
            SetWindowPos(Handle, 0, 0, 0, 0, 0, SWP_NOSIZE OR SWP_NOMOVE);
            MessageDlg(e.Message, mtInformation, [mbOK], 0);
            exit;
        end;
     end;
     ClipCursor(nil);
     Screen.Cursor := crDefault;
     //Close;
end;

//Возвращает количество индивидуальных полисов
function TInsLeben2.GetIndividualSum(Plan : string; Period, filter : string) : double;
begin
    WorkQuery.SQL.Clear;
    WorkQuery.SQL.Add('select SUM(Payment) As cnt from private2 where ');
    WorkQuery.SQL.Add('InsType = ''' + Plan + ''' AND ');
    WorkQuery.SQL.Add('(Private2."Group" <= 1 OR Private2."Group" is null) AND ');
    if Length(Period) > 0 then
       WorkQuery.SQL.Add('(' + Period + ') AND');
    WorkQuery.SQL.Add(Filter);
    WorkQuery.Open;
    GetIndividualSum := WorkQuery.FieldByName('CNT').AsFloat;
    WorkQuery.Close;
end;

function TInsLeben2.GetGroupSum(Plan : string; Period, filter : string) : double;
begin
    WorkQuery.SQL.Clear;
    WorkQuery.SQL.Add('select SUM(Payment) As cnt from private2 where ');
    WorkQuery.SQL.Add('InsType = ''' + Plan + ''' AND ');
    WorkQuery.SQL.Add('(Private2."Group" > 1) AND ');
    if Length(Period) > 0 then
       WorkQuery.SQL.Add('(' + Period + ') AND');
    WorkQuery.SQL.Add(Filter);
    WorkQuery.Open;
    GetGroupSum := WorkQuery.FieldByName('CNT').AsFloat;
    WorkQuery.Close;
end;

//Возвращает количество индивидуальных полисов
function TInsLeben2.GetIndividual(Plan : string; Period, filter : string) : integer;
begin
    WorkQuery.SQL.Clear;
    WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
    WorkQuery.SQL.Add('InsType = ''' + Plan + ''' AND ');
    WorkQuery.SQL.Add('(Private2."Group" <= 1 OR Private2."Group" is null) AND ');
    if Length(Period) > 0 then
       WorkQuery.SQL.Add('(' + Period + ') AND');
    WorkQuery.SQL.Add(Filter);
    WorkQuery.Open;
    GetIndividual := WorkQuery.FieldByName('CNT').AsInteger;
    WorkQuery.Close;
end;

function TInsLeben2.GetGroup(Plan : string; Period, filter : string) : integer;
var
    tmpInt : integer;
begin
    WorkQuery.SQL.Clear;
    WorkQuery.SQL.Add('select SUM(Private2."Group") As cnt from private2 where ');
    WorkQuery.SQL.Add('InsType = ''' + Plan + ''' AND ');
    WorkQuery.SQL.Add('(Private2."Group" > 1) AND');
    if Length(Period) > 0 then
       WorkQuery.SQL.Add('(' + Period + ') AND');
    WorkQuery.SQL.Add(Filter);
    WorkQuery.Open;
    tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
    WorkQuery.Close;

    GetGroup := tmpInt;
end;

function TInsLeben2.FillPlanGroup(Plan, s : string; StartIndex : integer) : integer;
var
      tmpInt : integer;
      AllSum : integer;
begin
              // 1 - 14 individual
              tmpInt := GetGroup(Plan, 'Duration between 1 and 14', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := tmpInt;
              Total[StartIndex + 1] := Total[StartIndex + 1] + tmpInt;

              // 15 - 30 individual
              tmpInt := GetGroup(Plan, 'Duration between 15 and 30', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 2] := Total[StartIndex + 2] + tmpInt;

              // 31 - 90 individual
              tmpInt := GetGroup(Plan, 'Duration between 31 and 90', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 3] := Total[StartIndex + 3] + tmpInt;

              // 91 - 365 individual
              tmpInt := GetGroup(Plan, 'Duration between 91 and 366', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 4] := Total[StartIndex + 4] + tmpInt;
              Total[StartIndex + 5] := Total[StartIndex + 5] + AllSum;
              MSWORD.Insert(IntToStr(AllSum));
              MSWORD.LineDown;

              FillPlanGroup := AllSum;
end;

//Возвращает сумму полисов для плана
function TInsLeben2.FillPlanIndividual(Plan, s : string; StartIndex : integer) : integer;
var
      tmpInt : integer;
      AllSum : integer;
begin
              // 1 - 14 individual
              tmpInt := GetIndividual(Plan, 'Duration between 1 and 14', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := tmpInt;
              Total[StartIndex + 1] := Total[StartIndex + 1] + tmpInt;

              // 15 - 30 individual
              tmpInt := GetIndividual(Plan, 'Duration between 15 and 30', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 2] := Total[StartIndex + 2] + tmpInt;

              // 31 - 90 individual
              tmpInt := GetIndividual(Plan, 'Duration between 31 and 90', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 3] := Total[StartIndex + 3] + tmpInt;

              // 91 - 365 individual
              tmpInt := GetIndividual(Plan, 'Duration between 91 and 366', s);
              MSWORD.Insert(IntToStr(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 4] := Total[StartIndex + 4] + tmpInt;
              Total[StartIndex + 5] := Total[StartIndex + 5] + AllSum;
              MSWORD.Insert(IntToStr(AllSum));
              MSWORD.LineDown;

              FillPlanIndividual := AllSum;
end;

function TInsLeben2.FillPlanGroupSum(Plan, s : string; StartIndex : integer) : double;
var
      tmpInt : double;
      AllSum : double;
begin
              // 1 - 14 individual
              tmpInt := GetGroupSum(Plan, 'Duration between 1 and 14', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := tmpInt;
              Total[StartIndex + 1] := Total[StartIndex + 1] + tmpInt;

              // 15 - 30 individual
              tmpInt := GetGroupSum(Plan, 'Duration between 15 and 30', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 2] := Total[StartIndex + 2] + tmpInt;

              // 31 - 90 individual
              tmpInt := GetGroupSum(Plan, 'Duration between 31 and 90', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 3] := Total[StartIndex + 3] + tmpInt;

              // 91 - 365 individual
              tmpInt := GetGroupSum(Plan, 'Duration between 91 and 366', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 4] := Total[StartIndex + 4] + tmpInt;
              Total[StartIndex + 5] := Total[StartIndex + 5] + AllSum;
              MSWORD.Insert(FTS(AllSum));
              MSWORD.LineDown;

              FillPlanGroupSum := AllSum;
end;

function TInsLeben2.FillPlanIndividualSum(Plan, s : string; StartIndex : integer) : double;
var
      tmpInt : double;
      AllSum : double;
begin
              // 1 - 14 individual
              tmpInt := GetIndividualSum(Plan, 'Duration between 1 and 14', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := tmpInt;
              Total[StartIndex + 1] := Total[StartIndex + 1] + tmpInt;

              // 15 - 30 individual
              tmpInt := GetIndividualSum(Plan, 'Duration between 15 and 30', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 2] := Total[StartIndex + 2] + tmpInt;

              // 31 - 90 individual
              tmpInt := GetIndividualSum(Plan, 'Duration between 31 and 90', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 3] := Total[StartIndex + 3] + tmpInt;

              // 91 - 365 individual
              tmpInt := GetIndividualSum(Plan, 'Duration between 91 and 366', s);
              MSWORD.Insert(FTS(tmpInt));
              MSWORD.LineDown;
              AllSum := AllSum + tmpInt;
              Total[StartIndex + 4] := Total[StartIndex + 4] + tmpInt;
              Total[StartIndex + 5] := Total[StartIndex + 5] + AllSum;
              MSWORD.Insert(FTS(AllSum));
              MSWORD.LineDown;

              FillPlanIndividualSum := AllSum;
end;

function TInsLeben2.GetPersons2(s, Plan, Period : string) : integer;
var
      tmpInt : integer;
begin
      //Количество полисов, где один человек застрахован
      WorkQuery.SQL.Clear;
      WorkQuery.SQL.Add('select count(*) As cnt from private2 where ');
      WorkQuery.SQL.Add(Plan + ' AND ');
      WorkQuery.SQL.Add('(private2."Group" <= 1 OR private2."Group" is null) AND ');
      if Length(Period) > 0 then
            WorkQuery.SQL.Add('(' + Period + ') AND');
      WorkQuery.SQL.Add(s);
      WorkQuery.Open;
      tmpInt := WorkQuery.FieldByName('CNT').AsInteger;
      WorkQuery.Close;

      //Сумма челове в полисе, где количество человек в полисе есть и полис групповой
      WorkQuery.SQL.Clear;
      WorkQuery.SQL.Add('select sum(private2."Group") As cnt from private2 where ');
      WorkQuery.SQL.Add(Plan + ' AND ');
      WorkQuery.SQL.Add('(private2."Group" > 1 AND private2."Group" is not null) AND ');
      if Length(Period) > 0 then
            WorkQuery.SQL.Add('(' + Period + ') AND');
      WorkQuery.SQL.Add(s);
      WorkQuery.Open;
      tmpInt := tmpInt + WorkQuery.FieldByName('CNT').AsInteger;
      WorkQuery.Close;
(*
      //Сумма группы в полисе, где количество человек в полисе отсутствует и полис групповой
      WorkQuery.SQL.Clear;
      WorkQuery.SQL.Add('select sum(Private."Count") As cnt from private where ');
      WorkQuery.SQL.Add(Plan + ' AND ');
      WorkQuery.SQL.Add('(Private."Count" > 1 AND Mans is null) AND ');
      if Length(Period) > 0 then
            WorkQuery.SQL.Add('(' + Period + ') AND');
      WorkQuery.SQL.Add(s);
      WorkQuery.Open;
      tmpInt := tmpInt + WorkQuery.FieldByName('CNT').AsInteger;
      WorkQuery.Close;
*)
      MSWORD.Insert(IntToStr(tmpInt));
      MSWORD.LineDown;

      GetPersons2 := tmpInt;
end;

procedure TInsLeben2.PreviewRptClick(Sender: TObject);
begin
    if LebenReport <> nil then exit;
    if MainQuery.RecordCount = 0 then exit;

    if LebenReport = nil then begin
        Application.CreateForm(TLebenReport, LebenReport);
    end;
    LebenReport.frReport.Dictionary.Variables['ITOGO'] := '''ИТОГО: ' + Memo.Lines[0] + '''';
    LebenReport.Show;
    LebenReport.BringToFront;

    Leben2Grid.DataSource := nil;
    MainQuery.First;
    LebenReport.frReport.ShowReport;
    Leben2Grid.DataSource := DataSourceLeben;
end;

procedure TInsLeben2.N7Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по страхованию жизни';
    About.ShowModal
end;

procedure TInsLeben2.N11Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

procedure TInsLeben2.N12Click(Sender: TObject);
begin
     MessageDlg('Эта функция уже не используется', mtInformation, [mbOk], 0);
     DaysForSellCurr.ShowModal
end;

procedure TInsLeben2.N13Click(Sender: TObject);
begin
     if MandSellCurr = nil then begin
         Application.CreateForm(TMandSellCurr, MandSellCurr);
         MandSellCurr.ShowModal;
     end;
end;

procedure TInsLeben2.ReservClick(Sender: TObject);
var
     List : TASKOneTaxList;
begin
     if InitRezervLeben.ShowModal = mrOk then begin
         Screen.Cursor := crHourGlass;
         FillOutTaxList(List, InitRezervLeben.OutFonds);
         InsLeben2.CreateReport(TControl(Sender).Tag = 1,
                                InitRezervLeben.RepData.Date,
                                InitRezervLeben.FizichTax.Value,
                                InitRezervLeben.UridichTax.Value,
                                InitRezervLeben.S,
                                InitRezervLeben.StringGrid.RowCount - 1,
                                InitRezervLeben.IsFast.Checked,
                                List);
         Screen.Cursor := crDefault;
         Application.BringToFront;
     end;
end;

procedure TInsLeben2.N15Click(Sender: TObject);
begin
     MoveCurr.ShowModal
end;

procedure TInsLeben2.CreateReport(IsRNPU : boolean; RepDate : TDateTime; FizichTax, UridichTax : double; S : RestPercentsDataArray; SSize : integer; IsFast : boolean; List : TASKOneTaxList);
var
    ReinsurData : TList;
    idx_curr, i, j : integer;
    ExcelLine, ExcelLineStartPart : integer;
    OutFileName : string;
    CellNumber : string;
    VAL : VARIANT;
    CellRate : string;
    str_big_summ : string;
    iCompany : integer;
begin
    ReinsurData := TList.Create;
    try
         for iCompany := 1 to SSize do begin
             StatusMsg.Caption := 'Резерв по полисам ' + AssistNames[iCompany - 1];
             FillDocsData(IsRNPU, ReinsurData, TRUNC(RepDate), FizichTax, UridichTax, S[iCompany].PutPercent, S[iCompany].ComisPercent, iCompany - 1, List);
         end;
         if not InitExcel(MSEXCEL) then begin
            MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
            exit;
         end;
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;
         WaitForm.ProgressBar.Position := 0;
         MSEXCEL.Visible := true;
         //TransferToExecl('Страхование жизни и здоровья, выезжающих за рубеж', RepDate, 'LEBEN2', ListData, MSEXCEL, IsFast);
         OutTo10Form(IsRNPU, 'Страхование жизни и здоровья, выезжающих за рубеж', MSEXCEL, RepDate, ReinsurData, IsFast, 'LEBEN2');
    except
        on E : Exception do begin
            Application.BringToFront;
            if WaitForm <> nil then WaitForm.Close;
            MessageDlg('Возникла ошибка при формировании отчёта'#13 + e.Message, mtInformation, [mbOk], 0);
        end;
    end;

    //for i := 0 to ListData.Count - 1 do Dispose(ListData.Items[i]);
    //ListData.Free;
    for i := 0 to ReinsurData.Count - 1 do Dispose(ReinsurData.Items[i]);
    ReinsurData.Free;
    ProgressBar.Position := 0;
    StatusPanel.Visible := false;
    if WaitForm <> nil then WaitForm.Close;
end;

// _3 потому что максимум может заполнить 2 валюты
// Для полисов, которые не перестраховаваются
// Не перестраховывается
// Отличается от PolandUnit ф-ии расчётом
procedure FillStructure_3(RepDate : TDateTime; IsUridich : boolean; FizichTax, UridichTax : double; var Info : ReportStructure_3; InpData : TDateTime; Summa : double; Curr : string; AgPercent, OneTax : double);
var
     SellData : TDateTime;
     CurrIndex : integer;
     SellPercent : double;
     SellSumma : double;
     AgentSumma : double;
begin
     if Curr = 'BRB' then begin
        CurrIndex := GetCurrIndex(Info, Curr);
        Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                         Summa;
        AgentSumma := Summa * AgPercent / 100;
        if IsUridich then
            AgentSumma := AgentSumma * (1 + UridichTax / 100)
        else
            AgentSumma := AgentSumma * (1 + FizichTax / 100);
        Info[CurrIndex].AgOtchislenie := Info[CurrIndex].AgOtchislenie +
                                       AgentSumma;
        Info[CurrIndex].OneTaxOtchislenie := Info[CurrIndex].OneTaxOtchislenie +
                                       Summa * OneTax / 100;
     end else begin
        CurrIndex := GetCurrIndex(Info, Curr);

        Info[CurrIndex].OneTaxOtchislenie := Info[CurrIndex].OneTaxOtchislenie +
                                       Summa * OneTax / 100;

        AgentSumma := Summa * AgPercent / 100;
        if IsUridich then
            AgentSumma := AgentSumma * (1 + UridichTax / 100)
        else
            AgentSumma := AgentSumma * (1 + FizichTax / 100);
        Info[CurrIndex].AgOtchislenie := Info[CurrIndex].AgOtchislenie +
                                       AgentSumma;
        //Дата продажи
        SellData := CalcSellDate(InpData, Curr);
        {if SellData >= RepDate then begin //Не продаём
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             Summa;
            AgentSumma := Summa * AgPercent / 100;
            if IsUridich then
                AgentSumma := AgentSumma * (1 + UridichTax / 100)
            else
                AgentSumma := AgentSumma * (1 + FizichTax / 100);
            Info[CurrIndex].Otchislenie := Info[CurrIndex].Otchislenie +
                                           AgentSumma;
        end
        else begin //Успели продать}
            SellPercent := GetSellPercent(SellData, Curr);
            SellSumma := Summa * SellPercent / 100;
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             (Summa - SellSumma);
            //Агенту ничего, т.к продали

            CurrIndex := GetCurrIndex(Info, 'BRB');
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             SellSumma * GetRateOfCurrency(Curr, SellData);

            //AgentSumma := Summa * AgPercent / 100;
            //if IsUridich then
            //    AgentSumma := AgentSumma * (1 + UridichTax / 100)
            //else
            //    AgentSumma := AgentSumma * (1 + FizichTax / 100);

            //Info[CurrIndex].Otchislenie := Info[CurrIndex].Otchislenie +
            //                               AgentSumma * GetRateOfCurrency(Curr, InpData);
        //end;
     end;
end;

function GetStructPtr(var Info : GroupRestrax_3; Curr : string) : PGroupRestrax;
var
     i : integer;
begin
     for i := 1 to 3 do begin
         if (Info[i].Curr = '') OR (Info[i].Curr = Curr) then begin
             Info[i].Curr := Curr;
             GetStructPtr := @(Info[i]);
             exit;
         end;
     end;

     raise Exception.Create('Ошибка в блоке GetStructPtr');
end;

procedure InitReinsurData(var Data : GroupRestrax);
begin
      Data.Curr := '';
      Data.Number := -1;
      Data.DateFrom := 0;
      Data.DateTo := 0;
      Data.BruttoPremium := 0;
      Data.AgentSum := 0;
      Data.FondSum := 0;
      Data.BasePremium := 0;
      Data.ReserveNZPSumm := 0;
      Data.RestraxPart := 0;
      Data.StraxPart := 0;
end;

//Возвращает, попал ли полис в перестрахование или нет
function DecideIsReinsur(RepDate, PayDate : TDateTime; var SendDate : TDateTime; var RestDok : string) : boolean;
begin
     if (FastDateFrom <> 0) AND (FastDateTo <> 0) AND
        (PayDate >= FastDateFrom) AND (PayDate <= FastDateTo) then begin
        DecideIsReinsur := true;
        exit;
     end;
     
     //InsLeben2.IsReinsurSQL.ParamByName('RepDate').AsDateTime := RepDate;
     InsLeben2.IsReinsurSQL.ParamByName('PayDate').AsDateTime := PayDate;
     InsLeben2.IsReinsurSQL.Open;
     SendDate := InsLeben2.IsReinsurSQL.FieldByName('PayDate').AsDateTime;
     //if SendDate < Date - 3650 then
     //    raise Exception.Create('Дата перестрахования ' + DateToStr(SendDate));
     RestDok := '';//InsLeben2.IsReinsurSQL.FieldByName('Doc').AsString;
     DecideIsReinsur := not InsLeben2.IsReinsurSQL.EOF;

     if not InsLeben2.IsReinsurSQL.EOF then begin
         FastDateFrom := InsLeben2.IsReinsurSQL.FieldByName('PeriodFrom').AsDateTime;
         FastDateTo := InsLeben2.IsReinsurSQL.FieldByName('PeriodTo').AsDateTime;
     end;

     InsLeben2.IsReinsurSQL.Close;
end;
{
//Расчёт не отдающихся в перестрахование
procedure TInsLeben2.FillDocsDataALBUR(var DataList : TList; RepDate : TDateTime; FizichTax, UridichTax : double);
var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : ^ReportStructure;
begin
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT');
          Add('   Seria, Number, WriteDate As RegDate, T."Begin" As DateFrom, T."End" As DateTo, Sum1 As PremiumPay, Sum1Curr As PremiumCurr, Sum2 As PremiumPay2, Sum2Curr As PremiumCurr2, Uridich, AgPercent');
          Add('FROM');
          Add('   PRIVATE2 T');
          Add('WHERE');
          Add('   WriteDate is not null AND Restrax=1 AND');
          Add('   (T."Begin" < :RepDate AND T."End" >= :RepDate OR ');
          Add('   (PaymentDate >= :RepDateSubMonth AND PaymentDate < :RepDate AND (T."End" - T."Begin" + 1) <= 30))');
          Add('ORDER BY');
          Add('   Seria, Number');
          ParamByName('RepDate').AsDateTime := TRUNC(RepDate);
          ParamByName('RepDateSubMonth').AsDateTime := IncDate(TRUNC(RepDate), 0, -1, 0);
          try
              Open;
              if EOF then begin
                 MessageDlg('Не обнаружено полисов ALBUR по указанному условию', mtInformation, [mbOk], 0);
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
                  FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'Y', FizichTax, UridichTax, Info, FieldByName('RegDate').AsDateTime, FieldByName('PremiumPay').AsFloat, FieldByName('PremiumCurr').AsString, FieldByName('AgPercent').AsFloat);
               if FieldByName('PremiumCurr2').AsString <> '' then
                  FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'Y', FizichTax, UridichTax, Info, FieldByName('RegDate').AsDateTime, FieldByName('PremiumPay2').AsFloat, FieldByName('PremiumCurr2').AsString, FieldByName('AgPercent').AsFloat);
             end;

             //Append Info to container
             for i := 1 to 3 do begin
                 if Info[i].Curr <> '' then begin
                     New(InfoPtr);
                     InfoPtr^ := Info[i];
                     InfoPtr.Number := WorkSQL.FieldByName('Number').AsFloat;
                     InfoPtr.RegDate := WorkSQL.FieldByName('RegDate').AsDateTime;
                     InfoPtr.FromDate := WorkSQL.FieldByName('DateFrom').AsDateTime;
                     InfoPtr.ToDate := WorkSQL.FieldByName('DateTo').AsDateTime;
                     DataList.Add(InfoPtr);
                 end;
             end;

             WorkSQL.Next;
         end;
     ProgressBar.Position := 0;
end;
}

procedure CalcRNP(var RNP, RESTRAX, STRAX : double;
                  localBasePremium, SellValuta : double;
                  RestPart : double;
                  Repdate, PayDate, DateFrom, DateTo, SendDate : TDateTime;
                  DurationIns, DurationsToRepDate : integer
                  );
var
    Kf : double;
begin
    if (RepDate < DateFrom) then begin
        RNP := localBasePremium;
        RESTRAX := 0;
        STRAX := localBasePremium;
    end;
    if (RepDate >= DateFrom) AND (RepDate <= DateTo) then begin
        if (SendDate > 0) AND (SendDate < Repdate) then begin
            Kf := RestPart / (localBasePremium + SellValuta);
            //перечислили до отчёта
            if DurationIns <= 30 then begin
                RNP := (localBasePremium + SellValuta) / 2;
            end
            else begin
                RNP := (localBasePremium + SellValuta) * DurationsToRepDate / DurationIns;
            end;

            RESTRAX := RNP * Kf;

            if DurationIns <= 30 then begin
                RNP := (localBasePremium) / 2;
            end
            else begin
                RNP := (localBasePremium) * DurationsToRepDate / DurationIns;
            end;

            STRAX := RNP - RESTRAX;
        end
        else begin
            //не перечислил денег фактически
            RESTRAX := 0;
            if DurationIns <= 30 then
                STRAX := (localBasePremium - RestPart) / 2
            else
                STRAX := (localBasePremium - RestPart) * DurationsToRepDate / DurationIns;
            STRAX := STRAX + RestPart;
            RNP := STRAX + RESTRAX;
        end;
    end;
    if (RepDate > DateTo) then begin
        if (SendDate > 0) AND (SendDate < Repdate) then begin
            //перечислили до отчёта
            if (DurationIns <= 30) AND (Repdate - PayDate <= 30) then begin
                Kf := RestPart / (localBasePremium + SellValuta);
                //перечислили до отчёта
                RNP := (localBasePremium) / 2;
                RESTRAX := RNP * Kf;
                STRAX := RNP - RESTRAX;
            end
            else begin
                STRAX := 0;
                RESTRAX := 0;
                RNP := 0;
            end;
        end
        else begin
            //не перечислил денег фактически
            STRAX := RestPart;
            RESTRAX := 0;
            RNP := RestPart;
        end;
    end;
end;

procedure TInsLeben2.FillDocsData(IsRNPU : boolean; var {DataList, }ReinsurList : TList; RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxComis : double; Company : integer; List : TASKOneTaxList);
var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : ^ReportStructure;
     IsNotReinsurPolis : boolean;

     unReinsurData : GroupRestrax_3;
     ReinsurDataPtr, ReinsurDataBRBPtr : PGroupRestrax;

     //PayPart : double;

     SellDate : TDateTime;

     SellValuta : double;

     DurationIns, DurationsToRepDate : integer;

     AgentSumma : double;

//     IsfirstCurrBRB : boolean;
//     IssecondCurrBRB : boolean;

     ListOfSumm : array[1..2] of double;
     ListOfCurr : array[1..2] of string[5];

     iter_i : integer;
     RestPartSumm : double;

     localStraxPart : double; 
     localBasePremium : double;
     localRestraxPart : double;
     localNZPSumm : double;
     KfRest : double;
     localTaxSumm : double;
     localBruttoPremium : double;
     //localRestPremium : double;
     localComis : double;
     RecordCountIsRS, RecordNumber : integer;
     SendDate : TDateTime;
     RestDok : string;

     StopNumber : integer;
     AgKf : double;
     TaxKf : double;
begin //  try
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT');
          Add('   Seria, AllSumma, ''USD'' AS AllSummaCurr, Insurer, Insurer2, Number, WriteDate As RegDate, T."Begin" As DateFrom, T."End" As DateTo, Sum1 As PaySumma, Sum1Curr As Curr, Sum2, Sum2Curr, PaymentDate As PayDate, Uridich, AgPercent');
          Add('FROM');
          Add('   PRIVATE2 T');
          Add('WHERE');
          Add('   WriteDate is not null AND ');
          Add('   (stopDate is null or stopDate > :RepDate) AND ');
          if Company = 0 then
              Add('(Restrax is null or Restrax =' + IntToStr(Company) + ')')
          else
              Add('(Restrax =' + IntToStr(Company) + ')');
          if IsRNPU then begin
                Add('   AND (PaymentDate between :RepDateSubMonth and :RepDate)');
                //Add('   AND (T."Begin" < :RepDate)')
          end
          else begin
                Add('   AND (PaymentDate < :RepDate AND T."End" >= :RepDate OR ');
                Add('   (PaymentDate >= :RepDateSubMonth AND PaymentDate < :RepDate AND (T."End" - T."Begin" + 1) <= 30))');
                Add('UNION');
                Add('SELECT');
                Add('   Seria, AllSumma, ''USD'' AS AllSummaCurr, Insurer, Insurer2, Number, WriteDate As RegDate, T."Begin" As DateFrom, T."End" As DateTo, Sum1 As PaySumma, Sum1Curr As Curr, Sum2, Sum2Curr, PaymentDate As PayDate, Uridich, AgPercent');
                Add('FROM');
                Add('   PRIVATE2 T,MOVETORESTRAX M');
                Add('WHERE');
                Add('   WriteDate is not null AND ');
                Add('   (stopdate is null or stopDate > :RepDate) AND ');
                Add('   PaymentDate < :RepDate AND ');
                if Company = 0 then
                    Add('(Restrax is null or Restrax =' + IntToStr(Company) + ')')
                else
                    Add('(Restrax =' + IntToStr(Company) + ')');
                Add(' and  T.PaymentDate between M.periodfrom and M.periodto ');
                Add(' and  (M.PayDate is null or M.PayDate >= :RepDate)');  //планируется оплатить
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
                 //MessageDlg('Не обнаружено полисов Mercur Assistance по указанному условию', mtInformation, [mbOk], 0);
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

     // Fill Structure
         ProgressBar.Position := 0;
         StatusPanel.Visible := true;
         StatusPanel.Update;
         RecordCountIsRS := WorkSQL.RecordCount;
         RecordNumber := 0;

         FastDateFrom := 0;
         FastDateTo := 0;
         while not WorkSQL.EOF do begin
             Summa := 0;
             InitStructure_3(Info);

             Inc(RecordNumber);
             ProgressBar.Position := TRUNC(RecordNumber * 100 / RecordCountIsRS);

             if WorkSQL.FieldByName('Number').AsInteger = 26172 then
                MessageBeep(0);

             for iter_i := 1 to 3 do
                 InitReinsurData(unReinsurData[iter_i]);

             IsNotReinsurPolis := true;
             if OutPercent > 0 then begin
                   //IsNotReinsurPolis := (WorkSQL.FieldByName('DateTo').AsDateTime - WorkSQL.FieldByName('DateFrom').AsDateTime + 1) <= 30;
                   //if not IsNotReinsurPolis then begin
                       //Проверить, отдали в перестрахование или нет
                       IsNotReinsurPolis := not DecideIsReinsur(RepDate, WorkSQL.FieldByName('PayDate').AsDateTime, SendDate, RestDok);
                   //end;
             end;

             if IsNotReinsurPolis then begin
                  MessageDlg('Не перестрахован полис ' + WorkSQL.FieldByName('Number').AsString + '. Нет расчёта', mtInformation, [mbOk], 0);
                  exit;
                  {
                  with WorkSQL do begin
                      if FieldByName('Curr').AsString <> '' then
                          FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'Y', FizichTax, UridichTax, Info, FieldByName('PayDate').AsDateTime, FieldByName('PaySumma').AsFloat, FieldByName('Curr').AsString, FieldByName('AgPercent').AsFloat, OneTax);
                      if FieldByName('Sum2Curr').AsString <> '' then
                          FillStructure_3(RepDate, FieldByName('Uridich').AsString = 'Y', FizichTax, UridichTax, Info, FieldByName('PayDate').AsDateTime, FieldByName('Sum2').AsFloat, FieldByName('Sum2Curr').AsString, FieldByName('AgPercent').AsFloat, OneTax);
                  end;

                   //Append Info to container
                  for i := 1 to 3 do begin
                       if Info[i].Curr <> '' then begin
                           New(InfoPtr);
                           InfoPtr^ := Info[i];
                           InfoPtr.Number := WorkSQL.FieldByName('Number').AsFloat;
                           InfoPtr.RegDate := WorkSQL.FieldByName('RegDate').AsDateTime;
                           InfoPtr.FromDate := WorkSQL.FieldByName('DateFrom').AsDateTime;
                           InfoPtr.ToDate := WorkSQL.FieldByName('DateTo').AsDateTime;
                           DataList.Add(InfoPtr);
                       end;
                   end; }
             end
             else begin
                with WorkSQL do begin
                   AgKf := 1;
                   TaxKf := 1;
                   //if FieldByName('DateFrom').AsDateTime >= RepDate then begin
                   //    AgKf := 0;
                   //    TaxKf := 0;
                   //end;
                   ListOfSumm[1] := 0;
                   ListOfSumm[2] := 0;
                   ListOfCurr[1] := '';
                   ListOfCurr[2] := '';
                   iter_i := 1;
                   if FieldByName('Curr').AsString <> '' then begin
                      ListOfSumm[iter_i] := FieldByName('PaySumma').AsFloat;
                      ListOfCurr[iter_i] := FieldByName('Curr').AsString;
                      Inc(iter_i);
                   end;
                   if FieldByName('Sum2Curr').AsString <> '' then begin
                      ListOfSumm[iter_i] := FieldByName('Sum2').AsFloat;
                      ListOfCurr[iter_i] := FieldByName('Sum2Curr').AsString;
                      Inc(iter_i);
                   end;

                   //Check for BRB and not BRB
                   DurationIns := Trunc(FieldByName('DateTo').AsDateTime) - Trunc(FieldByName('DateFrom').AsDateTime) + 1;
                   DurationsToRepDate := Trunc(FieldByName('DateTo').AsDateTime) - Trunc(RepDate) + 1;

                   for iter_i := 1 to 2 do begin
                       if ListOfCurr[iter_i] = '' then break;
                       if ListOfSumm[iter_i] <= 0.01 then begin
                          MessageDlg('По полису ' + WorkSQL.FieldByName('Number').AsString + ' есть валюта но нет суммы оплаты. Это не хорошо. Исправь.', mtInformation, [mbOk], 0);
                          break;
                       end;
                       if ListOfCurr[iter_i] <> 'BRB' then begin
                          ReinsurDataPtr := GetStructPtr(unReinsurData, ListOfCurr[iter_i]);
                          ReinsurDataBRBPtr := GetStructPtr(unReinsurData, 'BRB');
                          RestPartSumm := ListOfSumm[iter_i] * OutPercent / 100 * (1 - RestraxComis / 100);

                          SellDate := CalcSellDate(FieldByName('PayDate').AsDateTime,  ListOfCurr[iter_i]);
                          SellValuta := ListOfSumm[iter_i] * GetSellPercent(SellDate, ListOfCurr[iter_i]) / 100;

                          RoundValue(SellValuta);

                          localBruttoPremium := (ListOfSumm[iter_i] - SellValuta);
                          AgentSumma := AgKf * ListOfSumm[iter_i] * FieldByName('AgPercent').AsFloat / 100;

                          if WorkSQL.FieldByName('Uridich').AsString = 'Y' then
                              AgentSumma := AgentSumma * (1 + UridichTax / 100)
                          else
                              AgentSumma := AgentSumma * (1 + FizichTax / 100);
                          localTaxSumm := TaxKf * ListOfSumm[iter_i] * GetOneTax(List, WorkSQL.FieldByName('PayDate').AsDateTime) / 100;

                          RoundValue(AgentSumma);
                          RoundValue(localTaxSumm);

                          localBasePremium := localBruttoPremium - AgentSumma - localTaxSumm;

                          CalcRNP(localNZPSumm,
                                  localRestraxPart,
                                  localStraxPart,

                                  localBasePremium, SellValuta,
                                  RestPartSumm,

                                  RepDate,
                                  FieldByName('PayDate').AsDateTime,
                                  FieldByName('DateFrom').AsDateTime,
                                  FieldByName('DateTo').AsDateTime,
                                  SendDate, DurationIns, DurationsToRepDate);

                          if IsRNPU then begin
                              localNZPSumm := (localBasePremium + SellValuta) / 10;
                              KfRest := RestPartSumm / (localBasePremium+SellValuta);
                              localRestraxPart := localNZPSumm * KfRest;
                              RoundValue(localRestraxPart);
                              //localStraxPart := localNZPSumm * (1 - KfRest);
                              localNZPSumm := (localBasePremium) / 10;
                              RoundValue(localNZPSumm);
                              localStraxPart := localNZPSumm - localRestraxPart;
                          end;

                          ReinsurDataPtr.BruttoPremium := ReinsurDataPtr.BruttoPremium + localBruttoPremium;
                          ReinsurDataPtr.BasePremium := ReinsurDataPtr.BasePremium + localBasePremium;
                          ReinsurDataPtr.ReserveNZPSumm := ReinsurDataPtr.ReserveNZPSumm + localNZPSumm;
                          ReinsurDataPtr.RestraxPart := ReinsurDataPtr.RestraxPart + localRestraxPart;
                          ReinsurDataPtr.StraxPart := ReinsurDataPtr.StraxPart + localStraxPart;
                          ReinsurDataPtr.AgentSum := ReinsurDataPtr.AgentSum + AgentSumma;
                          ReinsurDataPtr.FondSum := ReinsurDataPtr.FondSum + localTaxSumm;

                          //РУБЛЁВАЯ ЧАСТЬ
                          //Продажа идёт в день прихода денег
                          localBruttoPremium := SellValuta * GetRateOfCurrency(ListOfCurr[iter_i], SellDate);
                          localBasePremium := localBruttoPremium;

                          ReinsurDataPtr := ReinsurDataBRBPtr;
                          ReinsurDataPtr.BruttoPremium := ReinsurDataPtr.BruttoPremium + localBruttoPremium;
                          ReinsurDataPtr.BasePremium := ReinsurDataPtr.BasePremium + localBasePremium;

                          if RepDate < FieldByName('DateFrom').AsDateTime then
                              localNZPSumm := localBasePremium
                          else
                          if RepDate <= FieldByName('DateTo').AsDateTime then begin
                              if DurationIns <= 30 then
                                  localNZPSumm := localBasePremium / 2
                              else
                                  localNZPSumm := localBasePremium * DurationsToRepDate / DurationIns;
                          end
                          else begin
                              localNZPSumm := 0;
                              if (DurationIns <= 30) AND (Repdate - FieldByName('PayDate').AsDateTime <= 30) then
                                    localNZPSumm := localBasePremium / 2.
                          end;

                          if IsRNPU then
                              localNZPSumm := localBasePremium / 10;
                          ReinsurDataPtr.ReserveNZPSumm := ReinsurDataPtr.ReserveNZPSumm + localNZPSumm;
                          ReinsurDataPtr.StraxPart := ReinsurDataPtr.StraxPart + localNZPSumm;
                      end
                      else begin
                          ReinsurDataBRBPtr := GetStructPtr(unReinsurData, 'BRB');
                          ReinsurDataPtr := ReinsurDataBRBPtr;

                          RestPartSumm := ListOfSumm[iter_i] * OutPercent / 100 * (1 - RestraxComis / 100);

                          SellValuta := 0;
                          localBruttoPremium := (ListOfSumm[iter_i] - SellValuta);
                          AgentSumma := AgKf * ListOfSumm[iter_i] * FieldByName('AgPercent').AsFloat / 100;
                          if WorkSQL.FieldByName('Uridich').AsString = 'Y' then
                              AgentSumma := AgentSumma * (1 + UridichTax / 100)
                          else
                              AgentSumma := AgentSumma * (1 + FizichTax / 100);
                          localTaxSumm := TaxKf * ListOfSumm[iter_i] * GetOneTax(List, WorkSQL.FieldByName('PayDate').AsDateTime) / 100;

                          localBasePremium := localBruttoPremium - AgentSumma - localTaxSumm;

                          CalcRNP(localNZPSumm,
                                  localRestraxPart,
                                  localStraxPart,

                                  localBasePremium, SellValuta,
                                  RestPartSumm,

                                  RepDate,
                                  FieldByName('PayDate').AsDateTime,
                                  FieldByName('DateFrom').AsDateTime,
                                  FieldByName('DateTo').AsDateTime,
                                  SendDate, DurationIns, DurationsToRepDate);

                          if IsRNPU then begin
                              localNZPSumm := localBasePremium / 10;
                              KfRest := RestPartSumm / (localBasePremium + SellValuta);
                              localRestraxPart := localNZPSumm * KfRest;
                              localStraxPart := localNZPSumm * (1 - KfRest);
                          end;
                          ReinsurDataPtr.BruttoPremium := ReinsurDataPtr.BruttoPremium + localBruttoPremium;
                          ReinsurDataPtr.BasePremium := ReinsurDataPtr.BasePremium + localBasePremium;
                          ReinsurDataPtr.ReserveNZPSumm := ReinsurDataPtr.ReserveNZPSumm + localNZPSumm;
                          ReinsurDataPtr.RestraxPart := ReinsurDataPtr.RestraxPart + localRestraxPart;
                          ReinsurDataPtr.StraxPart := ReinsurDataPtr.StraxPart + localStraxPart;
                          ReinsurDataPtr.AgentSum := ReinsurDataPtr.AgentSum + AgentSumma;
                          ReinsurDataPtr.FondSum := ReinsurDataPtr.FondSum + localTaxSumm;
                      end;
                  end; //cycle iter_i

                  for iter_i := 1 to 3 do begin
                      if unReinsurData[iter_i].Curr = '' then break;
                      New(ReinsurDataPtr);
                      ReinsurDataPtr^ := unReinsurData[iter_i];  
                      ReinsurDataPtr.Seria := FieldByName('Seria').AsString;
                      ReinsurDataPtr.Number := FieldByName('Number').AsFloat;
                      ReinsurDataPtr.DateFrom := FieldByName('DateFrom').AsDateTime;
                      ReinsurDataPtr.DateTo := FieldByName('DateTo').AsDateTime;
                      ReinsurDataPtr.Insurer := FieldByName('Insurer2').AsString;
                      ReinsurDataPtr.InsObject := FieldByName('Insurer').AsString;
                      ReinsurDataPtr.AllSumma := FieldByName('AllSumma').AsFloat;
                      ReinsurDataPtr.AllSummaCurr := FieldByName('AllSummaCurr').AsString;
                      ReinsurList.Add(ReinsurDataPtr);
                  end;
                end; //With
             end;

             WorkSQL.Next;
         end;
end;

procedure TInsLeben2.N16Click(Sender: TObject);
begin
     if KorrectPoland = nil then begin
         Application.CreateForm(TKorrectPoland, KorrectPoland);
     end;

     KorrectPoland.DbName := Leben2Database.DatabaseName;
     KorrectPoland.TableName := 'PRIVATE2';
     KorrectPoland.AgentCodeFld := 'Agent_Code';
     KorrectPoland.AgentPercentFld := 'AgPercent';
     KorrectPoland.UridichFizichFld := 'URIDICH';
     KorrectPoland.NULLFld := 'WriteDate';
     KorrectPoland.IsUridichFldStr := true;
     KorrectPoland.ShowModal;

     MainQuery.Refresh
end;

procedure TInsLeben2.MainQueryUridichGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
     if Sender.AsString = 'Y' then Text := 'Юридическое'
     else if Sender.AsString = 'N' then Text := 'Физическое'
     else Text := ''
end;

procedure TInsLeben2.Coris1Click(Sender: TObject);
begin
    if LebenReport <> nil then exit;
    if MainQuery.RecordCount = 0 then exit;

    if LebenReport = nil then begin
        Application.CreateForm(TLebenReport, LebenReport);
    end;
    LebenReport.frReportCoris.Dictionary.Variables['ITOGO'] := '''ИТОГО: ' + Memo.Lines[0] + '''';
    LebenReport.Show;
    LebenReport.BringToFront;

    Leben2Grid.DataSource := nil;
    MainQuery.First;
    LebenReport.frReportCoris.ShowReport;
    Leben2Grid.DataSource := DataSourceLeben;
end;

procedure TInsLeben2.N17Click(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then
         InitTimer.Enabled := true;
end;

procedure TInsLeben2.CorisExcelClick(Sender: TObject);
var
    MSEXCEL : Variant;
    nLine : integer;
    NMB : Variant;
    A1 : string;
begin
    if MainQuery.RecordCount = 0 then exit;
    try
        Screen.Cursor := crHourGlass;
        if not InitExcel(MSEXCEL) then exit;
        MSEXCEL.Visible := true;
        MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\CORIS.XLS');
        A1 := MSEXCEL.ActiveSheet.Range['A1', 'A1'].Value;
        MainQuery.First;
        nLine := 7;
        Leben2Grid.DataSource := nil;
        while not MainQuery.EOF do begin
            MSEXCEL.ActiveSheet.Range['A' + IntToStr(nLine), 'A' + IntToStr(nLine)].Value := MainQueryInsurer.AsString;
            MSEXCEL.ActiveSheet.Range['B' + IntToStr(nLine), 'B' + IntToStr(nLine)].Value := MainQuerySeria.AsString + '/' + MainQueryNumber.AsString;
            MSEXCEL.ActiveSheet.Range['C' + IntToStr(nLine), 'C' + IntToStr(nLine)].Value := MainQueryPeriod.AsString;
            NMB := MainQueryAllSumma.AsFloat;
            MSEXCEL.ActiveSheet.Range['D' + IntToStr(nLine), 'D' + IntToStr(nLine)].Value := NMB;
            NMB := MainQueryPayment.AsFloat;
            MSEXCEL.ActiveSheet.Range['E' + IntToStr(nLine), 'E' + IntToStr(nLine)].Value := NMB;
            MainQuery.Next;
            Inc(nLine);
            if (nLine MOD 2) = 1 then
                MSEXCEL.ActiveSheet.Range['A1', 'A1'].Value := ''
            else
                MSEXCEL.ActiveSheet.Range['A1', 'A1'].Value := 'ЖДИТЕ...'
        end;
        MSEXCEL.ActiveSheet.Range['A' + IntToStr(nLine), 'A' + IntToStr(nLine)].Value := 'ИТОГО';
        MSEXCEL.ActiveSheet.Range['E' + IntToStr(nLine), 'E' + IntToStr(nLine)].Value := '=SUM(E7:E' + IntToStr(nLine - 1) + ')';
        MSEXCEL.ActiveSheet.Range['A1', 'A1'].Value := A1;
        DeleteFile(GetAppDir + 'DOCS\CORIS.XLS');
        MSEXCEL.ActiveWorkbook.SaveAs(GetAppDir + 'DOCS\CORIS.XLS', 1);
    except
        on E : Exception do
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
    end;

    MainQuery.First;
    Leben2Grid.DataSource := DataSourceLeben;
    Screen.Cursor := crDefault;
end;

procedure TInsLeben2.Leben2GridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
     if MainQuerySt.AsString = '1' then
         Background := $A0A0FF;
     if MainQuerySt.AsString = '2' then
         Background := clLime
end;

procedure TInsLeben2.MagazineClick(Sender: TObject);
var
     Line : integer;
     LineStr : string;
     TempStr : string;
     OutFileName : string;
     ASH : VARIANT;
     LastLine : integer;
     ArrayData : VARIANT;
     IsNotReinsurPolis : boolean;
     OutPercent : double;
     OutKomis : double;
     SendDate : TDateTime;
     RestDok : string;
     StartDate, EndDate : TDateTime;
     Y, M, D : WORD;
     MonthName : string;
     AgKf : double;
begin
     if MagazineLeben.ShowModal <> mrOk then exit;

     if not InitExcel(MSEXCEL) then exit;
     MSEXCEL.Visible := true;
     MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\ЖУРНАЛ УЧЁТА ЗАКЛЮЧЁННЫХ ДОГОВОРОВ.XLS');
     ASH := MSEXCEL.ActiveSheet;
     ASH.Range['C4','C4'].Value := 'Страхование жизни';
     try
           DecodeDate(MagazineLeben.RepData.Date, Y, M, D);
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

           while(StartDate < TRUNC(MagazineLeben.RepData.Date)) do begin
               LineStr := InttoStr(Line);
               ASH.Range['A' + LineStr, 'A' + LineStr].Value := MonthName;
               Inc(Line);
               WorkSQL.SQL.Clear;
               WorkSQL.SQL.Add('SELECT T.* FROM PRIVATE2 T WHERE WRITEDATE IS NOT NULL AND PAYMENTDATE BETWEEN :D1 AND :D2');
               WorkSQL.SQL.Add('UNION');
               WorkSQL.SQL.Add('SELECT P.* FROM PRIVATE2 P,MOVETORESTRAX M WHERE P.WRITEDATE IS NOT NULL AND P.PAYMENTDATE BETWEEN M.PERIODFROM AND M.PERIODTO AND M.PAYDATE BETWEEN :D1 AND :D2 ORDER BY SERIA,NUMBER');
               WorkSQL.ParamByName('D1').AsdateTime := StartDate;
               WorkSQL.ParamByName('D2').AsdateTime := EndDate;
               WorkSQL.Open;


           ArrayData := VarArrayCreate([1, 1, 1, 27], varVariant);
           while not WorkSQL.EOF do begin
                 //Не испорчен и не дубликат
                 with WorkSQL do begin
                       ArrayData[1, 1] := FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                       ArrayData[1, 2] := FieldByName('BEGIN').AsString;
                       ArrayData[1, 3] := FieldByName('INSURER2').AsString;
                       ArrayData[1, 4] := FieldByName('INSURER').AsString;
                       ArrayData[1, 5] := FieldByName('AllSumma').AsFloat;
                       ArrayData[1, 6] := 'USD';
                       ArrayData[1, 7] := FieldByName('BEGIN').AsString;
                       ArrayData[1, 8] := FieldByName('END').AsString;

                       TempStr := FieldByName('PAYMENTDATE').AsString;
                       //if FieldByName('Draftnumber').AsFloat > 0.5 then
                       //    TempStr := TempStr + ', ' + FieldByName('Draftnumber').AsString;

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

                               if FieldByName('Uridich').AsString = 'Y' then
                                  AgKf := 1 + MagazineLeben.UridichTax.Value / 100
                               else
                                  AgKf := 1 + MagazineLeben.FizichTax.Value / 100;

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

                       IsNotReinsurPolis := true;
                       OutPercent := 0;
                       if MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1] <> '' then begin
                           OutPercent := StrToInt(MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1]);
                           OutKomis := StrtoInt(MagazineLeben.StringGrid.Cells[2, FieldByName('Restrax').AsInteger + 1]);
                       end;

                       if OutPercent > 0 then begin
                             IsNotReinsurPolis := (WorkSQL.FieldByName('End').AsDateTime - WorkSQL.FieldByName('Begin').AsDateTime + 1) <= 30;
                             if not IsNotReinsurPolis then begin
                                 //Проверить, отдали в перестрахование или нет
                                 IsNotReinsurPolis := not DecideIsReinsur(MagazineLeben.RepData.Date, WorkSQL.FieldByName('PaymentDate').AsDateTime, SendDate, RestDok);
                             end;
                       end;

                       ArrayData[1, 19] := '';
                       ArrayData[1, 20] := '';
                       ArrayData[1, 21] := '';
                       ArrayData[1, 22] := '';
                       ArrayData[1, 23] := '';
                       ArrayData[1, 24] := '';
                       ArrayData[1, 25] := '';
                       ArrayData[1, 26] := '';
                       ArrayData[1, 27] := '';
                       if not IsNotReinsurPolis then begin
                           ArrayData[1, 19] := RestDok;
                           ArrayData[1, 20] := MagazineLeben.StringGrid.Cells[0, FieldByName('Restrax').AsInteger + 1];

                           ArrayData[1, 21] := FieldByName('Sum1Curr').AsString;
                           ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                           ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                           ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                           ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;

                           ArrayData[1, 26] := FieldByName('AllSumma').AsFloat * OutPercent / 100;
                           ArrayData[1, 27] := OutPercent;
                       end;

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

                     if not IsNotReinsurPolis then begin
                         ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                         ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                         ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                         ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;
                     end;
                     ASH.Range['A' + InttoStr(Line), 'AA' + InttoStr(Line)] := ArrayData;
                     Inc(Line);
                 end;

                 WorkSQL.Next;
           end;

           Inc(Line);
           //LineStr := IntToStr(Line);
           //ASH.Range['L' + LineStr, 'L' + LineStr].Value := '=SUM(L22:L' + IntToStr(LastLine) + ')';
           //ASH.Range['N' + LineStr, 'N' + LineStr].Value := '=SUM(N22:N' + IntToStr(LastLine) + ')';
           //ASH.Range['P' + LineStr, 'P' + LineStr].Value := '=SUM(P22:P' + IntToStr(LastLine) + ')';
           //ASH.Range['R' + LineStr, 'R' + LineStr].Value := '=SUM(R22:R' + IntToStr(LastLine) + ')';

           VarClear(ArrayData);



               StartDate := EncodeDate(Y, M, D);
               MonthName := FormatDateTime('MMMM YYYY', StartDate);
               IncAMonth(Y, M, D, 1);
               EndDate := EncodeDate(Y, M, D) - 1;
           end;
{
           WorkSQL.SQL.Clear;
           WorkSQL.SQL.Add('SELECT * FROM PRIVATE2 T WHERE WRITEDATE IS NOT NULL AND (T."BEGIN" <= :D1 AND T."END" >= :D1 OR (T."END" - T."BEGIN" + 1) <= 30 AND PAYMENTDATE BETWEEN (:D2) AND :D1) ORDER BY SERIA,NUMBER');
           WorkSQL.ParamByName('D1').AsdateTime := MagazineLeben.RepData.Date;
           WorkSQL.ParamByName('D2').AsdateTime := IncMonth(MagazineLeben.RepData.Date, -1);
           WorkSQL.Open;

           //MainGrid.DataSource := nil;
           //MainQuery.First;
           Line := 22;
           ArrayData := VarArrayCreate([1, 1, 1, 27], varVariant);
           while not WorkSQL.EOF do begin
                 //Не испорчен и не дубликат
                 with WorkSQL do begin
                       ArrayData[1, 1] := FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                       ArrayData[1, 2] := FieldByName('BEGIN').AsString;
                       ArrayData[1, 3] := FieldByName('INSURER2').AsString;
                       ArrayData[1, 4] := FieldByName('INSURER').AsString;
                       ArrayData[1, 5] := FieldByName('AllSumma').AsFloat;
                       ArrayData[1, 6] := 'USD';
                       ArrayData[1, 7] := FieldByName('BEGIN').AsString;
                       ArrayData[1, 8] := FieldByName('END').AsString;

                       TempStr := FieldByName('PAYMENTDATE').AsString;
                       //if FieldByName('Draftnumber').AsFloat > 0.5 then
                       //    TempStr := TempStr + ', ' + FieldByName('Draftnumber').AsString;
                       ArrayData[1, 9] := TempStr;
                       ArrayData[1, 10] := FieldByName('Sum1Curr').AsString;
                       ArrayData[1, 11] := FieldByName('Sum1').AsFloat;
                       if FieldByName('Sum1Curr').AsString = 'BRB' then
                           ArrayData[1, 12] := FieldByName('Sum1').AsFloat
                       else
                           ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Sum1Curr').AsString, FieldByName('PAYMENTDATE').AsDateTime) * FieldByName('Sum1').AsFloat;

                       ArrayData[1, 13] := ArrayData[1, 11];
                       ArrayData[1, 14] := ArrayData[1, 12];
                       ArrayData[1, 15] := double(ArrayData[1, 11]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 16] := double(ArrayData[1, 12]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 17] := double(ArrayData[1, 13]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 18] := double(ArrayData[1, 14]) * FieldByName('AgPercent').AsFloat / 100;

                       IsNotReinsurPolis := true;
                       OutPercent := 0;
                       if MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1] <> '' then begin
                           OutPercent := StrToInt(MagazineLeben.StringGrid.Cells[1, FieldByName('Restrax').AsInteger + 1]);
                           OutKomis := StrtoInt(MagazineLeben.StringGrid.Cells[2, FieldByName('Restrax').AsInteger + 1]);
                       end;

                       if OutPercent > 0 then begin
                             IsNotReinsurPolis := (WorkSQL.FieldByName('End').AsDateTime - WorkSQL.FieldByName('Begin').AsDateTime + 1) <= 30;
                             if not IsNotReinsurPolis then begin
                                 //Проверить, отдали в перестрахование или нет
                                 IsNotReinsurPolis := not DecideIsReinsur(MagazineLeben.RepData.Date, WorkSQL.FieldByName('PaymentDate').AsDateTime, SendDate, RestDok);
                             end;
                       end;

                       ArrayData[1, 19] := '';
                       ArrayData[1, 20] := '';
                       ArrayData[1, 21] := '';
                       ArrayData[1, 22] := '';
                       ArrayData[1, 23] := '';
                       ArrayData[1, 24] := '';
                       ArrayData[1, 25] := '';
                       ArrayData[1, 26] := '';
                       ArrayData[1, 27] := '';
                       if not IsNotReinsurPolis then begin
                           ArrayData[1, 19] := RestDok;
                           ArrayData[1, 20] := MagazineLeben.StringGrid.Cells[0, FieldByName('Restrax').AsInteger + 1];

                           ArrayData[1, 21] := FieldByName('Sum1Curr').AsString;
                           ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                           ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                           ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                           ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;

                           ArrayData[1, 26] := FieldByName('AllSumma').AsFloat * OutPercent / 100;
                           ArrayData[1, 27] := OutPercent;
                       end;

                       LastLine := Line;
                 end;
                 ASH.Range['A' + InttoStr(Line), 'AA' + InttoStr(Line)] := ArrayData;
                 Inc(Line);

                 with WorkSQL do
                 if FieldByName('Sum2').AsFloat > 0.01 then begin
                     ArrayData[1, 10] := FieldByName('Sum2Curr').AsString;
                     ArrayData[1, 11] := FieldByName('Sum2').AsFloat;
                     if FieldByName('Sum2Curr').AsString = 'BRB' then
                         ArrayData[1, 12] := FieldByName('Sum2').AsFloat
                     else
                         ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Sum2Curr').AsString, FieldByName('PAYMENTDATE').AsDateTime) * FieldByName('Sum2').AsFloat;
                     ArrayData[1, 13] := ArrayData[1, 11];
                     ArrayData[1, 14] := ArrayData[1, 12];

                     ArrayData[1, 22] := double(ArrayData[1, 13]) * OutPercent / 100;
                     ArrayData[1, 23] := double(ArrayData[1, 22]) * GetRateOfCurrency(FieldByName('Sum1Curr').AsString, SendDate);

                     ArrayData[1, 24] := double(ArrayData[1, 22]) * OutKomis / 100;
                     ArrayData[1, 25] := double(ArrayData[1, 23]) * OutKomis / 100;

                     ASH.Range['A' + InttoStr(Line), 'AA' + InttoStr(Line)] := ArrayData;
                     Inc(Line);
                 end;

                 WorkSQL.Next;
           end;

           Inc(Line);
           LineStr := IntToStr(Line);
           ASH.Range['L' + LineStr, 'L' + LineStr].Value := '=SUM(L22:L' + IntToStr(LastLine) + ')';
           ASH.Range['N' + LineStr, 'N' + LineStr].Value := '=SUM(N22:N' + IntToStr(LastLine) + ')';
           ASH.Range['P' + LineStr, 'P' + LineStr].Value := '=SUM(P22:P' + IntToStr(LastLine) + ')';
           ASH.Range['R' + LineStr, 'R' + LineStr].Value := '=SUM(R22:R' + IntToStr(LastLine) + ')';

           VarClear(ArrayData);
}
           OutFileName := GetAppDir + 'DOCS\_ЖУРНАЛ ЖИЗНЬ УЗД.XLS';
           DeleteFile(OutFileName);
           MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
     except
           on E : Exception do
              MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
end;

procedure TInsLeben2.UbMnuClick(Sender: TObject);
begin
    UbPanel.Visible := not UbPanel.Visible;
    Splitter.Visible := UbPanel.Visible;
    UbMnu.Checked := UbPanel.Visible;
end;

procedure TInsLeben2.MainQueryAfterScroll(DataSet: TDataSet);
begin
    Seria.Text := MainQuerySeria.AsString;
    Number.Text := MainQueryNumber.AsString;
    //UbTable.Close;
    UbTable.Filter := 'SER=''' + MainQuerySeria.AsString + ''' AND NMB=' + MainQueryNumber.AsString;
    UbTable.Open;
end;

procedure TInsLeben2.AddBtnClick(Sender: TObject);
begin
    if (not UbTable.Active) then exit;
    UbitokFrm.DataQuery := nil;
    UbitokFrm.SaveQuery := UbTable;
    UbTable.Append;
    Leben2Grid.Options := Leben2Grid.Options + [dgRowSelect];
    UbitokFrm.MinDate := MainQueryBegin.AsDateTime;
    UbitokFrm.MaxDate := MainQueryEnd.AsDateTime;
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
    Leben2Grid.Options := Leben2Grid.Options - [dgRowSelect];
end;

procedure TInsLeben2.DelBtnClick(Sender: TObject);
begin
    if (not UbTable.Active) OR (UbTable.RecordCount = 0) then exit;
    if MessageDlg('Удалить убыток?', mtInformation, [mbYes, mbNo], 0) = mrYes then
      UbTable.Delete;
end;

procedure TInsLeben2.EditBtnClick(Sender: TObject);
begin
    if (not UbTable.Active) OR (UbTable.RecordCount = 0) then exit;
    UbitokFrm.DataQuery := UbTable;
    UbTable.Edit;
    UbitokFrm.SaveQuery := UbTable;
    UbitokFrm.MinDate := MainQueryBegin.AsDateTime;
    UbitokFrm.MaxDate := MainQueryEnd.AsDateTime;
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
        UbTable.Cancel;
end;

procedure TInsLeben2.SpeedButton1Click(Sender: TObject);
begin
    if not MainQuery.Locate('SERIA;NUMBER', vararrayof([Seria.Text,Number.Text]), []) then begin
        MessageDlg('Полис не найден', mtInformation, [mbOk], 0);
    end;    
end;

procedure TInsLeben2.N1Click(Sender: TObject);
var
    s : string;
begin
      try
          Screen.Cursor := crHourglass;
          s := 'Отчёт создаётся на основании фильтра';
          with WorkSQL, WorkSQL.SQL do begin
              //if RepDtFrom.Checked OR RepDtTo.Checked then begin
                  s := 'ОТЧЁТНЫЙ ПЕРИОД (ПО ДАТЕ ПЛАТЕЖА)';
                  if LebenFilter.DateFrom.Checked then
                      s := s + ' C ' + DateToStr(LebenFilter.DateFrom.Date);
                  if LebenFilter.DateTo.Checked then
                      s := s + ' ПО ' + DateToStr(LebenFilter.DateTo.Date);
                  if LebenFilter.DateFrom.Checked and LebenFilter.DateTo.Checked then
                      s := s + #13#10' за ' + IntToStr(Round(LebenFilter.DateTo.Date-LebenFilter.DateFrom.Date+1)) + ' дней'
                  else
                      s := s + #13#10'НЕ УКАЗАН';

                  //StatisticTxt.Lines.Add(s);
                  //StatisticTxt.Lines.Add('');
                  s := s + #13#10;
              //end;

              Close;
              Clear;
              Add('SELECT SUM(TARIF) AS S');
              Add('FROM PRIVATE2');
              Add('WHERE TARIF>0 AND ' + GetFilter());
              Open;
              s := s + #13#10'Тариф';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('S').AsString;
                  Next;
              end;

              Close;
              Clear;
              Add('SELECT SUM(SUM1) AS S, SUM1CURR AS CUR');
              Add('FROM PRIVATE2');
              Add('WHERE SUM1CURR<>'''' AND ' + GetFilter());
              Add('GROUP BY SUM1CURR');
              Open;
              s := s + #13#10;
              s := s + #13#10'Первая часть';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString;
                  Next;
              end;

              Close;
              Clear;
              Add('SELECT SUM(SUM2) AS S, SUM2CURR AS CUR');
              Add('FROM PRIVATE2');
              Add('WHERE SUM2CURR<>'''' AND ' + GetFilter());
              Add('GROUP BY SUM2CURR');
              Open;
              s := s + #13#10;
              s := s + #13#10'Вторая часть';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString;
                  Next;
              end;
              Close;

              Clear;
              Add('SELECT COUNT(*) AS C');
              Add('FROM PRIVATE2');
              Add('WHERE ST=''1'' AND ' + GetFilter());
              Open;
              s := s + #13#10;
              s := s + #13#10'Испорчено';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('C').AsString;
                  Next;
              end;
              Close;

              Clear;
              Add('SELECT COUNT(*) AS C');
              Add('FROM PRIVATE2');
              Add('WHERE ST=''2'' AND ' + GetFilter());
              Open;
              s := s + #13#10;
              s := s + #13#10'Расторгнуто';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('C').AsString;
                  Next;
              end;
              Close;

              Close;
              Clear;
              Add('SELECT SUM(RETSUM) AS S, RETCUR AS CUR');
              Add('FROM PRIVATE2');
              Add('WHERE ST=''2'' AND ' + GetFilter());
              Add('GROUP BY RETCUR');
              Open;
              s := s + #13#10;
              s := s + #13#10'Выплаты по расторжению';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('S').AsString + ' ' + FieldByName('CUR').AsString;
                  Next;
              end;
              Close;

              {Clear;
              Add('SELECT COUNT(*) AS C');
              Add('FROM PRIVATE2');
              Add('WHERE STATE=''3'' AND ' + GetFilter());
              Open;
              s := s + #13#10;
              s := s + #13#10 + 'Утеряных полисов';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('C').AsString;
                  Next;
              end;
              Close;
              Clear;
              Add('SELECT COUNT(*) AS C');
              Add('FROM PRIVATE2');
//                Add('WHERE STATE=''1'' AND ' + GetFilter());
              Add('WHERE PNUMBER>0 AND ' + GetFilter());
              Open;
              s := s + #13#10;
              s := s + #13#10 + 'Дубликатов';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('C').AsString;
                  Next;
              end;
              Close;
              Clear;
              Add('SELECT COUNT(*) AS C');
              Add('FROM PRIVATE2');
              Add('WHERE TEXT<>'''' AND ' + GetFilter());
              Open;
              s := s + #13#10;
              s := s + #13#10 + 'Сцепок';
              while not Eof do begin
                  s := s + #13#10 + FieldByName('C').AsString;
                  Next;
              end;
              }
              Clipboard.AsText := s;
              MessageDlg(s, mtInformation, [mbOk], 0);
          end;
      except
        on E : Exception do begin
          MessageDlg(e.Message , mtInformation, [mbOk], 0);
        end;
      end;
    Screen.Cursor := crDefault;
end;

function TInsLeben2.GetFilter() : string;
begin
    GetFilter := LebenFilter.getFilterStr()
end;

procedure TInsLeben2.InputSumClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT SUM1,SUM1CURR,PAYMENTDATE,NUMBER,PAYMENTDATE FROM PRIVATE2 WHERE PAYMENTDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT SUM2,SUM1CURR,PAYMENTDATE,NUMBER,PAYMENTDATE FROM PRIVATE2 WHERE PAYMENTDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkQuery);
end;

procedure TInsLeben2.N8Click(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    RepParamsAgnt.filtParam := 'LEBEN2PCNT > 0';
    if RepParamsAgnt.ShowModal <> mrOk then exit;

    DtFrom := min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    DtTo := max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date);
    sqlText := 'SELECT SUM1,SUM1CURR,PAYMENTDATE,NUMBER,AGPERCENT FROM PRIVATE2 WHERE AGENT_CODE=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PAYMENTDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT SUM2,SUM1CURR,PAYMENTDATE,NUMBER,AGPERCENT FROM PRIVATE2 WHERE AGENT_CODE=''' + RepParamsAgnt.agCodeList[RepParamsAgnt.listAgnt.ItemIndex] + ''' AND PAYMENTDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';

    BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Text, WorkSQL);
end;

function DecodeAssist(s : string) : string;
var
     INI : TIniFile;
begin
     INI := TIniFile.Create('blank.ini');
     if s = '' then s := '0';
     DecodeAssist := INI.ReadString('PRIVATE2', 'Assist' + s, '');
     INI.Free;
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
        if(StrToInt(y) < 99) then
            y := '19' + IntToStr(StrToInt(y));
        DecodeTxtDate := DecodeGlDate(EncodeDate(StrToInt(y), StrToInt(m), StrToInt(d)));
    except
    end;
end;

function DecodeNB(s : string) : string;
begin
    DecodeNB := '?';
    if s = '1' then DecodeNB := '1';
    if s = '0' then DecodeNB := '2';
end;

function DecodeFUYN(s : string; pcnt : double) : string;
begin
    DecodeFUYN := '?';
    if pcnt = 0 then DecodeFUYN := '3'
    else
    if s = 'N' then DecodeFUYN := '1'
    else
    if s = 'Y' then DecodeFUYN := '6';
end;

function DecodeAGNameL(s : string) : string;
begin
    if lastAgCode2 = s then begin
        DecodeAGNameL := lastAgName2;
        exit;
    end;
    InsLeben2.WorkQuery.Close;
    InsLeben2.WorkQuery.SQL.Clear;
    InsLeben2.WorkQuery.SQL.Add('SELECT NAME FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    InsLeben2.WorkQuery.Open;
    if InsLeben2.WorkQuery.Eof then lastAgName2 := s
    else lastAgName2 := InsLeben2.WorkQuery.FieldByName('NAME').AsString;
    DecodeAGNameL := lastAgName2;
    lastAgCode2 := s;
end;

procedure DecodeKS(ks, t : string; var List : TStringList);
const
   decoder1_src : array[1..19] of string = ('0101', '0201', '0301', '0401', '0501',
   '0601', '0602', '0630', '0604', '0605', '0606', '0607', '0608',
   '0701', '0801', '0901', '1001', '1101', '1201');
   decoder1_dest : array[1..19] of string = ('0101', '0201', '0202', '0300', '0400',
   '0501', '0502', '0503', '', '0600', '0700', '0800', '0900',
   '1000', '1200', '1100', '', '1300', ''
   );
   decoder2 : array[1..3] of string = ('12', '06', '13');
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
       fs := s;
       key := '?';
       if t[1] in ['Q', 'Y', 'S'] then begin
        List.Add(decoder2[StrToInt(Copy(s, 1, 2))] + '00');
        key := Format('K_YSQ_%d_%d', [StrToInt(Copy(s, 1, 2)), StrToInt(Copy(s, 3, 2))]);
        f := true;
       end
       else begin
        f := false;
        for x := 1 to 19 do
            if s = decoder1_src[x] then begin
                if decoder1_dest[x] = '' then begin
                    exit;
                    //raise Exception.Create('Не найден коэффициент ' + fs + ' ' + t + ' в ГЛОБО');
                end;
                List.Add(decoder1_dest[x]);
                key := Format('K%d_%d', [StrToInt(Copy(s, 1, 2)), StrToInt(Copy(s, 3, 2))]);
                f := true;
                break;
            end;
       end;

       INI := TIniFile.Create('blank.ini');
       s := INI.ReadString('PRIVATE2', key, '');
       if s = '' then begin
            INI.Free;
            if f then List.Delete(List.Count - 1);
            exit;
       end;
       List.Add(Copy(s, 1, Pos(',', s) - 1));
   end;

   INI.Free;
end;

function DecodeAddr(s1, s2, s : string) : string;
begin
    DecodeAddr := '';
    if AnsiUpperCase(ReplaceStr(Trim(s1), ' ', '')) = AnsiUpperCase(ReplaceStr(Trim(s2), ' ', '')) then DecodeAddr := s;
end;

procedure TInsLeben2.GloboClick(Sender: TObject);
var
    F1 : TextFile;
    F2 : TextFile;
    //F3 : TextFile;
    F6 : TextFile;
//    F9 : TextFile;
    F8 : TextFile;
//    F10 : TextFile;
    F11 : TextFile;
    F12 : TextFile;
    F13 : TextFile;
    F15 : TextFile;
    All, Curr : integer;
    IsError : boolean;
    UnloadN : string;
    UnloadDt : string;
    Division : string;
    DD, MM, YY : WORD;
    errStr : string;
    List : TStringList;
    i : integer;
const
    InsTypeCode : string = '08';
begin
    List := TStringList.Create;
    errStr := '';
    UnloadN := '00';
    Division := '00';
    DecodeDate(Date, YY, MM, DD);
    UnloadDt := IntToStr(MM);
    if MM < 10 then UnloadDt := '0' + UnloadDt;
    if DD < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(DD);
    IsError := false;
    AssignFile(F1, 'c:\' + Division + 'D' + UnloadDt + UnloadN + '.txt');
    ReWrite(F1);
    AssignFile(F2, 'c:\' + Division + 'L' + UnloadDt + UnloadN + '.txt');
    ReWrite(F2);
//    AssignFile(F3, 'c:\' + Division + 'T' + UnloadDt + UnloadN + '.txt');
//    ReWrite(F3);
    AssignFile(F6, 'c:\' + Division + 'C' + UnloadDt + UnloadN + '.txt');
    ReWrite(F6);
//    AssignFile(F9, 'c:\' + Division + 'S' + UnloadDt + UnloadN + '.txt');
//    ReWrite(F9);
    AssignFile(F8, 'c:\' + Division + 'P' + UnloadDt + UnloadN + '.txt');
    ReWrite(F8);
    //AssignFile(F10, 'c:\' + Division + 'H' + UnloadDt + UnloadN + '.txt');
    //ReWrite(F10);
    AssignFile(F11, 'c:\' + Division + 'Z' + UnloadDt + UnloadN + '.txt');
    ReWrite(F11);
    AssignFile(F12, 'c:\' + Division + 'G' + UnloadDt + UnloadN + '.txt');
    ReWrite(F12);
    AssignFile(F13, 'c:\' + Division + 'B' + UnloadDt + UnloadN + '.txt');
    ReWrite(F13);
    AssignFile(F15, 'c:\' + Division + 'K' + UnloadDt + UnloadN + '.txt');
    ReWrite(F15);

    Screen.Cursor := crHourGlass;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT A.divisioncode AS DIVISION,P.* FROM PRIVATE2 P LEFT OUTER JOIN AGENT A ON (P.AGENT_CODE=A.AGENT_CODE)');
    if Length(GetFilter) <> 0 then
        WorkSQL.SQL.Add(' WHERE ' + GetFilter);
    WorkSQL.Open;
    //if MainQuery.Active then
    //    All := MainQuery.RecordCount
    //else
        All := WorkSQL.RecordCount;
    Curr := 1;
    StatusPanel.Visible := true;
    StatusMsg.Visible := false;
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
            if FieldByName('ST').AsString = '' then
                raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' состояние не определено ' + FieldByName('ST').AsString);
            if FieldByName('ST').AsInteger in [1] then begin //испорчен
                WriteLn(F13, FieldByName('DIVISION').AsString + '|25|' +
                            DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                            FieldByName('SERIA').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            FieldByName('NUMBER').AsString + '|' +
                            DecodeGlDate(FieldByName('PAYMENTDATE').AsDateTime)
                )
            end
            {else
            if FieldByName('STATE').AsInteger in [3] then begin //утерян
                WriteLn(F13, FieldByName('DIVISION').AsString + '|24|' +
                            DecodeBlank(FieldByName('SERIA').AsString) + '|' +
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
                    if (FieldByName('URIDICH').AsString = '') OR not (FieldByName('URIDICH').AsString[1] in ['Y', 'N']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип агента не определён "' + FieldByName('URIDICH').AsString + '" ');
                    if (FieldByName('PAYMENTDATE').IsNull) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата перечисления не определена');
                    if (FieldByName('WRITEDATE').AsDateTime > FieldByName('BEGIN').AsDateTime) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата выписки после начала');
                    if (FieldByName('PAYMENTDATE').AsDateTime - 45 > FieldByName('WRITEDATE').AsDateTime) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата оплаты сдишком поздно');
                    //if ((FieldByName('WRITEDATE').AsDateTime - 45) > FieldByName('PAYMENTDATE').AsDateTime) then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' дата выписки позже даты оплаты на много');
                    if (FieldByName('INSURTYPE').AsString = '') OR not (FieldByName('INSURTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' тип страхователя не определён "' + FieldByName('INSURTYPE').AsString + '" ');
                    //if DecodePeriod(FieldByName('PERIOD').AsString) = '?' then
                    //    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' период не определён "' + FieldByName('PERIOD').AsString + '"');

                    WriteLn(F1, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeFUYN(FieldByName('URIDICH').AsString, -1) + '|' +
                                DecodeGlDate(FieldByName('WRITEDATE').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('DURATION').AsString) + '|' +
                                DecodeGlDate(FieldByName('BEGIN').AsDateTime) + '|' +
                                DecodeTimeX('00:00') + '|' +
                                DecodeGlDate(FieldByName('END').AsDateTime) + '|' +
                                FieldByName('INSTYPE').AsString + '|' + //Код варианта страхования
                                FieldByName('GROUP').AsString + '|' + //Количество застрахованных
                                FieldByName('INSURER2').AsString + '|' +
                                DecodeFU12(FieldByName('INSURTYPE').AsString) + '|' +
                                '1' + {DecodeR(FieldByName('RESIDENT').AsString) +} '|' +
                                '' + '|' + //PLACE_CODE
                                DecodeAddr(FieldByName('INSURER').AsString, FieldByName('INSURER2').AsString, FieldByName('ADDRESS').AsString) + '|' +
                                '' + '|' + //Форма собственности
                                DecodePayGraphic('N') + '|' + //График оплаты
                                DecodePayType('N') + '|' +
                                DecodeAgNameL(FieldByName('AGENT_CODE').AsString) + '|' +
                                FieldByName('ALLSUMMA').AsString + '|' + //Страховая сумма по полису (на весь период)
                                FieldByName('ALLSUMMA').AsString + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('USD') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('PAYMENT').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TARIF').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('USD') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' //Для заметок
                    );
                    if FieldByName('GROUP').AsInteger <= 1 then
                    WriteLn(F2, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeCurrency('EUR') + '|' + //Код валюты
                                FieldByName('ALLSUMMA').AsString + '|' + //Страховая сумма
                                FieldByName('INSURER').AsString + '|' +
                                DecodeTxtDate(FieldByName('BIRTHDATE').AsString) + '|' + //Дата рождения
                                DecodePassport(FieldByName('PASSPORT').AsString) + '|' + //Паспорт
                                FieldByName('ADDRESS').AsString
                    );
                    WriteLn(F15, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                DecodeAssist(FieldByName('RESTRAX').AsString)
                    );
                    {WriteLn(F3, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
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
                        DecodeKS(FieldByName('KS').AsString, FieldByName('INSTYPE').AsString, List);
                        for i := 0 to (List.Count DIV 2) - 1 do
                        WriteLn(F6, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                    DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                    FieldByName('SERIA').AsString + '|' +
                                    FieldByName('NUMBER').AsString + '|' +
                                    '|' + //Номер застрахованного не указан
                                    List[i * 2] + '|' + //Код коэффициента
                                    '4|' + //Единица измерения Коэффициент
                                    List[i * 2 + 1] + '|' //Значение коэффициента
                        );
                    end;
                    //if FieldByName('TEXT').AsString = '' then begin // НЕ Прицеп
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('CASH').AsString) + '|' + //Форма оплаты
                                '' + '|' +
                                DecodeGlDate(FieldByName('WRITEDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAYMENTDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('PAYMENTDATE').AsDateTime) + '|' + //Дата  проведено по бухгалтерии

                                FieldByName('SUM1').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('SUM1CURR').AsString) + '|' + //Код валюты
                                FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                DecodeAgNameL(FieldByName('AGENT_CODE').AsString) + '|' +
                                DecodeFUYN(FieldByName('URIDICH').AsString, FieldByName('AGPERCENT').AsFloat)
                    );
                    if FieldByName('SUM2').AsFloat > 0 then
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('CASH').AsString) + '|' + //Форма оплаты
                                '' + '|' +
                                DecodeGlDate(FieldByName('WRITEDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAYMENTDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('PAYMENTDATE').AsDateTime) + '|' + //Дата  проведено по бухгалтерии
                                FieldByName('SUM2').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('SUM2CURR').AsString) + '|' + //Код валюты
                                FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                DecodeAGNameL(FieldByName('AGENT_CODE').AsString) + '|' +
                                DecodeFUYN(FieldByName('URIDICH').AsString, FieldByName('AGPERCENT').AsFloat)
                    );
                    //end;
                    if FieldByName('ST').AsString = '2' then
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
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
                                '|' + //Комиссия
                                '|'
                    );
                    if FieldByName('ST').AsString = '2' then
                    WriteLn(F11, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
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
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                FieldByName('PRCPNMB').AsString
                    );}
                //end;
                //else begin //Дубликаты
                {if FieldByName('PSERIA').AsString <> '' then begin
                    if FieldByName('DUPTYPE').AsString = '0' then
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('PSERIA').AsString + '|' +
                                FieldByName('PNUMBER').AsString + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '12'
                    );
                    if FieldByName('DUPTYPE').AsString = '1' then
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|' + InsTypeCode + '|' +
                                DecodeBlank(FieldByName('SERIA').AsString) + '|' +
                                FieldByName('PSERIA').AsString + '|' +
                                FieldByName('PNUMBER').AsString + '|' +
                                FieldByName('SERIA').AsString + '|' +
                                FieldByName('NUMBER').AsString + '|' +
                                '9'
                    );
                end; }
            end; //if
            Next;
          except
            on E : Exception do begin

                //if Pos('Access', e.Message) <> 0 then
                //exit;

                errStr := errStr + #13#10;
                if Pos('/', e.Message) = 0 then errStr := errStr + FieldByName('NUMBER').AsString + ' ';
                errStr := errStr + e.Message + ' ' + FieldByName('AGENT_CODE').AsString;
                IsError := true;
                Next;
                //break;
            end
        end; //while
    end; //width
  end;
    StatusPanel.Visible := false;
    StatusMsg.Visible := true;
    CloseFile(F1);
    CloseFile(F2);
//    CloseFile(F3);
    CloseFile(F6);
//    CloseFile(F9);
    CloseFile(F8);
    //CloseFile(F10);
    CloseFile(F11);
    CloseFile(F12);
    CloseFile(F13);
    CloseFile(F15);
    ProgressBar.Position := 0;

    Screen.Cursor := crDefault;
    List.Free;

    if not IsError then
        MessageDlg('Экспорт выполнен успешно', mtInformation, [mbOk], 0)
    else begin
        ShowInfo.Caption := 'Ошибки в данных';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
end;

end.
