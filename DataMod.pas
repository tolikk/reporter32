unit DataMod;

interface

uses
  SysUtils, Windows, Messages, Classes, Graphics, Controls, Forms,
  Dialogs, DBTables, DB, comobj, comctrls, math, RxQuery, Grids;

type
  TaxPercents = class
  public
    FizPcntBefore, FizPcntAfter, CompPcntBefore : double;
    UrPcntBefore, UrPcntAfter, CompPcntAfter : double;
    UrDt, FizDt, CompPcntDt : TDateTime;
    function GetPcnt(Dt : TDateTime; IsFiz : boolean) : double;
    function GetCompPcnt(Dt : TDateTime) : double;
  end;

  PTStringGrid = ^TStringGrid;

  TASKDateAndPercent = record
    Dt : TDate;
    Pcnt : double;
  end;

  TASKOneTaxList = array[0..10] of TASKDateAndPercent;

  ReportStructure = record
    RegDate : TDateTime;
    Number : double;

    //Form 1
    Curr : string[5];
    BruttoPremium : double;
    AgOtchislenie : double;
    OneTaxOtchislenie : double; //предупред. мероприятия
    BasePremium_nouse : double;
    FromDate, ToDate : TDateTime;
  end;

  ReportStructure_3 = array[1..3] of ReportStructure;
  PReportStructure = ^ReportStructure;

  GroupRestrax = record
      Curr : string[10];
      Seria : string;
      Number : double;
      DateFrom, DateTo : TDateTime;
      BruttoPremium : double;
      AgentSum : double;
      FondSum : double;
      BasePremium : double;
      ReserveNZPSumm : double;
      RestraxPart : double;
      StraxPart : double;
      Insurer : string;
      InsObject : string;

      AllSummaCurr : string[5];
      AllSumma : double;
  end;

  PGroupRestrax = ^GroupRestrax;
  GroupRestrax_3 = array[1..3] of GroupRestrax;

var
  BLANK_INI : string;
  ArrStr : array[0..18] of string = (
                     'one',
                     'two',
                     'three',
                     'four',
                     'five',
                     'six',
                     'seven',
                     'eight',
                     'nine',
                     'ten',
                     'eleven',
                     'twelve',
                     'thirteen',
                     'fourteen',
                     'fivteen',
                     'sixteen',
                     'seventeen',
                     'eighteen',
                     'nineteen'
                  );

  Ten : array[0..7] of string = (
                 'twenty',
                 'thirty',
                 'fourty',
                 'fivty',
                 'sixty',
                 'seventy',
                 'eighty',
                 'ninety'
               );

type
  typeTextFile = file of char;
  typeValList = array[1..10] of string[5];
  typeSumList = array[1..10] of double;

  DecodeBlankData = record
    seria : string;
    kode : string
  end;

  PDecodeBlankData = ^DecodeBlankData;

  RateStruct = record
      Dt : TDAteTime;
      USD, EUR, RUR, DM_ : double;
  end;

  TTolikIniFile = class
  private
    Sections : TStringList;
    Offsets : TList;
    F : Integer;
  public
    constructor Create(filename : string);
    destructor Free;

    function ReadLn : string;
    function FilePos : integer;
    function FileSize : integer;
    function ReadString(Section, Ident, Default: string): string;
  end;

  TCustomerData = class(TDataModule)
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  CustomerData: TCustomerData;
  RateStructList : array [1..100] of RateStruct;

const
  COMPAMYNAME : string = 'ЗАСО "ТАСК"';
  //COMPAMYNAME : string = 'ЗАСО "Промтрансинвест"';
  //COMPAMYNAME : string = 'ЗАСО "БелНефтестрах"';


var
  last_Curr : string[10] = '???';
  FillIdx : integer = 1;
  last_Date : TDateTime;
  last_SellPercent : double;

function CalcAlvenaTime(T : integer) : string;
procedure BuildReportInputSum(RepDate1, RepDate2: TDateTime; sqlText, AgentName : string; WorkSQL : TQuery);
function PrepareStr(City, Street, House : string) : string;
function CalcPeriod(D1, D2 : TDateTime) : string;
function FillOutTaxList(var List : TASKOneTaxList; strList : TStringGrid) : boolean;
function GetOneTax(List : TASKOneTaxList; Dt : TDate) : double;
function ExpStr(s : string; var IsRussian, IsEnglish, IsDigit, IsBad : boolean; CheckMinus : boolean) : string;
procedure OutTo10Form(IsRNPU : boolean; InsTypeName : string; MSEXCEL : VARIANT; RepDate : TDateTime; ReinsurData : TList; IsFast : boolean; Dir : string);
procedure FillPolisesForReserve(IsRNPU : boolean; var DataList : TList; RepDate : TDateTime; taxobj : TaxPercents; List : TASKOneTaxList; sqlText : string; WorkSQL : TQuery; ProgressBar : TProgressBar; InsStart, InsEnd, CURRCODE : string);
procedure FillStructure_3(RepDate : TDateTime; IsUridich : boolean; FizichTax, UridichTax, CompanyPercent : double; var Info : ReportStructure_3; InpData : TDateTime; Summa : double; Curr : string; AgPercent : double; OneTax : double);
procedure InitStructure_3(var Info : ReportStructure_3);
function GetSellPercent(Dt : TDateTime; Curr : string) : double;
function CalcSellDate(InpData : TDateTime; Curr : string) : TDateTime;
function GetCurrIndex(var Info : ReportStructure_3; Curr : string) : integer;
procedure RoundValue(var Val : double);
function InitExcel(var MSEXCEL : VARIANT) : boolean;
function InitWord(var MSWORD : VARIANT) : boolean;
function GetAppDir : string;
procedure DatSetToExcel(DS : TDataSet; InsertNumbers : boolean = false);
function PrepareHiLo(RptTbl : TTable; WorkSQL : TQuery; IsServer : boolean; ProgressBar : TProgressBar) : boolean;
function OpenExcel2C(var MSEXCEL : VARIANT) : boolean;
procedure RunReport(MSEXCEL : VARIANT; N : integer; WorkSQL : TQuery);
function GetRateOfCurrency(Curr : string; D : TDate; msg : string = '') : double;
function PrintRecord(var MSWORD : VARIANT; Sect : string; tmplFileName, destFileName : string; DS : TDataSet) : boolean;
function NumberToWords(Value : double; CurrCode : string; Kop : boolean) : string;
procedure SetSumma(var dest : string; v, _currname : string);
procedure ExecLittleReport(INIFileName, SECTION, Params : string; RepIndex : integer; REPQuery : TRxQuery; CurrFilter : string);
function CalcPeriod2(D1, D2 : TDateTime) : Integer;
function GetAliasPath(Alias : string) : string;

implementation

uses MF, inifiles, Excel97, CurrRatesUnit, Variants, rxstrutils, dateutil,
  WaitFormUnit;

const
    strLitValues : array [0..9] of
                   array [0..2] of string = ( ('',       '',            ''),
                                              ('одна',   'десять',      'сто'),
                                              ('две',    'двадцать',    'двести'),
                                              ('три',    'тридцать',    'триста'),
                                              ('четыре', 'сорок',       'четыреста'),
                                              ('пять',   'пятьдесят',   'пятьсот'),
                                              ('шесть',  'шестьдесят',  'шестьсот'),
                                              ('семь',   'семьдесят',   'семьсот'),
                                              ('восемь', 'восемьдесят', 'восемьсот'),
                                              ('девять', 'деваносто',   'девятьсот')
                                             );

     _10_20 : array [0..9] of string = ('десять',
                                        'одиннадцать',
                                        'двенадцать',
                                        'тринадцать',
                                        'четырнадцать',
                                        'пятнадцать',
                                        'шестнадцать',
                                        'семнадцать',
                                        'восемьнадцать',
                                        'девятнадцать'
                                        );

function GetAppDir : string;
var
     i : integer;
begin
     for i := Length(Application.ExeName) downto 1 do
         if Application.ExeName[i] = '\' then break;
     GetAppDir := Copy(Application.ExeName, 1, i);
end;

function InitWord(var MSWORD : VARIANT) : boolean;
begin
     try
        if MSWORD.Visible then Sleep(0); //Check
        InitWord := true;
        exit;
     except
        try
            MSWORD := GetActiveOLEObject('Word.Application');
            if MSWORD.Visible then Sleep(0); //Check
            InitWord := true;
            exit;
        except
            MSWORD := CreateOLEObject('Word.Application');
        end
     end;

     try
         MSWORD.Visible := true;
         if MSWORD.Visible then Sleep(0); //Check
         InitWord := true;
     except
         InitWord := false;
     end;
end;

function InitExcel(var MSEXCEL : VARIANT) : boolean;
begin
     try
        if MSEXCEL.Visible then Sleep(0); //Check
        InitExcel := true;
        exit;
     except
        try
            MSEXCEL := GetActiveOLEObject('Excel.Application');
            if MSEXCEL.Visible then Sleep(0); //Check
            InitExcel := true;
            exit;
        except
            MSEXCEL := CreateOLEObject('Excel.Application');
        end
     end;

     try
         if MSEXCEL.Visible then Sleep(0); //Check
         InitExcel := true;
     except
         InitExcel := false;
     end;
end;

procedure DatSetToExcel(DS : TDataSet; InsertNumbers : boolean = false);
var
    MSEXCEL : VARIANT;
    ASH : VARIANT;
    Col : char;
    i, j : integer;
    prefix, s : string;
    ArrayData : VARIANT;
    ColCount : integer;
    LastCol : string;
    LIMIT, LinesCounter, Blinker : integer;
    Numberer : integer;
begin
    LIMIT := 100;
    Blinker := 0;
    if not InitExcel(MSEXCEL) then exit;
    MSEXCEL.WorkBooks.Add;
    MSEXCEL.Visible := true;
    ASH := MSEXCEL.ActiveSheet;
    try
        DS.DisableControls;
        DS.First;
        ColCount := 0;
        Col := 'A';
        if InsertNumbers then Col := 'B';
        prefix := '';
        for i := 0 to DS.Fields.Count - 1 do begin
          if DS.Fields[i].Visible then begin
            ASH.Range[prefix + Col + '1', prefix + Col + '1'].Value := DS.Fields[i].DisplayLabel;
            LastCol := prefix + Col;
            ColCount := ColCount + 1;
            if Col = 'Z' then begin
                 prefix := 'A';
                 Col := 'A';
             end
             else
                 Col := Chr(Ord(Col) + 1);
          end;
        end;
        if InsertNumbers then
            ArrayData := VarArrayCreate([1, LIMIT, 1, ColCount + 1], varVariant)
        else
            ArrayData := VarArrayCreate([1, LIMIT, 1, ColCount], varVariant);

        s := ASH.Range['A1', 'A1'].Value;

        i := 2;
        LinesCounter := 0;
        Numberer := 1;
        while not DS.EOF do begin
           ColCount := 1;
           if InsertNumbers then ColCount := 2;
           Inc(LinesCounter);
           ArrayData[LinesCounter, 1] := Numberer;
           Inc(Numberer);
           for j := 0 to DS.Fields.Count - 1 do begin
             if DS.Fields[j].Visible then begin
               ArrayData[LinesCounter, ColCount] := DS.Fields[j].AsString;
               Inc(ColCount);
              end;
           end;

           Inc(Blinker);

           DS.Next;

           if (LinesCounter = LIMIT) OR (DS.Eof) then begin
               //if InsertNumbers then
               //    ASH.Range['B' + IntToStr(i - LinesCounter + 1), LastCol + IntToStr(i)].Value := ArrayData
               //else
                   ASH.Range['A' + IntToStr(i - LinesCounter + 1), LastCol + IntToStr(i)].Value := ArrayData;
               LinesCounter := 0;
               if (Blinker MOD 2) > 0 then
                   ASH.Range['A1', 'A1'].Value := 'РАБОТАЮ...'
               else
                   ASH.Range['A1', 'A1'].Value := '';
           end;

           //ASH.Range['A' + IntToStr(i), LastCol + IntToStr(i)].Value := ArrayData;
           Inc(i);
         end;
         ASH.Range['A1', prefix + Col + '1'].Font.Bold := True;
         ASH.Columns['A:' + LastCol].EntireColumn.AutoFit;

     except
       on E : Exception do begin
         MessageDlg('Ошибка передачи данных' + E.Message, mtError, [mbOk], 0);
       end;
     end;
     ASH.Range['A1', 'A1'].Value := s;
     DS.EnableControls;
end;

function GetCELL(Cell : string; dx, dy : integer) : string;
var
    i : integer;
    part1, part2 : string;
begin
    for i := 1 to Length(Cell) do
        if Cell[i] in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] then
            part2 := part2 + Cell[i]
        else
            part1 := part1 + Cell[i];

    GetCELL := Chr(Ord(part1[1]) + dx) + IntToStr(StrToInt(part2) + dy);
end;

function ExpandSQL(Section : string; X, Y : integer; SQL : string; INI : TTolikIniFile) : string;
var
    key, val : string;
    StartPos, i, counter : integer;
begin
    INI := TTolikIniFile.Create('reports.ini');
    counter := 0;
    while Pos('~', SQL) <> 0 do begin
        Inc(counter);
        if(counter>1000) then raise Exception.Create('Произошло зацикливание при разборе маски');
        StartPos := Pos('~', SQL);
        key := '';
        for i := StartPos + 1 to Length(SQL) do begin
            if SQL[i] = '~' then break;
            key := key + SQL[i];
        end;
        if SQL[i]<>'~' then raise Exception.Create('Не закрыта маска ' + Key);

        val := INI.ReadString(Section, 'Query_' + IntToStr(X) + '_' + IntToStr(Y) + '_' + key, '');
        if val = '' then val := INI.ReadString(Section, 'Query_X_' + IntToStr(Y) + '_' + key, '');
        if val = '' then val := INI.ReadString(Section, 'Query_Х_' + IntToStr(Y) + '_' + key, '');
        if val = '' then val := INI.ReadString(Section, 'Query_' + IntToStr(X) + '_X_' + key, '');
        if val = '' then val := INI.ReadString(Section, 'Query_' + IntToStr(X) + '_Х_' + key, '');
        if val = '' then val := INI.ReadString(Section, key, '');
        if val = '' then val := INI.ReadString('MASK', key, '');
        if val = '' then raise Exception.Create('Не найдена маска для ' + Key);
        SQL := Copy(SQL, 1, StartPos - 1) + val + Copy(SQL, i + 1, Length(SQL));
    end;
    INI.Free;
    ExpandSQL := SQL;
end;

procedure AddValuta(var CurrList : typeValList; var SummList : typeSumList; Sum : double; Curr : string);
var
    i : integer;
begin
    for i := 1 to 10 do begin
      if CurrList[i] = '' then break;
      if CurrList[i] = Curr then break;
    end;
    CurrList[i] := Curr;
    SummList[i] := SummList[i] + ROUND(Sum * 100) / 100;
end;

procedure RunReport(MSEXCEL : VARIANT; N : integer; WorkSQL : TQuery);
var
    INI : TTolikIniFile;
    val, SCell : string;
    SHEET : VARIANT;
    i, j, k : integer;
    SQL : string;
    defRptType, qType, strRes : string;
    Alias : string;
    fVal : double;
    part : integer;
    CurrList : typeValList;
    SummList : typeSumList;
    CashNone : integer;
    temp_s : string;
begin
    for i := 1 to 10 do begin
      CurrList[i] := '';
      SummList[i] := 0;
    end;
    
    INI := TTolikIniFile.Create('reports.ini');
    SHEET := null;
    try
        //находим шит
        val := INI.ReadString('REPORT' + IntToStr(N), 'Sheet', '');
        Alias := INI.ReadString('REPORT' + IntToStr(N), 'ALIAS', '');
        SCell := INI.ReadString('REPORT' + IntToStr(N), 'StartCell', '');
        defRptType := INI.ReadString('REPORT' + IntToStr(N), 'Type', '');
        WorkSQL.Close;
        if Alias <> '' then WorkSQL.DatabaseName := Alias;
        //DX := INI.ReadInteger('REPORT' + IntToStr(N), 'DX', 0);
        //DY := INI.ReadInteger('REPORT' + IntToStr(N), 'DY', 0);
        for i := 1 to MSEXCEL.ActiveWorkBook.Worksheets.Count do
            if MSEXCEL.ActiveWorkBook.Worksheets[i].Name = val then begin
                MSEXCEL.ActiveWorkBook.Worksheets[i].Select;
                SHEET := MSEXCEL.ActiveWorkBook.Worksheets[i];
                break;
            end;
        if VarIsNull(SHEET) then raise Exception.Create('Не найдена страница ' + val);
        if SCell = '' then raise Exception.Create('Не найдена стартовая ячейка');
        //бомбим запросы по очереди и вставляем данные

        for i := 1 to 50 do begin
           CashNone := 0;
           for j := 1 to 50 do begin
              if CashNone > 5 then break;
              SQL := INI.ReadString('REPORT' + IntToStr(N), 'Query_' + IntToStr(i-1) + '_' + IntToStr(j-1), '');
              if SQL <> '' then begin
                  CashNone := 0;
                  part := 2;
                  while true do begin
                      temp_s := INI.ReadString('REPORT' + IntToStr(N), IntToStr(i-1) + '_' + IntToStr(j-1) + '_' + IntToStr(part), '');
                      temp_s := Trim(temp_s);
                      if temp_s = '' then break;
                      if SQL[Length(SQL)] = '\' then SQL := Copy(SQL, 1, Length(SQL)-1);
                      SQL := SQL + ' ' + temp_s;
                      Inc(part);
                  end;
                  SQL := ExpandSQL('REPORT' + IntToStr(N), i - 1, j - 1, SQL, INI);
                  WorkSQL.Close;
                  WorkSQL.SQL.Clear;
                  WorkSQL.SQl.Add(SQL);
                  try
                      strRes := '0';
                      SHEET.Range[GetCELL(SCell, i-1, j-1), GetCELL(SCell, i-1, j-1)] := 'ДУМАЮ...';
                      SHEET.Range[GetCELL(SCell, i-1, j-1), GetCELL(SCell, i-1, j-1)].Interior.ColorIndex := 3;
                      WorkSQL.Open;
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Type', defRptType);
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Query_' + IntToStr(i-1) + '_X_Type', qType);
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Query_' + IntToStr(i-1) + '_Х_Type', qType);
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Query_X_' + IntToStr(j-1) + '_Type', qType);
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Query_Х_' + IntToStr(j-1) + '_Type', qType);
                      qType := INI.ReadString('REPORT' + IntToStr(N), 'Query_' + IntToStr(i-1) + '_' + IntToStr(j-1) + '_Type', qType);
                      if qType = 'COUNT' then
                          strRes := WorkSQL.Fields[0].AsString;
                      if qType = 'STRING' then
                          strRes := WorkSQL.Fields[0].AsString;
                      if qType = 'SUMMA_RATE' then
                          while not WorkSQL.EOF do begin
                              fVal := WorkSQL.Fields[0].AsFloat * GetRateOfCurrency(WorkSQL.Fields[1].AsString, WorkSQL.Fields[2].AsDateTime);
                              strRes := FloatToStr(StrToFloat(strRes) + fVal);
                              WorkSQL.Next;
                          end;
                      if qType = 'SUMMA_VAL' then begin
                          while not WorkSQL.EOF do begin
                              AddValuta(CurrList, SummList, WorkSQL.Fields[0].AsFloat, WorkSQL.Fields[1].AsString);
                              WorkSQL.Next;
                          end;
                          strRes := '';
                          for k := 1 to 10 do begin
                            if CurrList[k] <> '' then begin
                                if strRes <> '' then strRes := strRes + ' ;';
                                strRes := strRes + FloatToStr(SummList[k]) + ' ' + CurrList[k];
                            end;
                          end;
                      end;
                      if qType = 'SUMMA' then
                          while not WorkSQL.EOF do begin
                              strRes := FloatToStr(StrToFloat(strRes) + WorkSQL.Fields[0].AsFloat);
                              WorkSQL.Next;
                          end;
                      SHEET.Range[GetCELL(SCell, i-1, j-1), GetCELL(SCell, i-1, j-1)] := strRes;
                      SHEET.Range[GetCELL(SCell, i-1, j-1), GetCELL(SCell, i-1, j-1)].Interior.ColorIndex := xlNone;
                  except
                    on E : Exception do begin
                      raise Exception.Create('Ошибка запроса Query_' + IntToStr(i-1) + '_' + IntToStr(j-1) + #13 + SQL + #13 + E.Message);
                    end;
                  end;
                  WorkSQL.Close;
              end
              else
                  CashNone := CashNone + 1;
           end;
        end;
    finally
        INI.Free;
    end;
end;

function PrepareHiLo(RptTbl : TTable; WorkSQL : TQuery; IsServer : boolean; ProgressBar : TProgressBar) : boolean;
var
    n, all, i : integer;
begin
    PrepareHiLo := true;
    try
        Screen.Cursor := crHourglass;
        RptTbl.Open;
        RptTbl.EmptyTable;
        ProgressBar.Position := 0;
        with WorkSQL, WorkSQL.SQL do begin
            Clear;
            if IsServer then
                Add('SELECT SER AS SERIA, NMB AS NUMBER, AUTOTPCD AS SUPER FROM SYSADM.VS_MANDATOR WHERE STATE<>1')
            else
                Add('SELECT SERIA, NUMBER, SUPER FROM MANDATOR WHERE STATE<>1');
            Open;
            all := RecordCount;
            n := 0;
            while not Eof do begin
                RptTbl.Append;
                RptTbl.FieldByName('S').AsString := FieldByName('Seria').AsString;
                RptTbl.FieldByName('N').AsFloat := FieldByName('Number').AsInteger;
                RptTbl.FieldByName('Hi').AsFloat := (FieldByName('SUPER').AsInteger AND $00ff0000) shr 16;
                RptTbl.FieldByName('Lo').AsFloat := (FieldByName('SUPER').AsInteger AND $000000ff);
                RptTbl.Post;
                ProgressBar.Position := Round(n * 100 / all);
                Inc(n);
                Next;
            end;
        end;
    except
      on E : Exception do begin
        MessageDlg('Ошибка получения отчёта 2-С'#13 + E.Message, mtError, [mbOk], 0);
        PrepareHiLo := false;
      end;
    end;
    ProgressBar.Position := 0;
    Screen.Cursor := crDefault;
    RptTbl.Close;
    WorkSQL.Close;
end;

function OpenExcel2C(var MSEXCEL : VARIANT) : boolean;
var
    i : integer;
    IsFound : boolean;
begin
    OpenExcel2C := true;
    if not InitExcel(MSEXCEL) then begin
        MessageDlg('Excel ен найден!!!', mtInformation, [mbOk], 0);
        OpenExcel2C := false;
        exit;
    end;
    MSEXCEL.Visible := true;
    IsFound := false;
    for i := 1 to MSEXCEL.Windows.Count do
        if MSEXCEL.Windows[i].Caption = '2с.xls' then begin
            MSEXCEL.Windows[i].Activate;
            IsFound := true;
            break;
         end;
    if not IsFound then
        MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\2с.xls');
end;
                                       
function GetRateOfCurrency(Curr : string; D : TDate; msg : string = '') : double;
var
    RateTable : TTable;
    fldName : string;
    i : integer;
    StartDate : TDate;
    Val : double;
begin
    if Curr = 'BRB' then begin
        GetRateOfCurrency := 1;
        exit;
    end;

    Val := 0;
    for i := 1 to 100 do begin
        if RateStructList[i].Dt = 0 then break;
        if RateStructList[i].Dt = D then begin
            //if Curr = 'DM' then Val := RateStructList[i].DM
            //else
            if Curr = 'EUR' then Val := RateStructList[i].EUR
            else if Curr = 'RUR' then Val := RateStructList[i].RUR
            else if Curr = 'USD' then Val := RateStructList[i].USD
            else raise Exception.Create(msg + ' Не правильный тип валюты "' + Curr + '"');

            GetRateOfCurrency := Val;
            if Val <= 0.1 then
                raise Exception.Create(msg + ' Не правильный курс (<0.1) ' + FloatToStr(Val) + ' валюты ' + Curr + ' на ' + DateToStr(D) + ' ' + CurrencyRates.RateTableName.Text);
            exit;
        end;
    end;

    if CurrencyRates.RateTableName.Text = '' then begin
        CurrencyRates.ShowModal;
    end;

    if Trim(CurrencyRates.RateTableName.Text) = '' then
        MessageDlg('Не указана таблица курсов валют', mtInformation, [mbOk], 0);

    if (CurrencyRates.FieldEUR.Text = '') OR
       (CurrencyRates.FieldRUR.Text = '') OR
       (CurrencyRates.FieldUSD.Text = '')
       then
        MessageDlg('Не указано поле из таблицы курсов валют', mtInformation, [mbOk], 0);

    StartDate := D;
    RateTable := TTable.Create(Application.MainForm);
    RateTable.TableName := CurrencyRates.RateTableName.Text;
    try
        RateTable.Open;
    except
        raise Exception.Create('Невозможно открыть таблицу курсов валют "' + CurrencyRates.RateTableName.Text + '"');
    end;

    for i := 1 to 14 do begin
        if RateTable.Locate(CurrencyRates.FieldDate.Text, D, []) then break;
        D := D - 1;
    end;

    if i > 13 then begin
        RateTable.Last;
        msg := msg + ' Last Date ' + RateTable.fieldByName(CurrencyRates.FieldDate.Text).AsString + ' ';
        RateTable.Free;
        raise Exception.Create(msg + ' Нет курса валюты ' + Curr + ' на ' + DateToStr(StartDate) + ' ±2 недели. ' + CurrencyRates.RateTableName.Text);
    end;

//    if Curr = 'DM' then fldName := CurrencyRates.FieldDM.Text
    //else
    if Curr = 'EUR' then fldName := CurrencyRates.FieldEUR.Text
    else if Curr = 'RUR' then fldName := CurrencyRates.FieldRUR.Text
    else if Curr = 'USD' then fldName := CurrencyRates.FieldUSD.Text
    else raise Exception.Create(msg + ' Не правильный тип валюты "' + Curr + '"');

    RateStructList[FillIdx].Dt := StartDate;
    try
        RateStructList[FillIdx].USD := RateTable.FieldByName(CurrencyRates.FieldUSD.Text).AsFloat;
        RateStructList[FillIdx].RUR := RateTable.FieldByName(CurrencyRates.FieldRUR.Text).AsFloat;
        RateStructList[FillIdx].EUR := RateTable.FieldByName(CurrencyRates.FieldEUR.Text).AsFloat;
        //RateStructList[FillIdx].DM := RateTable.FieldByName(CurrencyRates.FieldDM.Text).AsFloat;
    except
        raise Exception.Create('Возможно определены не все имена полей валют');
    end;
    FillIdx := FillIdx + 1;
    if FillIdx > 100 then FillIdx := 1;

    GetRateOfCurrency := RateTable.FieldByName(fldName).AsFloat;

    RateTable.Free;
end;

constructor TTolikIniFile.Create(filename : string);
var
    Path : array[0..256] of char;
    s : string;
begin
    Sections := TStringList.Create;
    Offsets := TList.Create;
    GetWindowsDirectory(Path, sizeof(Path));
//{$I-}
//    AssignFile(F, String(Path) + '\' + filename);   { File selected in dialog box }
//    Reset(F);
//    if IOResult <> 0 then begin
    F := FileOpen(String(Path) + '\' + filename, fmOpenRead or fmShareDenyWrite);
    if F = -1 then begin
        GetSystemDirectory(Path, sizeof(Path));
//        AssignFile(F, String(Path) + '\' + filename);   { File selected in dialog box }
//        Reset(F);
//        if IOResult <> 0 then exit;
        F := FileOpen(String(Path) + '\' + filename, fmOpenRead or fmShareDenyWrite);
        if F = -1 then exit;
    end;
                 
    while FilePos < FileSize do begin
        s := ReadLn;
        s := Trim(s);
        if (Length(s) > 3) and (s[1] = '[') and (s[Length(s)] = ']') then begin
            Sections.Add(Copy(s, 2, Length(s) - 2));
            Offsets.Add(Pointer(FilePos));
        end;
    end;
    Offsets.Add(Pointer(FileSize));
{$I+}
end;

destructor TTolikIniFile.Free;
begin
    Sections.Free;
    Offsets.Free;
    FileClose(F);
end;

function TTolikIniFile.ReadString(Section, Ident, Default: string): string;
var
    i : integer;
    s : string;
begin
    ReadString := Default;
    for i := 0 to Sections.Count - 1 do begin
        if Section = Sections[i] then break;
    end;
                    
    if i >= Sections.Count then exit;

    FileSeek(F, Integer(Offsets[i]), 0);

    Ident := UpperCase(Ident);
    while FilePos < Integer(Offsets[i + 1]) do begin
        s := ReadLn;
        s := Trim(s);
        if (Length(s) > 0) AND (Length(s) > Length(Ident)) then begin
            if Pos('=', s) >= Length(Ident) then begin
                if(UpperCase(Trim(Copy(s, 1, Pos('=', s) - 1)))) = Ident then begin
                    ReadString := Trim(Copy(s, Pos('=', s) + 1, Length(s)));
                    exit;
                end;
            end;
        end;
    end;
end;

function TTolikIniFile.ReadLn : string;
var
    s : string;
    ch : char;
    databuffer : array[0..1023] of Char;
    i, InitPos, Readed : integer;
begin
    InitPos := FilePos;
    s := '';
    ch := #0;
    while(true) do begin
        if ch = #13 then break;
        Readed := FileRead(f, databuffer, sizeof(databuffer));
        if Readed = 0 then break;
        for i := 0 to min(Readed, SizeOf(databuffer) - 1) do begin
            ch := databuffer[i];
            Inc(InitPos);
            if ch = #13 then break;
            if ch <> #10 then
              s := s + ch;
        end;
    end;
    FileSeek(f, InitPos, 0);
    ReadLn := s;
end;

function TTolikIniFile.FilePos : integer;
begin
    FilePos := FileSeek(f, 0, 1);
end;

function TTolikIniFile.FileSize : integer;
var
    currPosition : integer;
begin
    currPosition := FileSeek(f, 0, 1);
    FileSize := FileSeek(f, 0, 2);
    FileSeek(f, currPosition, 0);
end;

function GetFields(str : string; var sList : TStringList) : integer;
var
    s : string;
    i : integer;
    flag : boolean;
    IsLiteral : boolean;
begin
    IsLiteral := false;
    flag := true;
    for i := 1 to Length(str) do begin
        if (not ((str[i] in ['+', '''']) OR ((str[i] = ' ') AND (not IsLiteral)))) AND (not flag) then begin
            if(Trim(s) <> '')then
                sList.Add(Trim(s));
            s := '';
            break;
        end;
        if (str[i] in ['+']) OR ((str[i] = ' ') AND (not IsLiteral)) then begin
            if(Trim(s) <> '')then
                sList.Add(Trim(s));
            s := '';
            flag := str[i] = '+';
            continue;
        end;
        if str[i] = '''' then
            IsLiteral := not IsLiteral;
        s := s + str[i];
        flag := true;
    end;

    GetFields := i;
end;

function PrintRecord(var MSWORD : VARIANT; Sect : string; tmplFileName, destFileName : string; DS : TDataSet) : boolean;
var
    F : TextFile;
    str : string;
    txtKey, sValue : string;
    dbFieldLen, i : integer;
    D, M, Y : WORD;
    shape : integer;
    founded : boolean;
    s : string;
    sList : TStringList;
    INI : TIniFile;
begin
    MSWORD.Documents.Open(destFileName);
    AssignFile(F, tmplFileName);
    Reset(F);
    sList := TStringList.Create;
    try
        while not Eof(F) do begin
            sValue := '';
            sList.Clear;
            ReadLn(F, str);
            dbFieldLen := GetFields(str, sList);
            //dbFieldLen := ExtractWord(1, str, [' ']);
            str := 'f ' + Copy(str, dbFieldLen, 1024);
            txtKey := ExtractWord(2, str, [' ']);
            for i := 0 to sList.Count - 1 do begin
                if sList[i][1] = '''' then
                    sValue := sValue + Copy(sList[i], 2, Length(sList[i]) - 2)
                else
                    sValue := sValue + DS.FieldByName(sList[i]).AsString;
            end;

            if ExtractWord(3, str, [' ']) = 'DAY' then begin
                DecodeDate(DS.FieldByName(sList[0]).AsDateTime, Y, M, D);
                sValue := Format('%02d', [D]);
                if Length(sValue) < 2 then sValue := '0' + sValue;
            end;
            if ExtractWord(3, str, [' ']) = 'MONTH' then begin
                DecodeDate(DS.FieldByName(sList[0]).AsDateTime, Y, M, D);
                sValue := Format('%02d', [M]);
                if Length(sValue) < 2 then sValue := '0' + sValue;
            end;
            if ExtractWord(3, str, [' ']) = 'YEAR' then begin
                DecodeDate(DS.FieldByName(sList[0]).AsDateTime, Y, M, D);
                sValue := Format('%02d', [Y]);
            end;
            if ExtractWord(3, str, [' ']) = 'DECODE' then begin
                INI := TIniFile.Create(BLANK_INI);
                sValue := INI.ReadString(Sect, sValue, sValue);
                INI.Free;
            end;
            if ExtractWord(3, str, [' ']) = 'TEXT' then begin
                if DS.FieldByName(sList[0]).AsString = '' then
                    sValue := ''
                else
                    sValue := NumberToWords(DS.FieldByName(sList[0]).AsFloat,
                                            DS.FieldByName(ExtractWord(4, str, [' '])).AsString, true);
            end;
            if ExtractWord(3, str, [' ']) = 'TEXT_ENGL' then begin
                sValue := '';
                if DS.FieldByName(sList[0]).AsString <> '' then
                    SetSumma(sValue, DS.FieldByName(sList[0]).AsString, DS.FieldByName(ExtractWord(4, str, [' '])).AsString);
            end;

            MSWORD.Selection.HomeKey;
            MSWORD.Selection.Find.ClearFormatting;
            MSWORD.Selection.Find.Text := txtKey;
            MSWORD.Selection.Find.Replacement.Text := '';
            MSWORD.Selection.Find.Forward := True;
            if MSWORD.Selection.Find.Execute then
                MSWORD.Selection := sValue
            else begin
                founded := false;
                for shape := 1 to MSWORD.ActiveDocument.Shapes.Count do begin
                    if Trim(MSWORD.ActiveDocument.Shapes.Item(shape).TextFrame.TextRange.Text) = txtKey then begin
                        MSWORD.ActiveDocument.Shapes.Item(shape).TextFrame.TextRange.Text := sValue;
                        founded := true;
                        break;
                    end;
                end;
                if not founded then begin
                    MessageDlg('Ключ ' + txtKey + ' не найден', mtInformation, [mbOk], 0);
                    break;
                end;
            end;
        end;
    except
      on E : Exception do begin
        MessageDlg(E.Message, mtInformation, [mbOk], 0);
      end;
    end;

    sList.Free;
    CloseFile(F);
end;


function GetNextDigit(V : integer) : integer;
begin
     Result := V MOD 10;
end;

function IsPodrostok(V : integer) : boolean;
begin
     V := V MOD 100;
     Result := (V > 10) AND (V < 20);
end;


function NumberToWords(Value : double; CurrCode : string; Kop : boolean) : string;
label
    Calculate;
var
    i : integer;
    intValue : int64;
    s : string;
    Counter : integer;
    CurrName, LittleCurrName : string;
    sResult : string;
begin
    intValue := TRUNC(Value);
    s := '';
    Counter := 0;

    if CurrCode = 'USD' then begin
        CurrName := 'доллар США';
        if (intValue MOD 10) = 0 then CurrName := 'долларов США';
        if (intValue MOD 10) >= 2 then CurrName := 'доллара США';
        if (intValue MOD 10) >= 5 then CurrName := 'долларов США';
        LittleCurrName := 'цент'
    end;
    if CurrCode = 'BRB' then begin
        CurrName := 'рубль';
        if (intValue MOD 10) = 0 then CurrName := 'рублей';
        if (intValue MOD 10) >= 2 then CurrName := 'рубля';
        if (intValue MOD 10) >= 5 then CurrName := 'рублей';
        LittleCurrName := 'копейка'
    end;
    if CurrCode = 'RUR' then begin
        CurrName := 'рубль';
        if (intValue MOD 10) = 0 then CurrName := 'рублей';
        if (intValue MOD 10) >= 2 then CurrName := 'рубля';
        if (intValue MOD 10) >= 5 then CurrName := 'рублей';
        LittleCurrName := 'копейка'
    end;
    if CurrCode = 'EUR' then begin
        CurrName := 'евро';
        LittleCurrName := 'евроцент'
    end;

    while true do begin
          if ((Counter MOD 3) = 0) AND (intValue > 10) then begin
            if ((intValue - TRUNC(intValue / 100) * 100) > 10) AND
               ((intValue - TRUNC(intValue / 100) * 100) < 20) then begin
               s := _10_20[(intValue - TRUNC(intValue / 100) * 100) - 10] + s;
              Inc(Counter);
              intValue := TRUNC(intValue / 10);
            end
            else goto Calculate;
          end
          else
            Calculate : s := strLitValues[intValue - TRUNC(intValue / 10) * 10, Counter MOD 3] + ' ' + s;

          Inc(Counter);
          intValue := TRUNC(intValue / 10);
          if intValue > 0 then begin
            if Counter = 3 then begin
               if IsPodrostok(intValue) OR (GetNextDigit(intValue) >= 5) OR (GetNextDigit(intValue) = 0) then
                  s := ' тысяч ' + s
               else
               if GetNextDigit(intValue) >= 2 then
                  s := ' тысячи ' + s
               else
                  s := ' тысяча ' + s;
            end;
            if Counter = 6 then begin
               if IsPodrostok(intValue) OR (GetNextDigit(intValue) >= 5) OR (GetNextDigit(intValue) = 0) then
                  s := ' миллионов ' + s
               else
               if GetNextDigit(intValue) >= 2 then
                  s := ' миллиона ' + s
               else
                  s := ' миллион ' + s;
            end;
            if Counter = 9 then begin
               if IsPodrostok(intValue) OR (GetNextDigit(intValue) >= 5) OR (GetNextDigit(intValue) = 0) then
                  s := ' миллиардов ' + s
               else
               if GetNextDigit(intValue) >= 2 then
                  s := ' миллиарда ' + s
               else
                  s := ' миллиард ' + s;
            end
          end
          else
            Break;
    end;

    intValue := ROUND((Value - TRUNC(Value)) * 100);

    if (intValue >= 0.01) AND Kop then
       sResult := s + ' ' + CurrName + ' ' + NumberToWords(intValue, '', false) + ' ' + LittleCurrName
    else begin
       sResult := s + ' ' + CurrName;
       if intValue > 0 then
          sResult := s + ' ' + CurrName + ' ' + IntToStr(intValue) + ' ' + LittleCurrName;
    end;

    while Pos('  ', sResult) <> 0 do
        sResult := Copy(sResult, 1, Pos('  ', sResult)) + Copy(sResult, Pos('  ', sResult) + 2, 1024);
    Result := sResult;
end;


procedure strcat(var s1 : string; s2 : string);
begin
    s1 := s1 + s2;
end;

procedure _1000(var dest : string; value : integer; what : string);
var
   tmp : integer;
begin
   if(value>0)then
    begin
       tmp := value DIV 100;
       if(tmp>0)then
        begin
          strcat(dest, ArrStr[tmp - 1]);
          strcat(dest, ' hundred ');
        end;
       tmp := (value MOD 100);
       if(tmp>0)then
        begin
          if(tmp < 20)then
           begin
             strcat(dest, ArrStr[tmp - 1]);
             strcat(dest, ' ');
           end
          else
           begin
             strcat(dest, Ten[tmp DIV 10 - 2]);
             strcat(dest, ' ');
             if((tmp MOD 10)>0)then
              begin
                strcat(dest, ArrStr[tmp MOD 10 - 1]);
                strcat(dest, ' ');
              end
           end
        end;

       strcat(dest, what);
       strcat(dest, ' ');
    end
end;

procedure SetSumma(var dest : string; v, _currname : string);
var
   value, cent : integer;
   currname, currname2 : string;
begin
   value := TRUNC(StrToFloat(v));
   //strcat(dest, v);
   //strcat(dest, ' ( ');
   _1000(dest, value DIV 1000000, 'million');
   _1000(dest, (value MOD 1000000) DIV 1000, 'thousand');
   _1000(dest, value MOD 1000, '');

    if _currname = 'USD' then begin
        currname := 'dollar';
        currname2 := 'cent';
    end
    else
    if _currname = 'EUR' then begin
        currname := 'euro';
        currname2 := 'cent';
    end
    else
    if _currname = 'BRB' then begin
        currname := 'rubl';
        currname2 := 'kop';
    end
    else
    if _currname = 'RUR' then begin
        currname := 'rubl';
        currname2 := 'kop';
    end
    else begin
        currname := _currname;
        currname2 := '???';
    end;

    strcat(dest, currname);
   if((value MOD 10) <> 1)then strcat(dest, 's');
   strcat(dest, ' ');
   if(StrToFloat(v) - trunc(StrToFloat(v))>0.009)then
    begin
      cent := ROUND((StrToFloat(v) * 100 - TRUNC(StrToFloat(v)) * 100));
      dest := dest + IntToStr(cent);
      strcat(dest, ' ' + currname2);
      if((cent MOD 10) <> 1)then strcat(dest, 's');
    end;
   //strcat(dest, ' )');
end;

procedure ExecLittleReport(INIFileName, SECTION, Params : string; RepIndex : integer; REPQuery : TRxQuery; CurrFilter : string);
var
    INI : TIniFile;
    i : integer;
    s, param : string;
begin
    REPQuery.Close;
    REPQuery.SQL.Clear;
    INI := TIniFile.Create(INIFileName);
    for i := 1 to 100 do begin
        s := INI.ReadString(SECTION, 'QueryText' + IntToStr(RepIndex) + '_' + IntToStr(i), '');
        if s <> '' then
            REPQuery.SQL.Add(s)
        else
            break;
    end;
    //1.1.2003;2.2.1003
    for i := 0 to REPQuery.MacroCount - 1 do begin
        s := INI.ReadString(SECTION, REPQuery.Macros[i].Name, '');
        param := ExtractWord(i + 1, Params, [';']);
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
                MessageDlg('Дата ' + ExtractWord(i + 1, Params, [';']) + ' введена не правильно', mtInformation, [mbOk], 0);
                exit;
            end;
            param := '''' + param + '''';
        end;
        if s = 'S' then param := '''' + param + '''';
        REPQuery.MacroByName(REPQuery.Macros[i].Name).AsString := param;
    end;

    try
        if CurrFilter = '' then
            REPQuery.MacroByName('CURRFILTER').AsString := 'AND 1=1'
        else
        REPQuery.MacroByName('CURRFILTER').AsString := 'AND (' + CurrFilter + ')';
    except
    end;

    INI.WriteString(SECTION, 'LastParams' + IntToStr(RepIndex), Params);
    INI.Free;

    try
        Screen.Cursor := crHourGlass;
        REPQuery.Open;
    except
        on E : Exception do
            MessageDlg('Ошибка выполнения запроса'#13 + e.Message + #13 + REPQuery.Text, mtInformation, [mbOk], 0);
    end;
    Screen.Cursor := crDefault;
end;

procedure InitStructure_3(var Info : ReportStructure_3);
var
    i : integer;
begin
    for i := 1 to 3 do begin
        Info[i].RegDate := 0;
        Info[i].Number := 0;
        Info[i].Curr := '';
        Info[i].BruttoPremium := 0;
        Info[i].AgOtchislenie := 0;
        Info[i].BasePremium_nouse := 0;
        Info[i].OneTaxOtchislenie := 0;
        Info[i].FromDate := 0;
        Info[i].ToDate := 0;
    end;
end;

function GetCurrIndex(var Info : ReportStructure_3; Curr : string) : integer;
var
    i : integer;
begin
    for i := 1 to 3 do begin
        if (Info[i].Curr = Curr) OR (Info[i].Curr = '') then begin
            GetCurrIndex := i;
            Info[i].Curr := Curr;
            exit;
        end;
    end;

    raise Exception.Create('Не найден индекс валюты ' + Curr);

    GetCurrIndex := 1;
end;

procedure RoundValue(var Val : double);
var
   intVal : int64;
begin
    intVal := ROUND(Val * 100);
    Val := intVal / 100;
end;

function CalcSellDate(InpData : TDateTime; Curr : string) : TDateTime;
var
    DaysDelta : integer;
    Y, M, D : WORD;
begin
    CalcSellDate := InpData;
{    DecodeDate(InpData, Y, M, D);
    D := 1;
    IncAMonth(Y, M, D);
    CalcSellDate := EncodeDate(Y, M, D) - 1;
{    DaysDelta := GetDaysDelta(Curr);
    InpData := InpData + DaysDelta;

    if DayOfWeek(InpData) = 7 then //Суббота
        CalcSellDate := InpData + 2
    else
    if DayOfWeek(InpData) = 1 then //Воскресение
        CalcSellDate := InpData + 1
    else
        CalcSellDate := InpData;}
end;

function GetSellPercent(Dt : TDateTime; Curr : string) : double;
var
     SellTbl : TQuery;
begin
     if (last_Date = Dt) AND (last_Curr = Curr) then begin
        GetSellPercent := last_SellPercent;
        exit;
     end;

     SellTbl := TQuery.Create(Application.MainForm);

     SellTbl.SQL.Add('SELECT SaleDate, ' + Curr + ' FROM SaleCurr WHERE SaleDate <= :Dt ORDER BY SaleDate Desc');
     SellTbl.ParamByName('Dt').AsDateTime := Dt;
     try
        SellTbl.DataBaseName := 'RESTRAX';
        SellTbl.Open;
     except
     end;

     if not SellTbl.Active then begin
         SellTbl.DataBaseName := 'BASO';
         SellTbl.Open;
     end;

     if SellTbl.EOF OR SellTbl.FieldByName(Curr).IsNull then begin
        SellTbl.Free;
        raise Exception.Create('Не установлен процент продажи валюты ' + Curr + ' на ' + DateToStr(Dt));
     end;

     last_Date := Dt;
     last_Curr := Curr;
     last_SellPercent := SellTbl.FieldByName(Curr).AsFloat;
     GetSellPercent := SellTbl.FieldByName(Curr).AsFloat;

     SellTbl.Free;
end;

procedure FillStructure_3(RepDate : TDateTime; IsUridich : boolean; FizichTax, UridichTax, CompanyPercent : double; var Info : ReportStructure_3; InpData : TDateTime; Summa : double; Curr : string; AgPercent : double; OneTax : double);
var
     SellData : TDateTime;
     CurrIndex : integer;
     SellPercent : double;
     SellSumma : double;
     AgentSumma, Tax : double;
begin
     if Curr = 'BRB' then begin
        CurrIndex := GetCurrIndex(Info, Curr);
        Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                         Summa * CompanyPercent / 100;
        AgentSumma := Summa * AgPercent / 100;
        if IsUridich then
            AgentSumma := AgentSumma * (1 + UridichTax / 100)
        else
            AgentSumma := AgentSumma * (1 + FizichTax / 100);
        Info[CurrIndex].AgOtchislenie := Info[CurrIndex].AgOtchislenie +
                                       AgentSumma;

        Tax := Summa * CompanyPercent / 100 * OneTax / 100;
        RoundValue(Tax);
        Info[CurrIndex].OneTaxOtchislenie := Info[CurrIndex].OneTaxOtchislenie + Tax;
     end else begin
        CurrIndex := GetCurrIndex(Info, Curr);

        Info[CurrIndex].OneTaxOtchislenie := Info[CurrIndex].OneTaxOtchislenie +
                                             Summa * CompanyPercent / 100 * OneTax / 100;

        //Дата продажи
        SellData := CalcSellDate(InpData, Curr);
        {if SellData >= RepDate then begin //Не продаём
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             Summa * CompanyPercent / 100;
            AgentSumma := Summa * AgPercent / 100;
            if IsUridich then
                AgentSumma := AgentSumma * (1 + UridichTax / 100)
            else
                AgentSumma := AgentSumma * (1 + FizichTax / 100);
            Info[CurrIndex].Otchislenie := Info[CurrIndex].Otchislenie +
                                           AgentSumma;
        end
        else begin }//Успели продать
            SellPercent := GetSellPercent(SellData, Curr);
            SellSumma := Summa * CompanyPercent / 100 * SellPercent / 100;
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             (Summa * CompanyPercent / 100 - SellSumma);

            AgentSumma := Summa * AgPercent / 100;
            if IsUridich then
                AgentSumma := AgentSumma * (1 + UridichTax / 100)
            else
                AgentSumma := AgentSumma * (1 + FizichTax / 100);

            Info[CurrIndex].AgOtchislenie := Info[CurrIndex].AgOtchislenie +
                                           AgentSumma;

            CurrIndex := GetCurrIndex(Info, 'BRB');
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             SellSumma * GetRateOfCurrency(Curr, SellData);

        //end;
     end;
end;

procedure FillPolisesForReserve(IsRNPU : boolean; var DataList : TList; RepDate : TDateTime; taxobj : TaxPercents; List : TASKOneTaxList; sqlText : string; WorkSQL : TQuery; ProgressBar : TProgressBar; InsStart, InsEnd, CURRCODE : string);
var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : PGroupRestrax; //^ReportStructure;
     IsUridichAgent : boolean;
     _AgPercent, _OneTax : double;
begin
     //Те кторые действуют на дату или те которые Б 30 дней и оплачены в предшеств. отчёту месяц
     //Сумма берётся вся пришедшая до отчётной даты
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add(sqlText);
//          Add('SELECT');
//          Add('   Seria, Number, Repdate, StartDate, EndDate, PayDate, PremiumPay, PremiumPay2, PremiumCurr, PremiumCurr2, Uridich, AgPercent');
//          Add('FROM');
//          Add('   POLANDPL');
//          Add('WHERE');
//          Add('   PremiumVal is not null AND');
          if IsRNPU then begin
              Add('   (PayDate between :RepDateSubMonth and :RepDate)');
              Add(' AND ' + InsStart + ' < :RepDate')
          end
          else begin
              Add('   (PayDate <= :RepDate AND ' +  InsEnd + ' >= :RepDate OR'); //попал в период
              Add('   (PayDate >= :RepDateSubMonth AND PayDate < :RepDate AND (' + InsEnd + ' - ' + InsStart + ' + 1) <= 30))');
          end;
          Add('ORDER BY');
          Add('   Seria, Number');
          ParamByName('RepDate').AsDateTime := TRUNC(RepDate);
          if IsRNPU then
              ParamByName('RepDateSubMonth').AsDateTime := IncDate(TRUNC(RepDate), 0, -12, 0)
          else
              ParamByName('RepDateSubMonth').AsDateTime := IncDate(TRUNC(RepDate), 0, -1, 0);
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

         //if WorkSQL.FieldByName('Number').AsInteger = 100590077 then
         //MessageBeep(0);

        ProgressBar.Position := TRUNC(WorkSQL.RecNo * 100 / WorkSQL.RecordCount);
         with WorkSQL do begin
           if FieldByName('Uridich').DataType = ftBoolean then
                IsUridichAgent := FieldByName('Uridich').AsBoolean
           else
           if FieldByName('Uridich').DataType = ftString then
                IsUridichAgent := FieldByName('Uridich').AsString = 'U'
           else
                raise Exception.Create('Поле типа агента имеет неправильный тип');

           _AgPercent := FieldByName('AgPercent').AsFloat;
           _OneTax := GetOneTax(List, FieldByName('PayDate').AsDateTime);
           //if FieldByName('StartDate').AsDateTime >= RepDate then begin
           //     _AgPercent := 0;
           //     _OneTax := 0;
           //end;

           if FieldByName('PremiumCurr').AsString <> '' then
              FillStructure_3(RepDate, IsUridichAgent, taxobj.GetPcnt(FieldByName('PayDate').AsDateTime, true), taxobj.GetPcnt(FieldByName('PayDate').AsDateTime, false), taxobj.GetCompPcnt(FieldByName('PayDate').AsDateTime), Info, FieldByName('PayDate').AsDateTime, FieldByName('PremiumPay').AsFloat, FieldByName('PremiumCurr').AsString, _AgPercent, _OneTax);
           if FieldByName('PremiumCurr2').AsString <> '' then
              FillStructure_3(RepDate, IsUridichAgent, taxobj.GetPcnt(FieldByName('PayDate').AsDateTime, true), taxobj.GetPcnt(FieldByName('PayDate').AsDateTime, false), taxobj.GetCompPcnt(FieldByName('PayDate').AsDateTime), Info, FieldByName('PayDate').AsDateTime, FieldByName('PremiumPay2').AsFloat, FieldByName('PremiumCurr2').AsString, _AgPercent, _OneTax);
         end;

         //Append Info to container
         for i := 1 to 3 do begin
             if Info[i].Curr <> '' then begin
                 New(InfoPtr);
                 //InfoPtr^ := Info[i];
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

                  InfoPtr.AllSummaCurr := CURRCODE;
                  InfoPtr.AllSumma := 100000;

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

                  RoundValue(InfoPtr.ReserveNZPSumm);

                  if IsRNPU then
                      InfoPtr.ReserveNZPSumm := InfoPtr.BasePremium / 10;
                  InfoPtr.StraxPart := InfoPtr.ReserveNZPSumm;

                  DataList.Add(InfoPtr);
             end;
         end;

         WorkSQL.Next;
     end;
     ProgressBar.Position := 0;
end;

procedure OutTo10Form(IsRNPU : boolean; InsTypeName : string; MSEXCEL : VARIANT; RepDate : TDateTime; ReinsurData : TList; IsFast : boolean; Dir : string);
var
    i, j : integer;
    CurrList : array[1..10] of string[10];
    InfoPtr : PGroupRestrax;
    Rounder : double;
    VAL : VARIANT;
    ListOfRows : array [1..10] of integer;
    ExcelLine, ExcelLineStartPart, idx_curr, rows_idx : integer;
    CellRate, CellNumber : string;
    CntCurrs : integer;
    A : integer;
    G : double;
    Ie : double;
    H : double;
    Je : double;
    K : double;
    M : double;
    O : double;
    str_big_summ : string;
    fileName : string;
    ArrayData : VARIANT;
begin
    rows_idx := 1;
    try
         for i := 1 to 10 do CurrList[i] := '';

         for i := 0 to ReinsurData.Count - 1 do begin
             InfoPtr := PGroupRestrax(ReinsurData.Items[i]);
             RoundValue(InfoPtr.BruttoPremium);
             RoundValue(InfoPtr.AgentSum);
             RoundValue(InfoPtr.FondSum);
             RoundValue(InfoPtr.BasePremium);
             RoundValue(InfoPtr.ReserveNZPSumm);
             RoundValue(InfoPtr.RestraxPart);
             RoundValue(InfoPtr.StraxPart);
             for j := 1 to 10 do
                 if (CurrList[j] = InfoPtr.Curr) OR (CurrList[j] = '') then begin
                    CurrList[j] := InfoPtr.Curr;
                    break;
                 end;
         end;

         WaitForm.WorkName.Caption := 'РНП(10)';
         WaitForm.Update;
         if IsRNPU then
             MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\РНПУ.XLS')
         else
             MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\РНП(10).XLS');
         MSEXCEL.ActiveSheet.Range['B6', 'B6'].Value := 'на ' + DateToStr(RepDate);
         MSEXCEL.ActiveSheet.Range['C8', 'C8'].Value := InsTypeName;

         //Вставка Курсов валют
         Rounder := GetRateOfCurrency('USD', RepDate-1);
         RoundValue(Rounder);
         VAL := Rounder;
         MSEXCEL.ActiveSheet.Range['F2', 'F2'].Value := VAL;

         if RepDate <= EncodeDate(2002, 2, 1) then begin
             Rounder := GetRateOfCurrency('DM', RepDate-1);
             RoundValue(Rounder);
             VAL := Rounder;
             MSEXCEL.ActiveSheet.Range['F3', 'F3'].Value := VAL;
         end
         else
             MSEXCEL.ActiveSheet.Range['F3', 'F3'].Value := 'Не существует';

         Rounder := GetRateOfCurrency('EUR', RepDate-1);
         RoundValue(Rounder);
         VAL := Rounder;
         MSEXCEL.ActiveSheet.Range['F4', 'F4'].Value := VAL;

         Rounder := GetRateOfCurrency('RUR', RepDate-1);
         RoundValue(Rounder);
         VAL := Rounder;
         MSEXCEL.ActiveSheet.Range['F5', 'F5'].Value := VAL;

         ExcelLine := 18;

         for CntCurrs := 1 to 10 do
             if CurrList[CntCurrs] = '' then break;
         Dec(CntCurrs);

         ArrayData := VarArrayCreate([1, 1, 1, 11], varVariant);
         for idx_curr := 1 to 10 do begin
             if CurrList[idx_curr] = '' then break;
             CellRate := '';
             if CurrList[idx_curr] = 'USD' then CellRate := 'F2'; //USD
             if CurrList[idx_curr] = 'DM' then CellRate := 'F3';
             if CurrList[idx_curr] = 'EUR' then CellRate := 'F4';
             if CurrList[idx_curr] = 'RUR' then CellRate := 'F5';
             ExcelLineStartPart := ExcelLine;
             A := 0;
             G := 0;
             Ie := 0;
             H := 0;
             Je := 0;
             K := 0;
             M := 0;
             O := 0;
             for i := 0 to ReinsurData.Count - 1 do begin

                 WaitForm.ProgressBar.Position := ROUND((idx_curr - 1) * 100 / CntCurrs) +
                                                  ROUND(i / ReinsurData.Count * 100 / CntCurrs);

                 InfoPtr := PGroupRestrax(ReinsurData.Items[i]);
                 if InfoPtr.Curr <> CurrList[idx_curr] then continue;

                 if not IsFast then begin
                       ArrayData[1, 1] := InfoPtr.Seria + '/' + FloatToStr(InfoPtr.Number);
                       ArrayData[1, 2] := InfoPtr.AllSummaCurr;
                       ArrayData[1, 3] := InfoPtr.AllSumma;
                       ArrayData[1, 4] := DateToStr(InfoPtr.DateFrom);
                       ArrayData[1, 5] := DateToStr(InfoPtr.DateTo);
                       ArrayData[1, 6] := InfoPtr.Curr;
                       ArrayData[1, 7] := InfoPtr.BruttoPremium;
                       ArrayData[1, 8] := InfoPtr.FondSum;
                       ArrayData[1, 9] := InfoPtr.AgentSum;
                       ArrayData[1, 10] := InfoPtr.BasePremium;
                       ArrayData[1, 11] := InfoPtr.ReserveNZPSumm;

                       CellNumber := IntToStr(ExcelLine);
                       {MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := InfoPtr.Seria + '/' + FloatToStr(InfoPtr.Number);

                       MSEXCEL.ActiveSheet.Range['B' + CellNumber, 'B' + CellNumber].Value := InfoPtr.AllSummaCurr;
                       VAL := InfoPtr.AllSumma;
                       MSEXCEL.ActiveSheet.Range['C' + CellNumber, 'C' + CellNumber].Value := VAL;

                       MSEXCEL.ActiveSheet.Range['D' + CellNumber, 'D' + CellNumber].Value := DateToStr(InfoPtr.DateFrom);
                       MSEXCEL.ActiveSheet.Range['E' + CellNumber, 'E' + CellNumber].Value := DateToStr(InfoPtr.DateTo);
                       MSEXCEL.ActiveSheet.Range['F' + CellNumber, 'F' + CellNumber].Value := InfoPtr.Curr;
                       VAL := InfoPtr.BruttoPremium;
                       MSEXCEL.ActiveSheet.Range['G' + CellNumber, 'G' + CellNumber].Value := VAL;
                       VAL := InfoPtr.FondSum;
                       MSEXCEL.ActiveSheet.Range['H' + CellNumber, 'H' + CellNumber].Value := VAL;
                       VAL := InfoPtr.AgentSum;
                       MSEXCEL.ActiveSheet.Range['I' + CellNumber, 'I' + CellNumber].Value := VAL;
                       VAL := InfoPtr.BasePremium;
                       MSEXCEL.ActiveSheet.Range['J' + CellNumber, 'J' + CellNumber].Value := VAL;
                       VAL := InfoPtr.ReserveNZPSumm;
                       MSEXCEL.ActiveSheet.Range['K' + CellNumber, 'K' + CellNumber].Value := VAL;
                       }
                       MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'K' + CellNumber].Value := ArrayData;


                       //Курс по НБ
                       if CurrList[idx_curr] <> 'BRB' then
                           MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=K' + CellNumber + '*' + CellRate
                       else
                           MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=K' + CellNumber;
                       VAL := InfoPtr.RestraxPart;
                       MSEXCEL.ActiveSheet.Range['M' + CellNumber, 'M' + CellNumber].Value := VAL;
                       //Курс по НБ
                       if CurrList[idx_curr] <> 'BRB' then
                           MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=M' + CellNumber + '*' + CellRate
                       else
                           MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=M' + CellNumber;
                       VAL := InfoPtr.StraxPart;
                       MSEXCEL.ActiveSheet.Range['O' + CellNumber, 'O' + CellNumber].Value := VAL;
                       //Курс по НБ
                       if CurrList[idx_curr] <> 'BRB' then
                           MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=O' + CellNumber + '*' + CellRate
                       else
                           MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=O' + CellNumber;
                       Inc(ExcelLine);
                 end else begin
                       A := A + 1;
                       G := G + InfoPtr.BruttoPremium;
                       Ie := Ie + InfoPtr.AgentSum;
                       H := H + InfoPtr.FondSum;
                       Je := Je + InfoPtr.BasePremium;
                       K := K + InfoPtr.ReserveNZPSumm;
                       M := M + InfoPtr.RestraxPart;
                       O := O + InfoPtr.StraxPart;
                 end;
             end;

             CellNumber := IntToStr(ExcelLine);

             if not IsFast then begin
                 MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := 'ИТОГО по валюте';
                 MSEXCEL.ActiveSheet.Range['G' + CellNumber, 'G' + CellNumber].Value := '=SUM(G' + IntToStr(ExcelLineStartPart) + ':G' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['H' + CellNumber, 'H' + CellNumber].Value := '=SUM(H' + IntToStr(ExcelLineStartPart) + ':H' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['I' + CellNumber, 'I' + CellNumber].Value := '=SUM(I' + IntToStr(ExcelLineStartPart) + ':I' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['J' + CellNumber, 'J' + CellNumber].Value := '=SUM(J' + IntToStr(ExcelLineStartPart) + ':J' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['K' + CellNumber, 'K' + CellNumber].Value := '=SUM(K' + IntToStr(ExcelLineStartPart) + ':K' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=SUM(L' + IntToStr(ExcelLineStartPart) + ':L' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['M' + CellNumber, 'M' + CellNumber].Value := '=SUM(M' + IntToStr(ExcelLineStartPart) + ':M' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=SUM(N' + IntToStr(ExcelLineStartPart) + ':N' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['O' + CellNumber, 'O' + CellNumber].Value := '=SUM(O' + IntToStr(ExcelLineStartPart) + ':O' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=SUM(P' + IntToStr(ExcelLineStartPart) + ':P' + IntToStr(ExcelLine - 1) + ')';
                 ListOfRows[rows_idx] := ExcelLine;
                 rows_idx := rows_idx + 1;
             end
             else begin
                 MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := 'ИТОГО по ' + CurrList[idx_curr] + '  Количество ' + IntToStr(A);

                 VAL := G;
                 MSEXCEL.ActiveSheet.Range['G' + CellNumber, 'G' + CellNumber].Value := VAL;
                 VAL := Ie;
                 MSEXCEL.ActiveSheet.Range['I' + CellNumber, 'I' + CellNumber].Value := VAL;
                 VAL := H;
                 MSEXCEL.ActiveSheet.Range['H' + CellNumber, 'H' + CellNumber].Value := VAL;
                 VAL := Je;
                 MSEXCEL.ActiveSheet.Range['J' + CellNumber, 'J' + CellNumber].Value := VAL;
                 VAL := K;
                 MSEXCEL.ActiveSheet.Range['K' + CellNumber, 'K' + CellNumber].Value := VAL;
                 VAL := M;
                 MSEXCEL.ActiveSheet.Range['M' + CellNumber, 'M' + CellNumber].Value := VAL;
                 VAL := O;
                 MSEXCEL.ActiveSheet.Range['O' + CellNumber, 'O' + CellNumber].Value := VAL;

                 if CurrList[idx_curr] <> 'BRB' then begin
                     MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=K' + CellNumber + '*' + CellRate;
                     MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=M' + CellNumber + '*' + CellRate;
                     MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=O' + CellNumber + '*' + CellRate;
                 end
                 else begin
                     MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=K' + CellNumber;
                     MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=M' + CellNumber;
                     MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=O' + CellNumber;
                 end;
                 ListOfRows[rows_idx] := ExcelLine;
                 rows_idx := rows_idx + 1;
             end;

             Inc(ExcelLine, 2);
         end;

         VarClear(ArrayData);

         if rows_idx > 1 then begin
             CellNumber := IntToStr(ExcelLine);
             MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := 'ИТОГО по резерву';

             str_big_summ := '';
             for i := rows_idx - 1 downto 1 do begin
                 str_big_summ := str_big_summ + 'L' + IntToStr(ListOfRows[i]) + '+';
             end;
             MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=' + Copy(str_big_summ, 1, Length(str_big_summ) - 1);

             str_big_summ := '';
             for i := rows_idx - 1 downto 1 do begin
                 str_big_summ := str_big_summ + 'N' + IntToStr(ListOfRows[i]) + '+';
             end;
             MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=' + Copy(str_big_summ, 1, Length(str_big_summ) - 1);

             str_big_summ := '';
             for i := rows_idx - 1 downto 1 do begin
                 str_big_summ := str_big_summ + 'P' + IntToStr(ListOfRows[i]) + '+';
             end;
             MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=' + Copy(str_big_summ, 1, Length(str_big_summ) - 1);
         end;

         if not IsRNPU then
             fileName := GetAppDir + 'DOCS\' + Dir + '\РНП(10)_' + DateToStr(RepDate) + '.XLS'
         else
             fileName := GetAppDir + 'DOCS\' + Dir + '\РНПУ_' + DateToStr(RepDate) + '.XLS';
         SysUtils.DeleteFile(fileName);
         MSEXCEL.ActiveWorkbook.SaveAs(fileName);
    except
        on E : Exception do begin
            Application.BringToFront;
            if WaitForm <> nil then WaitForm.Close;
            MessageDlg('Возникла ошибка при формировании отчёта'#13 + e.Message, mtInformation, [mbOk], 0);
        end;
    end;
end;

function ExpStr(s : string; var IsRussian, IsEnglish, IsDigit, IsBad : boolean; CheckMinus : boolean) : string;
var
    i : integer;
    Rus : string;
    Eng : string;
    BadSymb : string;
begin
    IsBad := false;

    BadSymb := ';:.,"''!@#$%^&*()_=+<>?~`/[]{}\|';
    for i := 1 to Length(BadSymb) do
        if Pos(BadSymb[i], s) <> 0 then begin
            IsBad := true;
            break;
         end;

    if not IsBad then
        for i := 1 to Length(s) do begin
            if (s[i] >= 'а') AND (s[i] <= 'я') then begin
                IsBad := true;
                break;
            end;
        end;

    if not IsBad then
        for i := 1 to Length(s) do begin
            if (s[i] >= 'a') AND (s[i] <= 'z') then begin
                IsBad := true;
                break;
            end;
        end;

    if CheckMinus then
      if Pos('-', s) <> 0 then begin
          IsBad := true;
      end;

    if Copy(s, 1, 1) = ' ' then
        IsBad := true;

    if Length(s) > 0 then
        if Copy(s, Length(s), 1) = ' ' then
            IsBad := true;

    if Pos('  ', s) <> 0 then begin
        IsBad := true;
    end;

    IsDigit := false;
    for i := 0 to 9 do begin
        IsDigit := Pos(Chr(Ord('0') + i), s) <> 0;
        if IsDigit then break;
    end;

    IsRussian := false;
    for i := 1 to Length(s) do begin
        IsRussian := (s[i] >= 'А') AND (s[i] <= 'Я');
        if IsRussian then break;
    end;

    IsEnglish := false;
    for i := 1 to Length(s) do begin
        IsEnglish := (s[i] >= 'A') AND (s[i] <= 'Z');
        if IsEnglish then break;
    end;

    ExpStr := s;
end;

function FillOutTaxList(var List : TASKOneTaxList; strList : TStringGrid) : boolean;
var
    i : integer;
begin
    try
        FillOutTaxList := false;
        for i := 1 to 10 do begin
            if (strList.Cells[1, i] = '') and (i = 1) then exit;
            if (strList.Cells[0, i] = '')  then break;
            List[i - 1].Dt := StrToDate(Trim(strList.Cells[0, i]));
            List[i - 1].Pcnt := StrToFloat(Trim(strList.Cells[1, i]));
            List[i].Dt := 0;
            if (Trim(strList.Cells[0, i]) = '') and (Trim(strList.Cells[1, i]) <> '') or
               (Trim(strList.Cells[0, i]) <> '') and (Trim(strList.Cells[1, i]) = '') then exit;
        end;
        FillOutTaxList := true;
    except
    end
end;

function GetOneTax(List : TASKOneTaxList; Dt : TDate) : double;
var
    i : integer;
begin
    GetOneTax := 0;
    for i := 0 to 10 do begin
        if List[i].Dt = 0 then exit;
        if Dt >= List[i].Dt then begin
            GetOneTax := List[i].Pcnt;
            //exit;
        end;
    end
end;

function PrepareStr(City, Street, House : string) : string;
var
    s : string;
    i : integer;
begin
    s := City;
    if Street <> '' then begin
        if s <> '' then s := s + ' ';
        s := s + Street;
    end;
    if House <> '' then begin
        if s <> '' then s := s + ' ';
        s := s + House;
    end;
    //if Tel <> '' then begin
    //    if s <> '' then s := s + ' ';
    //    s := s + 'ТЕЛ ' + Tel;
    //end;

    for i := 1 to Length(s) do
        if s[i] in ['.', ',', '-', '/', '''', '"', '*'] then
            s[i] := ' ';

    while Pos('  ', s) <> 0 do
        s := Copy(s, 1, Pos('  ', s)) + Copy(s, Pos('  ', s) + 2, 1024);

    PrepareStr := AnsiUpperCase(s);
end;

function CalcPeriod(D1, D2 : TDateTime) : string;
var
    s : string;
begin
    if (D2 - D1) < 10 then
        raise Exception.Create('Период страхования слишком мал');
    if (D2 - D1) > 370 then
        raise Exception.Create('Период страхования слишком велик');

    s := '';
    if ROUND(D2 - D1 + 1) in [15] then s := '15 суток';
    if ROUND(D2 - D1 + 1) in [28, 29, 30, 31] then s := '1 месяц';
    if (ROUND(D2 - D1) >= 30*2-2) AND (ROUND(D2 - D1) <= 31*2)  then s := '2 месяца';
    if (ROUND(D2 - D1) >= 30*3-2) AND (ROUND(D2 - D1) <= 31*3)  then s := '3 месяца';
    if (ROUND(D2 - D1) >= 30*4-2) AND (ROUND(D2 - D1) <= 31*4)  then s := '4 месяца';
    if (ROUND(D2 - D1) >= 30*5-2) AND (ROUND(D2 - D1) <= 31*5)  then s := '5 месяцев';
    if (ROUND(D2 - D1) >= 30*6-2) AND (ROUND(D2 - D1) <= 31*6)  then s := '6 месяцев';
    if (ROUND(D2 - D1) >= 30*7-2) AND (ROUND(D2 - D1) <= 31*7)  then s := '7 месяцев';
    if (ROUND(D2 - D1) >= 30*8-2) AND (ROUND(D2 - D1) <= 31*8)  then s := '8 месяцев';
    if (ROUND(D2 - D1) >= 30*9-2) AND (ROUND(D2 - D1) <= 31*9)  then s := '9 месяцев';
    if (ROUND(D2 - D1) >= 30*10-2) AND (ROUND(D2 - D1) <= 31*10)  then s := '10 месяцев';
    if (ROUND(D2 - D1) >= 30*11-2) AND (ROUND(D2 - D1) <= 31*11)  then s := '11 месяцев';
    if (ROUND(D2 - D1 + 1) >= 365) AND (ROUND(D2 - D1 + 1) <= 366)  then s := '1 год';

    if s = '' then
        raise Exception.Create('Период страхования не определён');
    CalcPeriod := s;
end;

function CalcPeriod2(D1, D2 : TDateTime) : Integer;
var
    s : string;
begin
    if (D2 - D1) < 10 then
        raise Exception.Create('Период страхования слишком мал');
    if (D2 - D1) > 370 then
        raise Exception.Create('Период страхования слишком велик');

    s := '';
    if ROUND(D2 - D1 + 1) in [15] then s := '15';
    if ROUND(D2 - D1 + 1) in [28, 29, 30, 31] then s := '1';
    if (ROUND(D2 - D1) >= 30*2-2) AND (ROUND(D2 - D1) <= 31*2)  then s := '2';
    if (ROUND(D2 - D1) >= 30*3-2) AND (ROUND(D2 - D1) <= 31*3)  then s := '3';
    if (ROUND(D2 - D1) >= 30*4-2) AND (ROUND(D2 - D1) <= 31*4)  then s := '4';
    if (ROUND(D2 - D1) >= 30*5-2) AND (ROUND(D2 - D1) <= 31*5)  then s := '5';
    if (ROUND(D2 - D1) >= 30*6-2) AND (ROUND(D2 - D1) <= 31*6)  then s := '6';
    if (ROUND(D2 - D1) >= 30*7-2) AND (ROUND(D2 - D1) <= 31*7)  then s := '7';
    if (ROUND(D2 - D1) >= 30*8-2) AND (ROUND(D2 - D1) <= 31*8)  then s := '8';
    if (ROUND(D2 - D1) >= 30*9-2) AND (ROUND(D2 - D1) <= 31*9)  then s := '9';
    if (ROUND(D2 - D1) >= 30*10-2) AND (ROUND(D2 - D1) <= 31*10)  then s := '10';
    if (ROUND(D2 - D1) >= 30*11-2) AND (ROUND(D2 - D1) <= 31*11)  then s := '11';
    if (ROUND(D2 - D1 + 1) >= 365) AND (ROUND(D2 - D1 + 1) <= 366)  then s := '12';

    if s = '' then
        raise Exception.Create('Период страхования не определён');
    CalcPeriod2 := StrToInt(s);
end;

procedure BuildReportInputSum(RepDate1, RepDate2: TDateTime; sqlText, AgentName : string; WorkSQL : TQuery);
var
    CurrList : array[1..10] of string[5];
    SummList : array[1..10] of double;
    SummBRBList : array[1..10] of double;
    SummBRBListRep : array[1..10] of double;
    SummAgList : array[1..10] of double;
    i : integer;
    MSEXCEL, ASH : VARIANT;
begin
    try
        Screen.Cursor := crHourGlass;
        for i := 1 to 10 do SummList[i] := 0;
        for i := 1 to 10 do SummBRBList[i] := 0;
        for i := 1 to 10 do SummBRBListRep[i] := 0;
        for i := 1 to 10 do SummAgList[i] := 0;
        for i := 1 to 10 do CurrList[i] := '';
        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add(sqlText);
        WorkSQL.Open;
        while not WorkSQL.Eof do begin
            for i := 1 to 10 do
                if (WorkSQL.Fields[1].AsString = CurrList[i]) or (CurrList[i] = '') then begin
                    if (WorkSQL.Fields[0].AsFloat > 0) and (WorkSQL.Fields[1].AsString = '') then
                        raise Exception.Create('Обнаружена ' + WorkSQL.Fields[0].AsString + ' сумма без указания валюты. Полис ' + WorkSQL.Fields[3].AsString);
                    CurrList[i] := WorkSQL.Fields[1].AsString;
                    SummList[i] := SummList[i] + WorkSQL.Fields[0].AsFloat;
                    if WorkSQL.Fields.Count >= 5 then
                        SummAgList[i] := SummAgList[i] + WorkSQL.Fields[0].AsFloat * WorkSQL.Fields[4].AsFloat / 100;
                    if CurrList[i] <> '' then begin
                        SummBRBList[i] := SummBRBList[i] + WorkSQL.Fields[0].AsFloat * GetRateOfCurrency(CurrList[i], WorkSQL.Fields[2].AsDateTime);
                        if AgentName = '' then
                            SummBRBListRep[i] := SummBRBListRep[i] + WorkSQL.Fields[0].AsFloat * GetRateOfCurrency(CurrList[i], WorkSQL.Fields[4].AsDateTime);
                    end;
                    break;
                end;
            WorkSQL.Next
        end;

        if CurrList[1] = '' then begin
            Screen.Cursor := crDefault;
            MessageDlg('Нет Данных за указанный период', mtInformation, [mbOk], 0);
            exit;
        end;

        if not InitExcel(MSEXCEL) then exit;
        MSEXCEL.WorkBooks.Add;
        MSEXCEL.Visible := true;
        ASH := MSEXCEL.ActiveSheet;

        ASH.Range['A1', 'A1'].Value := DateToStr(RepDate1);
        ASH.Range['B1', 'B1'].Value := DateToStr(RepDate2);

        ASH.Range['A2', 'A2'].Value := 'Получено';
        if AgentName = '' then
            ASH.Range['C2', 'C2'].Value := 'По курсу на дату оплаты'
        else
            ASH.Range['C2', 'C2'].Value := 'По курсу на дату отчёта';
        ASH.Range['D2', 'D2'].Value := 'Курс на ' + DateToStr(RepDate2 + 1);
        ASH.Range['E2', 'E2'].Value := 'По курсу на конец периода';
        if AgentName = '' then begin
            ASH.Range['F2', 'F2'].Value := 'По курсу на дату отчёта';
            ASH.Range['G2', 'G2'].Value := 'Разность F-C';
        end
        else
            ASH.Range['F2', 'F2'].Value := AgentName;

        for i := 1 to 10 do begin
            if CurrList[i] <> '' then begin
                ASH.Range['A' + IntToStr(2 + i), 'A' + IntToStr(2 + i)].Value := VARIANT(SummList[i]);
                ASH.Range['B' + IntToStr(2 + i), 'B' + IntToStr(2 + i)].Value := CurrList[i];
                ASH.Range['C' + IntToStr(2 + i), 'C' + IntToStr(2 + i)].Value := VARIANT(SummBRBList[i]);
                ASH.Range['D' + IntToStr(2 + i), 'D' + IntToStr(2 + i)].Value := VARIANT(GetRateOfCurrency(CurrList[i], RepDate2+1));
                ASH.Range['E' + IntToStr(2 + i), 'E' + IntToStr(2 + i)].Value := '=A' + IntToStr(2 + i) + '*D' + IntToStr(2 + i);
                if AgentName = '' then begin
                    ASH.Range['F' + IntToStr(2 + i), 'F' + IntToStr(2 + i)].Value := VARIANT(SummBRBListRep[i]);
                    ASH.Range['G' + IntToStr(2 + i), 'G' + IntToStr(2 + i)].Value := '=F' + IntToStr(2 + i) + '-C' + IntToStr(2 + i);
                end
                else begin
                    ASH.Range['F' + IntToStr(2 + i), 'F' + IntToStr(2 + i)].Value := VARIANT(SummAgList[i]);
                    ASH.Range['G' + IntToStr(2 + i), 'G' + IntToStr(2 + i)].Value :=  '=F' + IntToStr(2 + i) + '*D' + IntToStr(2 + i);
                end
            end
            else begin
                ASH.Range['C' + IntToStr(2 + i), 'C' + IntToStr(2 + i)].Value := '=SUM(C2:C' + IntToStr(i+1) + ')';
                ASH.Range['E' + IntToStr(2 + i), 'E' + IntToStr(2 + i)].Value := '=SUM(E2:E' + IntToStr(i+1) + ')';
                ASH.Range['F' + IntToStr(2 + i), 'F' + IntToStr(2 + i)].Value := '=SUM(F2:F' + IntToStr(i+1) + ')';
                ASH.Range['G' + IntToStr(2 + i), 'G' + IntToStr(2 + i)].Value := '=SUM(G2:G' + IntToStr(i+1) + ')';
                break;
            end
        end;
        ASH.Columns['A:F'].EntireColumn.AutoFit
    except
        on E : Exception do begin
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
        end
    end;
    Screen.Cursor := crDefault;
end;

function CalcAlvenaTime(T : integer) : string;
var
    Hours, Minutes : integer;
    s : string;
begin
    Hours := T div 3600;
    Dec(T, LongInt(Hours)*3600);
    Minutes := T div 60;
    s := IntToStr(Hours);
    if Hours < 10 then s := '0' + s;
    s := s + ':';
    if Minutes < 10 then
        s := s + '0' + IntToStr(Minutes)
    else
        s := s + IntToStr(Minutes);
    CalcAlvenaTime := s;
end;

function TaxPercents.GetPcnt(Dt : TDateTime; IsFiz : boolean) : double;
begin
    if IsFiz then begin
        if Dt < FizDt then GetPcnt := FizPcntBefore
        else GetPcnt := FizPcntAfter;
    end
    else begin
        if Dt < UrDt then GetPcnt := UrPcntBefore
        else GetPcnt := UrPcntAfter;
    end;
end;

function TaxPercents.GetCompPcnt(Dt : TDateTime) : double;
begin
    if Dt < CompPcntDt then GetCompPcnt := CompPcntBefore
    else GetCompPcnt := CompPcntAfter;
end;

function GetAliasPath(Alias : string) : string;
var
    MyStringList : TStringList;
    s : string;
begin
    MyStringList := TStringList.Create;
    try
    Session.GetAliasParams(Alias, MyStringList);
    s := MyStringList.Values['PATH'];
    if s[Length(s)] <> '\' then s := s + '\';
    GetAliasPath := s;
    except
        on E : Exception do
                ShowMessage('Alias не найден. Настройте пусть к базе данных. ' + Alias);
    end;
    MyStringList.Free;
end;

{$R *.DFM}

end.
