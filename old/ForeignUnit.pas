unit ForeignUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, DBTables, Db, RxQuery, StdCtrls, Grids, DBGrids, RXDBCtrl;

type
  TInsForeign = class(TForm)
    MainMenu: TMainMenu;
    N1: TMenuItem;
    MenuCompany: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    FilterMenu: TMenuItem;
    MainQuery: TRxQuery;
    FOREIGNDB: TDatabase;
    DataSource: TDataSource;
    ForeignGrid: TRxDBGrid;
    Memo: TMemo;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryInsComp: TStringField;
    MainQueryRegDate: TDateField;
    MainQueryRegTime: TStringField;
    MainQueryFIO: TStringField;
    MainQueryState: TStringField;
    MainQueryPassport: TStringField;
    MainQuerySrok: TFloatField;
    MainQueryDateFrom: TDateField;
    MainQueryDateTo: TDateField;
    MainQueryInsurer: TStringField;
    MainQuerySumma: TFloatField;
    MainQueryFee: TFloatField;
    MainQueryPaySumma: TFloatField;
    MainQueryCurr: TStringField;
    MainQueryPayDate: TDateField;
    MainQueryIsNal: TFloatField;
    MainQueryAGENTCODE: TStringField;
    MainQueryAgPercent: TFloatField;
    MainQueryNAME: TStringField;
    MainQueryPaySummaStr: TStringField;
    N2: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    MainQueryUridich: TStringField;
    CalcSQL: TRxQuery;
    N11: TMenuItem;
    N12: TMenuItem;
    WorkSQL: TQuery;
    SNLock: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    MainQueryPeriod: TStringField;
    procedure N3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FilterMenuClick(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure N6Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure MainQueryUridichGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure N10Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure SNLockClick(Sender: TObject);
    procedure N14Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
  private
    { Private declarations }
    W : HWnd;
    R : TRECT;
    MSEXCEL : Variant;
    procedure FillInfo;
    procedure FillDocsData(var DataList, ReinsurList : TList; RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxComis : double);
  public
    { Public declarations }
    procedure CreateReport(RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxKomis : double);
  end;

var
  InsForeign: TInsForeign;

implementation

uses mf, ForeignFilterUnit, DaysForSellCurrUnit, CurrRatesUnit,
  MandSellCurrUnit, Sort, InpRezParamsFor, MoveCurrPeriods, datamod,
  PolandUnit, AboutUnit, ForeignRptUnit {Структуры и ф-ии для резерва} ;

{$R *.DFM}

procedure TInsForeign.N3Click(Sender: TObject);
begin
     Close
end;

procedure TInsForeign.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
     Action := caFree;
end;

procedure TInsForeign.FormCreate(Sender: TObject);
begin
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

     MainQuery.MacroByName('WHERE').AsString := '';
     MainQuery.MacroByName('ORDER').AsString := '';
     MainQuery.Open;
     FillInfo;
end;

procedure TInsForeign.FilterMenuClick(Sender: TObject);
begin
     if ForeignFilter.ShowModal = mrCancel then exit;
     if MainQuery.MacroByName('WHERE').AsString <> ForeignFilter.FilterText then begin
         try
             Screen.Cursor := crHourGlass;
             MainQuery.Close;
             MainQuery.MacroByName('WHERE').AsString := ForeignFilter.FilterText;
             MainQuery.Open;
             FillInfo;
         except
             on E : Exception do
                 MessageDlg('Ошибка запроса к БД'#13 + E.Message, mtInformation, [mbOK], 0);
         end
     end;
     Screen.Cursor := crDefault;
end;

procedure TInsForeign.MainQueryCalcFields(DataSet: TDataSet);
begin
     if MainQueryPaySumma.AsFloat > 0.01 then
         MainQueryPaySummaStr.AsString := MainQueryPaySumma.AsString + ' ' + MainQueryCurr.AsString;

     MainQueryPeriod.AsString := '';
     if not MainQueryDateFrom.IsNull then
         MainQueryPeriod.AsString := MainQueryDateFrom.AsString + ' - ' + MainQueryDateTo.AsString;
end;

procedure TInsForeign.N6Click(Sender: TObject);
begin
     DaysForSellCurr.ShowModal
end;

procedure TInsForeign.N7Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

procedure TInsForeign.N5Click(Sender: TObject);
begin
     if MandSellCurr = nil then begin
         Application.CreateForm(TMandSellCurr, MandSellCurr);
         MandSellCurr.ShowModal;
     end;
end;

procedure TInsForeign.N9Click(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrCancel then exit;
     if MainQuery.MacroByName('ORDER').AsString <> SortForm.GetOrder('ORDER BY') then begin
         try
             Screen.Cursor := crHourGlass;
             MainQuery.Close;
             MainQuery.MacroByName('ORDER').AsString := SortForm.GetOrder('ORDER BY');
             MainQuery.Open;
         except
             on E : Exception do
                 MessageDlg('Ошибка запроса к БД'#13 + E.Message, mtInformation, [mbOK], 0);
         end
     end;
     Screen.Cursor := crDefault;
end;

procedure TInsForeign.MainQueryUridichGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
     if Sender.AsString = 'Y' then Text := 'Юридическое'
     else if Sender.AsString = '' then Text := ''
     else Text := 'Физическое'
end;

procedure TInsForeign.FillInfo;
var
     s : string;
begin
     Memo.Clear;
     CalcSQL.Close;
     if MainQuery.MacroByName('WHERE').AsString = '' then
         CalcSQL.MacroByName('WHERE').AsString := '1=1'
     else begin
         s := MainQuery.MacroByName('WHERE').AsString;
         if Pos('WHERE', s) <> 0 then s := Copy(s, Pos('WHERE', s) + 5, Length(s));
         CalcSQL.MacroByName('WHERE').AsString := s;
     end;
     CalcSQL.Open;
     s := '';
     while not CalcSQL.EOF do begin
        if CalcSQL.FieldByName('S').AsFloat > 0.01 then
           s := s + CalcSQL.FieldByName('S').AsString + ' ' + CalcSQL.FieldByName('CURR').AsString + '; ';
           CalcSQL.Next;
     end;
     Memo.Lines.Add(s);
     Memo.Lines.Add('Полисов ' + IntToStr(MainQuery.RecordCount));
     CalcSQL.Close;
end;

procedure TInsForeign.N10Click(Sender: TObject);
begin
     InitRezervLeben.ShowModal
end;

procedure TInsForeign.N12Click(Sender: TObject);
begin
     MoveCurr.ShowModal
end;

procedure TInsForeign.CreateReport(RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxKomis : double);
begin
     MessageDlg('Ещё не придумали!', mtInformation, [mbOk], 0);
end;

{
var
    ListData : TList;
    ReinsurData : TList;
    CurrList : array[1..10] of string[10];
    idx_curr, i, j : integer;
    InfoPtr : PGroupRestrax;
    ExcelLine, ExcelLineStartPart : integer;
    OutFileName : string;
    CellNumber : string;
begin
    ListData := TList.Create;
    ReinsurData := TList.Create;
    try
         FillDocsData(ListData, ReinsurData, TRUNC(RepDate), FizichTax, UridichTax, OutPercent, RestraxKomis);
         if not InitExcel(MSEXCEL) then begin
            MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
            exit;
         end;
         MSEXCEL.Visible := true;
         TransferToExecl('Инострано туристо', RepDate, 'FOREIGN', ListData, MSEXCEL);

         for i := 1 to 10 do CurrList[i] := '';

         for i := 0 to ReinsurData.Count - 1 do begin
             InfoPtr := PGroupRestrax(ReinsurData.Items[i]);
             for j := 1 to 10 do
                 if (CurrList[j] = InfoPtr.Curr) OR (CurrList[j] = '') then begin
                    CurrList[j] := InfoPtr.Curr;
                    break;
                 end;
         end;

         MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\РНП(10).XLS');
         MSEXCEL.ActiveSheet.Range['B6', 'B6'].Value := 'на ' + DateToStr(RepDate);
         MSEXCEL.ActiveSheet.Range['C8', 'C8'].Value := 'Инострано туристо';

         ExcelLine := 18;
         for idx_curr := 1 to 10 do begin
             if CurrList[idx_curr] = '' then break;
             ExcelLineStartPart := ExcelLine;
             for i := 0 to ReinsurData.Count - 1 do begin
                 InfoPtr := PGroupRestrax(ReinsurData.Items[i]);
                 if InfoPtr.Curr <> CurrList[idx_curr] then continue;
                 CellNumber := IntToStr(ExcelLine);
                 MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := FloatToStr(InfoPtr.Number);
                 MSEXCEL.ActiveSheet.Range['F' + CellNumber, 'F' + CellNumber].Value := DateToStr(InfoPtr.DateFrom);
                 MSEXCEL.ActiveSheet.Range['G' + CellNumber, 'G' + CellNumber].Value := DateToStr(InfoPtr.DateTo);
                 MSEXCEL.ActiveSheet.Range['H' + CellNumber, 'H' + CellNumber].Value := InfoPtr.Curr;
                 MSEXCEL.ActiveSheet.Range['I' + CellNumber, 'I' + CellNumber].Value := FloatToStr(InfoPtr.BruttoPremium);
                 MSEXCEL.ActiveSheet.Range['J' + CellNumber, 'J' + CellNumber].Value := FloatToStr(InfoPtr.Comis);
                 MSEXCEL.ActiveSheet.Range['K' + CellNumber, 'K' + CellNumber].Value := FloatToStr(InfoPtr.BasePremium);
                 MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := FloatToStr(InfoPtr.ReserveNZPSumm);
                 MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := FloatToStr(InfoPtr.RestraxPart);
                 MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := FloatToStr(InfoPtr.StraxPart);
                 Inc(ExcelLine);
             end;
             CellNumber := IntToStr(ExcelLine);
             MSEXCEL.ActiveSheet.Range['A' + CellNumber, 'A' + CellNumber].Value := 'ИТОГО по валюте';
             MSEXCEL.ActiveSheet.Range['I' + CellNumber, 'I' + CellNumber].Value := '=SUM(I' + IntToStr(ExcelLineStartPart) + ':I' + IntToStr(ExcelLine - 1) + ')';
             MSEXCEL.ActiveSheet.Range['J' + CellNumber, 'J' + CellNumber].Value := '=SUM(J' + IntToStr(ExcelLineStartPart) + ':J' + IntToStr(ExcelLine - 1) + ')';
             MSEXCEL.ActiveSheet.Range['K' + CellNumber, 'K' + CellNumber].Value := '=SUM(K' + IntToStr(ExcelLineStartPart) + ':K' + IntToStr(ExcelLine - 1) + ')';
             MSEXCEL.ActiveSheet.Range['L' + CellNumber, 'L' + CellNumber].Value := '=SUM(L' + IntToStr(ExcelLineStartPart) + ':L' + IntToStr(ExcelLine - 1) + ')';
             MSEXCEL.ActiveSheet.Range['N' + CellNumber, 'N' + CellNumber].Value := '=SUM(N' + IntToStr(ExcelLineStartPart) + ':N' + IntToStr(ExcelLine - 1) + ')';
             MSEXCEL.ActiveSheet.Range['P' + CellNumber, 'P' + CellNumber].Value := '=SUM(P' + IntToStr(ExcelLineStartPart) + ':P' + IntToStr(ExcelLine - 1) + ')';
             Inc(ExcelLine, 2);
         end;
         MSEXCEL.ActiveWorkbook.SaveAs(GetAppDir + 'DOCS\' + 'FOREIGN' + '\РНП(10)_' + DateToStr(RepDate) + '.XLS', 1);
    except
        on E : Exception do begin
            Application.BringToFront;
            MessageDlg('Возникла ошибка при формировании отчёта'#13 + e.Message, mtInformation, [mbOk], 0);
        end;
    end;

    for i := 0 to ListData.Count - 1 do Dispose(ListData.Items[i]);
    for i := 0 to ReinsurData.Count - 1 do Dispose(ReinsurData.Items[i]);
    ListData.Free;
    ReinsurData.Free;
end;

// _2 потому что максимум может заполнить 2 валюты
// Не перестраховывается
procedure FillStructure_2(RepDate : TDateTime; IsUridich : boolean; FizichTax, UridichTax : double; var Info : ReportStructure_3; InpData : TDateTime; Summa : double; Curr : string; AgPercent : double);
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
        Info[CurrIndex].Otchislenie := Info[CurrIndex].Otchislenie +
                                       AgentSumma;
     end else begin
        CurrIndex := GetCurrIndex(Info, Curr);
        //Дата продажи
        SellData := CalcSellDate(InpData, Curr);
        if SellData >= RepDate then begin //Не продаём
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
        else begin //Успели продать
            SellPercent := GetSellPercent(SellData, Curr);
            SellSumma := Summa * SellPercent / 100;
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             (Summa - SellSumma);
            //Агенту ничего, т.к продали

            CurrIndex := GetCurrIndex(Info, 'BRB');
            Info[CurrIndex].BruttoPremium := Info[CurrIndex].BruttoPremium +
                                             SellSumma * GetRateOfCurrency(Curr, SellData);

            AgentSumma := Summa * AgPercent / 100;
            if IsUridich then
                AgentSumma := AgentSumma * (1 + UridichTax / 100)
            else
                AgentSumma := AgentSumma * (1 + FizichTax / 100);

            Info[CurrIndex].Otchislenie := Info[CurrIndex].Otchislenie +
                                           AgentSumma * GetRateOfCurrency(Curr, InpData);
        end;
     end;
end;

procedure InitReinsurData(var Data : PGroupRestrax);
begin
      Data.Curr := '';
      Data.Number := -1;
      Data.DateFrom := 0;
      Data.DateTo := 0;
      Data.BruttoPremium := 0;
      Data.Comis := 0;
      Data.BasePremium := 0;
      Data.ReserveNZPSumm := 0;
      Data.RestraxPart := 0;
      Data.StraxPart := 0;
end;

//Возвращает, попал ли полис в перестрахование или нет
function DecideIsReinsur(RepDate, PayDate : TDateTime) : boolean;
begin
     InsForeign.IsReinsurSQL.ParamByName('RepDate').AsDateTime := RepDate;
     InsForeign.IsReinsurSQL.ParamByName('PayDate').AsDateTime := PayDate;
     InsForeign.IsReinsurSQL.Open;
     DecideIsReinsur := not InsForeign.IsReinsurSQL.EOF;
     InsForeign.IsReinsurSQL.Close;
end;
}
procedure TInsForeign.FillDocsData(var DataList, ReinsurList : TList; RepDate : TDateTime; FizichTax, UridichTax, OutPercent, RestraxComis : double);
begin
end;
{var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : ^ReportStructure;
     IsNotReinsurPolis : boolean;

     ReinsurData : PGroupRestrax;
     ReinsurDataBRB : PGroupRestrax;

     PayPart : double;

     SellDate : TDateTime;

     SellValuta : double;

     DurationIns, DurationsToRepDate : double;

     AgentSumma : double;
begin
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT');
          Add('   Seria, Number, RegDate, DateFrom, DateTo, PaySumma, Curr, PayDate, Uridich, AgPercent');
          Add('FROM');
          Add('   "FOREIGN"');
          Add('WHERE');
          Add('   PayDate is not null AND');
          Add('   DateFrom < :RepDate AND DateTo >= :RepDate');
          Add('ORDER BY');
          Add('   Seria, Number');
          ParamByName('RepDate').AsDateTime := TRUNC(RepDate);
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

     // Fill Structure
         while not WorkSQL.EOF do begin
             Summa := 0;
             InitStructure_3(Info);
             ReinsurData := nil;
             ReinsurDataBRB := nil;

             IsNotReinsurPolis := (WorkSQL.FieldByName('DateTo').AsDateTime - WorkSQL.FieldByName('DateFrom').AsDateTime + 1) <= 30;
             if not IsNotReinsurPolis then begin
                 //Проверить, отдали в перестрахование или нет
                 IsNotReinsurPolis := not DecideIsReinsur(RepDate, WorkSQL.FieldByName('PayDate').AsDateTime);
             end;

             if IsNotReinsurPolis then begin
                  with WorkSQL do
                      FillStructure_2(RepDate, FieldByName('Uridich').AsBoolean, FizichTax, UridichTax, Info, FieldByName('PayDate').AsDateTime, FieldByName('PaySumma').AsFloat, FieldByName('Curr').AsString, FieldByName('AgPercent').AsFloat);

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
             end
             else begin
                with WorkSQL do begin
                   DurationIns := Trunc(FieldByName('DateTo').AsDateTime) - Trunc(FieldByName('DateFrom').AsDateTime) + 1;
                   DurationsToRepDate := Trunc(RepDate) - FieldByName('DateFrom').AsDateTime;
                   if FieldByName('Curr').AsString <> 'BRB' then begin
                      New(ReinsurData);
                      New(ReinsurDataBRB);
                      InitReinsurData(ReinsurData);
                      InitReinsurData(ReinsurDataBRB);

                      ReinsurData.Curr := FieldByName('Curr').AsString;
                      ReinsurDataBRB.Curr := 'BRB';

                      //Полис отдан в перестрахованиу
                      ReinsurData.BruttoPremium := FieldByName('PaySumma').AsFloat;
                      ReinsurData.Comis := 0;
                      //Сколько перечислили денег (с учётом процента)
                      PayPart := (ReinsurData.BruttoPremium * OutPercent / 100) / (1 - RestraxComis / 100);
                      //Дата продажи валюты
                      SellDate := CalcSellDate(FieldByName('PayDate').AsDateTime,  ReinsurData.Curr);
                      //Размер проданной валюты
                      SellValuta := (ReinsurData.BruttoPremium -
                                    PayPart) * GetSellPercent(SellDate, ReinsurData.Curr) / 100;

                      //Базовая премия
                      ReinsurData.BasePremium := ReinsurData.BruttoPremium -
                                                 SellValuta -
                                                 PayPart;

                      AgentSumma := ReinsurData.BruttoPremium * FieldByName('AgPercent').AsFloat / 100;
                      if WorkSQL.FieldByName('Uridich').AsString = 'Y' then
                          AgentSumma := AgentSumma * (1 + UridichTax / 100)
                      else
                          AgentSumma := AgentSumma * (1 + FizichTax / 100);

                      ReinsurData.ReserveNZPSumm := ReinsurData.BasePremium * (DurationIns - DurationsToRepDate) / DurationIns;

                      ReinsurData.RestraxPart := (ReinsurData.BruttoPremium * OutPercent / 100) /
                                                 (ReinsurData.BRUTTOPremium - AgentSumma) * ReinsurData.ReserveNZPSumm;

                      ReinsurData.StraxPart := ReinsurData.ReserveNZPSumm - ReinsurData.RestraxPart;

                      //РУБЛЁВАЯ ЧАСТЬ
                      //Продажа идёт в день прихода денег
                      ReinsurDataBRB.BruttoPremium := SellValuta * GetRateOfCurrency(ReinsurData.Curr, SellDate);
                      ReinsurDataBRB.Comis := AgentSumma * GetRateOfCurrency(ReinsurData.Curr, FieldByName('PayDate').AsDateTime);
                      //if WorkSQL.FieldByName('Uridich').AsString = 'Y' then
                      //    ReinsurDataBRB.Comis := ReinsurDataBRB.Comis * (1 + UridichTax / 100)
                      //else
                      //    ReinsurDataBRB.Comis := ReinsurDataBRB.Comis * (1 + FizichTax / 100);
                      ReinsurDataBRB.BasePremium := ReinsurDataBRB.BruttoPremium - ReinsurDataBRB.Comis;
                  end
                  else begin
                      New(ReinsurDataBRB);
                      InitReinsurData(ReinsurDataBRB);
                      ReinsurDataBRB.Curr := 'BRB';
                      ReinsurDataBRB.BruttoPremium := FieldByName('PaySumma').AsFloat;
                      ReinsurDataBRB.Comis := ReinsurDataBRB.BruttoPremium * FieldByName('AgPercent').AsFloat / 100;
                      if FieldByName('Uridich').AsString = 'Y' then
                          ReinsurDataBRB.Comis := ReinsurDataBRB.Comis * (1 + UridichTax / 100)
                      else
                          ReinsurDataBRB.Comis := ReinsurDataBRB.Comis * (1 + FizichTax / 100);

                      PayPart := (ReinsurDataBRB.BruttoPremium * OutPercent / 100) / (1 - RestraxComis / 100);
                      ReinsurDataBRB.BasePremium := ReinsurDataBRB.BruttoPremium -
                                                    ReinsurDataBRB.Comis -
                                                    PayPart;
                      ReinsurDataBRB.ReserveNZPSumm := ReinsurDataBRB.BasePremium * (DurationIns - DurationsToRepDate) / DurationIns;

                      ReinsurDataBRB.RestraxPart := (ReinsurDataBRB.BruttoPremium * OutPercent / 100) /
                                                 (ReinsurDataBRB.BRUTTOPremium - AgentSumma) * ReinsurDataBRB.ReserveNZPSumm;

                      ReinsurDataBRB.StraxPart := ReinsurDataBRB.ReserveNZPSumm - ReinsurDataBRB.RestraxPart;
                  end;

                  if ReinsurData <> nil then begin
                      ReinsurData.Number := FieldByName('Number').AsFloat;
                      ReinsurData.DateFrom := FieldByName('DateFrom').AsDateTime;
                      ReinsurData.DateTo := FieldByName('DateTo').AsDateTime;
                      ReinsurList.Add(ReinsurData);
                  end;
                  if ReinsurDataBRB <> nil then begin
                      ReinsurDataBRB.Number := FieldByName('Number').AsFloat;
                      ReinsurDataBRB.DateFrom := FieldByName('DateFrom').AsDateTime;
                      ReinsurDataBRB.DateTo := FieldByName('DateTo').AsDateTime;
                      ReinsurList.Add(ReinsurDataBRB);
                  end;

                end; //With
             end;

             WorkSQL.Next;
         end;
end;
}
procedure TInsForeign.SNLockClick(Sender: TObject);
begin
     SNLock.Checked := not SNLock.Checked;

     if SNLock.Checked then
        ForeignGrid.FixedCols := 2
     else
        ForeignGrid.FixedCols := 0;
end;

procedure TInsForeign.N14Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по инострано туристо';
    About.ShowModal
end;

procedure TInsForeign.N15Click(Sender: TObject);
begin
    if ForeignRpt = nil then begin
        Application.CreateForm(TForeignRpt, ForeignRpt);
    end;
    ForeignRpt.frReport.Dictionary.Variables['ITOGO'] := '''ИТОГО: ' + Memo.Lines[0] + '''';
    ForeignRpt.Show;
    ForeignRpt.BringToFront;

    ForeignGrid.DataSource := nil;
    MainQuery.First;
    ForeignRpt.frReport.ShowReport;
    ForeignGrid.DataSource := DataSource;
end;

end.
