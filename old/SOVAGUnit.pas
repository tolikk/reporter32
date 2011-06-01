unit SOVAGUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, Menus, Db, DBTables, RxQuery, Grids, DBGrids, RXDBCtrl,
  StdCtrls, Buttons, ComCtrls, CheckLst, Placemnt;

type
  TSOVAG = class(TForm)
    MainMenu: TMainMenu;
    Panel: TPanel;
    N1: TMenuItem;
    CompanyMenu: TMenuItem;
    SortMenu: TMenuItem;
    FilterMenu: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    MainQuery: TRxQuery;
    MainQuerySeria: TStringField;
    MainQueryNumber: TFloatField;
    MainQueryPSeria: TStringField;
    MainQueryPNumber: TFloatField;
    MainQueryPlace: TStringField;
    MainQueryRegDate: TDateField;
    MainQueryRegTime: TStringField;
    MainQueryOwner: TStringField;
    MainQueryAddr: TStringField;
    MainQueryDtFrom: TDateField;
    MainQueryDtTo: TDateField;
    MainQueryPeriod: TStringField;
    MainQueryAutoType: TFloatField;
    MainQueryLetter: TStringField;
    MainQueryAutoNmb: TStringField;
    MainQueryModel: TStringField;
    MainQueryTarif: TFloatField;
    MainQueryPayDate: TDateField;
    MainQueryPay1: TFloatField;
    MainQueryPay1C: TStringField;
    MainQueryPay2: TFloatField;
    MainQueryPay2C: TStringField;
    MainQueryAgCode: TStringField;
    MainQueryPcnt: TFloatField;
    MainQueryState: TFloatField;
    MainQuerySeria2: TStringField;
    MainQueryNumber2: TFloatField;
    MainQueryAutoNmb2: TStringField;
    MainQueryModel2: TStringField;
    DataSource: TDataSource;
    CounterMenu: TMenuItem;
    Timer: TTimer;
    BitBtn1: TBitBtn;
    ExcelMenu: TMenuItem;
    N2: TMenuItem;
    PageControl: TPageControl;
    ListTabSheet: TTabSheet;
    FilterTabSheet: TTabSheet;
    RxDBGrid: TRxDBGrid;
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
    IsSeria: TCheckBox;
    Label11: TLabel;
    AgentList: TCheckListBox;
    ListTemplates: TComboBox;
    SpeedButton: TSpeedButton;
    Label5: TLabel;
    Label6: TLabel;
    IsNumber: TCheckBox;
    NumberFrom: TEdit;
    NumberTo: TEdit;
    AgentsTbl: TQuery;
    AgentsTblAgent_code: TStringField;
    AgentsTblName: TStringField;
    AgentsTblSovag: TFloatField;
    Button1: TButton;
    MainQueryScepka: TStringField;
    MainQueryVState: TStringField;
    MainQueryVPeriod: TStringField;
    MainQueryPSN: TStringField;
    SOVAGTable: TTable;
    EditSERIA: TComboBox;
    FormStorage: TFormStorage;
    IsState: TCheckBox;
    State: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure SortMenuClick(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ExcelMenuClick(Sender: TObject);
    procedure FilterMenuClick(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure ListTemplatesChange(Sender: TObject);
  private
    { Private declarations }
    W : HWND;
    AgCodes : TStringList;
    procedure ReOpenQuery;
    function GetFilter(sCond : string) : string;
  public
    { Public declarations }
  end;

var
  SOVAG: TSOVAG;

implementation

uses Sort, mf, WaitFormUnit, polandunit, datamod, Excel97, Variants;//_TLB;

{$R *.DFM}

type
    dblList = array[1..10] of double;
    currList = array[1..10] of string[3];


procedure TSOVAG.FormCreate(Sender: TObject);
var
     R : TRect;
begin
     PageControl.ActivePage := ListTabSheet;
     AgCodes := TStringList.Create;
     AgentsTbl.Open;
     while not AgentsTbl.EOF do begin
           AgentList.Items.Add(AgentsTblName.AsString);
           AgCodes.Add(AgentsTblAgent_code.AsString);
           AgentsTbl.Next;
     end;
     AgentsTbl.Close;

     CompanyMenu.Caption := COMPAMYNAME;
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

procedure TSOVAG.FormDestroy(Sender: TObject);
begin
     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TSOVAG.SortMenuClick(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then begin
         ReOpenQuery;
     end;
end;

procedure TSOVAG.ReOpenQuery;
var
     sqlOrder : string;
begin
     MainQuery.Close;
     MainQuery.MacroByName('WHERE').AsString := GetFilter(' WHERE ');
     sqlOrder := SortForm.GetOrder('');
     if sqlOrder <> '' then
         sqlOrder := 'ORDER BY ' + sqlOrder;
     MainQuery.MacroByName('ORDER').AsString := sqlOrder;
     CounterMenu.Caption := 'Жди...';
     try
        Screen.Cursor := crHourGlass;
        MainQuery.Open;
        CounterMenu.Caption := 'Количество записей ' + IntToStr(MainQuery.RecordCount);
     except
        on E : Exception do begin
            CounterMenu.Caption := 'Ошибка';
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
        end;
     end;
     Screen.Cursor := crDefault;
end;


procedure TSOVAG.N5Click(Sender: TObject);
begin
     Close
end;

procedure TSOVAG.TimerTimer(Sender: TObject);
begin
     Timer.Enabled := false;
     ReOpenQuery
end;

procedure TSOVAG.BitBtn1Click(Sender: TObject);
begin
     Close
end;

procedure Connect(var s: string; cond : string);
begin
     if s <> '' then
         s := s + ' AND ';
     s := s + cond;
end;

function TSOVAG.GetFilter(sCond : string) : string;
var
     s, str, s2, s3 : string;
     i : integer;
begin
     s2 := '';
     s3 := '';

     if IsSeria.Checked then begin
        Connect(s, 'SERIA LIKE ''%' + EditSERIA.Text + '%''');
     end;

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
     end
     else begin
         if IsPayDate.Checked then begin
            if PayDateFrom.Checked then
                Connect(s2, 'PAYDATE>=''' + DateToStr(PayDateFrom.Date) + '''');
            if PayDateTo.Checked then
                Connect(s2, 'PAYDATE<=''' + DateToStr(PayDateTo.Date) + '''');

            //Для испорченного
            if PayDateFrom.Checked then
                Connect(s3, 'REGDATE>=''' + DateToStr(PayDateFrom.Date) + '''');
            if PayDateTo.Checked then
                Connect(s3, 'REGDATE<=''' + DateToStr(PayDateTo.Date) + '''');

            s2 := '(' + s2 + ' AND STATE<>2 OR ' + s3 + ' AND STATE=2)';
         end;
     end;

{     if IsStartDate.Checked then begin
        if StartDateFrom.Checked then
            Connect(s, 'DTFROM>=''' + DateToStr(StartDateFrom.Date) + '''');
        if StartDateTo.Checked then
            Connect(s, 'DTTO<=''' + DateToStr(StartDateTo.Date) + '''');
     end;
}
     if IsPayDate.Checked AND IsRegDate.Checked then begin
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

     if IsState.Checked then begin
         Connect(s, '(State = ' + IntToStr(State.ItemIndex) + ')');
     end;

{     if IsOwner.Checked AND (OwnerLike.Text <> '') then begin
         Connect(s, '(NAME LIKE ''%' + OwnerLike.Text + '%'' OR NAME LIKE ''%' + AnsiUpperCase(OwnerLike.Text) + '%'' OR NAME LIKE ''%' + AnsiLowerCase(OwnerLike.Text) + '%'')');
     end;

     if IsAutoNumber.Checked AND (AutoNumberLike.Text <> '') then begin
         Connect(s, '(AutoNmb LIKE ''%' + AutoNumberLike.Text + '%'' OR AutoNmb LIKE ''%' + AnsiUpperCase(AutoNumberLike.Text) + '%'' OR AutoNmb LIKE ''%' + AnsiLowerCase(AutoNumberLike.Text) + '%'')');
     end;
}

     if s2 <> '' then
         Connect(s, s2);

     if s <> '' then s := sCond + s;

     GetFilter := s;
end;

procedure TSOVAG.Button1Click(Sender: TObject);
begin
     ReOpenQuery;
end;

procedure AddVal(var Summs : dblList; var SummVals : currList; Sum : double; Curr : string);
var
    i : integer;
begin
    for i := 1 to 10 do begin
        if SummVals[i] = Curr then begin
           Summs[i] := Summs[i] + Sum;
           exit;
        end;
        if SummVals[i] = '' then
            break;
    end;

    Summs[i] := Sum;
    SummVals[i] := Curr;
end;

procedure TSOVAG.ExcelMenuClick(Sender: TObject);
var
     MSEXCEL, ASH, VAL : VARIANT;
     i, j, Line : integer;
     strI : string;
     s, s1, v : string;
     c : char;
     A1 : string;
     Timer : integer;
     pay_str : string;
     Summs : dblList;
     SummVals : currList;
     IsDupFnd : boolean;
     IsExcel2000 : boolean;
begin
     for i := 1 to 10 do begin
         Summs[i] := 0;
         SummVals[i] := '';
     end;

     try
         if not InitExcel(MSEXCEL) then exit;
         MSEXCEL.Visible := true;

         MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\SOVAG.XLS');

         v := MSEXCEL.Version;
         IsExcel2000 := StrToFloat(v[1]) >= 9;

         ASH := MSEXCEL.ActiveSheet;
         MainQuery.First;
         i := 6;
         Line := 1;
         A1 := ASH.Range['A1', 'A1'].Value;
         while not MainQuery.EOF do begin
           if (Timer MOD 2) = 1 then
               ASH.Range['A1', 'A1'].Value := 'ЖДИТЕ...'
           else
               ASH.Range['A1', 'A1'].Value := '';

           Inc(Timer);

           if MainQueryPSeria.AsString <> '' then begin //Дубликат пробросить
               MainQuery.Next;
               //Inc(i);
               //Inc(Line);
               continue;
           end;

           strI := IntToStr(i);
           ASH.Range['A' + strI, 'A'+ strI].Value := IntToStr(Line);
           ASH.Range['B' + strI, 'B'+ strI].Value := MainQuery.FieldByName('NUMBER').AsString;// + MainQuery.FieldByName('SERIA').AsString;

           if MainQueryState.AsInteger = 2 then begin
               ASH.Range['C' + strI, 'F'+ strI].MergeCells := true;
               ASH.Range['C' + strI, 'F'+ strI].HorizontalAlignment := xlCenter;
               ASH.Range['C' + strI, 'C'+ strI].Value := 'ИСПОРЧЕН';
           end
           else begin
               ASH.Range['C' + strI, 'C'+ strI].Value := MainQuery.FieldByName('REGDATE').AsString;
               ASH.Range['D' + strI, 'D'+ strI].Value := MainQuery.FieldByName('OWNER').AsString;
               ASH.Range['E' + strI, 'E'+ strI].Value := MainQuery.FieldByName('AUTONMB').AsString;
               ASH.Range['F' + strI, 'F'+ strI].Value := MainQuery.FieldByName('MODEL').AsString;
               ASH.Range['G' + strI, 'G'+ strI].Value := MainQuery.FieldByName('DTFROM').AsString + '-' + MainQuery.FieldByName('DTTO').AsString;
               ASH.Range['H' + strI, 'H'+ strI].Value := MainQuery.FieldByName('LETTER').AsString;
               ASH.Range['I' + strI, 'I'+ strI].Value := MainQuery.FieldByName('TARIF').AsString;
               if IsExcel2000 then
                   ASH.Range['I' + strI, 'I'+ strI].NumberFormat := '[$€-2] #,##0.00';
               pay_str := MainQueryPay1.AsString + ' ' + MainQueryPay1C.AsString;
               AddVal(Summs, SummVals, MainQueryPay1.AsFloat, MainQueryPay1C.AsString);
               if MainQueryPay2C.AsString <> '' then begin
                   pay_str := pay_str + '; ' + MainQueryPay2.AsString + ' ' + MainQueryPay2C.AsString;
                   AddVal(Summs, SummVals, MainQueryPay2.AsFloat, MainQueryPay2C.AsString);
               end;
               ASH.Range['J' + strI, 'J'+ strI].Value := pay_str;

               ASH.Range['C' + strI, 'C'+ strI].HorizontalAlignment := -4108;
               ASH.Range['D' + strI, 'D'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['E' + strI, 'E'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['F' + strI, 'F'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['G' + strI, 'G'+ strI].HorizontalAlignment := -4108;
               ASH.Range['H' + strI, 'H'+ strI].HorizontalAlignment := -4108;
           end;

           Inc(i);

           if MainQuery.FieldByName('SERIA2').AsString <> '' then begin
               strI := IntToStr(i);
               ASH.Range['B' + strI, 'B'+ strI].Value := 'Парный полис прицепа ' + MainQuery.FieldByName('SERIA2').AsString + '/' + MainQuery.FieldByName('NUMBER2').AsString;
               ASH.Range['B' + strI, 'D'+ strI].MergeCells := true;
               ASH.Range['B' + strI, 'D'+ strI].HorizontalAlignment := -4108; //center
               for c := 'E' to 'K' do begin
                   ASH.Range[c + strI, c + IntToStr(i - 1)].MergeCells := true;
                   ASH.Range[c + strI, c + IntToStr(i - 1)].VerticalAlignment := xlCenter;
               end;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].MergeCells := true;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].VerticalAlignment := xlCenter;

               Inc(i);
               strI := IntToStr(i);
               Inc(Line);
               ASH.Range['A' + strI, 'A'+ strI].Value := IntToStr(Line);
               ASH.Range['B' + strI, 'B'+ strI].Value := MainQuery.FieldByName('NUMBER2').AsString;// + MainQuery.FieldByName('SERIA').AsString;
               ASH.Range['C' + strI, 'C'+ strI].Value := MainQuery.FieldByName('REGDATE').AsString;
               ASH.Range['D' + strI, 'D'+ strI].Value := MainQuery.FieldByName('OWNER').AsString;
               ASH.Range['E' + strI, 'E'+ strI].Value := MainQuery.FieldByName('AUTONMB2').AsString;
               ASH.Range['F' + strI, 'F'+ strI].Value := MainQuery.FieldByName('MODEL2').AsString;
               ASH.Range['G' + strI, 'G'+ strI].Value := MainQuery.FieldByName('DTFROM').AsString + '-' + MainQuery.FieldByName('DTTO').AsString;

               ASH.Range['C' + strI, 'C'+ strI].HorizontalAlignment := -4108;
               ASH.Range['D' + strI, 'D'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['E' + strI, 'E'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['F' + strI, 'F'+ strI].HorizontalAlignment := xlLeft;
               ASH.Range['G' + strI, 'G'+ strI].HorizontalAlignment := -4108;
               ASH.Range['H' + strI, 'H'+ strI].HorizontalAlignment := -4108;

               s := MainQuery.FieldByName('LETTER').AsString;
               s1 := '';
               for j := 1 to Length(s) do
                   s1 := s[j] + s1;
               ASH.Range['H' + strI, 'H'+ strI].Value := s1;
               Inc(i);
               strI := IntToStr(i);
               ASH.Range['B' + strI, 'B'+ strI].Value := 'Парный полис прицепа ' + MainQuery.FieldByName('SERIA').AsString + '/' + MainQuery.FieldByName('NUMBER').AsString;
               ASH.Range['B' + strI, 'D'+ strI].MergeCells := true;
               ASH.Range['B' + strI, 'D'+ strI].HorizontalAlignment := -4108; //center
               ASH.Range['B' + strI, 'D'+ strI].VerticalAlignment := -4107; //vcenter
               for c := 'E' to 'K' do begin
                   ASH.Range[c + strI, c + IntToStr(i - 1)].MergeCells := true;
                   ASH.Range[c + strI, c + IntToStr(i - 1)].VerticalAlignment := xlCenter;
               end;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].MergeCells := true;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].VerticalAlignment := xlCenter;
               Inc(i);
           end
           else
           if MainQueryState.AsInteger = 1 then begin //Утерян
               strI := IntToStr(i);
               SOVAGTable.Open;
               IsDupFnd := true;
               if not SOVAGTable.Locate('PSERIA;PNUMBER', vararrayof([MainQuery.FieldByName('SERIA').AsString, MainQuery.FieldByName('NUMBER').AsInteger]), []) then begin
                   ASH.Range['B' + strI, 'B'+ strI].Value := 'Дубликат НЕ НАЙДЕН';
                   IsDupFnd := false;
               end
               else
                   ASH.Range['B' + strI, 'B'+ strI].Value := 'Дубликат ' + SOVAGTable.FieldByName('SERIA').AsString + '/' + SOVAGTable.FieldByName('NUMBER').AsString;

               ASH.Range['B' + strI, 'D'+ strI].MergeCells := true;
               ASH.Range['B' + strI, 'D'+ strI].HorizontalAlignment := -4108; //center
               for c := 'E' to 'K' do begin
                   ASH.Range[c + strI, c + IntToStr(i - 1)].MergeCells := true;
                   ASH.Range[c + strI, c + IntToStr(i - 1)].VerticalAlignment := xlCenter;
               end;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].MergeCells := true;
               ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].VerticalAlignment := xlCenter;
               Inc(i);

               if IsDupFnd then begin
                   strI := IntToStr(i);
                   Inc(Line);
                   ASH.Range['A' + strI, 'A'+ strI].Value := IntToStr(Line);
                   ASH.Range['B' + strI, 'B'+ strI].Value := SOVAGTable.FieldByName('NUMBER').AsString;
                   ASH.Range['C' + strI, 'C'+ strI].Value := SOVAGTable.FieldByName('REGDATE').AsString;
                   ASH.Range['D' + strI, 'D'+ strI].Value := SOVAGTable.FieldByName('OWNER').AsString;
                   ASH.Range['E' + strI, 'E'+ strI].Value := SOVAGTable.FieldByName('AUTONMB').AsString;
                   ASH.Range['F' + strI, 'F'+ strI].Value := SOVAGTable.FieldByName('MODEL').AsString;
                   ASH.Range['G' + strI, 'G'+ strI].Value := SOVAGTable.FieldByName('DTFROM').AsString + '-' + SOVAGTable.FieldByName('DTTO').AsString;

                   ASH.Range['C' + strI, 'C'+ strI].HorizontalAlignment := -4108;
                   ASH.Range['D' + strI, 'D'+ strI].HorizontalAlignment := xlLeft;
                   ASH.Range['E' + strI, 'E'+ strI].HorizontalAlignment := xlLeft;
                   ASH.Range['F' + strI, 'F'+ strI].HorizontalAlignment := xlLeft;
                   ASH.Range['G' + strI, 'G'+ strI].HorizontalAlignment := -4108;
                   ASH.Range['H' + strI, 'H'+ strI].HorizontalAlignment := -4108;
                   ASH.Range['H' + strI, 'H'+ strI].Value := SOVAGTable.FieldByName('LETTER').AsString;;
                   Inc(i);
                   strI := IntToStr(i);
                   ASH.Range['B' + strI, 'B'+ strI].Value := 'Утерянный ' + SOVAGTable.FieldByName('PSERIA').AsString + '/' + SOVAGTable.FieldByName('PNUMBER').AsString;
                   ASH.Range['B' + strI, 'D'+ strI].MergeCells := true;
                   ASH.Range['B' + strI, 'D'+ strI].HorizontalAlignment := -4108; //center
                   ASH.Range['B' + strI, 'D'+ strI].VerticalAlignment := -4107; //vcenter
                   for c := 'E' to 'K' do begin
                       ASH.Range[c + strI, c + IntToStr(i - 1)].MergeCells := true;
                       ASH.Range[c + strI, c + IntToStr(i - 1)].VerticalAlignment := xlCenter;
                   end;
                   ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].MergeCells := true;
                   ASH.Range['A' + strI, 'A' + IntToStr(i - 1)].VerticalAlignment := xlCenter;
                   Inc(i);
               end;
               SOVAGTable.Close;
           end;

           Inc(Line);
           MainQuery.Next
         end;

         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlInsideVertical].LineStyle := xlContinuous;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlInsideVertical].Weight := xlThin;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlInsideHorizontal].LineStyle := xlContinuous;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlInsideHorizontal].Weight := xlThin;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeLeft].LineStyle := xlContinuous;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeLeft].Weight := xlThin;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeRight].LineStyle := xlContinuous;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeRight].Weight := xlThin;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeBottom].LineStyle := xlContinuous;
         ASH.Range['A6', 'J' + IntToStr(i - 1)].Borders[xlEdgeBottom].Weight := xlThin;

         ASH.Range['A1', 'A1'].Value := A1;
         strI := IntToStr(i);
         ASH.Range['I' + strI, 'I' + strI].Value := '=SUM(I6:I' + IntToStr(i - 1) + ')';

         for Line := 1 to 10 do begin
             if SummVals[Line] = '' then break;
             ASH.Range['J' + IntToStr(i), 'J' + IntToStr(i)].Value := FloatToStr(Summs[Line]) + ' ' + SummVals[Line];
             Inc(i);
         end;

         DeleteFile(GetAppDir + 'Docs\SOVAG.XLS');
         ASH.SaveAs(GetAppDir + 'Docs\SOVAG.XLS');
     except
       on E : Exception do begin
         MessageDlg(E.Message, mtInformation, [mbOk], 0);
       end
     end;

     if WaitForm <> nil then WaitForm.Close;
end;

procedure TSOVAG.FilterMenuClick(Sender: TObject);
begin
    PageControl.ActivePage := FilterTabSheet;
end;

procedure TSOVAG.MainQueryCalcFields(DataSet: TDataSet);
begin
     if MainQuerySeria2.AsString <> '' then
         MainQueryScepka.AsString := MainQuerySeria2.AsString + '/' + MainQueryNumber2.AsString;

     case MainQueryState.AsInteger of
         0 : if MainQueryDtTo.AsDateTime < Date then
                 MainQueryVState.AsString := 'ЗАКОНЧИЛСЯ'
             else
                 MainQueryVState.AsString := 'ДЕЙСТВУЕТ';
         1 : MainQueryVState.AsString := 'УТЕРЯН';
         2 : MainQueryVState.AsString := 'ИСПОРЧЕН';
     end;

     if MainQueryState.AsInteger <> 2 then
         MainQueryVPeriod.AsString := MainQueryDtFrom.AsString + ' - ' + MainQueryDtTo.AsString;

     if MainQueryPSeria.AsString <> '' then
         MainQueryPSN.AsString := MainQueryPSeria.AsString + '/' + MainQueryPNumber.AsString;
end;

procedure TSOVAG.ListTemplatesChange(Sender: TObject);
var
    i : integer;
begin
     //if ListTemplates.ItemIndex = 0 then begin
       for i := 0 to AgentList.Items.Count - 1 do
            AgentList.Checked[i] := false;
       //exit;
     //end;


end;

end.
