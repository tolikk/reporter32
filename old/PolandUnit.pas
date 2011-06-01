unit PolandUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, ExtCtrls, Grids, DBGrids, RXDBCtrl, StdCtrls, ComCtrls, DateUtil, IniFiles,
  Menus, registry, RxQuery, mf, Mask, DBCtrls, RxDBComb, RxLookup, Buttons,
  RXSpin, bdeutils, datamod;

const
  NORMAL_POLIS : string = '0';
  LOST_POLIS : string = '1';
  BAD_POLIS : string = '2';
  BREAK_POLIS : string = '3';

type

  TPolandPolises = class(TForm)
    PolandGrid: TRxDBGrid;
    DBPOLAND: TDatabase;
    MainMenu: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ImportTable: TTable;
    N5: TMenuItem;
    EUROPA1: TMenuItem;
    FilterMenu: TMenuItem;
    PolandQuery: TRxQuery;
    PolandQuerySeria: TStringField;
    PolandQueryNumber: TFloatField;
    PolandQueryStartDate: TDateField;
    PolandQueryPeriod: TStringField;
    PolandQueryStartTime: TStringField;
    PolandQueryEndDate: TDateField;
    PolandQueryAutoNumber: TStringField;
    PolandQueryLetter: TStringField;
    PolandQueryName: TStringField;
    PolandQueryTraifCode: TStringField;
    PolandQueryPremiumVal: TFloatField;
    PolandQueryPremiumPay: TFloatField;
    PolandQueryPremiumCurr: TStringField;
    PolandQueryAgentCode: TStringField;
    PolandQueryAgent: TStringField;
    PolandQueryUpdateDate: TDateField;
    Report2: TMenuItem;
    N7: TMenuItem;
    TechCalcSQL: TRxQuery;
    InsertTable: TTable;
    PolandQueryPremiumPay2: TFloatField;
    PolandQueryPremiumCurr2: TStringField;
    N8: TMenuItem;
    WorkSQL: TQuery;
    MenuCompany: TMenuItem;
    PolandQueryRepDate: TDateField;
    SNLock: TMenuItem;
    PolandQueryAgPercent: TFloatField;
    PolandQueryUridich: TBooleanField;
    DataSourceAgent: TDataSource;
    N9: TMenuItem;
    ReserveMenu: TMenuItem;
    N11: TMenuItem;
    PanelTool: TPanel;
    btnDel: TButton;
    btnAdd: TButton;
    btnFilter: TButton;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N6: TMenuItem;
    N15: TMenuItem;
    Label3: TLabel;
    SeekNumber: TRxSpinEdit;
    SpeedButton1: TSpeedButton;
    Bevel1: TBevel;
    N16: TMenuItem;
    N17: TMenuItem;
    AgentTbl: TQuery;
    SortMenu: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N10: TMenuItem;
    N18: TMenuItem;
    IsDupOwnMnu: TMenuItem;
    CntPolisesMnu: TMenuItem;
    PolandQueryBODYNO: TStringField;
    PolandQueryMARKA: TStringField;
    PolandQueryCHARACT: TFloatField;
    ShowHideFlds: TMenuItem;
    PolandQueryAUTOOWNER: TStringField;
    PrintBtn: TButton;
    N21: TMenuItem;
    N111: TMenuItem;
    N22: TMenuItem;
    PolandQueryPayDate: TDateField;
    PolandSource: TDataSource;
    StatusBar: TStatusBar;
    ProgressBar: TProgressBar;
    btnForma: TButton;
    PolandQueryPNumber: TFloatField;
    PolandQueryState: TStringField;
    PolandTable: TTable;
    PolandQueryStateName: TStringField;
    PolandQueryIsDup: TStringField;
    PolandQueryInsType: TStringField;
    PolandQueryInsHouse: TStringField;
    PolandQueryInsFlat: TStringField;
    PolandQueryEngNmb: TStringField;
    PolandQueryAutoType: TStringField;
    PolandQueryInsCntry: TStringField;
    PolandQueryOWNTYPE: TStringField;
    PolandQueryOWNHOUSE: TStringField;
    PolandQueryOWNFLAT: TStringField;
    PolandQueryOWNCNTRY: TStringField;
    PolandQueryRetDate: TDateField;
    PolandQueryRetSum: TFloatField;
    PolandQueryRetCurr: TStringField;
    N23: TMenuItem;
    PolandQueryinscity: TStringField;
    PolandQueryinsstreet: TStringField;
    PolandQueryowncity: TStringField;
    PolandQueryownstreet: TStringField;
    Button1: TButton;
    PolandQueryauto_vid: TIntegerField;
    AutoVidTbl: TTable;
    DataSourceAutoVid: TDataSource;
    procedure CloseForm(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DelPolis(Sender: TObject);
    procedure NewPolis(Sender: TObject);
    procedure PolandQueryBeforePost(DataSet: TDataSet);
    procedure PolandQueryPeriodSetText(Sender: TField; const Text: String);
    procedure PolandQueryStartDateChange(Sender: TField);
    procedure PolandQueryPeriodChange(Sender: TField);
    procedure PolandQueryLetterSetText(Sender: TField; const Text: String);
    procedure PolandQueryLetterChange(Sender: TField);
    procedure ImportPolises(Sender: TObject);
    procedure PolandGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PolandGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure PolandQueryPremiumCurrSetText(Sender: TField;
      const Text: String);
    procedure EUROPA1Click(Sender: TObject);
    procedure FilterMenuClick(Sender: TObject);
    procedure Report2Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure N8Click(Sender: TObject);
    procedure PolandQueryPayDateSetText(Sender: TField;
      const Text: String);
    procedure SNLockClick(Sender: TObject);
    procedure PolandQueryBeforeInsert(DataSet: TDataSet);
    procedure N11Click(Sender: TObject);
    procedure PolandQueryUridichSetText(Sender: TField;
      const Text: String);
    procedure PolandQueryUridichGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure ReserveMenuClick(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure SortMenuClick(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure N18Click(Sender: TObject);
    procedure PolandQueryAgentCodeChange(Sender: TField);
    procedure IsDupOwnMnuClick(Sender: TObject);
    procedure PolandQueryAfterPost(DataSet: TDataSet);
    procedure ShowHideFldsClick(Sender: TObject);
    procedure PrintBtnClick(Sender: TObject);
    procedure N111Click(Sender: TObject);
    procedure btnFormaClick(Sender: TObject);
    procedure PolandQueryCalcFields(DataSet: TDataSet);
    procedure N23Click(Sender: TObject);
    procedure PolandQueryAfterScroll(DataSet: TDataSet);
    procedure PolandQueryAfterCancel(DataSet: TDataSet);
  private
    { Private declarations }

    MSEXCEL : Variant;
    MSWORD : Variant;

    W : HWnd;
    R : TRECT;
    procedure CalculateDates;
    procedure FillTarif;
    procedure UpdateCountPolises;
    procedure AbortFld(F : TField);
    function GetColumn(F : TField) : TColumn;
  public

    procedure CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List : TASKOneTaxList);
    { Public declarations }
  end;

var
  PolandPolises: TPolandPolises;

procedure TransferToExecl(InsurName : string; RepDate : TDateTime; SubDir : string; ListData : TList; MSEXCEL : VARIANT; IsFast : boolean);
procedure CreateSimpleMagazine(WorkSQL : TQuery; SQL : string; InsTypeName : string);

implementation

uses ImpPoland, BasoPolandUnit, PolandFilterUnit, TechCalc, ShowInfoUnit,
  GetFormulaUnit, BasoMonthRptUnit, KorrectPolandUnit,
  ComObj, CurrRatesUnit, DaysForSellCurrUnit, MandSellCurrUnit, Sort,
  InpRezParams, AboutUnit, WaitFormUnit, Variants, GetRepDateUnit,
  ForeignUnit, LebenUnit, PolPolisUnit1, GetPeriodUnit;


{$R *.DFM}


procedure TPolandPolises.CloseForm(Sender: TObject);
begin
     if PolandQuery.State in [dsEdit, dsInsert] then
        PolandQuery.Post;

     Close
end;

procedure TPolandPolises.FormCreate(Sender: TObject);
var
     IniFile: TRegIniFile;
begin
     ProgressBar.Parent := StatusBar;
     AutoVidTbl.Open;
     ProgressBar.Align := alRight;
     MenuCompany.Caption := COMPAMYNAME;
     Application.Title := Caption;
     IniFile := TRegIniFile.Create('POLAND');
     PolandGrid.RestoreLayoutReg(IniFile);
     IniFile.Free;

     try
         DBPOLAND.Open;
         PolandQuery.Open;
     except
         on E : Exception do begin
             MessageDlg('Ошибка запуска приложения'#13 + E.Message, mtInformation, [mbOK], 0);
             PostQuitMessage(0);
             exit;
         end;
     end;

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

     UpdateCountPolises;
end;

procedure TPolandPolises.FormClose(Sender: TObject;
  var Action: TCloseAction);
var
     IniFile: TRegIniFile;
begin
     if PolandQuery.State in [dsEdit, dsInsert] then
        PolandQuery.Post;

     IniFile := TRegIniFile.Create('POLAND');
     PolandGrid.SaveLayoutReg(IniFile);
     IniFile.Free;

     BdeFlushBuffers;

     PolandQuery.Close;
     DBPOLAND.Close;

     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TPolandPolises.DelPolis(Sender: TObject);
begin
     if PolandQuery.State = dsInsert then
        PolandQuery.Cancel
     else
     if MessageDlg('Вы хотите удалить полис ' + PolandQuerySeria.AsString + '/' + PolandQueryNumber.AsString, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
         PolandQuery.Delete;

     UpdateCountPolises;         
end;

procedure TPolandPolises.NewPolis(Sender: TObject);
var
     Seria : string;
     Agent : string;
     Ownertype : string;
     OwnerCountry : string;
     Owner : string;
     Address : string;
     City : string;
     Ur : boolean;
     AgPer : double;
     Start, Pay, Rep : TDateTime;
     N : integer;
     INI : TIniFile;
begin
     Seria := PolandQuerySeria.AsString;
     Agent := PolandQueryAgentCode.AsString;
     Ur := PolandQueryUridich.AsBoolean;
     Ownertype := PolandQueryOWNTYPE.AsString;
     OwnerCountry := PolandQueryOWNCNTRY.AsString;
     Owner := PolandQueryName.AsString;
     AgPer := PolandQueryAgPercent.AsFloat;
     Address := PolandQueryownstreet.AsString;
     City := PolandQueryowncity.AsString;
     Start := PolandQueryStartDate.AsDateTime;
     Pay := PolandQueryPayDate.AsDateTime;
     Rep := PolandQueryRepDate.AsDateTime;
     N := PolandQueryNumber.AsInteger;

     if PolandQuery.State in [dsEdit, dsInsert] then
        PolandQuery.Post;

     PolandQuery.Next;
     PolandQuery.Insert;

     PolandQueryStartTime.AsString := FormatDateTime('hh:nn', Time);
     
     if TControl(Sender).Tag = 1 then begin
           PolandQueryNumber.AsInteger := N + 1;
           PolandQueryAgentCode.AsString := Agent;
           if IsDupOwnMnu.Checked then begin
               PolandQueryName.AsString := Owner;
               PolandQueryownstreet.AsString := Address;
               PolandQueryowncity.AsString := City;
               PolandQueryOWNTYPE.AsString := Ownertype;
               PolandQueryOWNCNTRY.AsString := OwnerCountry;
           end;
           PolandQueryStartDate.AsDateTime := Start;
           PolandQueryPayDate.AsDateTime := Pay;
           PolandQueryRepDate.AsDateTime := Rep;
           PolandQueryUridich.AsBoolean := Ur;
           PolandQueryAgPercent.AsFloat := AgPer;
     end
     else begin
           PolandQueryStartDate.AsDateTime := Date + 1;
           PolandQueryPayDate.AsDateTime := Date;
           PolandQueryRepDate.AsDateTime := Date;
           INI := TIniFile.Create('blank.ini');
           PolandQueryAgentCode.AsString := INI.ReadString('POLAND', 'DefAgent', '');
           PolandQueryOWNTYPE.AsString := Ownertype;
           INI.Free;
     end;

     PolandQuery.FieldByName('State').AsString := NORMAL_POLIS;
     PolandQuery.FieldByName('IsDup').AsString := 'N';
     PolandQuerySeria.AsString := Seria;
     PolandQueryInsCntry.AsString := 'BY';
     PolandQueryOWNCNTRY.AsString := 'BY';
     PolandQueryAuto_Vid.AsInteger := 2;
     PolandGrid.SetFocus;

     PolandQuery.Edit
end;

procedure TPolandPolises.AbortFld(F : TField);
begin
    if F.IsNull then begin
        MessageDlg('Заполни поле ' + F.DisplayName, mtInformation, [mbOk], 0);
        Abort;
    end;
end;

procedure CheckRussian(s : string);
var
    i : integer;
begin
    for i := 1 to Length(s) do begin
        if (s[i] >= 'А') AND (s[i] <= 'Я') OR
           (s[i] >= 'а') AND (s[i] <= 'я') then begin
            MessageDlg('Русские символы в ' + s, mtInformation, [mbOk], 0);
            Abort;
        end;
    end;
end;

procedure TPolandPolises.PolandQueryBeforePost(DataSet: TDataSet);
var
     S, A : string;
     N : integer;
     PayDate : TDate;
     WrDate : TDate;
     IsContinue : boolean;
begin
     if PolandQueryStartDate.IsNull then begin
        PayDate := PolandQueryPayDate.AsDateTime;
        WrDate := PolandQueryRepDate.AsDateTime;
        S := PolandQuerySeria.Value;
        N := PolandQueryNumber.AsInteger;
        A := PolandQueryAgentCode.Value;
        PolandQuery.ClearFields;
        PolandQuerySeria.Value := S;
        PolandQueryNumber.Value := N;
        PolandQueryAgentCode.Value := A;
        PolandQueryName.AsString := 'ИСПОРЧЕН';
        PolandQueryPayDate.AsDateTime := PayDate;
        PolandQueryRepDate.AsDateTime := WrDate;

        {PolandQueryMarka.Clear;
        PolandQueryBodyNo.Clear;
        PolandQueryInsCntry.Clear;
        PolandQueryCHARACT.Clear;
        PolandQueryName.Clear;
        PolandQueryAddress.Clear;
        PolandQueryAUTOOWNER.Clear;
        PolandQueryAUTOOWNERADDR.Clear;
        }
        PolandQueryState.AsString := BAD_POLIS;
     end
     else begin
        CheckRussian(PolandQueryName.AsString);
        CheckRussian(PolandQueryMARKA.AsString);
        CheckRussian(PolandQueryAUTOOWNER.AsString);
        CheckRussian(PolandQueryInsCntry.AsString);
        CheckRussian(PolandQueryInsStreet.AsString);
        CheckRussian(PolandQueryOwnCntry.AsString);
        CheckRussian(PolandQueryOwnStreet.AsString);

        IsContinue := Pos('Продление', PolandQueryPeriod.AsString) <> 0;
        if (PolandQueryRepDate.AsDateTime > PolandQueryStartDate.AsDateTime) then begin
           if PolandQueryIsDup.AsString = 'N' then begin
               MessageDlg('Дата регистрации д.б. до или равна дате начала действия', mtInformation, [mbOK], 0);
               Abort;
           end;
        end;

        if PolandQueryRepDate.AsDateTime <= PolandQueryStartDate.AsDateTime then begin
           if PolandQueryIsDup.AsString = 'Y' then begin
               MessageDlg('Дата регистрации дубликата обычно после даты начала действия', mtInformation, [mbOK], 0);
           end;
        end;

        if PolandQueryStartDate.AsDateTime > PolandQueryEndDate.AsDateTime then begin
           MessageDlg('Дата начала действия д.б. до даты окончания действия', mtInformation, [mbOK], 0);
           Abort;
        end;
        if PolandQueryPremiumPay2.AsFloat < 0.01 then begin
            PolandQueryPremiumPay2.AsFloat := 0;
            PolandQueryPremiumCurr2.AsString := '';
        end
        else
            if PolandQueryPremiumCurr2.AsString = '' then Abort;
        AbortFld(PolandQueryEndDate);
        AbortFld(PolandQueryLetter);
        AbortFld(PolandQueryTraifCode);
        if PolandQueryIsDup.AsString = 'N' then begin
            AbortFld(PolandQueryPremiumVal);
            AbortFld(PolandQueryPremiumPay);
            AbortFld(PolandQueryPremiumCurr)
        end
        else begin
            PolandQueryPremiumVal.Clear;
            PolandQueryPremiumPay.Clear;
            PolandQueryPremiumCurr.Clear;
            PolandQueryPremiumPay2.Clear;
            PolandQueryPremiumCurr2.Clear;
        end;
        AbortFld(PolandQueryAgPercent);
        AbortFld(PolandQueryUridich);
        AbortFld(PolandQueryAutoNumber);
        AbortFld(PolandQueryMarka);
        AbortFld(PolandQueryBodyNo);
        //AbortFld(PolandQueryCOUNTRYSYMB);
        //AbortFld(PolandQueryCHARACT);
        AbortFld(PolandQueryName);
        AbortFld(PolandQueryinscity);
        AbortFld(PolandQueryinsstreet);
        AbortFld(PolandQueryowncity);
        AbortFld(PolandQueryownstreet);
        AbortFld(PolandQueryAUTOOWNER);
        //AbortFld(PolandQueryAUTOOWNERADDR);

        PolandQueryOWNCNTRY.AsString := Copy(PolandQueryOWNCNTRY.AsString, 1, 2);
        PolandQueryInsCntry.AsString := Copy(PolandQueryInsCntry.AsString, 1, 2);

        if (IsContinue) OR (PolandQueryIsDup.AsString = 'Y') then
            if PolandQueryPNumber.AsInteger < 1 then begin
               MessageDlg('Не указано, какой полис продлён/утерян', mtInformation, [mbOK], 0);
               Abort;
            end;

        if PolandQueryPNumber.AsInteger > 0 then begin
            PolandTable.Open;
            if not PolandTable.Locate('Seria;Number', VarArrayOf([PolandQuerySeria.AsString, PolandQueryPNumber.AsInteger]), []) then begin
               if PolandQueryIsDup.AsString = 'Y' then begin
                   MessageDlg('Полис ' + PolandQueryPNumber.AsString + ' не найден', mtInformation, [mbOK], 0);
                   Abort;
               end
               else begin
                   if MessageDlg('Продляемый полис ' + PolandQueryPNumber.AsString + ' не найден. Продолжить?', mtInformation, [mbYes, mbNo], 0) = mrNo then
                       Abort;
               end
            end;

            if (PolandQueryIsDup.AsString = 'Y') AND (PolandTable.FieldByName('State').AsString <> LOST_POLIS) then begin
                if MessageDlg('Полис ' + PolandQueryPNumber.AsString +  ' утерян?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
                    PolandTable.Edit;
                    PolandTable.FieldByName('State').AsString := LOST_POLIS;
                    PolandTable.Post;
                end
                else
                    Abort;
            end;
        end;

        if PolandQueryOWNTYPE.AsString = '' then begin
               MessageDlg('Не указан вид владельца (Физ/Юр лицо)', mtInformation, [mbOK], 0);
               Abort;
        end;

        if PolandQueryInsType.AsString = '' then begin
               MessageDlg('Не указан вид страхователя (Физ/Юр лицо)', mtInformation, [mbOK], 0);
               Abort;
        end;

        if not PolandQueryRetDate.IsNull then begin
            if PolandQueryStartDate.AsDateTime > PolandQueryRetDate.AsDateTime then begin
               MessageDlg('Дата расторжения д.б. после даты начала действия', mtInformation, [mbOK], 0);
               Abort;
            end;
            if PolandQueryEndDate.AsDateTime < PolandQueryRetDate.AsDateTime then begin
               MessageDlg('Дата расторжения д.б. до даты окончания действия', mtInformation, [mbOK], 0);
               Abort;
            end;
            if PolandQueryRetSum.AsFloat <= 0.01 then begin
               MessageDlg('Не указана сумма возврата', mtInformation, [mbOK], 0);
               Abort;
            end;
            if PolandQueryRetCurr.AsString = '' then begin
               MessageDlg('Не указана валюта возврата', mtInformation, [mbOK], 0);
               Abort;
            end;
            PolandQueryState.AsString := BREAK_POLIS;
        end
        else begin
            PolandQueryRetSum.Clear;
            PolandQueryRetCurr.Clear;

            if PolandQueryState.AsString = BREAK_POLIS then
                PolandQueryState.AsString := NORMAL_POLIS;
        end
     end;

     PolandQueryUpdateDate.AsDateTime := Date
end;

procedure TPolandPolises.PolandQueryPeriodSetText(Sender: TField;
  const Text: String);
begin
     if Text = '15' then
        PolandQueryPeriod.AsString := '15 дней'
     else
     if Text = '+15' then
        PolandQueryPeriod.AsString := 'Продление на 15 дней'
     else
     if (Text = '30') then
        PolandQueryPeriod.AsString := '30 дней'
     else
     if (Text = '+30') OR (Text = '+3') then
        PolandQueryPeriod.AsString := 'Продление на 30 дней'
     else
     if Text = '2' then
        PolandQueryPeriod.AsString := '2 месяца'
     else
     if Text = '3' then
        PolandQueryPeriod.AsString := '3 месяца'
     else
     if Text = '4' then
        PolandQueryPeriod.AsString := '4 месяца'
     else
     if Text = '5' then
        PolandQueryPeriod.AsString := '5 месяцев'
     else
     if Text = '6' then
        PolandQueryPeriod.AsString := '6 месяцев'
     else
     if Text = '7' then
        PolandQueryPeriod.AsString := '7 месяцев'
     else
     if Text = '8' then
        PolandQueryPeriod.AsString := '8 месяцев'
     else
     if Text = '9' then
        PolandQueryPeriod.AsString := '9 месяцев'
     else
     if Text = '10' then
        PolandQueryPeriod.AsString := '10 месяцев'
     else
     if Text = '11' then
        PolandQueryPeriod.AsString := '11 месяцев'
     else
     if Text = '1' then
        PolandQueryPeriod.AsString := '1 год'
     else
     if Text = '*2' then
        PolandQueryPeriod.AsString := 'Продление на 2 месяца'
     else
     if Text = '*3' then
        PolandQueryPeriod.AsString := 'Продление на 3 месяца'
     else
     if Text = '*4' then
        PolandQueryPeriod.AsString := 'Продление на 4 месяца'
     else
     if Text = '*5' then
        PolandQueryPeriod.AsString := 'Продление на 5 месяцев'
     else
     if Text = '*6' then
        PolandQueryPeriod.AsString := 'Продление на 6 месяцев'
     else
     if Text = '*7' then
        PolandQueryPeriod.AsString := 'Продление на 7 месяцев'
     else
     if Text = '*8' then
        PolandQueryPeriod.AsString := 'Продление на 8 месяцев'
     else
     if Text = '*9' then
        PolandQueryPeriod.AsString := 'Продление на 9 месяцев'
     else
     if Text = '*10' then
        PolandQueryPeriod.AsString := 'Продление на 10 месяцев'
     else
     if Text = '*11' then
        PolandQueryPeriod.AsString := 'Продление на 11 месяцев'
     else
     if Text = '' then
        PolandQueryPeriod.AsString := ''
     else
        PolandQueryPeriod.AsString := Text;

     if (PolandQueryPeriod.AsString <> '15 дней') AND
        (PolandQueryPeriod.AsString <> 'Продление на 15 дней') AND
        (PolandQueryPeriod.AsString <> '30 дней') AND
        (PolandQueryPeriod.AsString <> 'Продление на 30 дней') AND
        (PolandQueryPeriod.AsString <> '1 год') AND
        (PolandQueryPeriod.AsString <> '2 месяца') AND
        (PolandQueryPeriod.AsString <> '3 месяца') AND
        (PolandQueryPeriod.AsString <> '4 месяца') AND
        (PolandQueryPeriod.AsString <> '5 месяцев') AND
        (PolandQueryPeriod.AsString <> '6 месяцев') AND
        (PolandQueryPeriod.AsString <> '7 месяцев') AND
        (PolandQueryPeriod.AsString <> '8 месяцев') AND
        (PolandQueryPeriod.AsString <> '9 месяцев') AND
        (PolandQueryPeriod.AsString <> '10 месяцев') AND
        (PolandQueryPeriod.AsString <> '11 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 2 месяца') AND
        (PolandQueryPeriod.AsString <> 'Продление на 3 месяца') AND
        (PolandQueryPeriod.AsString <> 'Продление на 4 месяца') AND
        (PolandQueryPeriod.AsString <> 'Продление на 5 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 6 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 7 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 8 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 9 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 10 месяцев') AND
        (PolandQueryPeriod.AsString <> 'Продление на 11 месяцев') AND
        (PolandQueryPeriod.AsString <> '') then begin
        MessageBeep(0);
        PolandQueryPeriod.Clear;
     end
end;

procedure TPolandPolises.CalculateDates;
var
     D, M, Y : integer;
begin
     D := 0;
     Y := 0;
     M := 0;
     if (PolandQueryPeriod.AsString <> '') AND (PolandQueryStartDate.AsString <> '') then begin
        if PolandQueryPeriod.AsString = '15 дней' then D := 15;
        if PolandQueryPeriod.AsString = 'Продление на 15 дней' then D := 15;
        if PolandQueryPeriod.AsString = '30 дней' then M := 1;
        if PolandQueryPeriod.AsString = 'Продление на 30 дней' then M := 1;
        if PolandQueryPeriod.AsString = '1 год' then Y := 1;

        if PolandQueryPeriod.AsString = '2 месяца' then M := 2;
        if PolandQueryPeriod.AsString = '3 месяца' then M := 3;
        if PolandQueryPeriod.AsString = '4 месяца' then M := 4;
        if PolandQueryPeriod.AsString = '5 месяцев' then M := 5;
        if PolandQueryPeriod.AsString = '6 месяцев' then M := 6;
        if PolandQueryPeriod.AsString = '7 месяцев' then M := 7;
        if PolandQueryPeriod.AsString = '8 месяцев' then M := 8;
        if PolandQueryPeriod.AsString = '9 месяцев' then M := 9;
        if PolandQueryPeriod.AsString = '10 месяцев' then M := 10;
        if PolandQueryPeriod.AsString = '11 месяцев' then M := 11;

        if PolandQueryPeriod.AsString = 'Продление на 2 месяца' then M := 2;
        if PolandQueryPeriod.AsString = 'Продление на 3 месяца' then M := 3;
        if PolandQueryPeriod.AsString = 'Продление на 4 месяца' then M := 4;
        if PolandQueryPeriod.AsString = 'Продление на 5 месяцев' then M := 5;
        if PolandQueryPeriod.AsString = 'Продление на 6 месяцев' then M := 6;
        if PolandQueryPeriod.AsString = 'Продление на 7 месяцев' then M := 7;
        if PolandQueryPeriod.AsString = 'Продление на 8 месяцев' then M := 8;
        if PolandQueryPeriod.AsString = 'Продление на 9 месяцев' then M := 9;
        if PolandQueryPeriod.AsString = 'Продление на 10 месяцев' then M := 10;
        if PolandQueryPeriod.AsString = 'Продление на 11 месяцев' then M := 11;

        PolandQueryEndDate.AsDateTime := IncDate(PolandQueryStartDate.AsDateTime, D, M, Y) - 1;
     end;
end;

procedure TPolandPolises.PolandQueryStartDateChange(Sender: TField);
begin
     CalculateDates
end;

procedure TPolandPolises.PolandQueryPeriodChange(Sender: TField);
begin
     CalculateDates;
     FillTarif
end;

procedure TPolandPolises.PolandQueryLetterSetText(Sender: TField;
  const Text: String);
begin
     if Text <> '' then begin
         if Text[1] in ['o', 'O', 'о', 'О'] then
             PolandQueryLetter.AsString := 'O'
         else
         if Text[1] in ['m', 'M', 'м', 'М'] then
             PolandQueryLetter.AsString := 'M'
         else
         if Text[1] in ['p', 'P', 'р', 'Р'] then
             PolandQueryLetter.AsString := 'P'
         else
         if Text[1] in ['c', 'C', 'с', 'С'] then
             PolandQueryLetter.AsString := 'C'
         else
         if Text[1] in ['n', 'N', 'н', 'Н'] then
             PolandQueryLetter.AsString := 'N'
         else
         if Text[1] in ['a', 'A', 'а', 'А'] then begin
             if (Length(Text) > 1) AND (Text[2] <> ' ') then
                 PolandQueryLetter.AsString := 'AT'
             else
                 PolandQueryLetter.AsString := 'A';
         end
         else begin
             MessageBeep(0);
             PolandQueryLetter.AsString := '';
         end
     end
     else
         PolandQueryLetter.AsString := '';
end;

procedure TPolandPolises.FillTarif;
var
    INI : TIniFile;
    s, s2, subkey, subkey2: string;
    divider : integer;
    divider2 : integer;
    tarifCode : string;
    cnt : integer;
    IsContinue : boolean;
begin
    if PolandQueryPeriod.AsString = '' then exit;
    if PolandQueryLetter.AsString = '' then exit;

    subkey := '';
    subkey2 := '';
    cnt := 0;
    if PolandQueryPeriod.AsString = '15 дней' then subkey := '15';
    if PolandQueryPeriod.AsString = 'Продление на 15 дней' then subkey := '+15';
    if PolandQueryPeriod.AsString = '30 дней' then subkey := '30';
    if PolandQueryPeriod.AsString = 'Продление на 30 дней' then subkey := '+30';
    if PolandQueryPeriod.AsString = '1 год' then subkey := '1';

    if PolandQueryPeriod.AsString = '2 месяца' then begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 1;
    end;
    if PolandQueryPeriod.AsString = '3 месяца' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 2;
    end;
    if PolandQueryPeriod.AsString = '4 месяца' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 3;
    end;
    if PolandQueryPeriod.AsString = '5 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 4;
    end;
    if PolandQueryPeriod.AsString = '6 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 5;
    end;
    if PolandQueryPeriod.AsString = '7 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 6;
    end;
    if PolandQueryPeriod.AsString = '8 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 7;
    end;
    if PolandQueryPeriod.AsString = '9 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 8;
    end;
    if PolandQueryPeriod.AsString = '10 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 9;
    end;
    if PolandQueryPeriod.AsString = '11 месяцев' then  begin
        subkey := '+30';
        subkey2 := '30';
        cnt := 10;
    end;

    if PolandQueryPeriod.AsString = 'Продление на 2 месяца' then begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 1;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 3 месяца' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 2;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 4 месяца' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 3;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 5 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 4;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 6 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 5;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 7 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 6;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 8 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 7;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 9 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 8;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 10 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 9;
    end;
    if PolandQueryPeriod.AsString = 'Продление на 11 месяцев' then  begin
        subkey := '+30';
        subkey2 := '+30';
        cnt := 10;
    end;

    IsContinue := Pos('Продление', PolandQueryPeriod.AsString) <> 0;

    INI := TIniFile.Create('blank.ini');
    s := INI.ReadString('POLAND', 'Tarif_' + PolandQueryLetter.AsString + '_' + subkey, '');
    s2 := INI.ReadString('POLAND', 'Tarif_' + PolandQueryLetter.AsString + '_' + subkey2, '');
    divider := Pos(',', s);
    divider2 := Pos(',', s2);
    if (s <> '') and (divider <> 0) then begin
        tarifCode := Copy(s, 1, divider);
        Trim(tarifCode);
        PolandQueryTraifCode.AsString := tarifCode;
        try
            s := Copy(s, divider + 1, 1024); //Сумма тарифа
            Trim(s);

            if cnt > 0 then begin
                s2 := Copy(s2, divider2 + 1, 1024);
                s := FloatToStr(StrToFLoat(s) * cnt + StrToFLoat(s2));
            end;

            if PolandQueryIsDup.AsString = 'Y' then begin
                PolandQueryPremiumVal.Clear;
                PolandQueryPremiumPay.Clear;
                PolandQueryPremiumCurr.Clear;
            end
            else begin
                PolandQueryPremiumVal.AsFloat := StrToFloat(s);
                PolandQueryPremiumPay.AsFloat := StrToFloat(s);
                PolandQueryPremiumCurr.AsString := 'USD';
                PolandGrid.SelectedField := PolandQueryPremiumPay;
            end;
        except
        end
    end else
        MessageDlg('Отсутствует или неправильный тариф для ' + PolandQueryPeriod.AsString, mtInformation, [mbOK], 0);
    INI.Free;
end;


procedure TPolandPolises.PolandQueryLetterChange(Sender: TField);
begin
     FillTarif
end;

function GetPolandCurr(Code : integer) : string;
begin
    if Code = 2 then
        Result := 'USD'
    else
    if Code = 1 then
        Result := 'BRB'
    else
    if Code = 3 then
        Result := 'DM'
    else
        Result := '';
end;

function GetAutoLetter(s : string) : string;
begin
     GetAutoLetter := '';
     case StrToInt(s[2]) of 
          1 : GetAutoLetter := 'O';
          2 : GetAutoLetter := 'A';
          3 : GetAutoLetter := 'AT';
          4 : GetAutoLetter := 'M';
          5 : GetAutoLetter := 'P';
          6 : GetAutoLetter := 'C';
          7 : GetAutoLetter := 'N';
     end;
end;

procedure TPolandPolises.ImportPolises(Sender: TObject);
var
    s : string;
    sn : string;
    ErrPolises : string;
    Line : integer;
    CounterNew : integer;
    CounterUpdate : integer;
    CntRecords : integer;
    IgnoreRecords : integer;
//    hTmpCur : integer;//hDBIObj;
//    TableName : array[0..30] of char;
begin
     if ImportPoland.ShowModal = mrCancel then exit;

     ImportTable.TableName := ImportPoland.FileName.Text;
     ErrPolises := '';
     Line := 0;

     try
         Screen.Cursor := crHourGlass;
         ImportTable.Open;
         InsertTable.Open;

         CounterNew := 0;
         CounterUpdate := 0;
         IgnoreRecords := 0;
         CntRecords := InsertTable.RecordCount;

         while not ImportTable.EOF do begin
             Inc(Line);
             sn := ImportTable.FieldByName('SERIJA').AsString + '/' + ImportTable.FieldByName('NOM_POLIS').AsString;

             //Не зелёная карта
             if ImportTable.FieldByName('IS_ZELKART').AsString <> 'False' then begin
                ImportTable.Next;
                IgnoreRecords := IgnoreRecords + 1;
                continue;
             end;

             if InsertTable.Locate('Seria;Number', VarArrayOf([ImportTable.FieldByName('SERIJA').AsString, ImportTable.FieldByName('NOM_POLIS').AsFloat]), []) then begin
                 if InsertTable.FieldByName('AgentCode').AsString <> ImportPoland.Agents.FieldByName('AGENT_CODE').AsString then
                    raise Exception.Create(sn + #13'Зарегистрирован в БД для другого агента');
                 InsertTable.Edit;
                 CounterUpdate := CounterUpdate + 1;
             end
             else begin
                 InsertTable.Append;
                 CounterNew := CounterNew + 1;
             end;

//             if ImportTable.FieldByName('IS_ZELKART').AsString = 'False' then begin
                InsertTable.FieldByName('Seria').AsString := ImportTable.FieldByName('SERIJA').AsString;
                InsertTable.FieldByName('Number').AsFloat := ImportTable.FieldByName('NOM_POLIS').AsInteger;
                InsertTable.FieldByName('RegDate').AsDateTime := ImportTable.FieldByName('DAT_PODPIS').AsDateTime;
                InsertTable.FieldByName('RepDate').AsDateTime := ImportTable.FieldByName('DAT_REPORT').AsDateTime;
                if ImportTable.FieldByName('SUM_TARIF1').AsFloat > 0.01 then begin
                    InsertTable.FieldByName('StartDate').AsDateTime := ImportTable.FieldByName('DAT_NACH').AsDateTime;
                    InsertTable.FieldByName('StartTime').AsString := IntToStr(ROUND(ImportTable.FieldByName('TIM_NACH').AsInteger / 100)) +
                                                                     ':' +
                                                                     IntToStr(ImportTable.FieldByName('TIM_NACH').AsInteger MOD 100);
                    s := '';
                    if ImportTable.FieldByName('L_PROLONG').AsString = 'True' then
                        s := 'Продление на ';

                    if ImportTable.FieldByName('N_SROK').AsInteger = 28 then
                        InsertTable.FieldByName('Period').AsString := s + '30 дней'
                    else
                    if ImportTable.FieldByName('N_SROK').AsInteger in [15, 30] then
                        InsertTable.FieldByName('Period').AsString := s + ImportTable.FieldByName('N_SROK').AsString + ' дней'
                    else
                    if ImportTable.FieldByName('N_SROK').AsInteger in [16, 31] then
                        InsertTable.FieldByName('Period').AsString := s + IntToStr(ImportTable.FieldByName('N_SROK').AsInteger - 1) + ' дней'
                    else begin
                        MessageDlg(sn + #13'Неизвестный период страхования ' + ImportTable.FieldByName('N_SROK').AsString + #13'Полис будет пропущен', mtInformation, [mbOK], 0);
                        ErrPolises := ErrPolises + sn + ' Строка ' + IntToStr(Line) + #13;
                        InsertTable.Cancel;
                        ImportTable.Next;
                        continue;
                    end;

                    InsertTable.FieldByName('EndDate').AsDateTime := ImportTable.FieldByName('DAT_KONEC').AsDateTime;
                    InsertTable.FieldByName('AutoNumber').AsString := ImportTable.FieldByName('REGIST_NUM').AsString;
                    InsertTable.FieldByName('Letter').AsString := GetAutoLetter(ImportTable.FieldByName('COD_TARIF').AsString);
                    InsertTable.FieldByName('Name').AsString := ImportTable.FieldByName('FIO_CLIENT').AsString;
                    InsertTable.FieldByName('Address').AsString := ImportTable.FieldByName('ADR_CLIENT').AsString;
                    InsertTable.FieldByName('TraifCode').AsString := ImportTable.FieldByName('COD_TARIF').AsString;
                    InsertTable.FieldByName('PremiumVal').AsFloat := ImportTable.FieldByName('SUM_TARIF1').AsFloat;
                    if ImportTable.FieldByName('COD_VALTF1').AsInteger <> 2 then begin
                        raise Exception.Create(sn + #13'Неизвестный вид валюты тарифа');
                    end;

                    InsertTable.FieldByName('PremiumPay').AsFloat := ImportTable.FieldByName('SUM_VZNOS').AsFloat;
                    InsertTable.FieldByName('PremiumPay2').Clear;
                    if ImportTable.FieldByName('SUM_VZNOS2').AsFloat > 0.01 then
                        InsertTable.FieldByName('PremiumPay2').AsFloat := ImportTable.FieldByName('SUM_VZNOS2').AsFloat;

                    InsertTable.FieldByName('PremiumCurr').AsString := GetPolandCurr(ImportTable.FieldByName('COD_VALUT').AsInteger);
                    if InsertTable.FieldByName('PremiumCurr').AsString = '' then
                        raise Exception.Create(sn + #13'Неизвестный вид валюты 1 платежа');

                    if InsertTable.FieldByName('PremiumPay2').AsFloat > 0.01 then begin
                        InsertTable.FieldByName('PremiumCurr2').AsString := GetPolandCurr(ImportTable.FieldByName('COD_VALUT2').AsInteger);
                        if InsertTable.FieldByName('PremiumCurr2').AsString = '' then
                            raise Exception.Create(sn + #13'Неизвестный вид валюты 2 платежа');
                    end;
                end
                else begin
                    InsertTable.FieldByName('NAME').AsString := 'ИСПОРЧЕН';
                end;
                InsertTable.FieldByName('AgentCode').AsString := ImportPoland.Agents.FieldByName('AGENT_CODE').AsString;
                InsertTable.FieldByName('UpdateDate').AsDateTime := Date;
                InsertTable.FieldByName('AgPercent').AsFloat := ImportPoland.AgPercent.Value;
                InsertTable.FieldByName('Uridich').AsBoolean := ImportPoland.IsUridich.Checked;
                InsertTable.Post;
//             end;

             ImportTable.Next;
         end;
     except
         on E : Exception do
             MessageDlg('Ошибка импорта данных'#13'Импорт прерван'#13 + E.Message, mtInformation, [mbOK], 0);
     end;

     if ErrPolises <> '' then begin
        ShowInfo.InfoPanel.Lines.Text := ImportPoland.FileName.Text + #13 + ErrPolises + #13 + 'Исправьте файл и повторите импорт';
        ShowInfo.Caption := 'Пропущенные полисы';
        ShowInfo.ShowModal
     end;

     MessageDlg('Вставлено ' + IntToStr(CounterNew) + #13 +
                'Обновлено ' + IntToStr(CounterUpdate) + #13 +
                'Пропущено из импортир. файла ' + IntToStr(IgnoreRecords) + #13 +
                'В БД было ' + IntToStr(CntRecords) + ' записей' + #13 +
                'В БД стало ' + IntToStr(InsertTable.RecordCount) + ' записей' + #13 +
                'ИТОГО разница ' + IntToStr(InsertTable.RecordCount - CntRecords) + ' записей'
                , mtInformation, [mbOk], 0);

     InsertTable.Close;
     ImportTable.Close;
     PolandQuery.Refresh;
     Screen.Cursor := crDefault;

     UpdateCountPolises;
     BdeFlushBuffers;
end;

procedure TPolandPolises.PolandGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if (Key = Ord('Y')) AND ((GetAsyncKeyState(VK_CONTROL) AND $8000) <> 0) AND ((GetAsyncKeyState(VK_MENU) AND $8000) <> 0) then
         MessageDlg('Привет Настя!!!', mtInformation, [mbOk], 0);
     if (Key = 222) AND ((GetAsyncKeyState(VK_CONTROL) AND $8000) <> 0) AND ((GetAsyncKeyState(VK_MENU) AND $8000) <> 0) then
         MessageDlg('Здарова Эдик!!!', mtInformation, [mbOk], 0);
     if (Key = VK_RETURN) then
         btnFormaClick(nil);

     if (Key = VK_DELETE) AND (ssCtrl in Shift) then begin
         MessageBeep(0);
         Key := 0;
     end
     else
         inherited
end;

procedure TPolandPolises.PolandGridGetCellProps(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor);
begin
     if Field = PolandQueryPremiumVal then
         AFont.Color := $00AA;

     if Field = PolandQueryPremiumPay then
         AFont.Color := $00FF;
     if Field = PolandQueryPremiumCurr then
         AFont.Color := $00FF;

     if Field = PolandQueryPeriod then
         AFont.Color := $00FF0000;
     if Field = PolandQueryLetter then
         AFont.Color := $00FF0000;

     if Field = PolandQueryName then
         AFont.Color := $007700;
end;

procedure TPolandPolises.PolandQueryPremiumCurrSetText(Sender: TField;
  const Text: String);
begin
     if Text <> '' then begin
        if Text[1] in ['e', 'E'] then
            Sender.AsString := 'EUR';
        if Text[1] in ['b', 'B'] then
            Sender.AsString := 'BRB';
        if Text[1] in ['r', 'R'] then
            Sender.AsString := 'RUR';
        if Text[1] in ['u', 'U'] then
            Sender.AsString := 'USD';
        if Text[1] in ['d', 'D'] then
            Sender.AsString := 'DM';
     end
     else
         Sender.AsString := '';
end;

procedure TPolandPolises.EUROPA1Click(Sender: TObject);
begin
    if BasoPolandReportsForm <> nil then exit;

    if BasoPolandReportsForm = nil then begin
        Application.CreateForm(TBasoPolandReportsForm, BasoPolandReportsForm);
    end;

    BasoPolandReportsForm.Show;
    BasoPolandReportsForm.BringToFront;

    BasoPolandReportsForm.frReportEUROPA.Dictionary.Variables['REPORT_PERIOD'] := '';
    if PolandFilter.GetPeriodText <> '' then
        BasoPolandReportsForm.frReportEUROPA.Dictionary.Variables['REPORT_PERIOD'] := '''за период ' + PolandFilter.GetPeriodText + '''';

    PolandGrid.DataSource := nil;
    PolandQuery.First;
    BasoPolandReportsForm.frReportEUROPA.ShowReport;
    PolandGrid.DataSource := PolandSource;
end;

procedure TPolandPolises.FilterMenuClick(Sender: TObject);
begin
     if PolandFilter.ShowModal = mrCancel then exit;
     if PolandQuery.MacroByName('WHERE').AsString <> PolandFilter.FilterText then begin
         try
             Screen.Cursor := crHourGlass;
             PolandQuery.Close;
             PolandQuery.MacroByName('WHERE').AsString := PolandFilter.FilterText;
             PolandQuery.Open;
         except
             on E : Exception do
                 MessageDlg('Ошибка запроса к БД'#13 + E.Message, mtInformation, [mbOK], 0);
         end
     end;
     UpdateCountPolises;
     Screen.Cursor := crDefault; 
end;

procedure TPolandPolises.Report2Click(Sender: TObject);
begin
    if BasoPolandReportsForm <> nil then exit;

    if BasoPolandReportsForm = nil then begin
        Application.CreateForm(TBasoPolandReportsForm, BasoPolandReportsForm);
    end;
    BasoPolandReportsForm.Show;
    BasoPolandReportsForm.BringToFront;
    BasoPolandReportsForm.frReportSVOD.Dictionary.Variables['REPORT_PERIOD'] := '';
    if PolandFilter.GetPeriodText <> '' then
        BasoPolandReportsForm.frReportSVOD.Dictionary.Variables['REPORT_PERIOD'] := '''за период ' + PolandFilter.GetPeriodText + '';

    PolandGrid.DataSource := nil;
    PolandQuery.First;
    BasoPolandReportsForm.frReportSVOD.ShowReport;
    PolandGrid.DataSource := PolandSource;
end;

procedure TPolandPolises.N7Click(Sender: TObject);
begin
     TechCalculation.TextData.Lines.Clear;
     Screen.Cursor := crHourGlass;
     try
         with TechCalcSQL, TechCalcSQL.SQL do begin
              Clear;

              Add('SELECT SUM(PremiumVal) AS S FROM POLANDPL WHERE PremiumVal IS NOT NULL');
              if PolandFilter.FilterText <> '' then begin
                  Add(' AND ');
                  Add(Copy(PolandFilter.FilterText, 7, Length(PolandFilter.FilterText) - 6));
              end;
              Open;
              TechCalculation.Premium := FieldByName('S').AsFloat;
              Close;

              Clear;
              Add('SELECT COUNT(*) AS S FROM POLANDPL WHERE StartDate IS NOT NULL');
              if PolandFilter.FilterText <> '' then begin
                  Add(' AND ');
                  Add(Copy(PolandFilter.FilterText, 7, Length(PolandFilter.FilterText) - 6));
              end;
              Open;
              TechCalculation.Good := FieldByName('S').AsInteger;
              Close;

              Clear;
              Add('SELECT COUNT(*) AS S FROM POLANDPL WHERE StartDate IS NULL');
              if PolandFilter.FilterText <> '' then begin
                  Add(' AND ');
                  Add(Copy(PolandFilter.FilterText, 7, Length(PolandFilter.FilterText) - 6));
              end;
              Open;
              TechCalculation.Bad := FieldByName('S').AsInteger;
              Close;

              TechCalculation.Period := PolandFilter.GetPeriodText;
         end;
         TechCalculation.ShowModal
     except
         on E : Exception do begin
            MessageDlg('Ошибка технического расчёта'#13 + E.Message , mtInformation, [mbOK], 0);
         end;
     end;
     Screen.Cursor := crDefault;
end;

procedure TPolandPolises.N8Click(Sender: TObject);
var
     NewPolises, BadPolises : integer;
     PremiumSum : double;
     Ostatok : integer;
     filtStr : string;
begin
     if not PolandFilter.IsWasActive.Checked then begin
        MessageDlg('Вы не указали в фильтре период, за который необходимо подготовить отчёт', mtInformation, [mbOK], 0);
        FilterMenuClick(Sender);
        if not PolandFilter.IsWasActive.Checked then exit;
     end;

     if BasoPolandReportsForm <> nil then exit;

     if GetFormula.ShowModal <> mrOk then exit;

     if BasoPolandReportsForm = nil then begin
         Application.CreateForm(TBasoPolandReportsForm, BasoPolandReportsForm);
     end;

     PolandFilter.IsWasActive.Checked := false;
     filtStr := Copy(PolandFilter.FilterText, 7, Length(PolandFilter.FilterText));
     PolandFilter.IsWasActive.Checked := true;

     try
         Screen.Cursor := crHourGlass;
         with WorkSQL, WorkSQL.SQL do begin
              Clear;
              Add('SELECT COUNT(*) As CNT, SUM(PREMIUMVAL) AS S FROM PolandPl');
              Add('WHERE PREMIUMVAL IS NOT NULL AND STARTDATE BETWEEN ''' + DateToStr(PolandFilter.WasActiveFrom.Date) + ''' AND ''' + DateToStr(PolandFilter.WasActiveTo.Date) + '''');
              if filtStr <> '' then
                  Add(' AND ' + filtStr);
              Open;
              NewPolises := FieldByName('CNT').AsInteger;
              PremiumSum := FieldByName('S').AsFloat;
              Close;

              Clear;
              Add('SELECT COUNT(*) As CNT FROM PolandPl');
              Add('WHERE PREMIUMVAL IS NULL AND REPDATE BETWEEN ''' + DateToStr(PolandFilter.WasActiveFrom.Date) + ''' AND ''' + DateToStr(PolandFilter.WasActiveTo.Date) + '''');
              if filtStr <> '' then
                  Add(' AND ' + filtStr);
              Open;
              BadPolises := FieldByName('CNT').AsInteger;
              Close;

              Clear;
              Add('SELECT COUNT(*) As CNT FROM PolandPl');
              Add('WHERE PREMIUMVAL IS NOT NULL');// AND STARTDATE BETWEEN ''' + DateToStr(PolandFilter.WasActiveFrom.Date) + ''' AND ''' + DateToStr(PolandFilter.WasActiveTo.Date) + '''');
              Add('AND ENDDATE >= ''' + DateToStr(PolandFilter.WasActiveTo.Date) + '''');
              Add('AND STARTDATE <= ''' + DateToStr(PolandFilter.WasActiveTo.Date) + '''');
              if filtStr <> '' then
                  Add(' AND ' + filtStr);
              Open;
              Ostatok := FieldByName('CNT').AsInteger;
              Close;

//              MessageDlg('Квартальный отчёт за период ' + DateToStr(PolandFilter.WasActiveFrom.Date) + ' - ' + DateToStr(PolandFilter.WasActiveTo.Date) + #13 +
//                         '    Заключено ' + IntToStr(NewPolises) + #13 +
//                         '    Испорчено ' + IntToStr(BadPolises) + #13 +
//                         '    Премия     ' + FloatToStr(PremiumSum) + #13 +
//                         'На ' + DateToStr(PolandFilter.WasActiveTo.Date + 1) + ' действует ' + IntToStr(Ostatok), mtInformation, [mbOK], 0);
              with BasoPolandReportsForm.frReportQuarter.Dictionary do begin
              Variables['Period'] := '''' + DateToStr(PolandFilter.WasActiveFrom.Date) + ' - ' + DateToStr(PolandFilter.WasActiveTo.Date) + '''';
              Variables['Summary1'] := '''' + IntToStr(NewPolises) + '''';
              Variables['Summary2'] := '''' + FloatToStr(PremiumSum) + '''';
              Variables['Summary3'] := '''' + IntToStr(BadPolises) + '''';

              Variables['CheckDate'] := '''' + DateToStr(PolandFilter.WasActiveTo.Date + 1) + '''';
              Variables['Summary1_1'] := '''' + IntToStr(Ostatok) + '''';
              Variables['Summary1_2'] := '[Summary1_1]' + GetFormula.Formula.Text;
              Variables['Summary1+'] := '[Summary1]' + GetFormula.Formula.Text;
              BasoPolandReportsForm.Show;
              BasoPolandReportsForm.BringToFront;
              BasoPolandReportsForm.frReportQuarter.ShowReport;
              end;
         end;
     except
         on E : Exception do begin
            MessageDlg(E.Message, mtInformation, [mbOK], 0);
         end;
     end;
     Screen.Cursor := crDefault;
end;

procedure TPolandPolises.PolandQueryPayDateSetText(Sender: TField;
  const Text: String);
var
     IsOk : boolean;
     T : string;
begin
     if Text = '' then begin
         Sender.Clear;
         exit;
     end;

     T := Text;
     IsOk := false;
     try
         StrToDate(Text);
         IsOk := true;
     except
     end;

     if not IsOk then begin
         try
             while Pos('/', T) <> 0 do
                 T[Pos('/', T)] := '.';
             StrToDate(T);
             Sender.AsString := T;
             IsOk := true;
         except
         end;
     end;

     if not IsOk then begin
         try
             while Pos('.', T) <> 0 do
                 T[Pos('.', T)] := '/';
             StrToDate(T);
             Sender.AsString := T;
             IsOk := true;
         except
         end;
     end;

     if IsOk then
         Sender.AsString := T;
end;

procedure TPolandPolises.SNLockClick(Sender: TObject);
begin
     SNLock.Checked := not SNLock.Checked;

     if SNLock.Checked then
        PolandGrid.FixedCols := 2
     else
        PolandGrid.FixedCols := 0;
end;

procedure TPolandPolises.PolandQueryBeforeInsert(DataSet: TDataSet);
begin
     if SNLock.Checked then
         Abort;

//     if PolPolisData <> nil then
//         PolPolisData.IsEditNumber := true;
end;

procedure TPolandPolises.N11Click(Sender: TObject);
begin
     if KorrectPoland = nil then begin
         Application.CreateForm(TKorrectPoland, KorrectPoland);
     end;

     KorrectPoland.DbName := DBPOLAND.DatabaseName;
     KorrectPoland.TableName := 'POLANDPL';
     KorrectPoland.AgentCodeFld := 'AgentCode';
     KorrectPoland.AgentPercentFld := 'AgPercent';
     KorrectPoland.UridichFizichFld := 'URIDICH';
     KorrectPoland.NULLFld := 'PremiumVal';
     KorrectPoland.IsUridichFldStr := false;
     KorrectPoland.ShowModal;

     PolandQuery.Refresh
end;

procedure TPolandPolises.PolandQueryUridichSetText(Sender: TField;
  const Text: String);
begin
     if Text = '' then begin
         MessageBeep(0);
         exit;
     end;
     if Text[1] in ['ю', 'Ю'] then Sender.AsBoolean := True
     else if Text[1] in ['ф', 'Ф'] then Sender.AsBoolean := False
     else MessageBeep(0);
end;

procedure TPolandPolises.PolandQueryUridichGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
     if not Sender.IsNull then
         if Sender.AsBoolean = true then Text := 'Юридическое'
         else Text := 'Физическое'
end;

procedure TPolandPolises.ReserveMenuClick(Sender: TObject);
var
    List : TASKOneTaxList;
begin
     exit;
{     if InitRezerv.ShowModal = mrOk then begin
         Screen.Cursor := crHourGlass;
         CreateReport(TControl(Sender).Tag = 1,
                      InitRezerv.RepData.Date,
                      InitRezerv.FizichTax1.Value,
                      InitRezerv.UridichTax1.Value,
                      InitRezerv.CompanyPercent.Value,
                      InitRezerv.IsFast.Checked,
                      List);
         Screen.Cursor := crDefault;
         Application.BringToFront;
     end;}
end;

//TACK

//Алгоритм расчёта резерва по польским полисам
//Для всех полисов, у которых дата начала < отчётной и дата окончания больше либо равна отчётной
//Для полисов продолжительностью <= 30 Формы 1, 5
//Для полисов продолжительностью > 30  Формы 1, 2
//Продажа валюты через определённое количество дней (для вида валюты)

//Заполнение Формы 1
//По ВАЛЮТЕ
//Полученная брутто премия = (пришедшая сумма / 2) - % продажи от (пришедшая сумма / 2)
//Отчислений нет
//Вознаграждение агенту только если не продали, т.е. дата оплаты после отчётной - DaysDelta
//Колонка 5 (Базовая) = 2 - 4

//По РУБЛЯМ (от продажи) (если продали)
//Полученная брутто премия = % продажи от (пришедшая сумма / 2)
//Отчислений нет
//Вознаграждение агенту (%от суммы ПРИШЕДШЕЙ!!! * Курс поступления)
//Колонка 5 (Базовая) = 2 - 4

//РУБЛЯМИ ПЛАТИЛИ
//Полученная брутто премия =(пришедшая сумма / 2)
//Вознаграждение агенту = пришедшая сумма * % агенту
//Колонка 5 (Базовая) = 2 - 4

procedure TransferToExecl(InsurName : string; RepDate : TDateTime; SubDir : string; ListData : TList; MSEXCEL : VARIANT; IsFast : boolean);
var
    i, j, idx_curr : integer;
    ExcelLine, ExcelLine2 : integer;
    CurrList : array[1..10] of string[10];
    InfoPtr : PReportStructure;
    OutFileName : string;
    CntCurrs : integer;

    VAL : VARIANT;

    A : integer;
    B, C, D, E : double;
begin
         MSEXCEL.Application.EnableEvents := false;
         for i := 1 to 10 do CurrList[i] := '';

         for i := 0 to ListData.Count - 1 do begin
             InfoPtr := PReportStructure(ListData.Items[i]);
             RoundValue(InfoPtr.BruttoPremium);
             RoundValue(InfoPtr.AgOtchislenie);
             RoundValue(InfoPtr.OneTaxOtchislenie);
             RoundValue(InfoPtr.BasePremium_nouse);
             for j := 1 to 10 do
                 if (CurrList[j] = InfoPtr.Curr) OR (CurrList[j] = '') then begin
                    CurrList[j] := InfoPtr.Curr;
                    break;
                 end;
         end;

         for CntCurrs := 1 to 10 do
             if CurrList[CntCurrs] = '' then break;

         Dec(CntCurrs);

         //ФОРМА 1
         WaitForm.WorkName.Caption := 'ФОРМА 1';
         WaitForm.Update;
         for idx_curr := 1 to 10 do begin
             if CurrList[idx_curr] = '' then break;

             MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\БСП(1).XLS');
             MSEXCEL.ActiveSheet.Range['A4', 'A4'].Value := 'на ' + DateToStr(RepDate);
             MSEXCEL.ActiveSheet.Range['B6', 'B6'].Value := InsurName;
             MSEXCEL.ActiveSheet.Range['B8', 'B8'].Value := CurrList[idx_curr];

             ExcelLine := 16;

             A := 0;
             B := 0;
             C := 0;
             D := 0;
             E := 0;
             //Более 30 дней
             for i := 0 to ListData.Count - 1 do begin
                 WaitForm.ProgressBar.Position := ROUND((idx_curr - 1) * 100 / CntCurrs) +
                                                  ROUND(i * 100 / ListData.Count / 2 / CntCurrs);
                 InfoPtr := PReportStructure(ListData.Items[i]);
                 if (InfoPtr.ToDate - InfoPtr.FromDate + 1) <= 30 then continue;
                 if CurrList[idx_curr] = InfoPtr.Curr then begin
                     if not IsFast then begin
                         MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := FloatToStr(InfoPtr.Number) + ', ' + DateToStr(InfoPtr.RegDate);
                         VAL := InfoPtr.BruttoPremium;
                         MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                         VAL := InfoPtr.OneTaxOtchislenie;
                         MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := VAL;
                         VAL := InfoPtr.AgOtchislenie;
                         MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := VAL;
                         MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=B' + IntToStr(ExcelLine) + '-C' + IntToStr(ExcelLine) + '-D' + IntToStr(ExcelLine);
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) < 0 then begin
                             MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '0';
                             MSEXCEL.ActiveSheet.Range['F' + IntToStr(ExcelLine), 'F' + IntToStr(ExcelLine)].Value := 'Не хватает брутто премии';
                         end;
                         Inc(ExcelLine);
                     end
                     else begin
                         A := A + 1;
                         B := B + InfoPtr.BruttoPremium;
                         C := C + InfoPtr.OneTaxOtchislenie;
                         D := D + InfoPtr.AgOtchislenie;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) > 0 then
                             E := E + (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie)
                     end
                 end;
             end;

             if IsFast AND (A > 0) then begin
                MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := IntToStr(A) + ' полисов';
                VAL := B;
                MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                VAL := C;
                MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := VAL;
                VAL := D;
                MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := VAL;
                VAL := E;
                MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := VAL;
                Inc(ExcelLine);
             end;

             if ExcelLine > 16 then begin
                 MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'ИТОГО';
                 MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '=SUM(B16:B' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := '=SUM(C16:C' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := '=SUM(D16:D' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=SUM(E16:E' + IntToStr(ExcelLine - 1) + ')';

 //                MSEXCEL.ActiveSheet.Range['C16', 'C' + IntToStr(ExcelLine - 1)].NumberFormat := '$#,##0.00';
//                 MSEXCEL.ActiveSheet.Range['C16', 'C' + IntToStr(ExcelLine - 1)].Select;
  //               MSEXCEL.Selection.NumberFormat := '$#,##0.00'
             end;

             if ExcelLine > 16 then
                 Inc(ExcelLine, 2);
             MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'Срок страхования до 30 дней, включительно';
             Inc(ExcelLine);
             ExcelLine2 := ExcelLine;

             //До 30 дней
             A := 0;
             B := 0;
             C := 0;
             D := 0;
             E := 0;
             for i := 0 to ListData.Count - 1 do begin
                 WaitForm.ProgressBar.Position := ROUND((idx_curr - 1) * 100 / CntCurrs) +
                                                  ROUND((i * 100 / ListData.Count / 2 + 0.5) / CntCurrs);
                 InfoPtr := PReportStructure(ListData.Items[i]);
                 if (InfoPtr.ToDate - InfoPtr.FromDate + 1) > 30 then continue;
                 if CurrList[idx_curr] = InfoPtr.Curr then begin
                     if not IsFast then begin
                         MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := FloatToStr(InfoPtr.Number) + ', ' + DateToStr(InfoPtr.RegDate);
                         VAL := InfoPtr.BruttoPremium;
                         MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                         VAL := InfoPtr.OneTaxOtchislenie;
                         MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := VAL;
                         VAL := InfoPtr.AgOtchislenie;
                         MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := VAL;
                         MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=B' + IntToStr(ExcelLine) + '-C' + IntToStr(ExcelLine) + '-D' + IntToStr(ExcelLine);
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) < 0 then begin
                            MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := 0;
                            MSEXCEL.ActiveSheet.Range['F' + IntToStr(ExcelLine), 'F' + IntToStr(ExcelLine)].Value := 'Не хватает брутто премии';
                         end;
                         Inc(ExcelLine);
                     end
                     else begin
                         A := A + 1;
                         B := B + InfoPtr.BruttoPremium;
                         C := C + InfoPtr.OneTaxOtchislenie;
                         D := D + InfoPtr.AgOtchislenie;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) > 0 then
                             E := E + (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie)
                     end
                 end;
             end;

             if IsFast AND (A > 0) then begin
                MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := IntToStr(A) + ' полисов';
                VAL := B;
                MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                VAL := C;
                MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := VAL;
                VAL := D;
                MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := VAL;
                VAL := E;
                MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := VAL;
                Inc(ExcelLine);
             end;

             if ExcelLine > ExcelLine2 then begin
                 MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'ИТОГО';
                 MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '=SUM(B' + IntToStr(ExcelLine2) + ':B' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := '=SUM(C' + IntToStr(ExcelLine2) + ':C' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := '=SUM(D' + IntToStr(ExcelLine2) + ':D' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=SUM(E' + IntToStr(ExcelLine2) + ':E' + IntToStr(ExcelLine - 1) + ')';
             end;

             if (ExcelLine2 > 17) AND (ExcelLine > ExcelLine2) then begin
                 Inc(ExcelLine, 2);
                 MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'ИТОГО';
                 MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '=B' + IntToStr(ExcelLine2 - 3) + '+B' + IntToStr(ExcelLine - 2);
                 MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := '=C' + IntToStr(ExcelLine2 - 3) + '+C' + IntToStr(ExcelLine - 2);
                 MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := '=D' + IntToStr(ExcelLine2 - 3) + '+D' + IntToStr(ExcelLine - 2);
                 MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=E' + IntToStr(ExcelLine2 - 3) + '+E' + IntToStr(ExcelLine - 2);
             end;

             OutFileName := GetAppDir + 'DOCS\' + SubDir + '\БСП(1)_' + CurrList[idx_curr] + '_' + DateToStr(RepDate) + '.XLS';
             DeleteFile(OutFileName);
             MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
         end;

         //ФОРМА 2 (Больше 30 дней)
         WaitForm.WorkName.Caption := 'ФОРМА 2';
         WaitForm.Update;
         for idx_curr := 1 to 10 do begin
             if CurrList[idx_curr] = '' then break;

             MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\РНП(2).XLS');

             MSEXCEL.ActiveSheet.Range['A5', 'A5'].Value := 'на ' + DateToStr(RepDate);
             MSEXCEL.ActiveSheet.Range['B7', 'B7'].Value := InsurName;
             MSEXCEL.ActiveSheet.Range['B9', 'B9'].Value := CurrList[idx_curr];

             ExcelLine := 16;

             A := 0;
             B := 0;
             C := 0;
             D := 0;
             E := 0;
             for i := 0 to ListData.Count - 1 do begin
                 WaitForm.ProgressBar.Position := ROUND((idx_curr - 1) * 100 / CntCurrs) +
                                                  ROUND(i * 100 / ListData.Count / CntCurrs);
                 InfoPtr := PReportStructure(ListData.Items[i]);
                 if (InfoPtr.ToDate - InfoPtr.FromDate + 1) <= 30 then continue;
                 if CurrList[idx_curr] = InfoPtr.Curr then begin
                     if not IsFast then begin
                         MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := FloatToStr(InfoPtr.Number) + ', ' + DateToStr(InfoPtr.RegDate);
                         VAL := InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie;
                         MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) < 0 then begin
                             MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '0';
                             MSEXCEL.ActiveSheet.Range['F' + IntToStr(ExcelLine), 'F' + IntToStr(ExcelLine)].Value := 'Не хватает брутто премии';
                         end;
                         MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := Integer(TRUNC(InfoPtr.ToDate) - TRUNC(InfoPtr.FromDate) + 1);
                         MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := Integer(TRUNC(RepDate) - TRUNC(InfoPtr.FromDate));
                         MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=B' + IntToStr(ExcelLine) + '*(C' + IntToStr(ExcelLine) + '-D' + IntToStr(ExcelLine) + ')/C' + IntToStr(ExcelLine);
                         Inc(ExcelLine);
                     end else begin
                         A := A + 1;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie) > 0 then begin
                             B := B + (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie);
                             C := Integer(TRUNC(InfoPtr.ToDate) - TRUNC(InfoPtr.FromDate) + 1);
                             D := Integer(TRUNC(RepDate) - TRUNC(InfoPtr.FromDate));
                             E := E + (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) * (C - D) / C;
                         end;
                     end;
                 end;
             end;

             if IsFast AND (A > 0) then begin
                MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := IntToStr(A) + ' полисов';
                VAL := B;
                MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                VAL := E;
                MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := VAL;
                Inc(ExcelLine);
             end;

             if ExcelLine > 16 then begin
                 MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'ИТОГО';
                 MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '=SUM(B16:B' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['E' + IntToStr(ExcelLine), 'E' + IntToStr(ExcelLine)].Value := '=SUM(E16:E' + IntToStr(ExcelLine - 1) + ')';
             end;

             OutFileName := GetAppDir + 'DOCS\' + SubDir + '\РНП(2)_' + CurrList[idx_curr] + '_' + DateToStr(RepDate) + '.XLS';
             DeleteFile(OutFileName);
             MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
             if ExcelLine = 16 then begin
                 MSEXCEL.ActiveWorkbook.Close;
                 DeleteFile(OutFileName);
             end;
         end;

         //ФОРМА 5 (Меньше или 30 дней)

         WaitForm.WorkName.Caption := 'ФОРМА 5';
         for idx_curr := 1 to 10 do begin
             if CurrList[idx_curr] = '' then break;

             MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\РНП(5).XLS');

             MSEXCEL.ActiveSheet.Range['A5', 'A5'].Value := 'на ' + DateToStr(RepDate);
             MSEXCEL.ActiveSheet.Range['B7', 'B7'].Value := InsurName;
             MSEXCEL.ActiveSheet.Range['B9', 'B9'].Value := CurrList[idx_curr];

             ExcelLine := 16;

             A := 0;
             B := 0;
             C := 0;
             D := 0;
             E := 0;

             for i := 0 to ListData.Count - 1 do begin
                 WaitForm.ProgressBar.Position := ROUND((idx_curr - 1) / CntCurrs * 100) + ROUND(i / ListData.Count / CntCurrs * 100);
                 InfoPtr := PReportStructure(ListData.Items[i]);
                 if (InfoPtr.ToDate - InfoPtr.FromDate + 1) > 30 then continue;
                 if CurrList[idx_curr] = InfoPtr.Curr then begin
                     if not IsFast then begin
                         MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := FloatToStr(InfoPtr.Number) + ', ' + DateToStr(InfoPtr.RegDate);
                         VAL := InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie;
                         MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) < 0 then begin
                             MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '0';
                             MSEXCEL.ActiveSheet.Range['D' + IntToStr(ExcelLine), 'D' + IntToStr(ExcelLine)].Value := 'Не хватает брутто премии';
                         end;
                         MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := '=B' + IntToStr(ExcelLine) + '/2';
                         Inc(ExcelLine);
                     end else begin
                         A := A + 1;
                         if (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie) > 0 then
                             B := B + (InfoPtr.BruttoPremium - InfoPtr.AgOtchislenie - InfoPtr.OneTaxOtchislenie);
                     end
                 end;
             end;

             if IsFast AND (A > 0) then begin
                MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := IntToStr(A) + ' полисов';
                VAL := B;
                MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := VAL;
                VAL := B / 2;
                MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := VAL;
                Inc(ExcelLine);
             end;

             if ExcelLine > 16 then begin
                 MSEXCEL.ActiveSheet.Range['A' + IntToStr(ExcelLine), 'A' + IntToStr(ExcelLine)].Value := 'ИТОГО';
                 MSEXCEL.ActiveSheet.Range['B' + IntToStr(ExcelLine), 'B' + IntToStr(ExcelLine)].Value := '=SUM(B16:B' + IntToStr(ExcelLine - 1) + ')';
                 MSEXCEL.ActiveSheet.Range['C' + IntToStr(ExcelLine), 'C' + IntToStr(ExcelLine)].Value := '=SUM(C16:C' + IntToStr(ExcelLine - 1) + ')';
             end;

             OutFileName := GetAppDir + 'DOCS\' + SubDir + '\РНП(5)_' + CurrList[idx_curr] + '_' + DateToStr(RepDate) + '.XLS';
             DeleteFile(OutFileName);
             MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
             if ExcelLine = 16 then begin
                 MSEXCEL.ActiveWorkbook.Close;
                 DeleteFile(OutFileName);
             end;
         end;
         MSEXCEL.Application.EnableEvents := true;
end;

procedure TPolandPolises.CreateReport(IsRNPU : boolean; RepDate : TDateTime; obj : TaxPercents; IsFast : boolean; List : TASKOneTaxList);
var
    ListData : TList;
    OutFileName : string;
    i : integer;
begin
    ListData := TList.Create;
    try
         FillPolisesForReserve(IsRNPU, ListData, TRUNC(RepDate), obj, List,
          'SELECT  Seria, Number, Repdate, StartDate, EndDate, PayDate, PremiumPay, PremiumPay2, PremiumCurr, PremiumCurr2, Uridich, AgPercent ' +
          'FROM POLANDPL WHERE PremiumVal is not null AND',
                              WorkSQL, ProgressBar, 'StartDate', 'EndDate', 'EUR');

         //FillDocsData(IsRNPU, ListData, TRUNC(RepDate), FizichTax, UridichTax, CompanyPercent, OneTax);
         if not InitExcel(MSEXCEL) then begin
            MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
            exit;
         end;
         Application.CreateForm(TWaitForm, WaitForm);
         WaitForm.Update;
         MSEXCEL.Visible := true;

         OutTo10Form(IsRNPU, 'Страхование гражданской ответственности выезжающих в республику Польша', MSEXCEL, RepDate, ListData, IsFast, 'POLAND');
         //TransferToExecl('Страхование гражданской ответственности выезжающих в республику Польша', RepDate, 'POLAND', ListData, MSEXCEL, IsFast);
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

function GetDaysDelta(Curr : string) : integer;
begin
    if Curr = 'USD' then GetDaysDelta := ROUND(DaysForSellCurr.USD.Value)
    else if Curr = 'DM' then GetDaysDelta := ROUND(DaysForSellCurr.DM.Value)
    else if Curr = 'RUR' then GetDaysDelta := ROUND(DaysForSellCurr.RUR.Value)
    else if Curr = 'EUR' then GetDaysDelta := ROUND(DaysForSellCurr.EUR.Value)
    else raise Exception.Create('Продана неизвестная валюта ' + Curr);
end;

procedure TPolandPolises.N15Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

{
procedure TPolandPolises.FillDocsData(IsRNPU : boolean; var DataList : TList; RepDate : TDateTime; FizichTax, UridichTax, CompanyPercent : double; OneTax : double);
var
     Summa : double;
     Info : ReportStructure_3;
     i : integer;
     InfoPtr : PGroupRestrax; //^ReportStructure;
begin
     //Те кторые действуют на дату или те которые Б 30 дней и оплачены в предшеств. отчёту месяц
     //Сумма берётся вся пришедшая до отчётной даты
     With WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT');
          Add('   Seria, Number, Repdate, StartDate, EndDate, PayDate, PremiumPay, PremiumPay2, PremiumCurr, PremiumCurr2, Uridich, AgPercent');
          Add('FROM');
          Add('   POLANDPL');
          Add('WHERE');
          Add('   PremiumVal is not null AND'); 
          if IsRNPU then
              Add('   (PayDate between :RepDateSubMonth and :RepDate)')
          else begin
              Add('   (PayDate <= :RepDate AND EndDate >= :RepDate OR'); //попал в период
              Add('   (PayDate >= :RepDateSubMonth AND PayDate < :RepDate AND (EndDate - StartDate + 1) <= 30))');
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

        ProgressBar.Position := TRUNC(WorkSQL.RecNo * 100 / WorkSQL.RecordCount);
         with WorkSQL do begin
           if FieldByName('PremiumCurr').AsString <> '' then
              FillStructure_3(RepDate, FieldByName('Uridich').AsBoolean, FizichTax, UridichTax, CompanyPercent, Info, FieldByName('PayDate').AsDateTime, FieldByName('PremiumPay').AsFloat, FieldByName('PremiumCurr').AsString, FieldByName('AgPercent').AsFloat, OneTax);
           if FieldByName('PremiumCurr2').AsString <> '' then
              FillStructure_3(RepDate, FieldByName('Uridich').AsBoolean, FizichTax, UridichTax, CompanyPercent, Info, FieldByName('PayDate').AsDateTime, FieldByName('PremiumPay2').AsFloat, FieldByName('PremiumCurr2').AsString, FieldByName('AgPercent').AsFloat, OneTax);
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

                  InfoPtr.AllSummaCurr := 'EUR';
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
}
procedure TPolandPolises.SpeedButton1Click(Sender: TObject);
begin
     if PolandQuery.State in [dsEdit, dsInsert] then
         PolandQuery.Post;

     if not PolandQuery.Locate('Number', SeekNumber.Value, []) then
         MessageDlg('Полис ' + SeekNumber.Text + ' не найден.', mtInformation, [mbOk], 0);
end;

procedure TPolandPolises.N16Click(Sender: TObject);
begin
     DaysForSellCurr.ShowModal
end;

procedure TPolandPolises.N17Click(Sender: TObject);
begin
     if MandSellCurr = nil then begin
         Application.CreateForm(TMandSellCurr, MandSellCurr);
         MandSellCurr.ShowModal;
     end;
end;

procedure TPolandPolises.SortMenuClick(Sender: TObject);
begin
     SortForm.m_Query := PolandQuery;
     if SortForm.ShowModal = mrCancel then exit;
     if PolandQuery.MacroByName('ORDER').AsString <> SortForm.GetOrder('ORDER BY') then begin
         try
             Screen.Cursor := crHourGlass;
             PolandQuery.Close;
             PolandQuery.MacroByName('ORDER').AsString := SortForm.GetOrder('ORDER BY');
             PolandQuery.Open;
         except
             on E : Exception do
                 MessageDlg('Ошибка запроса к БД'#13 + E.Message, mtInformation, [mbOK], 0);
         end
     end;
     Screen.Cursor := crDefault;
end;

procedure TPolandPolises.N20Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по польским полисам';
    About.ShowModal
end;

procedure TPolandPolises.N18Click(Sender: TObject);
begin
     MessageDlg('Если нажать кнопку Добавить, то будет создан полис с датой начала = текущей дате. Если при этом держать нажатой кнопку Shift, то будет создан полис с номером, большим чем текущий, с датами и владельцем и агентом текущего полиса.'#13 +
                'Для ввода испорченного необходимо не заполнять дату начала периода страхования'#13 + 
                'Быстрый набор СРОКА СТРАХОВАНИЯ'#13 +
                 '15 = 15 дней'#13 +
                 '+15 = Продление на 15 дней'#13 +
                 '30 = 30 дней'#13 +
                 '+30 или +3 = Продление на 30 дней'#13 +
                 '2..11 = 2..11 месяцев'#13 +
                 '1 = 1 год'#13 +
                 '*2..*11 = Продление на 2..11 месяцев'#13,
                 mtInformation, [mbOk], 0); 
end;

procedure TPolandPolises.PolandQueryAgentCodeChange(Sender: TField);
begin
     PolandQueryAgPercent.AsFloat := AgentTbl.FieldByName('PolandPcnt').AsFloat;
end;

procedure TPolandPolises.IsDupOwnMnuClick(Sender: TObject);
begin
     IsDupOwnMnu.Checked := not IsDupOwnMnu.Checked; 
end;

procedure TPolandPolises.PolandQueryAfterPost(DataSet: TDataSet);
begin
     UpdateCountPolises;
     PolPolisData.ShowStore;
//     if PolPolisData <> nil then
//        PolPolisData.IsEditNumber := false;
end;

procedure TPolandPolises.UpdateCountPolises;
begin
     CntPolisesMnu.Caption := 'Найдено ' + IntToStr(PolandQuery.RecordCount);
end;

function TPolandPolises.GetColumn(F : TField) : TColumn;
var
    i : integer;
begin
    for i := 0 to PolandGrid.Columns.Count - 1 do begin
        if PolandGrid.Columns.Items[i].FieldName = F.FieldName then begin
            GetColumn := PolandGrid.Columns[i];
            exit
        end
    end;
    GetColumn := nil;
end;

procedure TPolandPolises.ShowHideFldsClick(Sender: TObject);
begin
    GetColumn(PolandQueryPayDate).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryRepDate).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryName).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryinscity).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryinsstreet).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryMARKA).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryBODYNO).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryAutoNumber).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryAgent).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryAgPercent).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryAUTOOWNER).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryowncity).Visible := not ShowHideFlds.Checked;
    GetColumn(PolandQueryownstreet).Visible := not ShowHideFlds.Checked;

    if ShowHideFlds.Checked then
        ShowHideFlds.Caption := 'Показать все поля в списке'
    else
        ShowHideFlds.Caption := 'Спрятать поля в списке';

    ShowHideFlds.Checked := not ShowHideFlds.Checked;
end;

procedure TPolandPolises.PrintBtnClick(Sender: TObject);
begin
    if PolandQuery.State in [dsEdit, dsInsert] then
        PolandQuery.Post;

    if not FileExists(GetAppDir + 'Templates\Poland.tmpl') then begin
        MessageDlg('Файл ' + GetAppDir + 'Templates\Poland.tmpl' + ' не найден', mtInformation, [mbOk], 0);
        exit;
    end;

    if not FileExists(GetAppDir + 'Templates\Poland.doc') then begin
        MessageDlg('Файл ' + GetAppDir + 'Templates\Poland.doc' + ' не найден', mtInformation, [mbOk], 0);
        exit;
    end;

    if InitWord(MSWORD) then begin
        PrintRecord(MSWORD, 'POLAND', GetAppDir + 'Templates\Poland.tmpl', GetAppDir + 'Templates\Poland.doc', PolandQuery);
        MSWORD.ActiveDocument.SaveAs(GetAppDir + 'DOCS\Polis ' + PolandQuerySeria.AsString + PolandQueryNumber.AsString, 0);
        MSWORD.Activate;
    end
    else
        MessageDlg('MS Word не найден на вашем компьютере', mtInformation, [mbOk], 0);
end;

procedure TPolandPolises.N111Click(Sender: TObject);
begin
    CreateSimpleMagazine(WorkSQL, 'SELECT SERIA,NUMBER,STARTDATE INSSTART,ENDDATE INSEND,AUTOOWNER INSURER, MARKA INSURER2, 1 ALLSUMMA, '''' ALLSUMMACURr, ' +
                                  'premiumpay sum1, premiumcurr sum1curr, premiumpay2 sum2, premiumcurr2 sum2curr, Paydate PAYMENTDATE, AgPercent, Uridich AgUr FROM POLANDPL T WHERE PremiumVal is not null and PAYDATE BETWEEN :D1 AND :D2', 'ЕВРОПА');
    Application.BringToFront;
end;


procedure TPolandPolises.btnFormaClick(Sender: TObject);
begin
    PolPolisData.Visible := not PolPolisData.Visible
end;

procedure TPolandPolises.PolandQueryCalcFields(DataSet: TDataSet);
begin
    PolandQueryStateName.AsString := 'НЕИЗВЕСТНО';
    if PolandQueryState.AsString <> '' then
        case PolandQueryState.AsString[1] of
        '0' : PolandQueryStateName.AsString := 'НОРМАЛЬНЫЙ';
        '1' : PolandQueryStateName.AsString := 'УТЕРЯН';
        '2' : PolandQueryStateName.AsString := 'ИСПОРЧЕН';
        '3' : PolandQueryStateName.AsString := 'РАСТОРГНУТ';
        end;
end;

procedure ParseAutoFields(Str : string; var MARKA, MODEL : string);
var
    i : integer;
begin
    Str := Trim(Str);
    i := Pos(' ', Str);
    if i = 0 then begin
        MARKA := Str;
        MODEL := '';
    end
    else begin
        MARKA := Trim(Copy(Str, 1, i));
        MODEL := Trim(Copy(Str, i + 1, 1024));
    end;
end;

//procedure ParseFields(Typ, Nm, City, Addr : string; var NAME, SURNAME, CITY, STREET : string);
procedure ParseFields(Typ, Nm : string; var NAME, SURNAME : string);
var
    i : integer;
begin
    if Typ = 'U' then begin
        NAME := Nm;
        SURNAME := Nm;
    end
    else begin
        Nm := Trim(Nm);
        i := Pos(' ', Nm);
        if i = 0 then begin
            NAME := Nm;
            SURNAME := '';
        end
        else begin
            SURNAME := Trim(Copy(Nm, 1, i));
            NAME := Trim(Copy(Nm, i + 1, 1024));
        end;
    end;

{    Addr := Trim(Addr);
    i := Pos(' ', Addr);
    if i = 0 then begin
        CITY := Addr;
        STREET := '';
    end
    else begin
        CITY := Trim(Copy(Nm, 1, i));
        STREET := Trim(Copy(Nm, i + 1, 1024));
    end;}
//    CITY := Trim(Copy(Nm, 1, i));
//    STREET := Trim(Copy(Nm, i + 1, 1024));
end;

function GetUSD(Summa : double; Curr : string; DT : TDateTime) : double;
var
    Val : double;
begin
    if Curr = 'USD' then begin
        GetUSD := Summa;
    end
    else begin
        Val := Summa * GetRateOfCurrency(Curr, DT) / GetRateOfCurrency('USD', DT);
        RoundValue(Val);
        GetUSD := Val;
    end;
end;

procedure TPolandPolises.N23Click(Sender: TObject);
var
    ASH : VARIANT;
    sLine : string;
    _NAME, SURNAME, CITY, STREET, MARKA, MODEL : string;
    Division : integer;
    INI : TIniFile;
begin
    INI := TIniFile.Create('blank.ini');
    Division := INI.ReadInteger('POLAND', 'Division', -1);
    INI.Free;

    if Division < 0 then begin
        MessageDlg('Не определён код подразделения', mtInformation, [mbOk], 0);
        exit;
    end;

    if PolandQuery.State in [dsEdit, dsInsert] then
        PolandQuery.Post;
        
    if GetPeriod.ShowModal <> mrOk then exit;

    if not InitExcel(MSEXCEL) then begin
       MessageDlg('MicroSoft Excel не обнаружен на вашем компьютере.', mtError, [mbOk], 0);
       exit;
    end;
    MSEXCEL.Visible := true;
    MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\POLANDAUTO.XLS');
    ASH := MSEXCEL.ActiveSheet;

    InsertTable.Close;
    InsertTable.Open;

    try
        with WorkSQL, WorkSQL.SQL do begin
            Clear;
            Close;
            Add('SELECT * FROM POLANDPL');
            if GetPeriod.ExpMode.ItemIndex = 0 then begin
                Add('WHERE EXPDATE IS NULL AND UPDATEDATE>=:SD');
                if GetPeriod.EndDate.Checked then
                    Add('AND UPDATEDATE<=:ED');
                ParamByName('SD').AsDateTime := GetPeriod.StartDate.Date;
                if GetPeriod.EndDate.Checked then
                    ParamByName('ED').AsDateTime := GetPeriod.EndDate.Date;
                end
            else
            if GetPeriod.ExpMode.ItemIndex = 1 then begin
                Add(PolandFilter.FilterText);
                Add('AND EXPDATE IS NULL AND UPDATEDATE>=''01.01.2004''');
            end
            else
            if GetPeriod.ExpMode.ItemIndex = 2 then begin
                Add(PolandFilter.FilterText);
                Add('WHERE EXPDATE IS NULL AND UPDATEDATE>=''01.01.2004''');
            end
            else begin
                if GetPeriod.ListExports.ItemIndex = -1 then exit;
                Add('WHERE EXPDATE =''' +  GetPeriod.ListExports.Items[GetPeriod.ListExports.ItemIndex] + '''');
            end;

            Screen.Cursor := crHourGlass;
            Open;

            if Eof then begin
                Application.BringToFront;
                MessageDlg('Нет данных за указанный период.', mtInformation, [mbOk], 0);
            end
            else begin
                sLine := '4';
                while not Eof do begin
                    ASH.Range['A' + sLine, 'A' + sLine].Value := FieldByName('Seria').AsString;
                    ASH.Range['B' + sLine, 'B' + sLine].Value := FieldByName('Number').AsString;
                    if FieldByName('State').AsString = '' then
                        raise Exception.Create('Не определено состояние у полиса ' + FieldByName('Number').AsString);
                    case FieldByName('State').AsString[1] of
                        '0', '1' : begin //normal
                              if FieldByName('ISDUP').AsString = 'Y' then
                                  ASH.Range['C' + sLine, 'C' + sLine].Value := '2'
                              else
                              if FieldByName('PNUMBER').AsInteger > 0 then
                                  ASH.Range['C' + sLine, 'C' + sLine].Value := '1'
                              else
                                  ASH.Range['C' + sLine, 'C' + sLine].Value := '0';
                              end;
                        '2' : //BAD_POLIS : string = '2';
                              ASH.Range['C' + sLine, 'C' + sLine].Value := '3';
                        '3' : //BREAK_POLIS : string = '3';
                              ASH.Range['C' + sLine, 'C' + sLine].Value := '4';
                    end;

                    ASH.Range['D' + sLine, 'D' + sLine].Value := FormatDateTime('YYYY-MM-DD', FieldByName('RepDate').AsDateTime);

                    if (FieldByName('State').AsString <> BAD_POLIS) AND (FieldByName('State').AsString <> BREAK_POLIS) then begin
                        ASH.Range['E' + sLine, 'E' + sLine].Value := FormatDateTime('YYYY-MM-DD', FieldByName('StartDate').AsDateTime);
                        ASH.Range['F' + sLine, 'F' + sLine].Value := FormatDateTime('YYYY-MM-DD', FieldByName('EndDate').AsDateTime);
                    end;

                    if FieldByName('State').AsString = BREAK_POLIS then begin
                        ASH.Range['G' + sLine, 'G' + sLine].Value := -GetUSD(FieldByName('RetSum').AsFloat, FieldByName('RetCurr').AsString, FieldByName('RetDate').AsDateTime);
                    end
                    else
                        ASH.Range['G' + sLine, 'G' + sLine].Value := FieldByName('PremiumVal').AsString;

                    if FieldByName('State').AsString = BREAK_POLIS then
                        ASH.Range['J' + sLine, 'J' + sLine].Value := FormatDateTime('YYYY-MM-DD', FieldByName('RetDate').AsDateTime);
                        
                    if (FieldByName('State').AsString <> BREAK_POLIS) then begin
                        ASH.Range['H' + sLine, 'H' + sLine].Value := FieldByName('TraifCode').AsString;

                        if FieldByName('PNumber').AsInteger > 0 then
                            ASH.Range['I' + sLine, 'I' + sLine].Value := FieldByName('PNumber').AsString;

                        //ЗАСТРАХОВАННЫЙ
                        if FieldByName('State').AsString <> BAD_POLIS then begin
                            if FieldByName('OwnType').AsString = 'U' then
                                ASH.Range['K' + sLine, 'K' + sLine].Value := '1'
                            else
                                ASH.Range['K' + sLine, 'K' + sLine].Value := '2';

                            ParseFields(FieldByName('OwnType').AsString,
                                        FieldByName('AutoOwner').AsString,
                                        _NAME, SURNAME);
                            ASH.Range['L' + sLine, 'L' + sLine].Value := _NAME;
                            ASH.Range['M' + sLine, 'M' + sLine].Value := SURNAME;
                            ASH.Range['N' + sLine, 'N' + sLine].Value := FieldByName('OwnCity').AsString;
                            ASH.Range['O' + sLine, 'O' + sLine].Value := FieldByName('OwnStreet').AsString;
                            ASH.Range['P' + sLine, 'P' + sLine].Value := FieldByName('OwnHouse').AsString;
                            ASH.Range['Q' + sLine, 'Q' + sLine].Value := FieldByName('OwnFlat').AsString;
                            ASH.Range['R' + sLine, 'R' + sLine].Value := FieldByName('OwnCntry').AsString;
                        end;

                        //АВТО
                        ParseAutoFields(FieldByName('Marka').AsString, MARKA, MODEL);
                        ASH.Range['S' + sLine, 'S' + sLine].Value := FieldByName('AutoNumber').AsString;
                        ASH.Range['T' + sLine, 'T' + sLine].Value := FieldByName('BodyNo').AsString;
                        ASH.Range['U' + sLine, 'U' + sLine].Value := FieldByName('EngNmb').AsString;
                        ASH.Range['V' + sLine, 'V' + sLine].Value := MARKA;
                        ASH.Range['W' + sLine, 'W' + sLine].Value := MODEL;
                        ASH.Range['X' + sLine, 'X' + sLine].Value := FieldByName('AutoType').AsString;
                        ASH.Range['Y' + sLine, 'Y' + sLine].Value := FieldByName('Auto_Vid').AsString;

                        //СТРАХОВАТЕЛЬ
                        if FieldByName('State').AsString <> BAD_POLIS then begin
                            if FieldByName('InsType').AsString = 'U' then
                                ASH.Range['Z' + sLine, 'Z' + sLine].Value := '1'
                            else
                                ASH.Range['Z' + sLine, 'Z' + sLine].Value := '2';

                            ParseFields(FieldByName('InsType').AsString,
                                        FieldByName('Name').AsString,
                                        _NAME, SURNAME);
                            ASH.Range['AA' + sLine, 'AA' + sLine].Value := _NAME;
                            ASH.Range['AB' + sLine, 'AB' + sLine].Value := SURNAME;
                            ASH.Range['AC' + sLine, 'AC' + sLine].Value := FieldByName('InsCity').AsString;
                            ASH.Range['AD' + sLine, 'AD' + sLine].Value := FieldByName('InsStreet').AsString;
                            ASH.Range['AE' + sLine, 'AE' + sLine].Value := FieldByName('InsHouse').AsString;
                            ASH.Range['AF' + sLine, 'AF' + sLine].Value := FieldByName('InsFlat').AsString;
                            ASH.Range['AG' + sLine, 'AG' + sLine].Value := FieldByName('InsCntry').AsString;
                        end;
                    end;
                    
                    Next;
                    sLine := IntToStr(StrToInt(sLine) + 1);

                    if not InsertTable.Locate('Seria;Number', VarArrayOf([FieldByName('SERIA').AsString, FieldByName('NUMBER').AsFloat]), []) then
                        raise Exception.Create('Ошибка обновления сессии импорта');

                    if InsertTable.FieldByName('EXPDATE').AsString = '' then begin
                        InsertTable.Edit;
                        InsertTable.FieldByName('EXPDATE').AsDateTime := Date;
                        InsertTable.FieldByName('Division').AsInteger := Division;
                        InsertTable.Post;
                    end;
                end;
            end;
        end;
        Application.BringToFront;
        MessageDlg('Отчёт создан успешно', mtInformation, [mbOk], 0);
    except
        on E : Exception do begin
            Application.BringToFront;
            MessageDlg('Произошла ошибка ' + e.Message, mtInformation, [mbOk], 0);
        end;
    end;
    Screen.Cursor := crDefault;
    InsertTable.Close;
end;

procedure TPolandPolises.PolandQueryAfterScroll(DataSet: TDataSet);
begin
//     if PolPolisData <> nil then
//         PolPolisData.IsEditNumber := false;
end;

procedure TPolandPolises.PolandQueryAfterCancel(DataSet: TDataSet);
begin
//     if PolPolisData <> nil then
//         PolPolisData.IsEditNumber := false;
end;

end.

