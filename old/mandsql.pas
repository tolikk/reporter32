unit MandSQL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, Db, DBTables, Grids, DBGrids, RXDBCtrl, ExtCtrls, StdCtrls, RXSpin,
  Buttons, RxQuery, inifiles, ComCtrls, FileUtil, BdeUtils, FR_DSet,
  FR_DBSet, FR_Class, FR_View, rxstrutils, strutils;

type
   AvKType = record
        Val : double;
        K : string;
   end;

   ErrorTypes = ( eeBadPolisNumber, eeBadSymbName, eeEnglSymbAddr, eeEnglSymbName, eeBadFizUr,
                  eeStartFromDigit, eeRussLetters, eeBadDiscount, eePayAndVer,
                  eeBadSymbAddr, eeBadVladen, eeBadAutoMark, eeBadTC, eeBadSymbMarka,
                  eeBadCharAuto, eeBadSymbBodyNo, eeRussBodyShassi, eeBadState,
                  eePayAfterStart, eeRegAndPay, eeBadPay2, eeStrangeDate, eeBadStopDate,
                  eeBadCurr, eeBadBaseAT, eeBadDecide, eeBadDoc, eeNoPay, eeBadBadMan, eeNoSumma,
                  eeSumNoSum, eeBadAvReg, eeBadPayAvDate, eeEmptyAddr, eeBadSymbSNP, eeSellAfterStart,
                  eeBadDateEnd, eeNoDivision, eeNoStopDate, eeNoNeedPay, eeBadAvK );

   AMrk = record
       IsImport : integer;
       Marka : string;
   end;

   PolisNmbsRange = record
       Seria : string[10];
       Start, _End : integer;
   end;

   DupRecord = record
       Code : string[32];
       DivCode : integer;
       DupCount : integer;
   end;

   PDupRecord = ^DupRecord;

   CarOwnerDescription_ = record
       Name : string[50];
       Addr_code : string[3];
       Address : string[50];
       Owner_Type : integer;
       R_N : string[2];
       F_L : string[2];
       DISCOUNT_2 : double;
       DVR_EXPR : integer;
       DIVISION : integer;
   end;

   SerCarDescription_ = record
       Marka : string[30];
       TypeTC : string[3];
       Characteristic : double;
       DIVISION : integer;
   end;

   CarDescription_ = record
       AutoNumber : string[10];
       BodyNumber : string[20];
       ShassiNumber : string[20];
       BaseCarCode : integer;
       DIVISION : integer;
   end;

  CodeControl = class
       DivCodes : TList; //Список кодов подразделений
       DivMaxCodesOwn : TList; //Список максимальных кодов
       DivMaxCodesBaseCar : TList; //Список максимальных кодов
       DivMaxCodesCar : TList; //Список максимальных кодов
       DivMaxCodesAV : TList; //Список максимальных кодов

       DivListCodesOwn : TList; //Список списков кодов дырок
       DivListCodesBaseCar : TList; //Список списков кодов дырок
       DivListCodesCar : TList; //Список списков кодов дырок
       DivListCodesAV : TList; //Список списков кодов дырок

       function GetNextOwnerCode(d : integer) : integer;
       function GetNextBaseCarCode(d : integer) : integer;
       function GetNextCarCode(d : integer) : integer;
       function GetNextAvCode(d : integer) : integer;

       constructor Create;
       destructor Free;
  end;

  TMandSQL = class(TForm)
    MainMenu: TMainMenu;
    AVList: TMenuItem;
    N2: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    FilterMn: TMenuItem;
    N9: TMenuItem;
    MainDataSource: TDataSource;
    MenuCompany: TMenuItem;
    N4: TMenuItem;
    N7: TMenuItem;
    TopPanel: TPanel;
    FilterBtn: TButton;
    Button2: TButton;
    N12: TMenuItem;
    StateText: TStaticText;
    N13: TMenuItem;
    MenuItemBuro: TMenuItem;
    StatusBar: TStatusBar;
    ProgressBar: TProgressBar;
    KorrectMenuItem: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N21: TMenuItem;
    PrintList: TMenuItem;
    N27: TMenuItem;
    N33: TMenuItem;
    MICalcSumm: TMenuItem;
    InitMenu: TMenuItem;
    N3: TMenuItem;
    ReservMenu: TMenuItem;
    Magazine: TMenuItem;
    N8: TMenuItem;
    N11: TMenuItem;
    SQLDB: TDatabase;
    MandSQLQuery: TRxQuery;
    Timer: TTimer;
    HandBookSQL: TRxQuery;
    LOCALDB: TDatabase;
    MandSQLQuerySER: TStringField;
    MandSQLQueryNMB: TIntegerField;
    MandSQLQueryPSER: TStringField;
    MandSQLQueryPNMB: TIntegerField;
    MandSQLQueryNAME: TStringField;
    MandSQLQueryADDR: TStringField;
    MandSQLQueryFRDT: TDateField;
    MandSQLQueryTODT: TDateField;
    MandSQLQueryTRF: TFloatField;
    MandSQLQueryPAY1: TFloatField;
    MandSQLQueryPAY1DT: TDateField;
    MandSQLQueryPAY1DOC: TFloatField;
    MandSQLQueryISFEE2: TStringField;
    MandSQLQueryFEE2DT: TDateField;
    MandSQLQueryISPAY2: TStringField;
    MandSQLQueryPAY2: TFloatField;
    MandSQLQueryPAY2DT: TDateField;
    MandSQLQueryAGNT: TStringField;
    MandSQLQuerySTATE: TSmallintField;
    MandSQLQueryPAY1CURR: TStringField;
    MandSQLQueryPAY2CURR: TStringField;
    MandSQLQueryPERIOD: TStringField;
    MandSQLQueryVPERIOD: TStringField;
    MandSQLQueryVPAY1: TStringField;
    MandSQLQueryVPAY2: TStringField;
    MandSQLQueryPAY2DOC: TFloatField;
    MandSQLQueryVSTATE: TStringField;
    MandSQLQueryREGDATE: TDateField;
    MandSQLQueryVPREV: TStringField;
    MandSQLQueryVSN: TStringField;
    MandSQLQueryMARKA: TStringField;
    WorkSQL: TRxQuery;
    ToBuroBtn: TButton;
    MandSQLQueryVER: TStringField;
    ExportBUROSQL: TRxQuery;
    ExportOwners: TRxQuery;
    ExportAvariasGLOBO: TRxQuery;
    MandSQLQueryK1: TFloatField;
    MandSQLQueryK2: TFloatField;
    MandSQLQueryTRF0: TFloatField;
    MandSQLQueryVTARUF: TStringField;
    TableAg: TTable;
    TableAgAgent_code: TStringField;
    TableAgDivisionCode: TFloatField;
    WorkLocalSQL: TQuery;
    N1: TMenuItem;
    LostPay: TMenuItem;
    N10: TMenuItem;
    LOG: TMemo;
    Splitter1: TSplitter;
    PageControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheetPrint: TTabSheet;
    MandGrid: TRxDBGrid;
    frReport: TfrReport;
    frDBDataSet: TfrDBDataSet;
    frPreview: TfrPreview;
    Panel2: TPanel;
    Button1: TButton;
    Tb5: TTable;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    Animate: TEdit;
    Timer2: TTimer;
    N17: TMenuItem;
    Excel1: TMenuItem;
    MandSQLQueryAUTONMB: TStringField;
    N20: TMenuItem;
    N22: TMenuItem;
    PolisNmb: TTable;
    PolisNmbS: TStringField;
    PolisNmbSTART: TFloatField;
    PolisNmbEND: TFloatField;
    Globo: TMenuItem;
    ExportAvariasBURO: TRxQuery;
    SaveDialog: TSaveDialog;
    OpenDialog: TOpenDialog;
    procedure N9Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SortClick(Sender: TObject);
    procedure MandGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure MenuItemBuroClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure KorrectMenuItemClick(Sender: TObject);
    procedure InitMenuClick(Sender: TObject);
    procedure MandSQLQueryCalcFields(DataSet: TDataSet);
    procedure FilterBtnClick(Sender: TObject);
    procedure FilterMnClick(Sender: TObject);
    procedure LostPayClick(Sender: TObject);
    procedure N10Click(Sender: TObject);
    procedure LOGKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure PrintListClick(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure N16Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure N17Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure Excel1Click(Sender: TObject);
    procedure N22Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure GloboClick(Sender: TObject);
  private
    W : HWnd;
    AnimRight : boolean;
    AnimateText : string;
    
//    lastAgCode : string;
//    lastDivCode : integer;

//    ListAgCodes : TStringList;
//    ListAgDiv : TList;
    ErrorCount : integer;
    CodeControl_ : CodeControl;

    { Private declarations }

    procedure AddError(S : string; N : integer; txt : string; Code : ErrorTypes; var ErrStrList : string);
  public
    { Public declarations }
    MSWORD : Variant;
    MSEXCEL : VARIANT;

//    MaxOwnerCode : integer;
//    MaxSerTCCode : integer;
//    MaxTCCode : integer;
//    MaxAVCode : integer;

//    OwnerCode_Holes : TList;
//    SerTCCode_Holes : TList;
//    TCCode_Holes : TList;
//    AVCode_Holes : TList;

    ExpSQL : TDataSet;
    DeltaDays : integer;

    procedure OpenQuery;
    procedure FindSN(Seria : string; Number : integer);
    function DropDublicate(SrcTable : TTable; tableName, MaxFld1, MaxFld2, MaxFld3  : string) : boolean;
    function  GetFilter : string;
    function  GetOrder : string;
    procedure CreateReportBuro();
//    function GetDivisionCode2(AgCode : string) : string;
    function GetVladenType(N : integer) : string;
    function GetFizUrid(s : string) : string;
    function GetDiscount2(N : double) : string;
    function GetCharacteristicAuto(s : string; Volume, Tonnage,Quantity, Power : double; IsResident : boolean) : double;
    function ExpTime(s : string) : integer;
    function ExpStateFunc(state : integer) : string;
    function ExpCurr(Curr : string) : string;
    function ExpZeroNmb(N : Integer) : string;
    function GetResident(N : boolean) : string;
    function GetDecision(N : integer) : string;
    function GetPaySN(s : string; mand : boolean) : string;

    //CODE GENERATION
    procedure InitExportBuro;
    function GetMaxCodeFor(Table, Field, Division : string; IsLinkMand : boolean) : integer;
    function GetNextCodeOwner(Code : integer; S : CarOwnerDescription_; var NeedUpdate : boolean) : string;
    function GetNextCodeSerTC(Code : integer; S : SerCarDescription_; var NeedUpdate : boolean) : string;
    function GetNextCodeTC(Code : integer; S : CarDescription_; var NeedUpdate : boolean) : string;
    function GetNextCodeAv(Code : integer; D : integer; var NeedUpdate : boolean) : string;
    procedure SetMaxCodeFor(Seria : string; Number : integer; Table, Field1, Field2, Val1, Val2 : string; IsVal2Int : boolean);
  end;

  function DecodeMsgId(str : string) : string;
  function CodeMsgId(str : string) : string;
  function DecodeBlank(s : string) : string;
  function DecodeGlDate(dt : TDate) : string;
  function DecodeGlDate2(dt, dt1, dt2, dt3 : TField) : string;
  function DecodePeriod(s : string) : string;
  function DecodeFU12(s : string) : string;
  function DecodeFU16(s : string; pcnt : double) : string;
  function DecodeTimeX(s : string) : string;
  //function DecodeR(s : string) : string;
  function DecodePayGraphic(s : string) : string;
  function DecodePayType(s : string) : string;
  function DecodeCurrency(s : string) : string;
  function DecodeAgentName(s : string) : string;
  function GetAvK(val : double) : string;
  function GetK(val : double) : string;
  function DecodeVladen(s : string) : string;
  function DecodeAvState(s : string; DtCl, DtFree : TField) : string;
  function DecodeSN(s : string) : string;
  function DecodeInsSum(d : TDate) : string;
  function DecodeBodyShas(s : string) : string;
  function DecodeCharact2(charact, s : string) : string;
  function DecodeVariant(s : string) : string;
  function DecodeR(s : string) : string;
  function DecodeOwnTypeForm(s : string) : string;
  function DecodeNone(s, s1 : string) : string;
  function DecodeCharact(s, v, t, c, p : string) : string;
  function DecodeNB(s, paydoc : string) : string;
  function KillZero(s : string) : string;
  function DecodeOneDate(F1, F2, F3 : TField) : string;

const
  CodeTable : string = '0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM';
                       //backup '0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM';
                       //йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ

var
  MandSQLForm : TMandSQL;

implementation

uses Sort, FilterMandatory, MyPreview, ProgressDlg, MandatoryRpt, Export, ComObj,
  BeforeImp, Explorer, ReportDates, Korrect,
  BasoMonthRptUnit, DateUtil, GetFormulaUnit, ShowInfoUnit, AboutUnit,
  KorrectPolandUnit, MandReservDlgUnit, AvarUnit, WaitFormUnit,
  GurnalUbitokUnit, ArariasFilterUnit, FilterMandSQL, PassDlg, ExpBuroUnit,
  KorrectSQL, frDecoder, ExpErrsUnit, datamod, AvarSQLFiltUnit, PolisNmbs, Variants,
  AvarSQLUnit;

type
    PolisVid = class
    public
       Auto_Type : string[5];
       Count : integer;
       In_Fee_Sum : double;
       AvariaCount : integer;
       Pay_Sum_1 : double;
       Pay_Count : integer;
       Pay_Sum_2 : double;
    end;

    AutoOwnerStruct = record
       FIO : string[64];
       Code : string[10];
    end;

{$R *.DFM}


procedure TMandSQL.N9Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по обязательному страхованию'#13'транспортных средств';
    About.ShowModal;
end;

procedure TMandSQL.N6Click(Sender: TObject);
begin
     Close
end;

function TMandSQL.GetOrder : string;
begin
     Result := SortForm.GetOrder('');
end;

function TMandSQL.GetFilter : string;
begin
     Result := MandSQLFilterDlg.GetFilter
end;

procedure TMandSQL.FormCreate(Sender: TObject);
var
     R : TRect;
begin
     AnimRight := true;
     AnimateText:= COMPAMYNAME;
     Animate.Text := COMPAMYNAME;
        
     MenuCompany.Caption := COMPAMYNAME;
//     OwnerCode_Holes := TList.Create;
//     SerTCCode_Holes := TList.Create;
//     TCCode_Holes := TList.Create;
//     AVCode_Holes := TList.Create;

//     ListAgCodes := TStringList.Create;
//     ListAgDiv := TList.Create;

     KorrectMenuItem.Visible := GetPrivateProfileInt('MANDATORY', 'CAN_KORRECT_DATA', 0, 'BLANK.INI') = 1;

     ExpSQL := nil;

 //    MaxOwnerCode := -1;

     ProgressBar.Parent := StatusBar;

     Application.Title := 'Отчёт по обязятельному страхованию';
     MandGrid.Align := alClient;
     MandGrid.Visible := true;
     Menu := MainMenu;

     Timer.Enabled := true;

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

procedure TMandSQL.OpenQuery;
begin
         MandSQLQuery.Close;
         if Length(GetFilter) <> 0 then begin
            MandSQLQuery.MacroByName('WHERE').AsString := GetFilter;
         end
         else
            MandSQLQuery.MacroByName('WHERE').AsString := '1=1';

         if Length(GetOrder) > 0 then begin
            MandSQLQuery.MacroByName('ORDER').AsString := GetOrder;
         end
         else
            MandSQLQuery.MacroByName('ORDER').AsString := 'SER,NMB';

         try
           Screen.Cursor := crHourGlass;
           MandSQLQuery.Open;
           Timer2.Enabled := false;

           StateText.Caption := 'Полисов ' + IntToStr(MandSQLQuery.RecordCount);

           if PageControl.ActivePage = TabSheetPrint then begin
              frReport.EMFPages.Clear;
              frPreview.Clear;
              frReport.ShowReport;
           end;

         except
           on e : Exception do begin
              MessageDlg('Ошибка при выполнении запроса.'#13 + e.Message, mtInformation, [mbOK], 0);
           end
         end;
         Timer2.Enabled := true;
         Screen.Cursor := crDefault;
end;

procedure TMandSQL.TimerTimer(Sender: TObject);
begin
     Timer.Enabled := false;

     if ParamCount >= 2 then begin
         PasswordDlg.AutoConnectStr := ParamStr(2)
     end;

     if PasswordDlg.MyShowModal <> mrOK then begin
        Close;
        exit;
     end;
     Application.CreateForm(TMandSQLFilterDlg, MandSQLFilterDlg);
     //FilterBtnClick(nil);
     Timer2.Enabled := true;
end;

procedure TMandSQL.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//     OwnerCode_Holes.Free;
//     SerTCCode_Holes.Free;
//     TCCode_Holes.Free;
//     AVCode_Holes.Free;
//     ListAgCodes.Free;
//     ListAgDiv.Free;
     Timer2.Enabled := false;

     if W <> 0 then begin
       ShowWindow(W, SW_SHOW);
       EnableWindow(W, true);
       SetActiveWindow(W);
     end;
end;

procedure TMandSQL.SortClick(Sender: TObject);
begin
     SortForm.m_Query := MandSQLQuery;
     if SortForm.ShowModal = mrOk then
         OpenQuery
end;

procedure TMandSQL.MandGridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
     if Field.IsNull then exit;

     if Field = MandSQLQueryVPAY2 then
        if MandSQLQueryISFEE2.AsString = 'Y' then
           if MandSQLQueryISPAY2.AsString = 'N' then
              if Date > MandSQLQueryFEE2DT.AsDateTime then begin
                 Background := clRed;
                 AFont.Color := clYellow
              end
              else begin
                 Background := clBlue;
                 AFont.Color := clWhite
              end;

end;

procedure TMandSQL.FindSN(Seria : string; Number : integer);
begin
     Screen.Cursor := crHourGlass;
     if not MandSQLQuery.Locate('Ser;Num', vararrayof([Seria, Number]), []) then
        MessageDlg('Полис ' + Seria + '/' + IntToStr(Number) + ' не найден.', mtInformation, [mbOK], 0);
     Screen.Cursor := crDefault;
end;

function CodeMsgId(str : string) : string;
var
    Value : integer;
    retCode : string;
    SystemBase : integer;
begin
    Value := StrToInt(str);
    SystemBase := Length(CodeTable);
    retCode := '';

    while(Value > 0) do begin
        retCode := CodeTable[(Value MOD SystemBase) + 1] + retCode;
        Value := TRUNC(Value / SystemBase);
    end;

    CodeMsgId := retCode;
end;

function DecodeYN(s : string) : integer;
begin
     if s = 'Y' then DecodeYN := 1
     else
     if s = 'N' then DecodeYN := 0
     else
     DecodeYN := -1;
end;


function CheckAutoLetter(s : string) : boolean;
var
     i : integer;
begin
     CheckAutoLetter := false;
     if s = '' then exit;
     if s = 'X' then exit;
     for i := 0 to Length(s) do begin
       if (s[i] < '0') AND (s[i] > '9') then
         if (s[i] >= 'А') AND (s[i] <= 'Я') then
             exit;
     end;
     CheckAutoLetter := true;
end;

procedure TMandSQL.MenuItemBuroClick(Sender: TObject);
var
     oldTitle : string;
begin
     if ExpBUROForm.ShowModal = mrCancel then
        exit;
     InitExportBuro();
     CodeControl_ := CodeControl.Create;
     oldTitle := Caption;
     CreateReportBuro();
     Caption := oldTitle;
     CodeControl_.Free;
     CodeControl_ := nil;
end;

function TMandSQL.DropDublicate(SrcTable : TTable; TableName, MaxFld1, MaxFld2, MaxFld3 : string) : boolean;
var
     TBL : TTable;
     i, j : integer;
begin
     try
         Screen.Cursor := crHourGlass;
         TBL := TTable.Create(self);
         TBL.TableName := TableName;
         TBL.Open;
         if not SrcTable.Active then
             SrcTable.Open;
              WorkLocalSQL.Close;
              WorkLocalSQL.SQL.Clear;

              WorkLocalSQL.SQL.Add('SELECT ');
              for i := 0 to SrcTable.Fields.Count - 1 do begin
                  if i > 0 then WorkLocalSQL.SQL.Add(',');

                  if (SrcTable.Fields[i].FieldName <> 'UPDATE') {AND (SrcTable.Fields[i].FieldName <> 'COMPANY')} AND (SrcTable.Fields[i].FieldName <> MaxFld1) AND (SrcTable.Fields[i].FieldName <> MaxFld2) AND (SrcTable.Fields[i].FieldName <> MaxFld3) then
                      WorkLocalSQL.SQL.Add(SrcTable.Fields[i].FieldName)
                  else begin
                      //if SrcTable.Fields[i].FieldName = 'COMPANY' then
                      //    WorkLocalSQL.SQL.Add('MIN(T."' + SrcTable.Fields[i].FieldName + '") AAA_' + SrcTable.Fields[i].FieldName)
                      //else
                          WorkLocalSQL.SQL.Add('MAX(T."' + SrcTable.Fields[i].FieldName + '") AAA_' + SrcTable.Fields[i].FieldName);
                  end;
              end;

              WorkLocalSQL.SQL.Add('FROM ''' + SrcTable.TableName + ''' T');
              WorkLocalSQL.SQL.Add('GROUP BY ');
              j := 0;
              for i := 0 to SrcTable.Fields.Count - 1 do begin
                  if (SrcTable.Fields[i].FieldName <> 'UPDATE') {AND (SrcTable.Fields[i].FieldName <> 'COMPANY')} AND (SrcTable.Fields[i].FieldName <> MaxFld1) AND (SrcTable.Fields[i].FieldName <> MaxFld2) AND (SrcTable.Fields[i].FieldName <> MaxFld3) then begin
                      if j > 0 then WorkLocalSQL.SQL.Add(',');
                      WorkLocalSQL.SQL.Add(SrcTable.Fields[i].FieldName);
                      Inc(j);
                   end;
              end;

              //WorkLocalSQL.SQL.Add('SELECT DISTINCT * FROM ''' + SrcTable.TableName + '''');

              //MessageDlg(WorkLocalSQL.SQL.Text, mtInformation, [mbOk], 0);
              WorkLocalSQL.Open;

              while not WorkLocalSQL.EOF do begin
                 if not WorkLocalSQL.FieldByName(WorkLocalSQL.Fields[0].FieldName).IsNull then begin
                    TBL.Append;
                    for i := 0 to WorkLocalSQL.Fields.Count - 1 do
                      if Copy(WorkLocalSQL.Fields[i].FieldName, 1, 4) <> 'AAA_' then
                        TBL.FieldByName(WorkLocalSQL.Fields[i].FieldName).Value := WorkLocalSQL.FieldByName(WorkLocalSQL.Fields[i].FieldName).Value
                      else
                        TBL.FieldByName(Copy(WorkLocalSQL.Fields[i].FieldName, 5, 100)).Value := WorkLocalSQL.FieldByName(WorkLocalSQL.Fields[i].FieldName).Value;
                    TBL.Post;
                 end;
                 WorkLocalSQL.Next;
              end;
              WorkLocalSQL.Close;

         TBL.Close;
         DropDublicate := true;
     except
         on E : Exception do begin
               MessageDlg('Дубликаты не удалены. Данные нельзя передавать в БЮРО!!!' + #13 + E.Message, mtInformation, [mbOk], 0);;
               DropDublicate := false;
         end;
     end;

     Screen.Cursor := crDefault;
     TBL.Free;
     WorkLocalSQL.Close;
end;

function DecodeMsgId(str : string) : string;
var
    Value : integer;
    SystemBase : integer;
    Multer : integer;
begin
    Value := 0;
    SystemBase := Length(CodeTable);
    Multer := 1;

    while(str > '') do begin
        Value := Value + Multer * (Pos(str[Length(str)], CodeTable) - 1);

        Multer := Multer * SystemBase;
        str := Copy(str, 1, Length(str) - 1);
    end;

    DecodeMsgId := IntToStr(Value);
end;

function IsValidPolisNumber(var _array : array of PolisNmbsRange; Seria : string; Nmb : integer) : boolean;
var
    i : integer;
begin
    IsValidPolisNumber := false;
    for i := 0 to Length(_array) - 1 do begin
        if Seria = _array[i].Seria then
            if (Nmb >= _array[i].Start) AND (Nmb <= _array[i]._End) then begin
                IsValidPolisNumber := true;
                exit;
            end;
    end;
end;

function IsMarka(List : TStringList; AutoClass : string; var AutoMark : array of AMrk; marka : string) : integer;
var
   i : integer;
begin
    if marka = '' then begin
        Result := 2;
        exit;
    end;

    for i := 0 to List.Count - 1 do
        if List[i] = AutoClass then begin
            Result := 2;
            exit;
        end;

    for i := 0 to Length(AutoMark) - 1 do
        if Pos(AutoMark[i].Marka, Marka) <> 0 then begin
            Result := AutoMark[i].IsImport;
            exit;
        end;

    Result := -1; //Нету
end;

procedure TMandSQL.CreateReportBuro;
var
   INI : TIniFile;
   TemplateDir : string;
   DestDir : string;
   ResidentStr : string;

   DBFileName1 : string;
   DBFileName2 : string;
   DBFileName3 : string;
   DBFileName4 : string;
   DBFileName5 : string;
   DBFileName6 : string;

   DBFileName2b : string;
   DBFileName3b : string;
   DBFileName4b : string;
   DBFileName5b : string;

   Table1 : TTable;
   Table2 : TTable;
   Table3 : TTable;
   Table4 : TTable;
   Table5 : TTable;
   Table6 : TTable;

   Appendix : string;
   PeriodStr : string;

   ExpState : boolean;

   i, find : integer;

   COMPANY_CODE : string;
//   DIVISION_CODE : string;
   DIV_CODE : Integer;
   SERIA : string[10];
   NUMBER : integer;
   CharAuto : double;

   AOwn : array of AutoOwnerStruct;
   OwnCounter : integer;
   CURR_CAR_CODE : string[10];
   CarOwnDescr_ : CarOwnerDescription_;
   SerCarDesrc_ : SerCarDescription_;
   CarDescr_ : CarDescription_;

   NeedUpdate : boolean;

   StartTime, CurrTime : integer;
   TimeStr : string;

   _ExceptSN : string;
   _ErrCollectStr : string;

   CountPolises : integer;

   IsRussian, IsEnglish, IsDigit, IsBad : boolean;

   fi, j, DefType : integer;

   AutoMarka : array of AMrk;
   s : string;

   PolisNmbsArray : array of PolisNmbsRange;
   IsMarkaValue : integer;

   ListNoCkeckAN : TStringList;
   amsg : MSG;
   fldList : string;
begin
    ErrorCount := 0;
    ListNoCkeckAN := TStringList.Create;

   INI := TIniFile.Create('blank.ini');
   s := INI.ReadString('MANDATORY', 'NumbersFilePath', '');
    if s = '' then begin
        PolisNmb.TableName := 'POLISNMB.db';
    end
    else begin
        if s[Length(s)] <> '\' then s := s + '\';
        PolisNmb.TableName := s + 'POLISNMB.db';
    end;

    PolisNmb.Open;
    i := 0;
    SetLength(PolisNmbsArray, PolisNmb.RecordCount);
    while not PolisNmb.Eof do begin
        PolisNmbsArray[i].Seria := PolisNmbS.AsString;
        PolisNmbsArray[i].Start := PolisNmbStart.AsInteger;
        PolisNmbsArray[i]._End := PolisNmbEnd.AsInteger;
        PolisNmb.Next;
        Inc(i);
    end;
    PolisNmb.Close;

   s := INI.ReadString('MANDATORY', 'AutoMarkFile', 'blank.ini');
   INI.Free;
   INI := TIniFile.Create(s);

   for i := 0 to 1024 do begin
       if INI.ReadString('MANDATORY', 'NoCheckAutoType' + IntToStr(i), '') = '' then break;
       ListNoCkeckAN.Add(INI.ReadString('MANDATORY', 'NoCheckAutoType' + IntToStr(i), ''));
   end;

   for i := 0 to 1024 do begin
       if INI.ReadString('MANDATORY', 'AutoType' + IntToStr(i), '') = '' then break;
       SetLength(AutoMarka, i + 1);
       AutoMarka[i].IsImport := INI.ReadInteger('MANDATORY', 'AutoType' + IntToStr(i) + 'IMP', 1);
       AutoMarka[i].Marka := INI.ReadString('MANDATORY', 'AutoType' + IntToStr(i), '');
   end;
   DefType := INI.ReadInteger('MANDATORY', 'DefAutoTypeHideIMP', 0);
   for j := 0 to 1024 do begin
       if INI.ReadString('MANDATORY', 'AutoType' + IntToStr(j) + 'Hide', '') = '' then break;
       SetLength(AutoMarka, i + j + 1);
       AutoMarka[i + j].IsImport := INI.ReadInteger('MANDATORY', 'AutoType' + IntToStr(j) + 'HideIMP', DefType);
       AutoMarka[i + j].Marka := INI.ReadString('MANDATORY', 'AutoType' + IntToStr(j) + 'Hide', '');
       while Pos('_', AutoMarka[i + j].Marka) <> 0 do
           AutoMarka[i + j].Marka[Pos('_', AutoMarka[i + j].Marka)] := ' ';
   end;
   INI.Free;

   ExpState := true;
   _ErrCollectStr := '';
   INI := TIniFile.Create('blank.ini');
   TemplateDir := INI.ReadString('MANDATORY', 'BuroDirectory', '');
   COMPANY_CODE := INI.ReadString('MANDATORY', 'BuroCompany', '');
   INI.Free;

   if (COMPANY_CODE < '1') OR (COMPANY_CODE > '9') then begin
      MessageDlg('Не определен код компании, необходимый для БЮРО', mtInformation, [mbOK], 0);
      exit;
   end;

   if Length(TemplateDir) = 0 then begin
      MessageDlg('Не указан путь к файлам-шаблонам', mtInformation, [mbOK], 0);
      exit;
   end;

   if TemplateDir[Length(TemplateDir)] <> '\' then
      TemplateDir := TemplateDir + '\';

   DestDir := ExpBUROForm.OutDirectory.Text;
   if DestDir[Length(DestDir)] <> '\' then
      DestDir := DestDir + '\';

   if not FileExists(TemplateDir + 'Report1.DBF') OR
      not FileExists(TemplateDir + 'Report2.DBF') OR
      not FileExists(TemplateDir + 'Report3.DBF') OR
      not FileExists(TemplateDir + 'Report4.DBF') OR
      not FileExists(TemplateDir + 'Report5.DBF') OR
      not FileExists(TemplateDir + 'Report6.DBF') then begin
      MessageDlg('Не существует файл-шаблон в директории ' + TemplateDir, mtInformation, [mbOK], 0);
      exit;
      end;

    Appendix := 'TASK';

    DBFileName1 := DestDir + 'N1' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName2b := DestDir + 'N2' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName3b := DestDir + 'N3' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName4b := DestDir + 'N4' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName5b := DestDir + 'N5' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName6 := DestDir + 'N6' + COMPANY_CODE + '_' + Appendix + '.DBF';

    DBFileName2 := DestDir + 'X2' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName3 := DestDir + 'X3' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName4 := DestDir + 'X4' + COMPANY_CODE + '_' + Appendix + '.DBF';
    DBFileName5 := DestDir + 'X5' + COMPANY_CODE + '_' + Appendix + '.DBF';

    StatusBar.Panels[0].Text := 'Копированте файлов...';
    StatusBar.Update;

    if FileExists(DBFileName1) AND not DeleteFile(DBFileName1) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName1, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName2) AND not DeleteFile(DBFileName2) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName2, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName3) AND not DeleteFile(DBFileName3) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName3, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName4) AND not DeleteFile(DBFileName4) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName4, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName5) AND not DeleteFile(DBFileName5) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName5, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName6) AND not DeleteFile(DBFileName6) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName6, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName2b) AND not DeleteFile(DBFileName2b) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName2b, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName3b) AND not DeleteFile(DBFileName3b) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName3b, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName4b) AND not DeleteFile(DBFileName4b) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName4b, mtInformation, [mbOK], 0);
      exit;
    end;
    if FileExists(DBFileName5b) AND not DeleteFile(DBFileName5b) then begin
      MessageDlg('Ошибка удаления файла' + DBFileName5b, mtInformation, [mbOK], 0);
      exit;
    end;
    try
      CopyFile(TemplateDir + 'Report1.DBF', DBFileName1, ProgressBar);
      CopyFile(TemplateDir + 'Report2.DBF', DBFileName2, ProgressBar);
      CopyFile(TemplateDir + 'Report3.DBF', DBFileName3, ProgressBar);
      CopyFile(TemplateDir + 'Report4.DBF', DBFileName4, ProgressBar);
      CopyFile(TemplateDir + 'Report5.DBF', DBFileName5, ProgressBar);
      CopyFile(TemplateDir + 'Report6.DBF', DBFileName6, ProgressBar);
      CopyFile(TemplateDir + 'Report2.DBF', DBFileName2b, ProgressBar);
      CopyFile(TemplateDir + 'Report3.DBF', DBFileName3b, ProgressBar);
      CopyFile(TemplateDir + 'Report4.DBF', DBFileName4b, ProgressBar);
      CopyFile(TemplateDir + 'Report5.DBF', DBFileName5b, ProgressBar);
    except
      ExpState := false;
    end;
    StatusBar.Panels[0].Text := '';
    StatusBar.Update;

    if not ExpState then begin
       MessageDlg('Ошибка копирования файлов', mtInformation, [mbOK], 0);
       exit;
    end;

    StatusBar.Panels[0].Text := 'Подготовка данных. Ждите...';
    StatusBar.Update;

    if ExpBUROForm.PeriodFrom.Checked then
        PeriodStr := ' UPDDT>=''' + DateToStr(ExpBUROForm.PeriodFrom.Date) + '''';
    if ExpBUROForm.PeriodTo.Checked then begin
        if PeriodStr <> '' then
            PeriodStr := PeriodStr + ' AND ';
        PeriodStr := PeriodStr + ' UPDDT<=''' + DateToStr(ExpBUROForm.PeriodTo.Date) + '''';
    end;
    if PeriodStr = '' then
        PeriodStr := '0=0';

    try
        SQLDB.StartTransaction;
        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add('UPDATE SYSADM.MANDATOR M SET DIVISION=(SELECT DIVISIONCODE FROM AGENT A WHERE A.AGENT_CODE=M.AGNT) WHERE DIVISION IS NULL');
        WorkSQL.ExecSQL;

        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add('DELETE FROM SYSADM.EXPORTERRS');
        WorkSQL.ExecSQL;

        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add('UPDATE SYSADM.MANDATOR SET CARCODE=0,NMBODY='''',UPDDT=SYSDATE WHERE NMBODY=''БН'' ');
        WorkSQL.ExecSQL;

        SQLDB.Commit;

        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add('SELECT COUNT(*) FROM SYSADM.MANDATOR WHERE DIVISION IS NULL AND ' + PeriodStr);
        WorkSQL.Open;
        if WorkSQL.Fields[0].AsInteger > 0 then begin
            MessageDlg('В выборке для экспорта обнаружены неопределённые подразделения', mtInformation, [mbOk], 0);
            exit;
        end;
        WorkSQL.Close;
    except
      on E : Exception do begin
        SQLDB.Rollback;
        MessageDlg('Ошибка очистки ' + E.Message, mtError, [mbOk], 0);
        ExpState := false;
        exit;
      end;
    end;

    try
        StatusBar.Panels[0].Text := 'Поиск полисов. Ждите...';
        StatusBar.Update;

        ExportBUROSQL.MacroByName('WHERE').AsString := PeriodStr;
        ExportBUROSQL.Open;

        fldList := 'Список полей'#13#10;
        for i := 1 to ExportBUROSQL.FieldCount do
            fldList := fldList + ExportBUROSQL.Fields[i - 1].FieldName + #13#10;
        //MessageDlg(fldList, mtInformation, [mbOk], 0);

        StatusBar.Panels[0].Text := 'Подсчёт количества полисов. Ждите...';
        StatusBar.Update;
        CountPolises := ExportBUROSQL.RecordCount;
    except
        on E : Exception do begin
            CountPolises := 1;
            MessageDlg('Ошибка создания запроса или определения количества найденных полисов. Нажми кнопку для продолжения работы.'#13 + e.Message, mtInformation, [mbOk], 0);
            ExpState := false;
        end
        else
            MessageDlg('Критическая Ошибка (2) создания запроса.', mtInformation, [mbOk], 0);
        exit;
    end;

    StatusBar.Panels[0].Text := 'Открытие таблиц. Ждите...';
    StatusBar.Update;

    Table1 := TTable.Create(self);
    Table2 := TTable.Create(self);
    Table3 := TTable.Create(self);
    Table4 := TTable.Create(self);
    Table5 := TTable.Create(self);
    Table6 := TTable.Create(self);

    Table1.TableName := DBFileName1;
    Table2.TableName := DBFileName2;
    Table3.TableName := DBFileName3;
    Table4.TableName := DBFileName4;
    Table5.TableName := DBFileName5;
    Table6.TableName := DBFileName6;

    try
        Table1.Open;
        Table2.Open;
        Table3.Open;
        Table4.Open;
        Table5.Open;
        Table6.Open;

        SetLength(AOwn, 64);

        StatusBar.Panels[0].Text := 'Экспортирую ' + IntToStr(CountPolises) + ' полисов...';
        StatusBar.Update;

        ProgressBar.Max := 100;
        StartTime := GetCurrentTime;
        i := 0;

        if ExpBUROForm.IsBackgroundMode.Checked then
            WindowState := wsMinimized;

        while not ExportBUROSQL.EOF do begin
            Inc(i);
            if(GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 then begin
                while (GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 do begin
                end;
                while PeekMessage(amsg, 0, WM_KEYDOWN, WM_KEYDOWN, PM_REMOVE) do;
                while PeekMessage(amsg, 0, WM_KEYUP, WM_KEYUP, PM_REMOVE) do;

                if ExpBUROForm.IsBackgroundMode.Checked then
                    WindowState := wsNormal;
                if MessageDlg('Вы хотите прервать процесс экспорта данных?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
                    ExpState := false;
                    break;
                end;
                if ExpBUROForm.IsBackgroundMode.Checked then
                    WindowState := wsMinimized;
            end;

            SERIA := ExportBUROSQL.FieldByName('SER').AsString;
            NUMBER := ExportBUROSQL.FieldByName('NMB').AsInteger;
            _ExceptSN := '';

            if ExportBUROSQL.FieldByName('STATE').AsString = '2' then
                if ExportBUROSQL.FieldByName('STOPDATE').IsNull then
                    AddError(SERIA, NUMBER, _ExceptSN + ' Код агента = ''' + ExportBUROSQL.FieldByName('AGNT').AsString + '''', eeNoStopDate, _ErrCollectStr);

            if ExportBUROSQL.FieldByName('STATE').AsString = '5' then
                if ExportBUROSQL.FieldByName('ISPAY2').AsString = 'Y' then
                    AddError(SERIA, NUMBER, _ExceptSN + ' Код агента = ''' + ExportBUROSQL.FieldByName('AGNT').AsString + '''', eeNoNeedPay, _ErrCollectStr);

            if not IsValidPolisNumber(PolisNmbsArray, SERIA, NUMBER) then
                 AddError(SERIA, NUMBER, _ExceptSN, eeBadPolisNumber, _ErrCollectStr);

            if ExportBUROSQL.FieldByName('DIVISION').IsNull then begin
                AddError(SERIA, NUMBER, _ExceptSN + ' Код агента = ''' + ExportBUROSQL.FieldByName('AGNT').AsString + '''', eeNoDivision, _ErrCollectStr);
                ExpState := false;
                DIV_CODE := -1
            end
            else
                DIV_CODE := ExportBUROSQL.FieldByName('DIVISION').AsInteger;

            ProgressBar.Position := ROUND(i * 100 / CountPolises);

            CurrTime := ROUND((GetCurrentTime - StartTime) / 1000);

            if CurrTime > 3600 then begin
               TimeStr := IntToStr(TRUNC(CurrTime / 3600)) + ' час ';
               CurrTime := CurrTime - TRUNC(CurrTime / 3600) * 3600;
               TimeStr := TimeStr + IntToStr(TRUNC(CurrTime / 60)) + ' мин';
            end else
            if CurrTime > 60 then begin
               TimeStr := IntToStr(TRUNC(CurrTime / 60)) + ' мин ';
               CurrTime := CurrTime - TRUNC(CurrTime / 60) * 60;
               TimeStr := TimeStr + IntToStr(CurrTime) + ' сек';
            end else
               TimeStr := IntToStr(CurrTime) + ' сек';

            StatusBar.Panels[1].Text := 'N' + IntToStr(i + 1) + ', ' + TimeStr + ', Колич. ошибок ' + IntToStr(ErrorCount);
            if ExpBUROForm.IsBackgroundMode.Checked then
                Caption := StatusBar.Panels[1].Text;
            StatusBar.Update;

            OwnCounter := 0;

            //Table 3 (Данные о владельце)
            //Part 1 (other owners)
            ExportOwners.Close;
            ExportOwners.MacroByName('WHERE').AsString := 'SER=''' + SERIA + ''' AND NMB=' + IntToStr(NUMBER);
            ExportOwners.Open;
            ExpSQL := ExportOwners;
            _ExceptSN := {SERIA + '/' + IntToStr(NUMBER) + }' Доп. Владельцы';
            while not ExpSQL.Eof do begin
                Table3.Append;
                Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table3.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

                CarOwnDescr_.Name := ExpSQL.FieldByName('NAME').AsString;
                CarOwnDescr_.Addr_code := ExpSQL.FieldByName('PLACE_CODE').AsString;
                CarOwnDescr_.Address := ExpStr(ExpSQL.FieldByName('ADDR').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                CarOwnDescr_.DIVISION := DIV_CODE;
                if IsBad then
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbName, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                if IsEnglish then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeEnglSymbAddr, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильный АДРЕС (англ. буквы) ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                end;
                CarOwnDescr_.Owner_Type := ExpSQL.FieldByName('PRAVO').AsInteger;
                CarOwnDescr_.R_N := 'Y';
                CarOwnDescr_.F_L := GetFizUrid(ExpSQL.FieldByName('TYPE').AsString);
                CarOwnDescr_.DIVISION := DIV_CODE;
                if CarOwnDescr_.F_L = '?' then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadFizUr, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружен неправильный признак Физ/Юр ' + ExpSQL.FieldByName('TYPE').AsString +
                   //                           ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                   ExpState := false;
                end;
                CarOwnDescr_.DISCOUNT_2 := ExpSQL.FieldByName('K').AsFloat;
                CarOwnDescr_.DVR_EXPR := ExpSQL.FieldByName('YEARPERM').AsInteger;

                Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNRCODE').AsInteger, CarOwnDescr_, NeedUpdate)); //OWNERCODE
                Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('NAME').AsString, IsRussian, IsEnglish, IsDigit, IsBad, false);
                if IsBad then
                    AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbName, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                //Английские буквы
                //if IsEnglish then begin
                //   errCollectStr := errCollectStr + ExceptSN + 'Не правильное ИМЯ (англ. буквы) (Владельцы) ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                //end;
                //Начинается с цифры
                if Length(Table3.FieldByName('NAME').AsString) > 0 then
                    if ('0' >= Table3.FieldByName('NAME').AsString[1]) AND (Table3.FieldByName('NAME').AsString[1] <= '9') then begin
                        AddError(SERIA, NUMBER, _ExceptSN, eeStartFromDigit, _ErrCollectStr);
                       //errCollectStr := errCollectStr + ExceptSN  + ' Не правильное ИМЯ (начин. с цифры) ' +
                       //                                 ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                    end;
                Table3.FieldByName('ADDR_CODE').AsString := ExpStr(ExpSQL.FieldByName('PLACE_CODE').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                if IsRussian then begin
                   AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.FieldByName('PLACE_CODE').AsString, eeRussLetters, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильный код (русск. буквы) ' +
                   //                           ' ' + ExpSQL.FieldByName('PLACE_CODE').AsString + #13#10;
                end;
                Table3.FieldByName('ADDRESS').AsString := ExpSQL.FieldByName('ADDR').AsString;
                Table3.FieldByName('OWNER_TYPE').AsString := GetVladenType(ExpSQL.FieldByName('PRAVO').AsInteger);

                if Table3.FieldByName('OWNER_TYPE').AsString = '?' then begin
                    Table3.FieldByName('OWNER_TYPE').AsString := 'D'; //доверенность
                end;
                Table3.FieldByName('R_N').AsString := 'R'; //Always resident
                Table3.FieldByName('F_L').AsString := GetFizUrid(ExpSQL.FieldByName('TYPE').AsString);
                if Table3.FieldByName('F_L').AsString = '?' then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadFizUr, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружен неправильный признак Физ/Юр ' + ExpSQL.FieldByName('OWNTYPE').AsString +
                   //                                            ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                   ExpState := false;
                end;
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('K').AsFloat);
                if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadDiscount, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружен неправильный тип скидки ' + ExpSQL.FieldByName('K').AsString +
                   //                                 ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                   ExpState := false;
                end;
                Table3.FieldByName('DRV_EXPR').Value := ExpSQL.FieldByName('YEARPERM').Value;
                Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
                if(NeedUpdate) then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDOWN', 'OWNRCODE', 'NAME', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), CarOwnDescr_.Name, false);
                Table3.Post;

                AOwn[OwnCounter].FIO := Table3.FieldByName('NAME').AsString;
                AOwn[OwnCounter].Code := Table3.FieldByName('MSG_ID').AsString;
                Inc(OwnCounter);

                //Table 2 (Данные о других владельцах)
                Table2.Append;
                Table2.FieldByName('POLICY_NUM').AsString := SERIA + ExpZeroNmb(NUMBER);
                Table2.FieldByName('OWNER_ID').AsString := Table3.FieldByName('MSG_ID').AsString;
                Table2.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table2.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);
                Table2.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
                Table2.Post;

                ExpSQL.Next;
            end;
            //END Part 1 (other owners)
            //End Table 3, Table 2 (other owners)

            ExpSQL.Close;
            //Заполнение данных о транспортном ср-ве + основной владелец + авто

            ExpSQL := ExportBUROSQL;
            _ExceptSN := {SERIA + '/' + IntToStr(NUMBER) + }' Полис';

            if ExportBUROSQL.FieldByName('PSER').AsString = '' then
                if (ExportBUROSQL.FieldByName('PAY1DT').AsDateTime < EncodeDate(2000, 7, 1)) AND
                   (ExportBUROSQL.FieldByName('VER').AsInteger <> 1) OR
                   (ExportBUROSQL.FieldByName('PAY1DT').AsDateTime >= EncodeDate(2000, 7, 1)) AND
                   (ExportBUROSQL.FieldByName('VER').AsInteger = 1) then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eePayAndVer, _ErrCollectStr);
                   //errCollectStr := errCollectStr +  ExceptSN + ' Несовпадение даты оплаты и версии полиса' + #13#10;
                   ExpState := false;
                end;

            /////////////////////////////////////////////////////////////////////////////////////////////
            //Главный владелец
            Table3.Append;
            Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table3.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

            INI := TIniFile.Create('blank.ini');
            ResidentStr := INI.ReadString('MANDATORY', 'TarifTable_Is_Resident' + ExpSQL.FieldByName('TRFTBL').AsString, '');
            INI.Free;

            CarOwnDescr_.Name := ExpSQL.FieldByName('NAME').AsString;
            CarOwnDescr_.Addr_code := ExpSQL.FieldByName('PLACE_CODE').AsString;
            CarOwnDescr_.Address := ExpStr(ExpSQL.FieldByName('ADDR').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if Length(Trim(CarOwnDescr_.Address)) < 3 then
               AddError(SERIA, NUMBER, _ExceptSN, eeEmptyAddr, _ErrCollectStr);
            CarOwnDescr_.DIVISION := DIV_CODE;

            if IsBad then
               AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbAddr, _ErrCollectStr);
            if IsEnglish then begin
                AddError(SERIA, NUMBER, _ExceptSN, eeEnglSymbAddr, _ErrCollectStr);
            end;
            CarOwnDescr_.Owner_Type := ExpSQL.FieldByName('VLADEN').AsInteger;
            if ResidentStr = 'YES' then
                CarOwnDescr_.R_N := GetResident(true)
            else
                CarOwnDescr_.R_N := GetResident(false);
            //CarOwnDescr.R_N := 'Y';
            CarOwnDescr_.F_L := GetFizUrid(ExpSQL.FieldByName('OWNTYPE').AsString);
            if CarOwnDescr_.F_L = '?' then begin
               if ExpSQL.FieldByName('VER').AsInteger = 1 then
                   CarOwnDescr_.F_L := ''
               else begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadFizUr, _ErrCollectStr); 
                   ExpState := false;
               end;
            end;
            CarOwnDescr_.DISCOUNT_2 := ExpSQL.FieldByName('K').AsFloat;
            CarOwnDescr_.DVR_EXPR := -1;

            Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNERCODE').AsInteger, CarOwnDescr_, NeedUpdate)); //OWNERCODE;
            Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('NAME').AsString, IsRussian, IsEnglish, IsDigit, IsBad, false);
            if IsBad then
                AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbName, _ErrCollectStr);
               //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
            //Английские буквы
            //if IsEnglish then begin
            //   errCollectStr := errCollectStr + ExceptSN + ' Не правильное ИМЯ (англ. буквы) ' +
            //                                    ExpSQL.FieldByName('NAME').AsString + #13#10;
            //end;
            //Начинается с цифры
            if Length(Table3.FieldByName('NAME').AsString) > 0 then
                if ('0' >= Table3.FieldByName('NAME').AsString[1]) AND (Table3.FieldByName('NAME').AsString[1] <= '9') then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeStartFromDigit, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильное ИМЯ (начин. с цифры) ' +
                   //                                 ExpSQL.FieldByName('NAME').AsString + #13#10;
                end;

            Table3.FieldByName('ADDR_CODE').AsString := ExpStr(ExpSQL.FieldByName('PLACE_CODE').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if IsRussian then begin
               AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.FieldByName('PLACE_CODE').AsString, eeRussLetters, _ErrCollectStr);
               //errCollectStr := errCollectStr + ExceptSN + ' Не правильный код (русск. буквы) ' +
               //                                 ExpSQL.FieldByName('PLACE_CODE').AsString + #13#10;
            end;
            Table3.FieldByName('ADDRESS').AsString := ExpSQL.FieldByName('ADDR').AsString;
            Table3.FieldByName('OWNER_TYPE').AsString := GetVladenType(ExpSQL.FieldByName('VLADEN').AsInteger);
            if Table3.FieldByName('OWNER_TYPE').AsString = '?' then begin
               if ExpSQL.FieldByName('VER').AsInteger = 1 then
                   Table3.FieldByName('OWNER_TYPE').AsString := 'S' //Основной владелец
               else begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadVladen, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Неправильный признак основания на владение'#13#10;
                   ExpState := false;
               end;
            end;

            if (ResidentStr <> 'YES') AND (ResidentStr <> 'NO') then begin
               //errCollectStr := errCollectStr + 'Не определён признак (резидент, не резидент) таблицы тарифов N' + ExpSQL.FieldByName('TRFTBL').AsString + #13#10;
               MessageDlg('Не определён признак (резидент, не резидент) таблицы тарифов N' + ExpSQL.FieldByName('TRFTBL').AsString + #13#10, mtInformation, [mbOk], 0);
               ExpState := false;
            end;

            if ResidentStr = 'YES' then
                Table3.FieldByName('R_N').AsString := GetResident(true)
            else
                Table3.FieldByName('R_N').AsString := GetResident(false);

            Table3.FieldByName('F_L').AsString := CarOwnDescr_.F_L ;//GetFizUrid(ExpSQL.FieldByName('OWNTYPE').AsString);
            if ExpSQL.FieldByName('VER').AsInteger = 1 then
                Table3.FieldByName('DISCOUNT_2').AsString := '0'
            else
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('K').AsFloat);
            if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
               AddError(SERIA, NUMBER, _ExceptSN, eeBadDiscount, _ErrCollectStr); 
               //errCollectStr := errCollectStr + ExceptSN + ' Обнаружен неправильный тип скидки ' + ExpSQL.FieldByName('K').AsString + #13#10;
               ExpState := false;
            end;
            Table3.FieldByName('DRV_EXPR').Clear;
            Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDATOR', 'OWNERCODE', '', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), '', false);
            Table3.Post;

            AOwn[OwnCounter].FIO := Table3.FieldByName('NAME').AsString;
            AOwn[OwnCounter].Code := Table3.FieldByName('MSG_ID').AsString;
            Inc(OwnCounter);
            //END Главный владелец

            // Данные о серийном ТС
            Table5.Append;
            Table5.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table5.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

            SerCarDesrc_.Marka := ExpSQL.FieldByName('MARKA').AsString;
            SerCarDesrc_.DIVISION := DIV_CODE;

            IsMarkaValue := IsMarka(ListNoCkeckAN, ExpSQL.FieldByName('LTR').AsString, AutoMarka, SerCarDesrc_.Marka);
            if IsMarkaValue = -1 then begin //не найдена!!!
                AddError(SERIA, NUMBER, _ExceptSN, eeBadAutoMark, _ErrCollectStr);
                //errCollectStr := errCollectStr + ExceptSN + ' не известная марка авто ' + SerCarDesrc.Marka + #13#10;
                ExpState := false;
            end;
            SerCarDesrc_.TypeTC := ExpSQL.FieldByName('LTR').AsString;
            if SerCarDesrc_.TypeTC = '' then begin
                //errCollectStr := errCollectStr + ExceptSN + ' не опреден тип ТС' + #13#10;
                AddError(SERIA, NUMBER, _ExceptSN, eeBadTC, _ErrCollectStr);
                ExpState := false;
            end;
            SerCarDesrc_.Characteristic := GetCharacteristicAuto(SerCarDesrc_.TypeTC,
                                              ExpSQL.FieldByName('VOLUME').AsFloat,
                                              ExpSQL.FieldByName('TONNAGE').AsFloat,
                                              ExpSQL.FieldByName('CNTPAS').AsFloat,
                                              ExpSQL.FieldByName('POWER').AsFloat,
                                              ResidentStr = 'YES');

            if SerCarDesrc_.Characteristic < 0 then begin
                AddError(SERIA, NUMBER, _ExceptSN, eeBadTC, _ErrCollectStr);
                //errCollectStr := errCollectStr + ExceptSN + ' не определена характеристика ТС' + #13#10;
                ExpState := false;
            end;

            Table5.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeSerTC(ExpSQL.FieldByName('BASECARCODE').AsInteger, SerCarDesrc_, NeedUpdate)); //BASECARCODE
            Table5.FieldByName('VEH_MARK').AsString := ExpStr(ExpSQL.FieldByName('MARKA').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if IsBad then
                AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbMarka, _ErrCollectStr);
               //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('MARKA').AsString + #13#10;
            Table5.FieldByName('VEH_B_TYPE').AsString := ExpStr(ExpSQL.FieldByName('LTR').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if IsRussian then begin
                AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.FieldByName('LTR').AsString, eeRussLetters, _ErrCollectStr);
                //errCollectStr := errCollectStr + ExceptSN + ' не правильное (русск. буквы) ТИП ТС ' + ExpSQL.FieldByName('LTR').AsString + #13#10;
                ExpState := false;
            end;

            CharAuto := SerCarDesrc_.Characteristic;
            if CharAuto > 0 then
                Table5.FieldByName('CHAR_VALUE').AsFloat := CharAuto;
            if CharAuto < 0 then begin
                AddError(SERIA, NUMBER, _ExceptSN, eeBadCharAuto, _ErrCollectStr);
                //errCollectStr := errCollectStr + ExceptSN + ' не правильная характеристика ТС ' + ExpSQL.FieldByName('LTR').AsString + #13#10;
                ExpState := false;
            end;

            Table5.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDATOR', 'BASECARCODE', '', DecodeMsgId(Table5.FieldByName('MSG_ID').AsString), '', false);
            Table5.Post;
            //END Данные о серийном ТС

            // Данные о ТС
            Table4.Append;
            Table4.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table4.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

            CarDescr_.AutoNumber := ExpSQL.FieldByName('AUTONMB').AsString;
            CarDescr_.BodyNumber := ExpSQL.FieldByName('NMBODY').AsString;
            CarDescr_.ShassiNumber := ExpSQL.FieldByName('CHASSIS').AsString;
            CarDescr_.DIVISION := DIV_CODE;
            if (Pos('.', CarDescr_.BodyNumber) <> 0) OR
               (Pos(',', CarDescr_.BodyNumber) <> 0) OR
               (Pos('=', CarDescr_.BodyNumber) <> 0) then
                   CarDescr_.BodyNumber := '';
            if (Pos('.', CarDescr_.ShassiNumber) <> 0) OR
               (Pos(',', CarDescr_.ShassiNumber) <> 0) OR
               (Pos('=', CarDescr_.ShassiNumber) <> 0) then
                   CarDescr_.ShassiNumber := '';

            CarDescr_.BaseCarCode := StrToInt(DecodeMsgId(Table5.FieldByName('MSG_ID').AsString));

            Table4.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeTC(ExpSQL.FieldByName('CARCODE').AsInteger, CarDescr_, NeedUpdate)); //CARCODE
            CURR_CAR_CODE := Table4.FieldByName('MSG_ID').AsString;

            Table4.FieldByName('NUM_SIGN').AsString := ExpSQL.FieldByName('AUTONMB').AsString;
            if (Table4.FieldByName('NUM_SIGN').AsString = '') OR (Table4.FieldByName('NUM_SIGN').AsString = '-') then
                Table4.FieldByName('NUM_SIGN').AsString := 'Б/Н';

            //В стороке м.б. K1=1.2, K2=3
            if(ExpSQL.FieldByName('NMBODY').AsString <> '') AND
              (Pos('=', ExpSQL.FieldByName('NMBODY').AsString) = 0) AND
              (Pos('.', ExpSQL.FieldByName('NMBODY').AsString) = 0) AND
              (Pos(',', ExpSQL.FieldByName('NMBODY').AsString) = 0) then
                 Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('NMBODY').AsString;

            if(ExpSQL.FieldByName('CHASSIS').AsString <> '') AND
              (Pos('=', ExpSQL.FieldByName('CHASSIS').AsString) = 0) AND
              (Pos('.', ExpSQL.FieldByName('CHASSIS').AsString) = 0) AND
              (Pos(',', ExpSQL.FieldByName('CHASSIS').AsString) = 0) then
                 Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('CHASSIS').AsString;

            ExpStr(Table4.FieldByName('BODY_NO').AsString, IsRussian, IsEnglish, IsDigit, IsBad, false);
            if IsBad then
                AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbBodyNo, _ErrCollectStr);
               //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + Table4.FieldByName('BODY_NO').AsString + #13#10;
            if IsRussian AND (IsMarkaValue = 1) then begin //русские не в русских машинах!!!
                AddError(SERIA, NUMBER, _ExceptSN, eeRussBodyShassi, _ErrCollectStr);
                //errCollectStr := errCollectStr + ExceptSN + ' не правильное (русск. буквы) номер кузова/шасси' + #13#10;
                ExpState := false;
            end;

            Table4.FieldByName('BASECAR_ID').AsString := Table5.FieldByName('MSG_ID').AsString;
            Table4.FieldByName('COUNTRY').Clear;

            Table4.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDATOR', 'CARCODE', '', DecodeMsgId(Table4.FieldByName('MSG_ID').AsString), '', false);
            Table4.Post;
            //END Данные о ТС

            //Договор
            Table1.Append;
            Table1.FieldByName('POLICY_NUM').AsString := SERIA + ExpZeroNmb(NUMBER);
            //if ResidentStr = 'YES' then
                Table1.FieldByName('INS_DOC').AsString := 'S';
            //else
            //    Table1.FieldByName('INS_DOC').AsString := 'P';
            Table1.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table1.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

            if ExpSQL.FieldByName('PNMB').AsInteger > 0 then begin //Дубликат
                if ExpSQL.FieldByName('REGDATE').AsDateTime >= (ExpSQL.FieldByName('TODT').AsDateTime) then begin
                    AddError(SERIA, NUMBER, _ExceptSN, eeSellAfterStart, _ErrCollectStr);
                    ExpState := false;
                end;
            end
            else begin
                DeltaDays := 0;
                if ExpSQL.FieldByName('OWNTYPE').AsString = 'U' then DeltaDays := 15;
                if ExpSQL.FieldByName('REGDATE').AsDateTime > (ExpSQL.FieldByName('FRDT').AsDateTime + DeltaDays) then begin
                    AddError(SERIA, NUMBER, _ExceptSN, eeSellAfterStart, _ErrCollectStr);
                    ExpState := false;
                end;
            end;
            Table1.FieldByName('DATE_DISTR').AsDateTime := ExpSQL.FieldByName('REGDATE').AsDateTime;
            Table1.FieldByName('DATE_BEG').AsDateTime := ExpSQL.FieldByName('FRDT').AsDateTime;
            Table1.FieldByName('TIME_BEG').AsFloat := ExpTime(ExpSQL.FieldByName('FRTM').AsString);
            if ExpSQL.FieldByName('TODT').AsDateTime > Date + 400 then begin
                AddError(SERIA, NUMBER, _ExceptSN, eeBadDateEnd, _ErrCollectStr);
                ExpState := false;
            end;
            Table1.FieldByName('DATE_END').AsDateTime := ExpSQL.FieldByName('TODT').AsDateTime;
            if ExpStateFunc(ExpSQL.FieldByName('State').AsInteger) = '' then begin
                //errCollectStr := errCollectStr + ExceptSN + ' не правильное состояние' + #13#10;
                AddError(SERIA, NUMBER, _ExceptSN, eeBadState, _ErrCollectStr);
                ExpState := false;
            end;

            //Проверка дат
            if ExpSQL.FieldByName('PAY1DT').AsDateTime > ExpSQL.FieldByName('FRDT').AsDateTime then begin
                //errCollectStr := errCollectStr + ExceptSN + ' Оплата после начала' + #13#10;
                AddError(SERIA, NUMBER, _ExceptSN, eePayAfterStart, _ErrCollectStr);
                ExpState := false;
            end;

            if ExpSQL.FieldByName('PNMB').AsInteger = 0 then begin
	            if (ExpSQL.FieldByName('REGDATE').AsDateTime > (ExpSQL.FieldByName('PAY1DT').AsDateTime + 45)) OR (ExpSQL.FieldByName('REGDATE').AsDateTime < ExpSQL.FieldByName('PAY1DT').AsDateTime) then begin
                    //errCollectStr := errCollectStr + ExceptSN + ' дата регистрации и 1й оплаты не соотносятся' + #13#10;
                    AddError(SERIA, NUMBER, _ExceptSN, eeRegAndPay, _ErrCollectStr);
                    ExpState := false;
                end;
            end;

            if not ExpSQL.FieldByName('PAY2DT').IsNull then begin
	            if (ExpSQL.FieldByName('PAY2DT').AsDateTime <= ExpSQL.FieldByName('PAY1DT').AsDateTime) OR (ExpSQL.FieldByName('PAY2DT').AsDateTime >= ExpSQL.FieldByName('TODT').AsDateTime) then begin
                    //errCollectStr := errCollectStr + ExceptSN + ' дата 2й оплаты не в диапазоне' + #13#10;
                    AddError(SERIA, NUMBER, _ExceptSN, eeBadPay2, _ErrCollectStr);
                    ExpState := false;
                end;
            end;

            for fi := 0 to ExpSQL.Fields.Count - 1 do begin
                if (ExpSQL.Fields[fi].DataType = ftDate) AND (not ExpSQL.Fields[fi].IsNull) then begin
                    if (ExpSQL.Fields[fi].AsDateTime < EncodeDate(1999, 7, 1)) OR (ExpSQL.Fields[fi].AsDateTime > EncodeDate(2020, 1, 1)) then begin
                        //errCollectStr := errCollectStr + ExceptSN + ' стрянная дата в поле ' + ExpSQL.Fields[fi].FieldName + #13#10;
                        AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.Fields[fi].FieldName, eeStrangeDate, _ErrCollectStr);
                        ExpState := false;
                    end;
                end;
            end;

            Table1.FieldByName('POL_STATE').AsString := ExpStateFunc(ExpSQL.FieldByName('State').AsInteger);
            Table1.FieldByName('DATE_TERM').Value := ExpSQL.FieldByName('STOPDATE').Value;

            if not ExpSQL.FieldByName('STOPDATE').IsNull then
                if (ExpSQL.FieldByName('STOPDATE').AsDateTime < ExpSQL.FieldByName('FRDT').AsDateTime) OR
                   (ExpSQL.FieldByName('STOPDATE').AsDateTime > ExpSQL.FieldByName('TODT').AsDateTime) then begin
                        //errCollectStr := errCollectStr + ExceptSN + ' дата остановки не в диапазоне '#13#10;
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadStopDate, _ErrCollectStr);
                        ExpState := false;
                 end;

            Table1.FieldByName('INSUR_ID').AsString := Table3.FieldByName('MSG_ID').AsString;
            Table1.FieldByName('VEH_TYPE').AsString := ExpSQL.FieldByName('LTR').AsString;
            ExpStr(Table1.FieldByName('VEH_TYPE').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if IsRussian then begin
                //errCollectStr := errCollectStr + ExceptSN + ' не правильное тип ТС (русск. буквы)' + #13#10;
                AddError(SERIA, NUMBER, _ExceptSN, eeBadTC, _ErrCollectStr);
                ExpState := false;
            end;

            if ExpSQL.FieldByName('DISCOUNT').AsInteger > 0 then
                Table1.FieldByName('DISCOUNT_1').AsString := 'H'
            else
                Table1.FieldByName('DISCOUNT_1').AsString := 'F';

            if ExpSQL.FieldByName('VER').AsInteger = 1 then
                Table1.FieldByName('DISCOUNT_3').AsString := 'A0'
            else
                Table1.FieldByName('DISCOUNT_3').AsString := GetAvK(ExpSQL.FieldByName('K2').AsFloat);//ExpSQL.FieldByName('CLAV').AsString;

            if Table1.FieldByName('DISCOUNT_3').AsString = '?' then
                AddError(SERIA, NUMBER, _ExceptSN, eeBadAvK, _ErrCollectStr);

            ExpStr(Table1.FieldByName('DISCOUNT_3').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
            if IsRussian then begin
                //errCollectStr := errCollectStr + ExceptSN + ' не правильная скидка за аварийность (русск. буквы)' + #13#10;
                AddError(SERIA, NUMBER, _ExceptSN, eeBadDiscount, _ErrCollectStr);
                ExpState := false;
            end;

            Table1.FieldByName('CAR_ID').AsString := Table4.FieldByName('MSG_ID').AsString;

            Table1.FieldByName('DATE_PREM1').Value := ExpSQL.FieldByName('PAY1DT').Value;
            Table1.FieldByName('PREM_SUM1').Value := ExpSQL.FieldByName('PAY1').Value;
            Table1.FieldByName('DATE_PREM2').Value := ExpSQL.FieldByName('PAY2DT').Value;
            if not ExpSQL.FieldByName('PAY2DT').IsNull then begin
                if (abs(ExpSQL.FieldByName('FRDT').AsDateTime - ExpSQL.FieldByName('PAY2DT').AsDateTime) -
                    abs(ExpSQL.FieldByName('TODT').AsDateTime - ExpSQL.FieldByName('PAY2DT').AsDateTime)) > 200 then begin
                    AddError(SERIA, NUMBER, _ExceptSN, eeBadPay2, _ErrCollectStr);
                    ExpState := false;
                end
            end;
            Table1.FieldByName('PREM_SUM2').Value := ExpSQL.FieldByName('PAY2').Value;
            if ExpCurr(ExpSQL.FieldByName('PAY1CURR').AsString) = '' then begin
                AddError(SERIA, NUMBER, _ExceptSN, eeBadCurr, _ErrCollectStr);
                ExpState := false;
            end;
            Table1.FieldByName('CURRENCY').AsString := ExpCurr(ExpSQL.FieldByName('PAY1CURR').AsString);
            Table1.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;

            Table1.Post;
            //END Договор

            _ExceptSN := {SERIA + '/' + IntToStr(NUMBER) + }' Аварии';
            //ВЫПЛАТЫ
            ExportAvariasBURO.Close;
            ExportAvariasBURO.MacroByName('WHERE').AsString := 'SER=''' + SERIA + ''' AND NMB=' + IntToStr(NUMBER);
            ExportAvariasBURO.Open;
            ExpSQL := ExportAvariasBURO;

            while not ExpSQL.EOF do begin
                //Потерпевший
                Table3.Append;
                Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table3.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

                CarOwnDescr_.Name := ExpSQL.FieldByName('FIO').AsString;
                CarOwnDescr_.Addr_code := ExpSQL.FieldByName('ADRCODE').AsString;
                CarOwnDescr_.Address := ExpStr(ExpSQL.FieldByName('ADDR').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                CarOwnDescr_.DIVISION := DIV_CODE;
                if IsBad then
                    AddError(SERIA, NUMBER, _ExceptSN, eeEnglSymbAddr, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('ADDR').AsString + #13#10;
                if IsEnglish then begin
                    AddError(SERIA, NUMBER, _ExceptSN, eeEnglSymbAddr, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильный АДРЕС (англ. буквы) ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                end;
                CarOwnDescr_.Owner_Type := ExpSQL.FieldByName('VLADEN').AsInteger;
                CarOwnDescr_.R_N := ExpSQL.FieldByName('ISRSD').AsString;
                CarOwnDescr_.F_L := ExpSQL.FieldByName('ISFIZ').AsString;
                CarOwnDescr_.DISCOUNT_2 := ExpSQL.FieldByName('K').AsFloat;
                CarOwnDescr_.DVR_EXPR := -1;

                Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNRCODE').AsInteger, CarOwnDescr_, NeedUpdate)); //OWNERCODE
                Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('FIO').AsString, IsRussian, IsEnglish, IsDigit, IsBad, false);
                if IsBad then
                    AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbName, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                //Английские буквы
                if IsEnglish then begin
                    AddError(SERIA, NUMBER, _ExceptSN, eeEnglSymbName, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильное ИМЯ (англ. буквы) ' +
                   //                           ' ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                end;
                //Начинается с цифры
                if Length(Table3.FieldByName('NAME').AsString) > 0 then
                    if ('0' >= Table3.FieldByName('NAME').AsString[1]) AND (Table3.FieldByName('NAME').AsString[1] <= '9') then begin
                        AddError(SERIA, NUMBER, _ExceptSN, eeStartFromDigit, _ErrCollectStr);
                       //errCollectStr := errCollectStr + ExceptSN + ' Не правильное ИМЯ (начин. с цифры) ' +
                       //                                 '(аварии) ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                    end;
                Table3.FieldByName('ADDR_CODE').AsString := ExpStr(ExpSQL.FieldByName('ADRCODE').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                if IsRussian then begin
                    AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.FieldByName('ADRCODE').AsString, eeRussLetters, _ErrCollectStr);
                   //errCollectStr := errCollectStr + ExceptSN + ' Не правильный код (русск. буквы) ' +
                   //                           ' ' + ExpSQL.FieldByName('ADRCODE').AsString + #13#10;
                end;
                Table3.FieldByName('ADDRESS').AsString := ExpSQL.FieldByName('ADDR').AsString;
                Table3.FieldByName('OWNER_TYPE').AsString := GetVladenType(ExpSQL.FieldByName('VLADEN').AsInteger);
                if Table3.FieldByName('OWNER_TYPE').AsString = '?' then begin
                   Table3.FieldByName('OWNER_TYPE').AsString := 'S'; //собственник
                end;
                Table3.FieldByName('R_N').AsString := GetResident(ExpSQL.FieldByName('ISRSD').AsString = 'Y');
                Table3.FieldByName('F_L').AsString := GetFizUrid(ExpSQL.FieldByName('ISFIZ').AsString);
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('K').AsFloat);
                if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
                   //errCollectStr := errCollectStr + ExceptSN + ' Обнаружен не правильный тип скидки ' + ExpSQL.FieldByName('K').AsString +
                   //                           ' ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadDiscount, _ErrCollectStr);
                   ExpState := false;
                end;
                Table3.FieldByName('DRV_EXPR').Clear;
                Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
                if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDAV', 'OWNRCODE', 'N', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString, true);
                Table3.Post;

                if ExpSQL.FieldByName('MARKA').AsString <> '' then begin
                    //Автомобиль(Базовый) потерпевшего
                    Table5.Append;
                    Table5.FieldByName('COMPANY').AsString := COMPANY_CODE;
                    Table5.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

                    SerCarDesrc_.Marka := ExpSQL.FieldByName('MARKA').AsString;
                    SerCarDesrc_.DIVISION := DIV_CODE;
                    IsMarkaValue := IsMarka(ListNoCkeckAN, ExpSQL.FieldByName('LTR').AsString, AutoMarka, SerCarDesrc_.Marka);
                    if IsMarkaValue = -1 then begin
                       //errCollectStr := errCollectStr + ExceptSN + ' Не известная Марка Авто ''' + SerCarDesrc.Marka + ''' ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                       AddError(SERIA, NUMBER, _ExceptSN, eeBadAutoMark, _ErrCollectStr);
                       ExpState := false;
                    end;

                    SerCarDesrc_.TypeTC := ExpSQL.FieldByName('LTR').AsString;
                    ExpStr(SerCarDesrc_.TypeTC, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    if IsRussian OR ((not CheckAutoLetter(SerCarDesrc_.TypeTC)) OR (SerCarDesrc_.TypeTC = '')) then begin
                       //errCollectStr := errCollectStr + ExceptSN + ' Не правильный Базовый тип Авто (возможно русские буквы) ''' + SerCarDesrc.TypeTC + ''' ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                       AddError(SERIA, NUMBER, _ExceptSN, eeBadBaseAT, _ErrCollectStr);
                       ExpState := false;
                    end;
                    SerCarDesrc_.Characteristic := ExpSQL.FieldByName('CHRCT').AsFloat;

                    Table5.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeSerTC(ExpSQL.FieldByName('BASECARCODE').AsInteger, SerCarDesrc_, NeedUpdate)); //BASECARCODE
                    Table5.FieldByName('VEH_MARK').AsString := ExpStr(ExpSQL.FieldByName('MARKA').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    if IsBad then
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbMarka, _ErrCollectStr);
                       //errCollectStr := errCollectStr + ExceptSN + ' Обнаружены недопустимые символы в ' + ExpSQL.FieldByName('MARKA').AsString + #13#10;
                    Table5.FieldByName('VEH_B_TYPE').AsString := ExpStr(ExpSQL.FieldByName('LTR').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    if IsRussian then begin
                        //errCollectStr := errCollectStr + ExceptSN + ' не правильное (русск. буквы) ТИП ТС ' + ExpSQL.FieldByName('LTR').AsString + #13#10;
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadTC, _ErrCollectStr);
                        ExpState := false;
                    end;
                    if ExpSQL.FieldByName('LTR').AsString = '' then begin
                       if ExpSQL.FieldByName('RS1B').AsFloat >= 0.01 then begin
                           //errCollectStr := errCollectStr + ExceptSN + ' Не правильный тип Авто ''' + SerCarDesrc.TypeTC + ''' ' + #13#10;
                           AddError(SERIA, NUMBER, _ExceptSN, eeBadTC, _ErrCollectStr);
                           ExpState := false;
                        end;
                    end;
                    Table5.FieldByName('CHAR_VALUE').AsFloat := ExpSQL.FieldByName('CHRCT').AsFloat;
                    Table5.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
                    if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDAV', 'BASECARCODE', 'N', DecodeMsgId(Table5.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString, true);
                    Table5.Post;

                   //Автомобиль потерпевшего
                    Table4.Append;
                    Table4.FieldByName('COMPANY').AsString := COMPANY_CODE;
                    Table4.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);

                    CarDescr_.AutoNumber := ExpSQL.FieldByName('AUTONMB').AsString;
                    CarDescr_.BodyNumber := ExpSQL.FieldByName('BODYSHS').AsString;
                    CarDescr_.ShassiNumber := ExpSQL.FieldByName('BODYSHS').AsString;
                    CarDescr_.DIVISION := DIV_CODE;
                    CarDescr_.BaseCarCode := StrToInt(DecodeMsgId(Table5.FieldByName('MSG_ID').AsString));

                    Table4.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeTC(ExpSQL.FieldByName('CARCODE').AsInteger, CarDescr_, NeedUpdate)); //CARCODE_P
                    Table4.FieldByName('NUM_SIGN').AsString := ExpSQL.FieldByName('AUTONMB').AsString;
                    if (Table4.FieldByName('NUM_SIGN').AsString = '') OR (Table4.FieldByName('NUM_SIGN').AsString = '-') then
                        Table4.FieldByName('NUM_SIGN').AsString := 'Б/Н';
                    Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('BODYSHS').AsString;
                    ExpStr(Table4.FieldByName('BODY_NO').AsString, IsRussian, IsEnglish, IsDigit, IsBad, false);
                    if IsRussian AND (IsMarkaValue = 1) then begin
                        //errCollectStr := errCollectStr + ExceptSN + ' авария не правильное (русск. буквы) номер кузова/шасси' + #13#10;
                        AddError(SERIA, NUMBER, _ExceptSN, eeRussBodyShassi, _ErrCollectStr);
                        ExpState := false;
                    end;
                    Table4.FieldByName('BASECAR_ID').AsString := Table5.FieldByName('MSG_ID').AsString;
                    Table4.FieldByName('COUNTRY').Clear;
                    Table4.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDDT').AsDateTime;
                    if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDAV', 'CARCODE', 'N', DecodeMsgId(Table4.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString, true);
                    Table4.Post;
                end; //экспорт авто потерпевшего

                Table6.Append;
                Table6.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table6.FieldByName('DIVISION').AsString := IntToStr(DIV_CODE);
                Table6.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeAv(ExpSQL.FieldByName('AVCODE').AsInteger, DIV_CODE, NeedUpdate)); //PAYMENTCODE
                Table6.FieldByName('DATE_INC').AsDateTime := ExpSQL.FieldByName('AVDT').AsDateTime;
                Table6.FieldByName('PLACE_CODE').AsString := ExpSQL.FieldByName('PLACE_CODE').AsString;
                Table6.FieldByName('PLACE_INC').AsString := ExpSQL.FieldByName('PLACE').AsString;
                Table6.FieldByName('INS_DOC_NO').AsString := ExpSQL.FieldByName('WRKNMB').AsString;
                Table6.FieldByName('DATE_STATE').AsDateTime := ExpSQL.FieldByName('REGDT').AsDateTime;
                Table6.FieldByName('DECISION').AsString := GetDecision(ExpSQL.FieldByName('DECISION').AsInteger);
                if ExpSQL.FieldByName('REGDT').AsDateTime < ExpSQL.FieldByName('AVDT').AsDateTime then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadAvReg, _ErrCollectStr);
                   ExpState := false;
                end;
                if Table6.FieldByName('DECISION').AsString = '' then begin
                   //errCollectStr := errCollectStr + ExceptSN + 'Решение по результатам рассмотрения не правильно ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadDecide, _ErrCollectStr);
                   ExpState := false;
                end;
                Table6.FieldByName('POL_NUM_P').AsString := GetPaySN(ExpSQL.FieldByName('DOCSN').AsString, false);
                if Table6.FieldByName('POL_NUM_P').AsString = '?' then begin
                   //errCollectStr := errCollectStr + ExceptSN + ' Номер документа потерпевшего не правильный ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                   AddError(SERIA, NUMBER, _ExceptSN + ' (' + ExpSQL.FieldByName('DOCSN').AsString + ')', eeBadDoc, _ErrCollectStr);
                   ExpState := false;
                end;

                ExpStr(Table6.FieldByName('POL_NUM_P').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                if IsEnglish then begin
                   AddError(SERIA, NUMBER, _ExceptSN, eeBadSymbSNP, _ErrCollectStr);
                end;

                //FK Section
                Table6.FieldByName('ID_P').AsString := Table3.FieldByName('MSG_ID').AsString;
                if ExpSQL.FieldByName('MARKA').AsString <> '' then
                    Table6.FieldByName('CAR_ID_P').AsString := Table4.FieldByName('MSG_ID').AsString
                else
                    Table6.FieldByName('CAR_ID_P').Clear;

                Table6.FieldByName('POL_NUM_V').AsString := SERIA + ExpZeroNmb(NUMBER);
                Table6.FieldByName('ID_V').AsString := '';
                for find := 0 to OwnCounter - 1 do
                    if ExpSQL.FieldByName('BADFIO').AsString = AOwn[find].FIO then begin
                        Table6.FieldByName('ID_V').AsString := AOwn[find].Code;
                        Table6.FieldByName('CAR_ID_V').AsString := CURR_CAR_CODE;
                        break;
                    end;

                if Table6.FieldByName('ID_V').AsString = '' then begin
                    //errCollectStr := errCollectStr + ExceptSN + ' В авариях записан виновник ' + ExpSQL.FieldByName('BADFIO').AsString + ' Среди владельцев такого нет!' + #13#10;
                    AddError(SERIA, NUMBER, _ExceptSN + ' ' + ExpSQL.FieldByName('BADFIO').AsString, eeBadBadMan, _ErrCollectStr);
                    ExpState := false;
                end;

                Table6.FieldByName('ACT_HARM1').Value := ExpSQL.FieldByName('FS1B').Value;
                Table6.FieldByName('PAYMENT1').Value := ExpSQL.FieldByName('RS1B').Value;
                Table6.FieldByName('DATE_PAY1').Value := ExpSQL.FieldByName('DT1').Value;
                Table6.FieldByName('ACT_HARM2').Value := ExpSQL.FieldByName('FS2B').Value;
                Table6.FieldByName('PAYMENT2').Value := ExpSQL.FieldByName('RS2B').Value;
                Table6.FieldByName('DATE_PAY2').Value := ExpSQL.FieldByName('DT2').Value;
                Table6.FieldByName('ACT_HARM3').Value := ExpSQL.FieldByName('FS3B').Value;
                Table6.FieldByName('PAYMENT3').Value := ExpSQL.FieldByName('RS3B').Value;
                Table6.FieldByName('DATE_PAY3').Value := ExpSQL.FieldByName('DT3').Value;

                if(not ExpSQL.FieldByName('DT1').IsNull) then
                    if ExpSQL.FieldByName('DT1').AsDateTime < ExpSQL.FieldByName('REGDT').AsDateTime then begin
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadPayAvDate, _ErrCollectStr);
                        ExpState := false;
                    end;
                if(not ExpSQL.FieldByName('DT2').IsNull) then
                    if ExpSQL.FieldByName('DT2').AsDateTime < ExpSQL.FieldByName('REGDT').AsDateTime then begin
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadPayAvDate, _ErrCollectStr);
                        ExpState := false;
                    end;
                if(not ExpSQL.FieldByName('DT3').IsNull) then
                    if ExpSQL.FieldByName('DT3').AsDateTime < ExpSQL.FieldByName('REGDT').AsDateTime then begin
                        AddError(SERIA, NUMBER, _ExceptSN, eeBadPayAvDate, _ErrCollectStr);
                        ExpState := false;
                    end;


                if (Table6.FieldByName('DECISION').AsString = 'Y') then begin
                    if( Table6.FieldByName('DATE_PAY1').IsNull AND
                        Table6.FieldByName('DATE_PAY2').IsNull AND
                        Table6.FieldByName('DATE_PAY3').IsNull ) then begin
                            //errCollectStr := errCollectStr + ExceptSN + ' Не определена дата выплаты!'#13#10;
                            AddError(SERIA, NUMBER, _ExceptSN, eeNoPay, _ErrCollectStr);
                             ExpState := false;
                    end;
                             
                    if ExpSQL.FieldByName('CURR').AsString <> '' then
                        Table6.FieldByName('CURRENCY').AsString := ExpCurr(ExpSQL.FieldByName('CURR').AsString)
                    else begin
                         //errCollectStr := errCollectStr + ExceptSN + ' Не определена валюта выплаты!'#13#10;
                         AddError(SERIA, NUMBER, _ExceptSN, eeBadCurr, _ErrCollectStr);
                         ExpState := false;
                    end;

                   if (Table6.FieldByName('PAYMENT1').AsFloat < 0.01) AND
                      (Table6.FieldByName('PAYMENT2').AsFloat < 0.01) AND
                      (Table6.FieldByName('PAYMENT3').AsFloat < 0.01) AND
                      (Table6.FieldByName('ACT_HARM1').AsFloat < 0.01) AND
                      (Table6.FieldByName('ACT_HARM2').AsFloat < 0.01) AND
                      (Table6.FieldByName('ACT_HARM3').AsFloat < 0.01) then begin
                         //errCollectStr := errCollectStr + ExceptSN + ' Не определена ни одна сумма выплаты!' + #13#10;
                         AddError(SERIA, NUMBER, _ExceptSN, eeNoSumma, _ErrCollectStr);
                         ExpState := false;
                      end;
                end
                else begin
                   if (Table6.FieldByName('PAYMENT1').AsFloat >= 0.01) OR
                      (Table6.FieldByName('PAYMENT2').AsFloat >= 0.01) OR
                      (Table6.FieldByName('PAYMENT3').AsFloat >= 0.01) OR
                      (Table6.FieldByName('ACT_HARM1').AsFloat >= 0.01) OR
                      (Table6.FieldByName('ACT_HARM2').AsFloat >= 0.01) OR
                      (Table6.FieldByName('ACT_HARM3').AsFloat >= 0.01) then begin
                         //errCollectStr := errCollectStr + ExceptSN + ' Cумма вырплаты есть а выплаты нету!' + #13#10;
                         AddError(SERIA, NUMBER, _ExceptSN, eeSumNoSum, _ErrCollectStr);
                         ExpState := false;
                      end;
                end;

                Table6.FieldByName('UPDATE').Value := ExpSQL.FieldByName('UPDDT').Value;
                if NeedUpdate then SetMaxCodeFor(Seria, Number, 'SYSADM.MANDAV', 'AVCODE', 'N', DecodeMsgId(Table6.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString, true);
                Table6.Post;

                ExpSQL.Next
            end;

         ExportBUROSQL.Next;

         end;
    except
      on E : Exception do begin
         WorkSQL.Close;
         ExpSQL.Close;
         if ExpBUROForm.IsBackgroundMode.Checked then
             WindowState := wsNormal;
         MessageDlg('Ошибка создания отчёта'#13#10 + SERIA + '/' + IntToStr(NUMBER) + _ExceptSN + #13#10 + E.Message, mtInformation, [mbOK], 0);
         ExpState := false;
      end
      else begin
         if ExpBUROForm.IsBackgroundMode.Checked then
             WindowState := wsNormal;
         MessageDlg('Критическая ошибка при формировании отчёта'#13#10 + SERIA + '/' + IntToStr(NUMBER) + _ExceptSN, mtInformation, [mbOk], 0);
       end;
    end;

    if ExpBUROForm.IsBackgroundMode.Checked then
        WindowState := wsNormal;
    ExportBUROSQL.Close;
    ExportOwners.Close;
    ExportAvariasBURO.Close;
    HandBookSQL.Close;
    WorkSQL.Close;

    if not( (_errCollectStr <> '') OR (ExpState = false) OR (ExpBUROForm.IsTestExport.Checked) ) then begin
      Table1.Close;
      Table2.Close;
      Table3.Close;
      Table4.Close;
      Table5.Close;
      Table6.Close;

      Table1.Open;
      Table2.Open;
      Table3.Open;
      Table4.Open;
      Table5.Open;
      Table6.Open;

      StatusBar.Panels[0].Text := 'Удаление дубликатов...';
      StatusBar.Update;

      //Drop dublicates
      if (not DropDublicate(Table2, DBFileName2b, '', '', '')) OR
         (not DropDublicate(Table3, DBFileName3b, 'DRV_EXPR', 'F_L', '')) OR
         (not DropDublicate(Table4, DBFileName4b, '', '', '')) OR
         (not DropDublicate(Table5, DBFileName5b, 'CHAR_VALUE', '', '')) then begin
         ExpState := false;
      end;
      //END Drop dublicates

      PackTable(Table1);
      PackTable(Table6);

      Table2.Close;
      Table3.Close;
      Table4.Close;
      Table5.Close;

      Table2.DeleteTable;
      Table3.DeleteTable;
      Table4.DeleteTable;
      Table5.DeleteTable;
    end;

    Table1.Close;
    Table6.Close;

    Table1.Free;
    Table2.Free;
    Table3.Free;
    Table4.Free;
    Table5.Free;
    Table6.Free;

    StatusBar.Panels[0].Text := '';

    if (_errCollectStr <> '') OR (ExpState = false) OR (ExpBUROForm.IsTestExport.Checked) then begin
       DeleteFile(DBFileName1);
       DeleteFile(DBFileName2);
       DeleteFile(DBFileName3);
       DeleteFile(DBFileName4);
       DeleteFile(DBFileName5);
       DeleteFile(DBFileName6);
       DeleteFile(DBFileName2b);
       DeleteFile(DBFileName3b);
       DeleteFile(DBFileName4b);
       DeleteFile(DBFileName5b);
    end;

    if _errCollectStr <> '' then begin
        ShowInfo.Caption := 'Ошибки в данных';
        ShowInfo.InfoPanel.Text := _errCollectStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
    else
    if ExpState then begin
        Application.MainForm.BringToFront;
        if ExpBUROForm.IsTestExport.Checked then
            MessageDlg('Тестовый прогон прошёл успешно! Можете запускать экспорт данных', mtInformation, [mbOK], 0)
        else begin
            MessageDlg('Данные подготовлены для передачи в БЮРО'#13'Они расположены в каталоге ' + ExpBUROForm.OutDirectory.Text,  mtInformation, [mbOK], 0);
            ProgressBar.Position := 0;
            //ExpErrsFlag.DirectoryEdit.Text := ExpBUROForm.OutDirectory.Text;
            //ExpErrsFlag.ShowModal;
        end
    end;

    ProgressBar.Position := 0;
    ListNoCkeckAN.Free;
end;

procedure TMandSQL.FormResize(Sender: TObject);
begin
     StatusBar.Panels[1].Width := 250;
     StatusBar.Panels[0].Width := Width - 200 - StatusBar.Panels[1].Width;
     ProgressBar.Top := 3;
     ProgressBar.Left := StatusBar.Panels.Items[0].Width + StatusBar.Panels[1].Width + 3;
     ProgressBar.Width := 200 - 12;
     Animate.Width := TopPanel.Width - Animate.Left - 10;
end;
{
function TMandSQL.GetDivisionCode2(AgCode : string) : string;
var
    i : integer;
    isfound : boolean;
begin
    if lastAgCode = AgCode then begin
        GetDivisionCode2 := IntToStr(lastDivCode);
        exit;
    end;
    GetDivisionCode2 := '?';
    isfound := false;

    for i := 0 to ListAgCodes.Count - 1 do
        if ListAgCodes[i] = AgCode then begin
           if Integer(ListAgDiv[i]) >= 0 then begin
               GetDivisionCode2 := IntToStr(Integer(ListAgDiv[i]));
               lastAgCode := AgCode;
               lastDivCode := Integer(ListAgDiv[i]);
               isfound := true;
               break;
           end;
        end;

    if not isfound then
        MessageDlg('Не определён код подразделения для агента ''' + AgCode + '''', mtInformation, [mbOK], 0);
end;
}

function TMandSQL.GetVladenType(N : integer) : string;
begin
     case N of
        1 : GetVladenType := 'S';
        2 : GetVladenType := 'H';
        3 : GetVladenType := 'O';
        4 : GetVladenType := 'A';
        5 : GetVladenType := 'D';
        6 : GetVladenType := 'I';
     else
         GetVladenType := '?';
     end;
end;

function TMandSQL.GetFizUrid(s : string) : string;
begin
     if s = 'F' then
        GetFizUrid := 'F'
     else
     if s = 'U' then
        GetFizUrid := 'L'
     else
        GetFizUrid := '?';
end;

function TMandSQL.GetDiscount2(N : double) : string;
var
      rounded : integer;
begin
      rounded := ROUND(N * 100);
      if rounded = 130 then
         GetDiscount2 := '3'
      else
      if rounded = 120 then
         GetDiscount2 := '2'
      else
      if rounded = 100 then
         GetDiscount2 := '0'
      else
      if rounded = 80 then
         GetDiscount2 := '8'
      else begin
          GetDiscount2 := '?';
      end;
end;

function TMandSQL.GetCharacteristicAuto(s : string; Volume, Tonnage,Quantity, Power : double; IsResident : boolean) : double;
begin
    if not IsResident then begin
       GetCharacteristicAuto := 0;
       exit;
    end;

    if (s = 'N1') OR (s = 'N2') OR (s = 'N3') OR (s = 'N4') OR (s = 'N5') OR
       (s = 'A1') OR (s = 'A2') OR (s = 'A3') OR (s = 'A4') OR (s = 'A5') OR
       (s = 'F1') OR (s = 'F2') OR (s = 'F3') then
       GetCharacteristicAuto := Volume
    else
    if (s = 'C0') OR (s = 'C1') OR (s = 'C2') OR (s = 'C3') OR (s = 'C4') OR
       (s = 'C5') OR (s = 'E1') OR (s = 'E2') OR (s = 'E3') then
       GetCharacteristicAuto := Tonnage
    else
    if (s = 'V1') OR (s = 'V2') OR (s = 'V3') then
       GetCharacteristicAuto := Power
    else
    if (s = 'L1') OR (s = 'L2') OR (s = 'L3') then
       GetCharacteristicAuto := Quantity
    else
    if (s = 'A6') OR (s = 'B1') OR (s = 'B2') OR (s = 'D') OR (s = 'M') OR
       (s = 'L4') OR (s = 'W') then
       GetCharacteristicAuto := 0
    else   
    //if (s = 'X') OR (s = 'Х') then
    //   GetCharacteristicAuto := 0
    //else
       GetCharacteristicAuto := -1;
end;

function TMandSQL.ExpTime(s : string) : integer;
var
     divPos : integer;
     Hour, Min : integer;
begin
     try
       ExpTime := 0;
       if s = '' then
          ExpTime := 0
       else begin
          if Pos('.', s) <> 0 then
              s[Pos('.', s)] := ':';
              
          divPos := Pos(':', s);
          if divPos = 0 then begin
               ExpTime := StrToInt(s);
          end
          else begin
               Hour := StrToInt(Copy(s, 1, divPos - 1));
               Min := StrToInt(Copy(s, divPos + 1, 100));
               ExpTime := Hour * 60 + Min;
          end;
       end;
     except
       MessageDlg('Ошибка преобразования времени ' + s, mtInformation, [mbOK], 0);
     end;
end;

function TMandSQL.ExpStateFunc(state : integer) : string;
begin
      if state in [0, 4] then
         ExpStateFunc := 'W'
      else
      if state = 3 then
         ExpStateFunc := 'D'
      else
      if state = 5 then
         ExpStateFunc := 'P'
      else
      if state = 2 then
         ExpStateFunc := 'O'
      else
         ExpStateFunc := '';
end;

function TMandSQL.ExpCurr(Curr : string) : string;
begin
     if Curr = 'BRB' then
         ExpCurr := 'B'
     else
     if Curr = 'DM' then
         ExpCurr := 'M'
     else
     if Curr = 'USD' then
         ExpCurr := 'S'
     else
     if Curr = 'RUR' then
         ExpCurr := 'R'
     else
     if Curr = 'EUR' then
         ExpCurr := 'E'
     else
         ExpCurr := '';
end;

function TMandSQL.ExpZeroNmb(N : Integer) : string;
var
      s : string;
begin
      s := IntToStr(N);
      while Length(s) < 7 do
          s := '0' + s;
      ExpZeroNmb := s;
end;

function TMandSQL.GetResident(N : boolean) : string;
begin
     if N then
        GetResident := 'R'
     else
        GetResident := 'N'
end;

function TMandSQL.GetDecision(N : integer) : string;
begin
     //см blank.ini
     if N = 0 then
        GetDecision := 'Y'
     else
     if N = 1 then
        GetDecision := 'N'
     else
        GetDecision := '';
end;

function TMandSQL.GetPaySN(s : string; mand : boolean) : string;
var
     divIdx : integer;
     resStr : string;
     i : integer;
     Rus : string;
begin
    Rus := 'йцукенгшщзхъфывапролджэячсмитьбюё';

    for i := 1 to Length(Rus) do begin
        if Pos(Rus[i], s) <> 0 then begin
            Result := '?';
            exit;
        end;
    end;

     s := Trim(s);
     if s = '' then
         resStr := ''
     else begin
         divIdx := Pos('/', s);
         if divIdx < 2 then begin
            resStr := '?'
         end
         else begin
            if (Length(s) = 3) and (not mand) then
                resStr := ''
            else begin
                resStr := Copy(s, 1, divIdx - 1);
                try
                    if StrToInt(Copy(s, divIdx + 1, 64)) <= 0 then
                        resStr := '?'
                    else begin
                        if StrToInt(Copy(s, divIdx + 1, 64)) = 0 then
                            resStr := '?'
                        else
                            resStr := resStr + ExpZeroNmb(StrToInt(Copy(s, divIdx + 1, 64)));
                    end;
                except
                    resStr := '?';
                end
            end;
         end;
     end;
     GetPaySN := resStr;
end;

procedure TMandSQL.KorrectMenuItemClick(Sender: TObject);
begin
     if KorrectDataSQL = nil then begin
        Application.CreateForm(TKorrectDataSQL, KorrectDataSQL);
     end;
     KorrectDataSQL.BringToFront
end;

function FTS(V : double) : string;
var
    s : string;
begin
    s := FloatToStr(V);
    FTS := '''' + s + '''';
end;

function TMandSQL.GetNextCodeOwner(Code : integer; S : CarOwnerDescription_; var NeedUpdate : boolean) : string;
var
     Flag : integer;
begin
     if ExpBUROForm.IsTestExport.Checked then begin
        NeedUpdate := false;
        GetNextCodeOwner := '1';
        exit;
     end;

     //function UpdateTabField
     NeedUpdate := true;
     if Code > 0 then begin
        NeedUpdate := false;
        GetNextCodeOwner := IntToStr(Code);
        exit;
     end;

     try
         with HandBookSQL, HandBookSQL.SQL do begin
              Flag := 0;
              Clear;
              Add('SELECT OWNERCODE');
              Add('FROM SYSADM.MANDATOR');
              Add('WHERE OWNERCODE>0');
              Add('AND NAME=''' + S.Name + '''');
              Add('AND ADDR=''' + S.Address + '''');
              Add('AND DIVISION=' + IntToStr(S.DIVISION));
              if S.Addr_code <> '' then
                  Add('AND PLACE_CODE=''' + S.Addr_code + '''')
              else
                  Add('AND (PLACE_CODE=''' + S.Addr_code + ''' OR PLACE_CODE IS NULL)');
              Add('AND VLADEN=' + IntToStr(S.Owner_Type));
              Add('AND K BETWEEN (:K - 0.00001) AND (:K + 0.00001)');
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if Fields[0].AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(FieldByName(Fields[0].FieldName).AsInteger);
                       HandBookSQL.Close;
                       exit;
                    end;
                    Next;
              end;

              //Поиск в других владельцах
              Flag := 1;
              Close;
              Clear;
              Add('SELECT O.OWNRCODE');
              Add('FROM SYSADM.MANDOWN O,SYSADM.MANDATOR M');
              Add('WHERE O.OWNRCODE >0 AND M.SER=O.SER AND M.NMB=O.NMB');
              Add('AND O.NAME=''' + S.Name + '''');
              Add('AND O.ADDR=''' + S.Address + '''');
              Add('AND M.DIVISION=' + IntToStr(S.DIVISION));
//              Add('AND PLACE_CODE=''' + S.Addr_code + '''');
              if S.Addr_code <> '' then
                  Add('AND O.PLACE_CODE=''' + S.Addr_code + '''')
              else
                  Add('AND (O.PLACE_CODE=''' + S.Addr_code + ''' OR O.PLACE_CODE IS NULL)');
              Add('AND O.PRAVO=' + IntToStr(S.Owner_Type));
              Add('AND O.K BETWEEN (:K - 0.00001) AND (:K + 0.00001)');
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if Fields[0].AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(Fields[0].AsInteger);
                       HandBookSQL.Close;
                       exit;
                    end;
                    Next;
              end;

              //Поиск в выплатах
              Flag := 2;
              Close;
              Clear;
              Add('SELECT A.OWNRCODE');
              Add('FROM SYSADM.MANDAV A, SYSADM.MANDATOR M');
              Add('WHERE OWNRCODE >0 AND M.SER=A.SER AND M.NMB=A.NMB');
              Add('AND A.FIO=''' + S.Name+ '''');
              Add('AND A.ADDR=''' + S.Address + '''');
              Add('AND M.DIVISION=' + IntToStr(S.DIVISION));
//              Add('AND ADRCODE=''' + S.Addr_code + '''');
              if S.Addr_code <> '' then
                  Add('AND A.ADRCODE=''' + S.Addr_code + '''')
              else
                  Add('AND (A.ADRCODE=''' + S.Addr_code + ''' OR A.ADRCODE IS NULL)');

              Add('AND A.VLADEN=' + IntToStr(S.Owner_Type));
              Add('AND A.K BETWEEN (:K - 0.00001) AND (:K + 0.00001)');
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if Fields[0].AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(Fields[0].AsInteger);
                       HandBookSQL.Close;
                       exit;
                    end;
                    Next;
              end;

              Close;
         end; // with
//     end;
     except
        on E : Exception do begin
           MessageDlg('GetNextCodeOwner N' + IntToStr(Flag) + #13 + e.Message, mtInformation, [mbOk], 0);
           raise;
        end;
     end;

     GetNextCodeOwner := IntToStr(CodeControl_.GetNextOwnerCode(S.DIVISION));
{     if OwnerCode_Holes.Count > 0 then begin
         GetNextCodeOwner := IntToStr(Integer(OwnerCode_Holes[0]));
         OwnerCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxOwnerCode);
     GetNextCodeOwner := IntToStr(MaxOwnerCode);}
end;


function TMandSQL.GetMaxCodeFor(Table, Field, Division : string; IsLinkMand : boolean) : integer;
begin
     try
         with HandBookSQL, HandBookSQL.SQL do begin
              Clear;
              Add('SELECT MAX(T.' + Field + ')');
              Add('FROM ' + Table + ' T');
              if IsLinkMand then
                Add(',SYSADM.MANDATOR M');
              Add('WHERE T.' + Field + ' IS NOT NULL AND Division=' + Division);
              if IsLinkMand then
                Add(' AND T.SER=M.SER AND T.NMB=M.NMB');
              Open;
              GetMaxCodeFor := FieldByName(Fields[0].FieldName).AsInteger;
              Close;
         end;
      except
         on E : Exception do begin
            MessageDlg('GetMaxCodeFor'#13 + e.Message, mtInformation, [mbOk], 0);
            raise;
         end;
      end
end;

procedure TMandSQL.SetMaxCodeFor(Seria : string; Number : integer; Table, Field1, Field2, Val1, Val2 : string; IsVal2Int : boolean);
begin
     try
         SQLDB.StartTransaction;
         with HandBookSQL, HandBookSQL.SQL do begin
              HandBookSQL.Close;
              HandBookSQL.SQL.Clear;
              Add('UPDATE ' + Table);
              Add('SET '  + Field1  + '=:VAL');
              Add('WHERE NMB=:N AND SER=:S');
              if Length(Field2) > 0 then
                 Add('AND ' + Field2 + '=:VAL2');

              ParamByName('S').AsString := Seria;
              ParamByName('N').AsInteger := Number;
              ParamByName('VAL').AsInteger := StrToInt(Val1);
              if Length(Field2) > 0 then begin
                  //try //Проверка строка или число
                     //StrToInt(Val2);
                  if IsVal2Int then
                     ParamByName('VAL2').AsInteger := StrToInt(Val2)
                  //except
                  else
                     ParamByName('VAL2').AsString := Val2;
                  //end;
              end;
              ExecSQL;
              SQLDB.Commit;
              HandBookSQL.Close;
         end;
     except
         on E : Exception do begin
             SQLDB.Rollback;
             MessageDlg('Ошибка обновления записи ' + Seria + IntToStr(Number) + '[' + Table + ',' + Field1 + ', ' + Field2 + ']', mtInformation, [mbOk], 0);
             raise;
         end;
     end;
end;

function TMandSQL.GetNextCodeSerTC(Code : integer; S : SerCarDescription_; var NeedUpdate : boolean) : string;
var
    indicator : string;
begin
     if ExpBUROForm.IsTestExport.Checked then begin
        NeedUpdate := false;
        GetNextCodeSerTC := '1';
        exit;
     end;

     NeedUpdate := true;
     if Code > 0 then begin
        NeedUpdate := false;
        GetNextCodeSerTC := IntToStr(Code);
        exit;
     end;
  try
     with HandBookSQL, HandBookSQL.SQL do begin
          Close;
          Clear;
          Add('SELECT BASECARCODE');
          Add('FROM SYSADM.MANDATOR');
          Add('WHERE BASECARCODE >0');
          Add('AND MARKA=''' + S.Marka + '''');
          Add('AND DIVISION=' + IntToStr(S.DIVISION));
          Add('AND LTR=''' + S.TypeTC + '''');
          Add('AND ((Tonnage = :CH) OR (CNTPAS = :CH) OR (Power = :CH) OR (Volume = :CH))');
          ParamByName('CH').AsFloat := S.Characteristic;
          indicator := '1';
          Open;

          while not EOF do begin
                if Fields[0].AsInteger > 0 then begin
                   GetNextCodeSerTC := IntToStr(FieldByName('BASECARCODE').AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;

          Close;
          Clear;
          Add('SELECT A.BASECARCODE');
          Add('FROM SYSADM.MANDAV A, SYSADM.MANDATOR M');
          Add('WHERE A.BASECARCODE >0 AND A.SER=M.SER AND A.NMB=M.NMB');
          Add('AND DIVISION=' + IntToStr(S.DIVISION));
          Add('AND A.MARKA=''' + S.Marka + '''');
          Add('AND A.LTR=''' + S.TypeTC + '''');
          Add('AND A.CHRCT=:CH');
          ParamByName('CH').AsFloat := S.Characteristic;
          indicator := '2';
          Open;

          while not EOF do begin
                if Fields[0].AsInteger > 0 then begin
                   GetNextCodeSerTC := IntToStr(Fields[0].AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;
     end; // with

     indicator := '3';
     GetNextCodeSerTC := IntToStr(CodeControl_.GetNextBaseCarCode(S.DIVISION));
  except
     on E : Exception do begin
        MessageDlg('GetNextCodeSerTC (' + indicator + ') DIVISION=' + IntToStr(S.DIVISION) + #13 + e.Message, mtInformation, [mbOk], 0);
        raise;
     end;
  end
end;

function TMandSQL.GetNextCodeTC(Code : integer; S : CarDescription_; var NeedUpdate : boolean) : string;
var
    Flag : integer;
begin
     if ExpBUROForm.IsTestExport.Checked then begin
        NeedUpdate := false;
        GetNextCodeTC := '1';
        exit;
     end;

     NeedUpdate := true;
     if Code > 0 then begin
        NeedUpdate := false;
        GetNextCodeTC := IntToStr(Code);
        exit;
     end;

  try
     //check with SerTCCode
     with HandBookSQL, HandBookSQL.SQL do begin
          Flag := 0;
          Close;
          Clear;
          Add('SELECT CARCODE');
          Add('FROM SYSADM.MANDATOR');
          Add('WHERE CARCODE >0');
          Add('AND DIVISION=' + IntToStr(S.DIVISION));
          Add('AND AUTONMB=''' + S.AutoNumber + '''');
          if S.BodyNumber <> '' then
              Add('AND Nmbody=''' + S.BodyNumber + '''')
          else
          if S.ShassiNumber <> '' then
              Add('AND CHASSIS=''' + S.ShassiNumber + '''');

          Add('AND BASECARCODE=:CODE');
          ParamByName('CODE').AsInteger := S.BaseCarCode;
          Open;

          while not EOF do begin
                if Fields[0].AsInteger > 0 then begin
                   GetNextCodeTC := IntToStr(Fields[0].AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;

          Flag := 1;
          Close;
          Clear;
          Add('SELECT A.CARCODE');
          Add('FROM SYSADM.MANDAV A, SYSADM.MANDATOR M');
          Add('WHERE A.CARCODE >0 AND A.SER=M.SER AND A.NMB=M.NMB');
          Add('AND DIVISION=' + IntToStr(S.DIVISION));
          Add('AND A.AUTONMB=''' + S.AutoNumber + '''');
          Add('AND A.BODYSHS=''' + S.BodyNumber + '''');
          Add('AND A.BASECARCODE=:CODE');
          ParamByName('CODE').AsInteger := S.BaseCarCode;
          Open;

          while not EOF do begin
                if Fields[0].AsInteger > 0 then begin
                   GetNextCodeTC := IntToStr(Fields[0].AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;
     end; // with
{
     if TCCode_Holes.Count > 0 then begin
         GetNextCodeTC := IntToStr(Integer(TCCode_Holes[0]));
         TCCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxTCCode);
     GetNextCodeTC := IntToStr(MaxTCCode);
}
     GetNextCodeTC := IntToStr(CodeControl_.GetNextCarCode(S.DIVISION));
  except
     on E : Exception do begin
        MessageDlg('GetNextCodeTC N' + IntToStr(Flag) + #13 + e.Message, mtInformation, [mbOk], 0);
        raise;
     end;
  end
end;

function TMandSQL.GetNextCodeAv(Code : integer; D : integer; var NeedUpdate : boolean) : string;
begin
     if ExpBUROForm.IsTestExport.Checked then begin
        NeedUpdate := false;
        GetNextCodeAv := '1';
        exit;
     end;

     NeedUpdate := true;
     if Code > 0 then begin
        NeedUpdate := false;
        GetNextCodeAv := IntToStr(Code);
        exit;
     end;
{
     if AVCode_Holes.Count > 0 then begin
         GetNextCodeAv := IntToStr(Integer(AVCode_Holes[0]));
         AVCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxAVCode);
     GetNextCodeAv := IntToStr(MaxAVCode);  }
     GetNextCodeAv := IntToStr(CodeControl_.GetNextAvCode(d));
end;

procedure TMandSQL.InitExportBuro;
var
    CurrCode : integer;
begin
//    try
//      Screen.Cursor := crHourGlass;

 //     OwnerCode_Holes.Clear;
//      SerTCCode_Holes.Clear;
//      TCCode_Holes.Clear;
//      AVCode_Holes.Clear;
      //ListAgCodes.Clear;
      //ListAgDiv.Clear;

//      StatusBar.SimpleText := 'Поиск кодов подразделений...';
//      StatusBar.Update;

//      ShowInfo.InfoPanel.Text := '';

//      TableAg.Open;
//      TableAg.First;
//      while not TableAg.EOF do begin
//          if not TableAgDivisionCode.IsNull then begin
//              ListAgCodes.Add(TableAgAgent_code.AsString);
//              ListAgDiv.Add(TObject(TableAgDivisionCode.AsInteger));
//          end;
//          TableAg.Next;
//      end;

      //OWNER CODE
{      StatusBar.SimpleText := 'Поиск кода владельца...';
      StatusBar.Update;

      MaxOwnerCode := GetMaxCodeFor('SYSADM.MANDATOR', 'OWNERCODE');
      CurrCode := GetMaxCodeFor('SYSADM.MANDOWN', 'OWNRCODE');
      if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;
      CurrCode := GetMaxCodeFor('SYSADM.MANDAV', 'OWNRCODE');
      if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;

      // Ser TC
      StatusBar.SimpleText := 'Поиск кода характеристики авто...';
      StatusBar.Update;
      MaxSerTCCode := GetMaxCodeFor('SYSADM.MANDATOR', 'BaseCarCode');
      CurrCode := GetMaxCodeFor('SYSADM.MANDAV', 'BaseCarCode');
      if CurrCode > MaxSerTCCode then MaxSerTCCode := CurrCode;

      // TC
      StatusBar.SimpleText := 'Поиск кода авто...';
      StatusBar.Update;
      MaxTCCode := GetMaxCodeFor('SYSADM.MANDATOR', 'CarCode');
      CurrCode := GetMaxCodeFor('SYSADM.MANDAV', 'CarCode');
      if CurrCode > MaxTCCode then MaxTCCode := CurrCode;

      //AVCode
      StatusBar.SimpleText := 'Поиск кода аварии...';
      StatusBar.Update;
      MaxAVCode := GetMaxCodeFor('SYSADM.MANDAV', 'AvCode');

      //Find Holes
      with WorkSQL, WorkSQL.SQL do begin
           StatusBar.SimpleText := 'Поиск пустых кодов (1)...';
           StatusBar.Update;

           Clear;
           Add('SELECT OWNERCODE FROM SYSADM.MANDATOR WHERE OWNERCODE > 0');
           Add('UNION');
           Add('SELECT OWNRCODE FROM SYSADM.MANDOWN WHERE OWNRCODE > 0');
           Add('UNION');
           Add('SELECT OWNRCODE FROM SYSADM.MANDAV WHERE OWNRCODE > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                     OwnerCode_Holes.Add(Pointer(CurrCode));
                     Inc(CurrCode);
                 end;
                 Inc(CurrCode);
                 Next;
           end;
           Close;

           //Base Car Code
           StatusBar.SimpleText := 'Поиск пустых кодов (2)...';
           StatusBar.Update;
           Clear;
           Add('SELECT BaseCarCode FROM SYSADM.MANDATOR WHERE BaseCarCode > 0');
           Add('UNION');
           Add('SELECT BaseCarCode FROM SYSADM.MANDAV WHERE BaseCarCode > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                     SerTCCode_Holes.Add(Pointer(CurrCode));
                     Inc(CurrCode);
                 end;
                 Inc(CurrCode);
                 Next;
           end;
           Close;

           //Car Code
           StatusBar.SimpleText := 'Поиск пустых кодов (3)...';
           StatusBar.Update;
           Clear;
           Add('SELECT CarCode FROM SYSADM.MANDATOR WHERE CarCode > 0');
           Add('UNION');
           Add('SELECT CarCode FROM SYSADM.MANDAV WHERE CarCode > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                     TCCode_Holes.Add(Pointer(CurrCode));
                     Inc(CurrCode);
                 end;
                 Inc(CurrCode);
                 Next;
           end;
           Close;

           //PaymentCode
           StatusBar.SimpleText := 'Поиск пустых кодов (4)...';
           StatusBar.Update;
           Clear;
           Add('SELECT DISTINCT AvCode FROM SYSADM.MANDAV WHERE AVCode > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                     AVCode_Holes.Add(Pointer(CurrCode));
                     Inc(CurrCode);
                 end;
                 Inc(CurrCode);
                 Next;
           end;
           Close;
      end;
    except
        Screen.Cursor := crDefault;
        raise;
    end;
}
//    StatusBar.SimpleText := '';
//    StatusBar.Update;
//    Screen.Cursor := crDefault;
end;

procedure TMandSQL.InitMenuClick(Sender: TObject);
begin
{     if KorrectPoland = nil then begin
         Application.CreateForm(TKorrectPoland, KorrectPoland);
     end;

     KorrectPoland.DbName := DB.DatabaseName;
     KorrectPoland.TableName := 'MANDATOR';
     KorrectPoland.AgentCodeFld := 'AgentCode';
     KorrectPoland.AgentPercentFld := 'AgPercent';
     KorrectPoland.UridichFizichFld := 'InsComp';
     KorrectPoland.IsUridichFldStr := true;
     KorrectPoland.UR_FIZ_VLA := 'UF';
     KorrectPoland.NULLFld := 'PayDate';
     KorrectPoland.RegDateFld := 'RegDate';
     KorrectPoland.ShowModal;

     MainQuery.Refresh}
end;

procedure TMandSQL.MandSQLQueryCalcFields(DataSet: TDataSet);
begin
     MandSQLQueryVSN.AsString := MandSQLQuerySER.AsString + '/' + MandSQLQueryNMB.AsString;

     MandSQLQueryVPREV.AsString := '';
     if MandSQLQueryPSER.AsString <> '' then
         MandSQLQueryVPREV.AsString := MandSQLQueryPSER.AsString + '/' + MandSQLQueryPNMB.AsString;

     MandSQLQueryVPERIOD.AsString := '';
     if MandSQLQuerySTATE.AsInteger <> 1 then //BAD FLAG
         MandSQLQueryVPERIOD.AsString := MandSQLQueryFRDT.AsString + ' - ' + MandSQLQueryTODT.AsString + ', ' + MandSQLQueryPERIOD.AsString;

     MandSQLQueryVPAY1.AsString := '';
     if MandSQLQuerySTATE.AsInteger <> 1 then begin //BAD FLAG
         MandSQLQueryVPAY1.AsString := MandSQLQueryPAY1DT.AsString  + ', ' + MandSQLQueryPAY1.AsString  + ' ' + MandSQLQueryPAY1CURR.AsString;
         if MandSQLQueryPAY1DOC.AsInteger > 0 then
             MandSQLQueryVPAY1.AsString := MandSQLQueryVPAY1.AsString + ', плтж. N' + MandSQLQueryPAY1DOC.AsString;
     end;

     MandSQLQueryVPAY2.AsString := '';
     if MandSQLQuerySTATE.AsInteger <> 1 then begin //BAD FLAG
         if MandSQLQueryISFEE2.AsString = 'Y' then begin
             if MandSQLQueryISPAY2.AsString = 'Y' then begin
                 MandSQLQueryVPAY2.AsString := MandSQLQueryPAY2DT.AsString  + ', ' + MandSQLQueryPAY2.AsString  + ' ' + MandSQLQueryPAY2CURR.AsString;
                 if MandSQLQueryPAY2DOC.AsInteger > 0 then
                     MandSQLQueryVPAY2.AsString := MandSQLQueryVPAY2.AsString + ', плтж. N' + MandSQLQueryPAY2DOC.AsString;
             end
             else begin //не оплачен
                 if MandSQLQuerySTATE.AsInteger = 0 then begin
                     if Date < MandSQLQueryFEE2DT.AsDateTime then
                         MandSQLQueryVPAY2.AsString := 'Ожидается ' + MandSQLQueryFEE2DT.AsString
                     else
                     if Date = MandSQLQueryFEE2DT.AsDateTime then
                         MandSQLQueryVPAY2.AsString := 'Сегодня должны оплатить'
                     else
                         MandSQLQueryVPAY2.AsString := 'ПРОСРОЧЕН c ' + MandSQLQueryFEE2DT.AsString;
                 end;
             end;
         end
         else
             MandSQLQueryVPAY2.AsString := 'Не нужен';
     end;

     MandSQLQueryVTARUF.Clear;
     if MandSQLQuerySTATE.AsInteger <> 1 then begin //BAD FLAG
         if MandSQLQueryVER.AsInteger = 1 then
             MandSQLQueryVTARUF.AsString := MandSQLQueryTRF.AsString
         else
             MandSQLQueryVTARUF.AsString := '(' + MandSQLQueryTRF0.AsString + ' * ' + MandSQLQueryK1.AsString + ' * ' + MandSQLQueryK2.AsString + ') ' + MandSQLQueryTRF.AsString;
     end;

     MandSQLQueryVSTATE.AsString := '?';
     case MandSQLQuerySTATE.AsInteger of
        0 : MandSQLQueryVSTATE.AsString := 'НОРМАЛЬНЫЙ';
        1 : MandSQLQueryVSTATE.AsString := 'ИСПОРЧЕН';
        2 : MandSQLQueryVSTATE.AsString := 'ДОСРОЧНО ЗАВЕРШЁН';
        3 : MandSQLQueryVSTATE.AsString := 'УТЕРЯН';
        4 : MandSQLQueryVSTATE.AsString := 'СРОК ДЕЙСТВИЯ ЗАКОНЧИЛСЯ';
        5 : MandSQLQueryVSTATE.AsString := 'ЗАВЕРШЁН ПО НЕУПЛАТЕ';
     end;
end;

procedure TMandSQL.FilterBtnClick(Sender: TObject);
var
     s : string;
begin
     s := GetFilter;
     MandSQLFilterDlg.ShowModal;
     if (s <> GetFilter) OR not MandSQLQuery.Active then begin
         if GetFilter = '' then begin
             if MessageDlg('Фильтр не установлен. Выборка данных может занять много времени. Выполнить?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
                 OpenQuery
             end
         end
         else
             OpenQuery
     end
end;

procedure TMandSQL.FilterMnClick(Sender: TObject);
begin
     FilterBtnClick(nil);
end;

procedure TMandSQL.LostPayClick(Sender: TObject);
begin
    try
     Screen.Cursor := crHourGlass;
     with WorkSQL, WorkSQL.SQL do begin
          Clear;
          Add('SELECT COUNT(*) FROM SYSADM.MANDATOR');
          Add('WHERE STATE <> 3'); //Не утерян
          Add('AND (ISFEE2=''Y'' AND ISPAY2=''N'' AND FEE2DT<:DT');
          //Завершён по неуплате в нужный срок
          Add('AND (STATE<>5 OR STOPDATE IS NULL OR STOPDATE<>(FEE2DT+1)))');

          ParamByName('DT').AsDateTime := (Date - 7);
          Open;
          Screen.Cursor := crDefault;

          if FieldByName(Fields[0].FieldName).AsInteger = 0 then begin
              MessageDlg('На данный момент нет просроченой оплаты 2й части взноса', mtInformation, [mbOK], 0);
              Close;
              exit;
          end;

          if MessageDlg('Обнаружено ' + FieldByName(Fields[0].FieldName).AsString + ' полисов.'#13 + 'Вы хотите сделать их ''ЗАВЕРШЁН ПО НЕУПЛАТЕ''?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then begin
             Close;
             exit;
          end;

          Screen.Cursor := crHourGlass;
          Clear;
          SQLDB.StartTransaction;
          Add('UPDATE SYSADM.MANDATOR');
          Add('SET STATE=5, UPDDT=:DT, STOPDATE=(FEE2DT+1)');
          Add('WHERE STATE <> 3 AND (ISFEE2=''Y'' AND ISPAY2=''N'' AND FEE2DT<:DT AND (STATE<>5 OR STOPDATE IS NULL OR STOPDATE<>(FEE2DT+1)))');
          ParamByName('DT').AsDateTime := (Date - 7);
          ExecSQL;
          SQLDB.Commit;

          MandSQLQuery.Close;
          MandSQLQuery.Open;
     end;
    except
          on E : Exception do begin
             SQLDB.Rollback;
             MessageDlg(E.Message, mtInformation, [mbOK], 0);
          end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TMandSQL.N10Click(Sender: TObject);
var
     All, Curr : integer;
     PSERFLD : TField;
     STATEFLD : TField;
     CheckEnd : boolean;
begin
     try
         PSERFLD := nil;
         Screen.Cursor := crHourGlass;
         WorkSQL.SQL.Clear;
         LOG.Text := '';
         WorkSQL.SQL.Add('SELECT SER, NMB, PSER, PNMB, STATE, REGDATE, STOPDATE, TODT, RETE, RETB FROM SYSADM.MANDATOR ORDER BY 1, 2');
         WorkSQL.Open;
         All := WorkSQL.RecordCount;
         Curr := 0;
         PSERFLD := WorkSQL.FieldByName('PSer');
         STATEFLD := WorkSQL.FieldByName('State');
         while not WorkSQL.Eof do begin
              CheckEnd := true;
              //BAD POLIS STATE 1
              //if MandTable.FieldByName('PayDate').IsNull then begin
              //   if MandTable.FieldByName('State').AsInteger <> 1 then begin //BAD
              //      MandTable.Edit;
              //      MandTable.FieldByName('State').AsInteger := 1;
              //      MandTable.FieldByName('UpdateDate').AsDateTime := Date;
              //      MandTable.Post;
              //      Output(MandTable.FieldByName('Seria').AsString + '/' +
              //             MandTable.FieldByName('Number').AsString + ' испорчен');
              //   end
              //end;

              //LOST POLISES
              if PSERFLD.AsString <> '' then begin
                 HandBookSQL.SQL.Clear;
                 HandBookSQL.SQL.Add('SELECT STATE,STOPDATE FROM SYSADM.MANDATOR WHERE SER=:S AND NMB=:N');
                 HandBookSQL.ParamByName('S').AsString := WorkSQL.FieldByName('PSer').AsString;
                 HandBookSQL.ParamByName('N').AsInteger := WorkSQL.FieldByName('PNMB').AsInteger;
                 HandBookSQL.Open;
                 if HandBookSQL.EOF then begin
                    LOG.Lines.Add(WorkSQL.FieldByName('PSer').AsString + '/' + WorkSQL.FieldByName('PNMB').AsString + ' утерян, но не найден');
                 end
                 else
                 if (HandBookSQL.FieldByName('state').AsInteger <> 3) OR (HandBookSQL.FieldByName('stopdate').AsDateTime <> WorkSQL.FieldByName('regdate').AsDateTime) then begin
                     HandBookSQL.SQL.Clear;
                     SQLDB.StartTransaction;
                     HandBookSQL.SQL.Add('UPDATE SYSADM.MANDATOR SET STATE=3,STOPDATE=:DT WHERE SER=:S AND NMB=:N');
                     HandBookSQL.ParamByName('S').AsString := WorkSQL.FieldByName('PSer').AsString;
                     HandBookSQL.ParamByName('N').AsInteger := WorkSQL.FieldByName('PNMB').AsInteger;
                     HandBookSQL.ParamByName('DT').AsDateTime := WorkSQL.FieldByName('regdate').AsDateTime;
                     HandBookSQL.ExecSQL;
                     SQLDB.Commit;
                     LOG.Lines.Add(WorkSQL.FieldByName('PSer').AsString + '/' + WorkSQL.FieldByName('PNMB').AsString + ' установлен как ''утерянный''');
                 end;
              end;
              //LOST POLISES 2
              if STATEFLD.AsInteger = 3 then begin
                 HandBookSQL.SQL.Clear;
                 HandBookSQL.SQL.Add('SELECT COUNT(*) FROM SYSADM.MANDATOR WHERE PSER=:S AND PNMB=:N');
                 HandBookSQL.ParamByName('S').AsString := WorkSQL.FieldByName('Ser').AsString;
                 HandBookSQL.ParamByName('N').AsInteger := WorkSQL.FieldByName('NMB').AsInteger;
                 HandBookSQL.Open;
                 if HandBookSQL.Fields[0].AsInteger <> 1 then begin
                    LOG.Lines.Add(WorkSQL.FieldByName('Ser').AsString + '/' + WorkSQL.FieldByName('NMB').AsString + ' в состоянии утерян, но числиться в заменах ' + HandBookSQL.Fields[0].AsString + ' раз');
                 end;
                 HandBookSQL.Close;

                 if WorkSQL.FieldByName('StopDate').IsNull then
                    LOG.Lines.Add(WorkSQL.FieldByName('Ser').AsString + '/' + WorkSQL.FieldByName('NMB').AsString + ' в состоянии утерян, но нет даты когда это произошло!');
              end;

              if not WorkSQL.FieldByName('StopDate').IsNull then begin
                if STATEFLD.AsInteger <> 5 then begin //Pay STOP
                 if STATEFLD.AsInteger <> 3 then begin //Утерян
                  if STATEFLD.AsInteger <> 2 then begin //ONLY STOP

                    if (WorkSQL.FieldByName('StopDate').AsDateTime < WorkSQL.FieldByName('ToDT').AsDateTime) AND
                       (WorkSQL.FieldByName('rete').AsFloat > 0) AND
                       (WorkSQL.FieldByName('retb').AsFloat > 0) then begin
                           HandBookSQL.SQL.Clear;
                           HandBookSQL.SQL.Add('UPDATE SYSADM.MANDATOR SET STATE=2 WHERE SER=:S AND NMB=:N');
                           HandBookSQL.ParamByName('S').AsString := WorkSQL.FieldByName('Ser').AsString;
                           HandBookSQL.ParamByName('N').AsInteger := WorkSQL.FieldByName('NMB').AsInteger;
                           SQLDB.StartTransaction;
                           HandBookSQL.ExecSQL;
                           SQLDB.Commit;
                           CheckEnd := false;
                           LOG.Lines.Add(WorkSQL.FieldByName('Ser').AsString + '/' + WorkSQL.FieldByName('NMB').AsString + ' досрочно закрыт');
                     end
                     else
                        LOG.Lines.Add(WorkSQL.FieldByName('Ser').AsString + '/' + WorkSQL.FieldByName('NMB').AsString + ' есть дата остановки, но состояние непонятное');
                  end
                 end 
                end
              end;

              //END ПОЛИС
              if CheckEnd AND (not WorkSQL.FieldByName('ToDt').IsNull) then
                if not (STATEFLD.AsInteger in [1, 2, 3, 4, 5]) then //Испорчен Досрочно, утерян, неуплата
                  if WorkSQL.FieldByName('ToDt').AsDateTime < Date then begin
                     //if MandTable.FieldByName('State').AsInteger <> 4 then begin
                        //MandTable.Edit;
                        //MandTable.FieldByName('State').AsInteger := 4;
                        //MandTable.FieldByName('UpdateDate').AsDateTime := Date;
                        //MandTable.Post;
                        //Output(MandTable.FieldByName('Seria').AsString + '/' +
                        //       MandTable.FieldByName('Number').AsString + ' закончился ' + MandTable.FieldByName('ToDate').AsString);
                     //end
                     HandBookSQL.SQL.Clear;
                     SQLDB.StartTransaction;
                     HandBookSQL.SQL.Add('UPDATE SYSADM.MANDATOR SET STATE=4 WHERE SER=:S AND NMB=:N');
                     HandBookSQL.ParamByName('S').AsString := WorkSQL.FieldByName('Ser').AsString;
                     HandBookSQL.ParamByName('N').AsInteger := WorkSQL.FieldByName('NMB').AsInteger;
                     HandBookSQL.ExecSQL;
                     SQLDB.Commit;
                     LOG.Lines.Add(WorkSQL.FieldByName('Ser').AsString + '/' + WorkSQL.FieldByName('NMB').AsString + ' закончился.');
                  end;

              Inc(Curr);
              ProgressBar.Position := ROUND(Curr * 100 / All);
              StatusBar.Panels[1].Text := 'N' + IntToStr(Curr);
              StatusBar.Update;
              //LabelPos.Caption := IntToStr(ROUND(Curr / 50) * 50);
              Update;
              WorkSQL.Next;

              //if (Curr MOD 250) = 0 then
              //    Output('Позиция Curr=' + IntToStr(Curr) + ' БД ' + IntToStr(MandTable.RecNo));

              if (GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 then begin
                  LOG.Lines.Add('Процесс остановлен по команде пользователя');
                  break;
              end
         end;
     except
         on E : Exception do begin
             if SQLDB.InTransaction then
                 SQLDB.Rollback;
             LOG.Lines.Add('ОШИБКА!!! ' + e.Message);
         end;
     end;
     LOG.Lines.Add('Проверка закончена');
     WorkSQL.Close;
     HandBookSQL.Close;
     ProgressBar.Position := 0;
     Screen.Cursor := crDefault;
end;

procedure TMandSQL.LOGKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if Key = vk_Delete then
         LOG.Text := '';
end;

procedure TMandSQL.PrintListClick(Sender: TObject);
begin
     if PageControl.ActivePage <> TabSheetPrint then begin
         PageControl.ActivePage := TabSheetPrint;
         frReport.ShowReport
     end;
end;

procedure TMandSQL.PageControlChange(Sender: TObject);
begin
     MandSQLQuery.First;
     if PageControl.ActivePage = TabSheetPrint then
         frReport.ShowReport
     else begin
         frReport.EMFPages.Clear;
         frPreview.Clear;
     end;

end;

procedure TMandSQL.Button1Click(Sender: TObject);
begin
     frPreview.Print
end;

procedure TMandSQL.N15Click(Sender: TObject);
begin
    DecoderFrm.ShowModal
end;

procedure TMandSQL.N16Click(Sender: TObject);
begin
    ExpErrsFlag.ShowModal
end;

procedure TMandSQL.Timer2Timer(Sender: TObject);
begin
    if AnimRight then begin                       
        Animate.Text := ' ' + Animate.Text;
        if Canvas.TextWidth(Animate.Text) > Animate.Width then AnimRight := not AnimRight;
    end
    else begin
        Animate.Text := Copy(Animate.Text, 2, 1024);
        if Animate.Text = AnimateText then AnimRight := not AnimRight;
    end;
end;

procedure TMandSQL.N17Click(Sender: TObject);
var
   INI : TIniFile;
   Tbl : TTable;
   TemplateDir, FilePath : string;
   i : integer;
   IsError : boolean;
begin
   INI := TIniFile.Create('blank.ini');
   TemplateDir := INI.ReadString('MANDATORY', 'BuroDirectory', '');
   INI.Free;

   if Length(TemplateDir) = 0 then begin
      MessageDlg('Не указан путь к файлам-шаблонам', mtInformation, [mbOK], 0);
      exit;
   end;

   if TemplateDir[Length(TemplateDir)] <> '\' then
      TemplateDir := TemplateDir + '\';

    if OpenDialog.Execute then begin
        for i := Length(OpenDialog.FileName) downto 1 do
            if OpenDialog.FileName[i] = '\' then break;
        FilePath := Copy(OpenDialog.FileName, 1, i);
        if Copy(OpenDialog.FileName, i + 1, 2) = 'N3' then begin
            if not FileExists(TemplateDir + 'Report3.DBF') then begin
                MessageDlg('Нет файла-шаблона', mtInformation, [mbOk], 0);
                exit;
            end;
            CopyFile(TemplateDir + 'Report3.DBF', FilePath + 'DESTTLK.DBF', ProgressBar);
            Tbl := TTable.Create(self);
            Tbl.TableName := OpenDialog.FileName;
            Screen.Cursor := crHourglass;
            IsError := false;
            if not DropDublicate(Tbl, FilePath + 'DESTTLK.DBF', 'DRV_EXPR', 'F_L', '') then begin
                MessageDlg('Ошибка удаления дубликатов', mtInformation, [mbOk], 0);
                IsError := true;
            end;
            Tbl.Free;

            if not IsError then begin
                DeleteFile(OpenDialog.FileName);
                RenameFile(FilePath + 'DESTTLK.DBF', OpenDialog.FileName);
            end
            else
                DeleteFile(FilePath + 'DESTTLK.DBF');

            Screen.Cursor := crDefault;
        end
        else
        if Copy(OpenDialog.FileName, i + 1, 2) = 'N4' then begin
            if not FileExists(TemplateDir + 'Report4.DBF') then begin
                MessageDlg('Нет файла-шаблона', mtInformation, [mbOk], 0);
                exit;
            end;
            CopyFile(TemplateDir + 'Report4.DBF', FilePath + 'DESTTLK.DBF', ProgressBar);

            WorkLocalSQL.SQL.Clear;
            WorkLocalSQL.SQL.Add('UPDATE "' + OpenDialog.FileName + '" t1 SET body_no=null where body_no in (''08'', ''12'', ''0'', ''1'', ''10'', ''09'')');
            WorkLocalSQL.ExecSQL;

            Tbl := TTable.Create(self);
            Tbl.TableName := OpenDialog.FileName;
            Screen.Cursor := crHourglass;
            IsError := false;
            if not DropDublicate(Tbl, FilePath + 'DESTTLK.DBF', 'BODY_NO', '', '') then begin
                MessageDlg('Ошибка удаления дубликатов', mtInformation, [mbOk], 0);
                IsError := true;
            end;
            Tbl.Free;

            if not IsError then begin
                DeleteFile(OpenDialog.FileName);
                RenameFile(FilePath + 'DESTTLK.DBF', OpenDialog.FileName);
            end
            else
                DeleteFile(FilePath + 'DESTTLK.DBF');

            Screen.Cursor := crDefault;
        end
        else
            MessageDlg('Такие файлы не суппортятся', mtInformation, [mbOk], 0);
    end;
end;

procedure TMandSQL.N11Click(Sender: TObject);
begin
    if ListAvariasSQL = nil then
        Application.CreateForm(TListAvariasSQL, ListAvariasSQL);
    ListAvariasSQL.ShowModal;
end;

procedure TMandSQL.Excel1Click(Sender: TObject);
begin
    DatSetToExcel(MandSQLQuery);
end;

procedure TMandSQL.N22Click(Sender: TObject);
begin
    PolisNmbsComp.ShowModal;
end;

procedure TMandSQL.AddError(S : string; N : integer; txt : string; Code : ErrorTypes; var ErrStrList : string);
var
    ErrStr : string;
    MaxCode : integer;
begin
    Inc(ErrorCount);
    ErrStr := '';
    if ErrStrList <> '' then ErrStrList := ErrStrList + #13#10;
    ErrStrList := ErrStrList + S + Format('/%07u', [N]) + ' ' + Txt + ' ';
    case Code of
    eeBadPolisNumber:  ErrStr := ErrStr + 'Не правильный серия/номер';
    eeBadSymbName:     ErrStr := ErrStr + ' Недопустимые символы в Имени/Названии';
    eeEnglSymbAddr:    ErrStr := ErrStr + ' Английские символы в Адресе';
    eeEnglSymbName:    ErrStr := ErrStr + ' Английские символы в Имени/Названии';
    eeBadFizUr:        ErrStr := ErrStr + ' Неправильный признак Физ/Юр ';
    eeStartFromDigit:  ErrStr := ErrStr + ' Начинается с цифры';
    eeRussLetters:     ErrStr := ErrStr + ' Встретились русские буквы';
    eeBadDiscount:     ErrStr := ErrStr + ' Неправильная скидка';
    eePayAndVer:       ErrStr := ErrStr + ' Несовпадение даты оплаты и версии полиса';
    eeBadSymbAddr:     ErrStr := ErrStr + ' Недопустимые символы в Адресе';
    eeBadSymbSNP:      ErrStr := ErrStr + ' Недопустимые символы в Серии номере потерпевшего';
    eeBadVladen:       ErrStr := ErrStr + ' Неправильный признак основания на владение';
    eeBadAutoMark:     ErrStr := ErrStr + ' Неизвестная марка авто';
    eeBadTC:           ErrStr := ErrStr + ' Не правильный тип Авто';
    eeBadSymbMarka:    ErrStr := ErrStr + ' Недопустимые символы в Марке Авто';
    eeBadCharAuto:     ErrStr := ErrStr + ' Неправильная характеристика Авто';
    eeBadSymbBodyNo:   ErrStr := ErrStr + ' Недопустимые символы в номере кузова или шасси';
    eeRussBodyShassi:  ErrStr := ErrStr + ' Русские символы в номере кузова или шасси';
    eeBadState:        ErrStr := ErrStr + ' Неправильное состояние';
    eePayAfterStart:   ErrStr := ErrStr + ' Оплата после начала';
    eeRegAndPay:       ErrStr := ErrStr + ' Дата регистрации и 1й оплаты не соотносятся';
    eeBadPay2:         ErrStr := ErrStr + ' Дата 2й оплаты не в диапазоне';
    eeStrangeDate:     ErrStr := ErrStr + ' Странная Дата в поле таблицы';
    eeBadStopDate:     ErrStr := ErrStr + ' Дата остановки не в диапазоне';
    eeBadCurr:         ErrStr := ErrStr + ' Не правильная валюта';
    eeBadBaseAT:       ErrStr := ErrStr + ' Не правильный Базовый тип Авто';
    eeBadDecide:       ErrStr := ErrStr + ' Не правильное решение';
    eeBadDoc:          ErrStr := ErrStr + ' Не допустимый серия или номер документа';
    eeNoPay:           ErrStr := ErrStr + ' Не определена дата выплаты';
    eeBadBadMan:       ErrStr := ErrStr + ' Виновник не найден во владельцах';
    eeNoSumma:         ErrStr := ErrStr + ' Не определена ни одна сумма выплаты';
    eeSumNoSum:        ErrStr := ErrStr + ' Сумма вырплаты есть а выплаты нету';
    eeBadAvReg:        ErrStr := ErrStr + ' Дата регистрации до даты аварии';
    eeBadPayAvDate:    ErrStr := ErrStr + ' Дата оплаты до даты регистрации заявления';
    eeEmptyAddr:       ErrStr := ErrStr + ' Адрес не заполнен';
    eeSellAfterStart:  ErrStr := ErrStr + ' Дата продажи после начала';
    eeBadDateEnd:      ErrStr := ErrStr + ' Дата окончания не правильная';
    eeNoDivision:      ErrStr := ErrStr + ' Не определён код подразделения';
    eeNoStopDate:      ErrStr := ErrStr + ' Дата расторжения не указана';
    eeNoNeedPay:       ErrStr := ErrStr + ' 2я оплата + досрочное расторжение по неуплате';
    eeBadAvK:          ErrStr := ErrStr + ' Коэфф. аварийности не правильный';
    else               ErrStr := ErrStr + ' Нет описания';
    end;
    ErrStrList := ErrStrList + ErrStr;
    try
        SQLDB.StartTransaction;
        with WorkSQL, WorkSQL.SQL do begin
            Clear;
            Add('SELECT MAX(N) FROM SYSADM.EXPORTERRS WHERE SER=:S AND NMB=:NMB');
            ParamByName('S').AsString := S;
            ParamByName('NMB').AsInteger := N;
            Open;
            MaxCode := Fields[0].AsInteger;
            Close;
            Clear;
            Add('INSERT INTO SYSADM.EXPORTERRS(SER, NMB, N, ERRCODE, MSG)VALUES(:SER, :NMB, :N, :ERRCODE, :MSG)');
            ParamByName('SER').AsString := S;
            ParamByName('NMB').AsInteger := N;
            ParamByName('N').AsInteger := MaxCode + 1;
            ParamByName('ERRCODE').AsInteger := ord(Code);
            ParamByName('MSG').AsString := ErrStr;
            ExecSQL;
        end;
        SQLDB.Commit;
    except
      on E : Exception do begin
            SQLDB.Rollback;
      end;
    end;
end;
{
  CodeControl = class
       DivCodes : TList; //Список кодов подразделений
       DivMaxCodesOwn : TList; //Список максимальных кодов
       DivMaxCodesBaseCar : TList; //Список максимальных кодов
       DivMaxCodesCar : TList; //Список максимальных кодов
       DivMaxCodesAV : TList; //Список максимальных кодов

       DivListCodesOwn : TList; //Список списков кодов дырок
       DivListCodesBaseCar : TList; //Список списков кодов дырок
       DivListCodesCar : TList; //Список списков кодов дырок
       DivListCodesAV : TList; //Список списков кодов дырок
}
function CodeControl.GetNextOwnerCode(d : integer) : integer;
var
    i, j : integer;
    isfound : boolean;
    L : Tlist;
    CurrCode : integer;
    MaxOwnerCode : integer;
begin
    isfound := false;
    for j := 0 to DivCodes.Count - 1 do
        if Integer(DivCodes[j]) = d then begin
            isfound := true;
            i := j;
            break;
        end;

     if (not isfound) then begin
         i := DivCodes.Count;
         DivCodes.Add(Pointer(d));
     end;

     if (not isfound) OR (DivListCodesOwn[i] = nil) then begin //Заполнить дырки
         L := Tlist.Create;

         if isfound then
             DivListCodesOwn[i] := L
         else begin
             DivListCodesOwn.Add(L);

             DivListCodesBaseCar.Add(nil);
             DivListCodesCar.Add(nil);
             DivListCodesAV.Add(nil);
             DivMaxCodesBaseCar.Add(0);
             DivMaxCodesCar.Add(0);
             DivMaxCodesAV.Add(0);
         end;

          with MandSQLForm.HandBookSQL, MandSQLForm.HandBookSQL.SQL do begin
               //StatusBar.SimpleText := 'Поиск пустых кодов (1)...';
               //StatusBar.Update;

               Clear;
               Add('SELECT OWNERCODE FROM SYSADM.MANDATOR WHERE OWNERCODE > 0 AND DIVISION=:D');
               Add('UNION');
               Add('SELECT OWNRCODE FROM SYSADM.MANDOWN A,SYSADM.MANDATOR M WHERE A.SER=M.SER AND A.NMB=M.NMB AND A.OWNRCODE > 0 AND DIVISION=:D');
               Add('UNION');
               Add('SELECT OWNRCODE FROM SYSADM.MANDAV A,SYSADM.MANDATOR M WHERE A.SER=M.SER AND A.NMB=M.NMB AND A.OWNRCODE > 0 AND DIVISION=:D');
               Add('ORDER BY 1');
               ParamByName('D').AsInteger := D;
               Open;
               CurrCode := 1;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Holes Owner ' + IntToStr(D) +#13#10;
               while not EOF do begin
                     while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                         L.Add(Pointer(CurrCode));
//                         _GLOBAL_TRACE := _GLOBAL_TRACE + IntToStr(CurrCode) +#13#10;
                         Inc(CurrCode);
                     end;
                     if CurrCode = FieldByName(Fields[0].FieldName).AsInteger then
                       Inc(CurrCode);
                     Next;
               end;
               Close;

               MaxOwnerCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDATOR', 'OWNERCODE', IntToStr(d), false);
               CurrCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDOWN', 'OWNRCODE', IntToStr(d), true);
               if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;
               CurrCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDAV', 'OWNRCODE', IntToStr(d), true);
               if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;
               if not isfound then
                   DivMaxCodesOwn.Add(Pointer(MaxOwnerCode))
               else
                   DivMaxCodesOwn[i] := Pointer(MaxOwnerCode);
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Max OwnerCode div=' + IntToStr(d) + ' code=' + IntToStr(MaxOwnerCode) + #13#10;
         end;
     end;

     if TList(DivListCodesOwn[i]).Count > 0 then begin
        GetNextOwnerCode := Integer(TList(DivListCodesOwn[i])[0]);
//        _GLOBAL_TRACE := _GLOBAL_TRACE + 'GetHole Owner div=' + IntToStr(d) + ' code=' + IntToStr(Integer(TList(DivListCodesOwn[i])[0])) + #13#10;
        TList(DivListCodesOwn[i]).Delete(0);
        exit;
     end;
     DivMaxCodesOwn[i] := Pointer(Integer(DivMaxCodesOwn[i]) + 1);
     GetNextOwnerCode := Integer(DivMaxCodesOwn[i]);
end;

function CodeControl.GetNextBaseCarCode(d : integer) : integer;
var
    isfound : boolean;
    i, j, CurrCode, MaxSerTCCode : integer;
    L : TList;
begin
    isfound := false;
    for j := 0 to DivCodes.Count - 1 do
        if Integer(DivCodes[j]) = d then begin
            isfound := true;
            i := j;
            break;
        end;

     if (not isfound) then begin
         i := DivCodes.Count;
         DivCodes.Add(Pointer(d));
     end;

     if (not isfound) OR (DivListCodesBaseCar[i] = nil) then begin //Заполнить дырки
         L := Tlist.Create;

         if isfound then
             DivListCodesBaseCar[i] := L
         else begin
             DivListCodesBaseCar.Add(L);

             DivListCodesOwn.Add(nil);
             DivListCodesCar.Add(nil);
             DivListCodesAV.Add(nil);
             DivMaxCodesOwn.Add(0);
             DivMaxCodesCar.Add(0);
             DivMaxCodesAV.Add(0);
         end;

          with MandSQLForm.HandBookSQL, MandSQLForm.HandBookSQL.SQL do begin
               Clear;
               Add('SELECT BaseCarCode FROM SYSADM.MANDATOR WHERE BaseCarCode > 0 and division=:D');
               Add('UNION');
               Add('SELECT A.BaseCarCode FROM SYSADM.MANDAV A,SYSADM.MANDATOR M WHERE A.SER=M.SER AND A.NMB=M.NMB AND A.BaseCarCode > 0 and division=:D');
               Add('ORDER BY 1');
               ParamByName('D').AsInteger := D;
               Open;
               CurrCode := 1;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Holes BaseCarCode ' + IntToStr(D) +#13#10;
               while not EOF do begin
                     while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                         L.Add(Pointer(CurrCode));
//                         _GLOBAL_TRACE := _GLOBAL_TRACE + IntToStr(CurrCode) +#13#10;
                         Inc(CurrCode);
                     end;
                     if CurrCode = FieldByName(Fields[0].FieldName).AsInteger then
                       Inc(CurrCode);
                     Next;
               end;
               Close;

               MaxSerTCCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDATOR', 'BaseCarCode', IntToStr(d), false);
               CurrCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDAV', 'BaseCarCode', IntToStr(d), true);
               if CurrCode > MaxSerTCCode then MaxSerTCCode := CurrCode;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Max BaseCarCode div=' + IntToStr(d) + ' code=' + IntToStr(MaxSerTCCode) + #13#10;

               if not isfound then
                   DivMaxCodesBaseCar.Add(Pointer(MaxSerTCCode))
               else
                   DivMaxCodesBaseCar[i] := Pointer(MaxSerTCCode);
         end;
     end;

     if TList(DivListCodesBaseCar[i]).Count > 0 then begin
        GetNextBaseCarCode := Integer(TList(DivListCodesBaseCar[i])[0]);
//        _GLOBAL_TRACE := _GLOBAL_TRACE + 'GetHole BaseCarCode div=' + IntToStr(d) + ' code=' + IntToStr(Integer(TList(DivListCodesBaseCar[i])[0])) + #13#10;
        TList(DivListCodesBaseCar[i]).Delete(0);
        exit;
     end;
     DivMaxCodesBaseCar[i] := Pointer(Integer(DivMaxCodesBaseCar[i]) + 1);
//     _GLOBAL_TRACE := _GLOBAL_TRACE + 'GetMax BaseCarCode div=' + IntToStr(d) + ' code=' + IntToStr(Integer(DivMaxCodesBaseCar[i])) + #13#10;
     GetNextBaseCarCode := Integer(DivMaxCodesBaseCar[i]);
end;

function CodeControl.GetNextCarCode(d : integer) : integer;
var
    isfound : boolean;
    i, CurrCode, MaxTCCode : integer;
    L : TList;
begin
    isfound := false;
    for i := 0 to DivCodes.Count - 1 do
        if Integer(DivCodes[i]) = d then begin
            isfound := true;
            break;
        end;

     if (not isfound) then begin
         i := DivCodes.Count;
         DivCodes.Add(Pointer(d));
     end;

     if (not isfound) OR (DivListCodesCar[i] = nil) then begin //Заполнить дырки
         L := Tlist.Create;

         if isfound then
             DivListCodesCar[i] := L
         else begin
             DivListCodesCar.Add(L);

             DivListCodesOwn.Add(nil);
             DivListCodesBaseCar.Add(nil);
             DivListCodesAV.Add(nil);
             DivMaxCodesOwn.Add(0);
             DivMaxCodesBaseCar.Add(0);
             DivMaxCodesAV.Add(0);
         end;

          with MandSQLForm.HandBookSQL, MandSQLForm.HandBookSQL.SQL do begin
               Clear;
               Add('SELECT CarCode FROM SYSADM.MANDATOR WHERE CarCode > 0 and division=:d');
               Add('UNION');
               Add('SELECT A.CarCode FROM SYSADM.MANDAV A,SYSADM.MANDATOR M WHERE A.SER=M.SER AND A.NMB=M.NMB AND A.CarCode > 0 and division=:d');
               Add('ORDER BY 1');
               ParamByName('D').AsInteger := D;
               Open;
               CurrCode := 1;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Holes CarCode ' + IntToStr(D) +#13#10;
               while not EOF do begin
                     while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                         L.Add(Pointer(CurrCode));
//                         _GLOBAL_TRACE := _GLOBAL_TRACE + IntToStr(CurrCode) +#13#10;
                         Inc(CurrCode);
                     end;
                     if CurrCode = FieldByName(Fields[0].FieldName).AsInteger then
                       Inc(CurrCode);
                     Next;
               end;
               Close;

                MaxTCCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDATOR', 'CarCode', IntToStr(d), false);
                CurrCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDAV', 'CarCode', IntToStr(d), true);
                if CurrCode > MaxTCCode then MaxTCCode := CurrCode;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Max CarCode div=' + IntToStr(d) + ' code=' + IntToStr(MaxTCCode) + #13#10;

               if not isfound then
                   DivMaxCodesCar.Add(Pointer(MaxTCCode))
               else
                   DivMaxCodesCar[i] := Pointer(MaxTCCode);
         end;
     end;

     if TList(DivListCodesCar[i]).Count > 0 then begin
        GetNextCarCode := Integer(TList(DivListCodesCar[i])[0]);
        TList(DivListCodesCar[i]).Delete(0);
        exit;
     end;
     DivMaxCodesCar[i] := Pointer(Integer(DivMaxCodesCar[i]) + 1);
     GetNextCarCode := Integer(DivMaxCodesCar[i]);
end;

function CodeControl.GetNextAvCode(d : integer) : integer;
var
    isfound : boolean;
    i, CurrCode, MaxAVCode : integer;
    L : TList;
begin
    isfound := false;
    for i := 0 to DivCodes.Count - 1 do
        if Integer(DivCodes[i]) = d then begin
            isfound := true;
            break;
        end;

     if (not isfound) then begin
         i := DivCodes.Count;
         DivCodes.Add(Pointer(d));
     end;

     if (not isfound) OR (DivListCodesAV[i] = nil) then begin //Заполнить дырки
         L := Tlist.Create;

         if isfound then
             DivListCodesAV[i] := L
         else begin
             DivListCodesAV.Add(L);

             DivListCodesOwn.Add(nil);
             DivListCodesBaseCar.Add(nil);
             DivListCodesCar.Add(nil);
             DivMaxCodesOwn.Add(0);
             DivMaxCodesBaseCar.Add(0);
             DivMaxCodesCar.Add(0);
         end;

          with MandSQLForm.HandBookSQL, MandSQLForm.HandBookSQL.SQL do begin
               Clear;
               Add('SELECT DISTINCT AVCODE FROM SYSADM.MANDAV A,SYSADM.MANDATOR M WHERE A.SER=M.SER AND A.NMB=M.NMB AND AVCODE > 0 and division=:d');
               Add('ORDER BY 1');
               ParamByName('D').AsInteger := D;
               Open;
               CurrCode := 1;
//               _GLOBAL_TRACE := _GLOBAL_TRACE + 'Holes Av ' + IntToStr(D) +#13#10;
               while not EOF do begin
                     while CurrCode < FieldByName(Fields[0].FieldName).AsInteger do begin
                         L.Add(Pointer(CurrCode));
//                         _GLOBAL_TRACE := _GLOBAL_TRACE + IntToStr(CurrCode) +#13#10;
                         Inc(CurrCode);
                     end;
                     if CurrCode = FieldByName(Fields[0].FieldName).AsInteger then
                       Inc(CurrCode);
                     Next;
               end;
               Close;

               MaxAVCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDAV', 'AVCODE', IntToStr(d), true);
               // CurrCode := MandSQLForm.GetMaxCodeFor('SYSADM.MANDAV', 'CarCode', IntToStr(d));
                //if CurrCode > MaxTCCode then MaxTCCode := CurrCode;

               if not isfound then
                   DivMaxCodesAv.Add(Pointer(MaxAVCode))
               else
                   DivMaxCodesAv[i] := Pointer(MaxAVCode);
         end;
     end;

     if TList(DivListCodesAv[i]).Count > 0 then begin
        GetNextAvCode := Integer(TList(DivListCodesAv[i])[0]);
        TList(DivListCodesAv[i]).Delete(0);
        exit;
     end;
     DivMaxCodesAv[i] := Pointer(Integer(DivMaxCodesAv[i]) + 1);
     GetNextAvCode := Integer(DivMaxCodesAv[i]);
end;

constructor CodeControl.Create;
begin
       DivCodes := TList.Create;
       DivMaxCodesOwn := TList.Create;
       DivMaxCodesBaseCar := TList.Create;
       DivMaxCodesCar := TList.Create;
       DivMaxCodesAV := TList.Create;

       DivListCodesOwn := TList.Create;
       DivListCodesBaseCar := TList.Create;
       DivListCodesCar := TList.Create;
       DivListCodesAV := TList.Create;
end;

destructor CodeControl.Free;
var
    i : integer;
begin
       DivCodes.Free;
       DivMaxCodesOwn.Free;
       DivMaxCodesBaseCar.Free;
       DivMaxCodesCar.Free;
       DivMaxCodesAV.Free;

       for i := 0 to DivListCodesOwn.Count - 1 do
           TList(DivListCodesOwn[i]).Free;
       DivListCodesOwn.Free;

       for i := 0 to DivListCodesBaseCar.Count - 1 do
           TList(DivListCodesBaseCar[i]).Free;
       DivListCodesBaseCar.Free;

       for i := 0 to DivListCodesCar.Count - 1 do
           TList(DivListCodesCar[i]).Free;
       DivListCodesCar.Free;

       for i := 0 to DivListCodesAV.Count - 1 do
           TList(DivListCodesAV[i]).Free;
       DivListCodesAV.Free;
end;

procedure TMandSQL.N13Click(Sender: TObject);
begin
     if AvSQLForm = nil then begin
        Application.CreateForm(TAvSQLForm, AvSQLForm);
        AvSQLForm.ShowModal;
     end;
end;

function NotExistsPay2(s, n : string) : boolean;
begin
    NotExistsPay2 := false;
    MandSQLForm.HandBookSQL.Close;
    MandSQLForm.HandBookSQL.SQL.Clear;
    MandSQLForm.HandBookSQL.SQL.Add('SELECT ISPAY2 FROM SYSADM.MANDATOR WHERE SER=''' + s + ''' AND NMB=' + n);
    MandSQLForm.HandBookSQL.Open;
    if not MandSQLForm.HandBookSQL.EOF then
        NotExistsPay2 := MandSQLForm.HandBookSQL.FieldByName('ISPAY2').AsString = 'N';
end;

function DecodeAgentName(s : string) : string;
begin
    if lastAgCode = s then begin
        DecodeAgentName := lastAgName;
        exit;
    end;
    lastAgCode := s;
    MandSQLForm.HandBookSQL.Close;
    MandSQLForm.HandBookSQL.SQL.Clear;
    MandSQLForm.HandBookSQL.SQL.Add('SELECT NAME FROM SYSADM.AGENT WHERE AGENT_CODE = ''' + s + '''');
    MandSQLForm.HandBookSQL.Open;
    if MandSQLForm.HandBookSQL.Eof then lastAgName := s
    else lastAgName := MandSQLForm.HandBookSQL.FieldByName('NAME').AsString;
    DecodeAgentName := lastAgName;
end;


procedure TMandSQL.GloboClick(Sender: TObject);
var
    F1 : TextFile;
    //F2 : TextFile;
    F3 : TextFile;
    F6 : TextFile;
    F8 : TextFile;
    F10 : TextFile;
    F11 : TextFile;
    //F12 : TextFile;
    F13 : TextFile;
    AVAR : TextFile;
    All, Curr : integer;
    IsError : boolean;
    UnloadDt : string;
    Division : string;
    DD, MM, YY : WORD;
    errStr : string;
    UnloadN, path : string;
    Hour, Min, Sec, MSec : word;
    Prefix : string;
const
    InsTypeCode : string = '30';    
begin
    Prefix := 'A.';
    if (GetAsyncKeyState(VK_SHIFT) and $8000) <> 0 then Prefix := '';

    errStr := '';
    DecodeTime(Now, Hour, Min, Sec, MSec);
    UnloadN := IntToStr(Hour);
    if Hour < 10 then UnloadN := '0' + UnloadN;
    DecodeDate(Date, YY, MM, DD);
    UnloadDt := IntToStr(DD);
    if DD < 10 then UnloadDt := '0' + UnloadDt;
    if MM < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(MM);
    Division := '00';
    IsError := false;

    SaveDialog.FileName := UnloadN + '.unload';
    if not SaveDialog.Execute then exit;
    path := Copy(SaveDialog.FileName, 1, Length(SaveDialog.FileName) - Pos('\', ReverseString(SaveDialog.FileName)) + 1);

    AssignFile(F1,  path + Division + 'D' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F1);
    //AssignFile(F2, path + Division + 'L' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    //ReWrite(F2);
    AssignFile(F3, path + Division + 'T' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F3);
    AssignFile(F6, path + Division + 'C' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F6);
    AssignFile(F8, path + Division + 'P' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F8);
    AssignFile(F10, path + Division + 'H' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F10);
    AssignFile(F11, path + Division + 'Z' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F11);
    //AssignFile(F12, path + Division + 'G' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    //ReWrite(F12);
    AssignFile(F13, path + Division + 'B' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F13);
    AssignFile(AVAR, path + Division + 'M' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(AVAR);

    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('UPDATE SYSADM.MANDATOR M SET DIVISION=(SELECT DIVISIONCODE FROM AGENT A WHERE A.AGENT_CODE=M.AGNT) WHERE DIVISION IS NULL ');
    if GetFilter <> '' then WorkSQL.SQL.Add(' AND ' + GetFilter);
    WorkSQL.ExecSQL;

    WorkSQL.Close;
    WorkSQL.SQL.Clear ;
    WorkSQL.SQL.Add('SELECT * FROM SYSADM.MANDATOR M ');
    if Length(GetFilter) <> 0 then
        WorkSQL.SQL.Add(' WHERE ' + GetFilter);
    WorkSQL.Open;
    All := MandSQLQuery.RecordCount;
    Curr := 1;
    with WorkSQL do begin
        while not Eof do begin
          try
            Curr := Curr + 1;
            StatusBar.Panels[0].Text := IntToStr(Curr) + ' шт.';
            StatusBar.Update;
            ProgressBar.Position := round(Curr * 100 / All);
            if FieldByName('DIVISION').AsString = '' then
                raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' нет кода подразделения');
            if FieldByName('STATE').AsInteger = 1 then begin //ИСПОРЧЕН
                WriteLn(F13, FieldByName('DIVISION').AsString + '|25|' +
                            DecodeBlank(FieldByName('SER').AsString) + '|' +
                            FieldByName('SER').AsString + '|' +
                            FieldByName('NMB').AsString + '|' +
                            FieldByName('NMB').AsString + '|' +
                            DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                )
            end
            else begin
                //if FieldByName('PSER').AsString = '' then begin
                    if Trim(FieldByName('NAME').AsString) = '' then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' имя страхователя не указано');
                    if (FieldByName('AGTYPE').AsString = '') OR not (FieldByName('AGTYPE').AsString[1] in ['F', 'U']) then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' тип агента не определён "' + FieldByName('AGTYPE').AsString + '"');
                    if DecodePeriod(FieldByName('PERIOD').AsString) = '?' then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' период не определён "' + FieldByName('PERIOD').AsString + '"');
                    if GetK(FieldByName('K').AsFloat) = '?' then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' коэффициент местности не правильный "' + FieldByName('K').AsString + '"');
                    if GetAVK(FieldByName('K2').AsFloat) = '?' then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' коэффициент аварийности не правильный "' + FieldByName('K2').AsString + '"');
                    if FieldByName('STATE').AsString = '2' then
                        if FieldByName('STOPDATE').IsNull then
                            raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' не указана дата расторжения');
                    if not FieldByName('STOPDATE').IsNull then
                        if FieldByName('STOPDATE').AsDateTime > FieldByName('TODT').AsDateTime then
                            raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' расторжение после окончания');
                    if FieldByName('STATE').AsString = '5' then
                        if FieldByName('ISPAY2').AsString = 'Y' then
                            raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' указана 2я оплата');
                    if DecodeVladen(FieldByName('VLADEN').AsString) = '?' then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' основание на вледение "' + FieldByName('VLADEN').AsString + '"');
                    //if Trim(FieldByName('AUTONMB').AsString) = '' then
                    //    raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' гос. номер не заполнен');
                    if Length(FieldByName('AUTONMB').AsString) > 10 then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' гос. номер очень длиный');
                    if (FieldByName('PAY1DT').AsDateTime < FieldByName('FRDT').AsDateTime - 45) or (FieldByName('PAY1DT').AsDateTime > FieldByName('FRDT').AsDateTime) then
                        raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' несоответствие оплаты и периода действия ' + FieldByName('PAY1DT').AsString + ' ' + FieldByName('FRDT').AsString + '-' + FieldByName('TODT').AsString);

                    if MandSQLFilterDlg.IsAvar.Checked then begin
                        ExportAvariasGLOBO.Close;
                        ExportAvariasGLOBO.MacroByName('WHERE').AsString := 'A.SER=''' + FieldByName('SER').AsString + ''' AND A.NMB=' + FieldByName('NMB').AsString;
                        ExportAvariasGLOBO.Open;
                        while not ExportAvariasGLOBO.Eof do begin
                            WriteLn(AVAR, ExportAvariasGLOBO.FieldByName('DIVISION').AsString + '|30|' +
                                          DecodeBlank(ExportAvariasGLOBO.FieldByName(Prefix + 'SER').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'SER').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'NMB').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'WRKNMB').AsString + '|' +
                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'AVDT').AsDateTime) + '|' +
                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'REGDT').AsDateTime) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'PLACE').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'PLACE_CODE').AsString + '|' +
                                          DecodeAvState(ExportAvariasGLOBO.FieldByName(Prefix + 'DECISION').AsString, nil, nil) + '|' +
                                          DecodeGlDate2(ExportAvariasGLOBO.FieldByName(Prefix + 'DTCLOSE'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'DT1'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'DT2'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'DT3')
                                                       ) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'FIO').AsString + '|' +
                                          DecodeFU12(ExportAvariasGLOBO.FieldByName(Prefix + 'ISFIZ').AsString) + '|' +
                                          DecodeRN(ExportAvariasGLOBO.FieldByName(Prefix + 'ISRSD').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'ADRCODE').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'ADDR').AsString + '|' +
                                          DecodeSN(ExportAvariasGLOBO.FieldByName(Prefix + 'DOCSN').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'LTR').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'MARKA').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'AUTONMB').AsString + '|' +
                                          '' + '|' +  //год выпуска
                                          DecodeBodyShas(ExportAvariasGLOBO.FieldByName(Prefix + 'BODYSHS').AsString) + '|' +
                                          DecodeCharact2(ExportAvariasGLOBO.FieldByName(Prefix + 'CHRCT').AsString, ExportAvariasGLOBO.FieldByName(Prefix + 'LTR').AsString) + '|' +
                                          DecodeCurrency(ExportAvariasGLOBO.FieldByName(Prefix + 'CURR').AsString) + '|' + //Код валюты (заявленная)
                                          Format('%0.2f', [ExportAvariasGLOBO.FieldByName(Prefix + 'DMG1').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DMG2').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DMG3').AsFloat]) + '|' + //Заявленная сумма
                                          DecodeCurrency(ExportAvariasGLOBO.FieldByName(Prefix + 'CURR').AsString) + '|' + //Код валюты выплаты
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'PAY').AsString + '|' + //Выплачено
                                          '2' + '|' + //Форма оплаты
                                          '' + '|' + //Номер платежного документа
                                          DecodeOneDate(ExportAvariasGLOBO.FieldByName(Prefix + 'DT1'), ExportAvariasGLOBO.FieldByName(Prefix + 'DT2'), ExportAvariasGLOBO.FieldByName(Prefix + 'DT3')) + '|'); //Дата платежного документа
                            if (ExportAvariasGLOBO.FieldByName(Prefix + 'DMG1').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DMG2').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DMG3').AsFloat) = 0 then
                                raise Exception.Create(FieldByName('SER').AsString + '/' + FieldByName('NMB').AsString + ' нет требований по суммам в аварии');
                            ExportAvariasGLOBO.Next;
                        end;
                    end;

                    WriteLn(F1, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, -1) + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString) + '|' +
                                DecodeGlDate(FieldByName('FRDT').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('FRTM').AsString) + '|' +
                                DecodeGlDate(FieldByName('TODT').AsDateTime) + '|' +
                                DecodeVariant(FieldByName('TRFTBL').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('NAME').AsString + '|' +
                                DecodeFU12(FieldByName('OWNTYPE').AsString) + '|' +
                                DecodeR(FieldByName('TRFTBL').AsString) + '|' +
                                FieldByName('PLACE_CODE').AsString + '|' +
                                FieldByName('ADDR').AsString + '|' +
                                DecodeOwnTypeForm(FieldByName('OWNTYPE').AsString) + '|' + //Форма собственности
                                DecodePayGraphic(FieldByName('ISFEE2').AsString) + '|' + //График оплаты
                                DecodePayType(FieldByName('PAY1TYPE').AsString) + '|' +
                                DecodeAgentName(FieldByName('AGNT').AsString) + '|' +
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма по полису (на весь период)
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('EUR') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('TRF').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TRF0').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('EUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' //Для заметок
                    );
                    {WriteLn(F2, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeCurrency('EUR') + '|' + //Код валюты
                                '10000' + '|' + //Страховая сумма
                                FieldByName('NAME').AsString + '|' +
                                '' + '|' + //Дата рождения
                                '' + '|' + //Паспорт
                                FieldByName('ADDR').AsString
                    );}
                    WriteLn(F3, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '1' + '|' + //Номер застрахованного
                                FieldByName('LTR').AsString + '|' + //Тип транспортного средства
                                FieldByName('MARKA').AsString + '|' +
                                DecodeNone(FieldByName('AUTONMB').AsString, 'НЕ УКАЗАНО') + '|' +
                                '' + '|' + //Год выпуска
                                FieldByName('NMBODY').AsString + '|' +
                                FieldByName('CHASSIS').AsString + '|' +
                                '' + '|' + //Серия техпаспорта
                                '' + '|' + //Номер техпаспорта
                                DecodeCharact(FieldByName('LTR').AsString,
                                              FieldByName('VOLUME').AsString,
                                              FieldByName('TONNAGE').AsString,
                                              FieldByName('CNTPAS').AsString,
                                              FieldByName('POWER').AsString) + '|' +
                                '' + '|' + //Стоимость
                                '' + '|' + //Код валюты (стоимости и страховой суммы)
                                '' + '|' + //Страховая сумма по объекту
                                DecodeVladen(FieldByName('VLADEN').AsString) + '|' +
                                '' + '|'//Символ страны прибытия
                    );
                    if FieldByName('DISCOUNT').AsFloat <> 0 then
                    WriteLn(F6, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '|' + //Номер застрахованного не указан
                                'H|' + //Код коэффициента
                                '3|' + //Единица измерения Коэффициент
                                FieldByName('DISCOUNT').AsString  + '|' //Значение коэффициента
                    );
                    if DecodeR(FieldByName('TRFTBL').AsString) = '1' then //резидент
                    WriteLn(F6, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '|' + //Номер застрахованного не указан
                                 GetK(FieldByName('K').AsFloat) + '|' + //Код коэффициента
                                '4|' + //Единица измерения Коэффициент
                                FieldByName('K').AsString  + '|' //Значение коэффициента
                    );
                    WriteLn(F6, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '|' + //Номер застрахованного не указан
                                 GetAvK(FieldByName('K2').AsFloat) + '|' + //Код коэффициента
                                '4|' + //Единица измерения Коэффициент
                                FieldByName('K2').AsString + '|'//Значение коэффициента
                    );
                    if FieldByName('PNMB').AsInteger = 0 then //не дубликат
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('PAY1TYPE').AsString, FieldByName('PAY1DOC').AsString) + '|' + //Форма оплаты
                                KillZero(FieldByName('PAY1DOC').AsString) + '|' +
                                DecodeGlDate(FieldByName('PAY1DT').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAY1DT').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('PAY1DT').AsDateTime) + '|' +
                                FieldByName('PAY1').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('PAY1CURR').AsString) + '|' + //Код валюты
                                FieldByName('AGPCNT').AsString + '|' + //Комиссия
                                DecodeAgentName(FieldByName('AGNT').AsString) + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('AGPCNT').AsFloat)
                    );
                    if FieldByName('ISPAY2').AsString = 'Y' then
                    if (FieldByName('PNMB').AsInteger = 0) or NotExistsPay2(FieldByName('PSER').AsString, FieldByName('PNMB').AsString) then
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('PAY2TYPE').AsString, FieldByName('PAY2DOC').AsString) + '|' + //Форма оплаты
                                KillZero(FieldByName('PAY2DOC').AsString) + '|' +
                                DecodeGlDate(FieldByName('PAY2DT').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAY2DT').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('PAY2DT').AsDateTime) + '|' +
                                FieldByName('PAY2').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('PAY2CURR').AsString) + '|' + //Код валюты
                                FieldByName('AGPCNT').AsString + '|' + //Комиссия
                                DecodeAgentName(FieldByName('AGNT').AsString) + '|' +
                                DecodeFU16(FieldByName('AGTYPE').AsString, FieldByName('AGPCNT').AsFloat)
                    );
                    if FieldByName('STATE').AsString = '2' then
                    WriteLn(F8, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '51' + '|' + //Код события
                                DecodeNB('N', '') + '|' + //Форма оплаты
                                '|' +
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                                FieldByName('RETE').AsString + '|' + //Сумма
                                DecodeCurrency('EUR') + '|' + //Код валюты
                                '|' + //Комиссия
                                '|'
                    );
                    {if (FieldByName('ISFEE2').AsString = 'Y' ) AND (FieldByName('ISPAY2').AsString = 'N') then
                    WriteLn(F12, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                DecodeGlDate(FieldByName('FEE2DT').AsDateTime) + '|' +
                                Format('%0.2f', [FieldByName('FEE2').AsFloat]) + '|' + //Сумма
                                DecodeCurrency('EUR') //Код валюты
                    );}
                //end
                //else begin
                if FieldByName('PSER').AsString <> '' then begin
                    WriteLn(F10, FieldByName('DIVISION').AsString + '|30|' +
                                DecodeBlank(FieldByName('PSER').AsString) + '|' +
                                FieldByName('PSER').AsString + '|' +
                                FieldByName('PNMB').AsString + '|' +
                                DecodeBlank(FieldByName('SER').AsString) + '|' +
                                FieldByName('SER').AsString + '|' +
                                FieldByName('NMB').AsString + '|' +
                                '9' + '|' +                
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                    );
                end;
                if FieldByName('STATE').AsString = '5' then
                WriteLn(F11, FieldByName('DIVISION').AsString + '|30|' +
                            DecodeBlank(FieldByName('SER').AsString) + '|' +
                            FieldByName('SER').AsString + '|' +
                            FieldByName('NMB').AsString + '|' +
                            '|' +
                            DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                            '2' + '|' +
                            '2' + '|' + //расторгнут по неуплате
                            '2'
                );
                if FieldByName('STATE').AsString = '2' then
                WriteLn(F11, FieldByName('DIVISION').AsString + '|30|' +
                            DecodeBlank(FieldByName('SER').AsString) + '|' +
                            FieldByName('SER').AsString + '|' +
                            FieldByName('NMB').AsString + '|' +
                            '|' +
                            DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                            '2' + '|' +
                            '2' + '|' + //расторгнут по неуплате
                            '42'
                );
            end; //if
            Next;
          except
            on E : Exception do begin
                errStr := errStr + #13#10 + e.Message + ' ' + DecodeAgentName(FieldByName('AGNT').AsString);
                //MessageDlg(e.Message, mtInformation, [mbOk], 0);
                IsError := true;
                Next;
            end
        end; //while
    end; //width}
  end;
  StatusBar.Panels[0].Text := '';
    CloseFile(F1);
//    CloseFile(F2);
    CloseFile(F3);
    CloseFile(F6);
    CloseFile(F8);
    CloseFile(F10);
    CloseFile(F11);
    //CloseFile(F12);
    CloseFile(F13);
    CloseFile(AVAR);
    ProgressBar.Position := 0;

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
