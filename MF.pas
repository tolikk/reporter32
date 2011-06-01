unit MF;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Menus, Db, DBTables, Grids, DBGrids, RXDBCtrl, ExtCtrls, StdCtrls, RXSpin,
  Buttons, RxQuery, inifiles, ComCtrls, FileUtil, BdeUtils, Excel97, math, strutils, shellapi;

const
    MAND_SECT : string = 'MANDATORY';
    KOMPLEX_SECT : string = 'KOMPLEX';

type
   AgReportPayStruct = record
        Summ : double;
        RubSumm : double;
   end;

   AgReportPayListStruct = array[1..10] of AgReportPayStruct;

   AgReportStruct = record
        AgCode : string;
        PageName : string;
        Curr : array[1..10] of string;
        LastRubSumm : double;
        NalSumm : AgReportPayListStruct;
        BezNalSumm : AgReportPayListStruct;
   end;

   PAgReportStruct = ^AgReportStruct;

   AvKType = record
        Val : double;
        K : string;
   end;

   ReportStructure = record
	    Curr : string[10];
	    BRUTTO_PREMUIM : double;
	    FOND_PREMUIM : double;
	    AGENT_PREMUIM : double;
	    BASE_PREMUIM : double;
	    RNP2_PREMUIM : double; //Dates
	    RNP5_PREMUIM : double; //1 / 2 BASE
   end;

   DupRecord = record
       Code : string[32];
       DivCode : integer;
       DupCount : integer;
   end;

   PDupRecord = ^DupRecord;

   CarOwnerDescription = record
       Name : string[50];
       Addr_code : string[3];
       Address : string[50];
       Owner_Type : integer;
       R_N : integer;
       F_L : integer;
       DISCOUNT_2 : double;
       DVR_EXPR : integer;
   end;

   SerCarDescription = record
       Marka : string[30];
       TypeTC : string[3];
       Characteristic : double;
   end;

   CarDescription = record
       AutoNumber : string[10];
       BodyNumber : string[20];
       BaseCarCode : double;
   end;

  TMainForm = class(TForm)
    MainMenu: TMainMenu;
    PolisesMenu: TMenuItem;
    N2: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    FilterMn: TMenuItem;
    N9: TMenuItem;
    MainQuery: TQuery;
    DB: TDatabase;
    MainDataSource: TDataSource;
    MainGrid: TRxDBGrid;
    MenuCompany: TMenuItem;
    PrintDialog: TPrintDialog;
    MainQuerySERIA: TStringField;
    MainQueryNUMBER: TFloatField;
    MainQueryNAME: TStringField;
    MainQueryFROMDATE: TDateField;
    MainQueryTODATE: TDateField;
    MainQueryPAYDATE: TDateField;
    MainQueryDRAFTNUMBER: TFloatField;
    MainQueryTARIF: TFloatField;
    MainQueryRATE: TFloatField;
    MainQueryDISCOUNT: TFloatField;
    MainQueryPAY1: TFloatField;
    MainQueryAGNAME: TStringField;
    MainQueryPAY1CURR: TStringField;
    MainQueryPAYWITHCURR: TStringField;
    MainQueryDISCOUNTTEXT: TStringField;
    Timer: TTimer;
    WorkSQL: TQuery;
    MainQueryISFEE2: TFloatField;
    MainQueryISPAY2: TFloatField;
    MainQueryFee2Symbol: TStringField;
    MainQueryPay2Symbol: TStringField;
    MainQueryPAY2DATE: TDateField;
    MainQueryLETTER: TStringField;
    N4: TMenuItem;
    AgentTbl: TTable;
    AgentTblAgent_code: TStringField;
    N7: TMenuItem;
    MainQueryFEE2DATE: TDateField;
    Panel1: TPanel;
    Button1: TButton;
    Button2: TButton;
    MainQueryAUTONUMBER: TStringField;
    N12: TMenuItem;
    ExplorerMnu: TMenuItem;
    StateText: TStaticText;
    AvForPolises: TMenuItem;
    MainQueryPREVSERIA: TStringField;
    MainQueryPREVNUMBER: TFloatField;
    MainQueryOWNEROFPOLIS: TStringField;
    MainQueryCHANGE: TStringField;
    MainQueryVERSION: TFloatField;
    MenuItemBuro: TMenuItem;
    StatusBar: TStatusBar;
    ProgressBar: TProgressBar;
    HandBookSQL: TQuery;
    KorrectMenuItem: TMenuItem;
    MainQueryADDRESS: TStringField;
    MainQueryMARKA: TStringField;
    ExpSQL_T: TTable;
    ExpSQL_Q: TQuery;
    TabUpdater: TTable;
    MainQueryPERIOD_FT: TStringField;
    MainQueryPERIOD: TStringField;
    MainQueryMR_PAY1_SUMVAL: TStringField;
    MainQueryMR_PAY1_RATE: TStringField;
    MainQueryMR_PAY1_SUMBRB: TStringField;
    MainQueryMR_PAY2_SUMVAL: TStringField;
    MainQueryMR_PAY2_RATE: TStringField;
    MainQueryMR_PAY2_SUMBRB: TStringField;
    MainQueryFEE2: TFloatField;
    MainQueryPAY2RATE: TFloatField;
    MainQueryPAY2SUMMA: TFloatField;
    MainQueryREGDATE: TDateField;
    N19: TMenuItem;
    N26: TMenuItem;
    MainQueryMR_PAY2_DATE: TStringField;
    MainQueryPAY2CURR: TStringField;
    MainQuerySTATE: TFloatField;
    ExpSQLPrep: TQuery;
    MainQuerySTATE_TXT: TStringField;
    InitMenu: TMenuItem;
    MainQueryAgPercent: TFloatField;
    N3: TMenuItem;
    ReservMenu: TMenuItem;
    MagazineLoss: TMenuItem;
    MagazineDocs: TMenuItem;
    MainQueryInsComp: TStringField;
    OpenDialogDBF: TOpenDialog;
    BatchMove: TBatchMove;
    QueryDistinct: TQuery;
    BatchMoveRest: TBatchMove;
    AvariaList: TMenuItem;
    MainQueryFULLDUPNAME: TStringField;
    N1: TMenuItem;
    Excel1: TMenuItem;
    MainQueryONE_1SU: TStringField;
    N14: TMenuItem;
    N15: TMenuItem;
    MainQueryRET1: TDateField;
    MainQueryRET2: TDateField;
    TaxiTbl: TTable;
    MainQueryTAXI: TStringField;
    N16: TMenuItem;
    RptTbl: TTable;
    RptTblS: TStringField;
    RptTblN: TFloatField;
    RptTblHi: TFloatField;
    RptTblLo: TFloatField;
    RepItog: TMenuItem;
    Alvena: TMenuItem;
    OpenDialogPx: TOpenDialog;
    _2C: TMenuItem;
    ImpParadox: TMenuItem;
    OpenDialogPx2: TOpenDialog;
    MainQueryK1: TFloatField;
    MainQueryK2: TFloatField;
    MainQueryNUMBERBODY: TStringField;
    InputSumm: TMenuItem;
    N20: TMenuItem;
    ToGlobo: TMenuItem;
    ExportAvariasGLOBO: TRxQuery;
    SaveDialog: TSaveDialog;
    N22: TMenuItem;
    ExportAvariasGLOBO2: TRxQuery;
    ToGloboBtn: TButton;
    mnuNoUnloadAvarias: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    AgentReport: TMenuItem;
    MainQueryBASECARCODE: TFloatField;
    mnuUnloadAllPolises: TMenuItem;
    MainQueryAGENTCODE: TStringField;
    N28: TMenuItem;
    OpenDialogNMB: TOpenDialog;
    To1C: TMenuItem;
    SaveDialog1C: TSaveDialog;
    Table1C: TTable;
    MainQueryTAXI_N: TFloatField;
    N29: TMenuItem;
    N30: TMenuItem;
    ExpectedPayMenu: TMenuItem;
    MonthlyReport: TMenuItem;
    StopInsInfo: TMenuItem;
    DivisionLabel: TLabel;
    MainQuerySMS: TStringField;
    MainQueryMOBTEL: TStringField;
    MainQueryHOMETEL: TStringField;
    MainQueryCONTACT: TStringField;
    K3Always: TMenuItem;
    MainQueryPAY2DRAFTNUMBER: TFloatField;
    MainQuery_1SU_CALC: TStringField;
    MainQueryKomplex: TQuery;
    StringField1: TStringField;
    FloatField1: TFloatField;
    StringField2: TStringField;
    DateField1: TDateField;
    DateField2: TDateField;
    StringField3: TStringField;
    DateField3: TDateField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    FloatField4: TFloatField;
    FloatField5: TFloatField;
    FloatField6: TFloatField;
    StringField4: TStringField;
    StringField5: TStringField;
    StringField6: TStringField;
    StringField7: TStringField;
    FloatField7: TFloatField;
    FloatField8: TFloatField;
    StringField8: TStringField;
    StringField9: TStringField;
    DateField4: TDateField;
    StringField10: TStringField;
    DateField5: TDateField;
    StringField11: TStringField;
    StringField12: TStringField;
    FloatField9: TFloatField;
    StringField13: TStringField;
    StringField14: TStringField;
    FloatField10: TFloatField;
    StringField15: TStringField;
    StringField16: TStringField;
    StringField17: TStringField;
    StringField18: TStringField;
    StringField19: TStringField;
    StringField20: TStringField;
    StringField21: TStringField;
    StringField22: TStringField;
    StringField23: TStringField;
    FloatField11: TFloatField;
    FloatField12: TFloatField;
    FloatField13: TFloatField;
    StringField24: TStringField;
    StringField25: TStringField;
    DateField6: TDateField;
    FloatField14: TFloatField;
    StringField26: TStringField;
    FloatField15: TFloatField;
    StringField27: TStringField;
    StringField28: TStringField;
    StringField29: TStringField;
    DateField7: TDateField;
    DateField8: TDateField;
    StringField30: TStringField;
    FloatField16: TFloatField;
    FloatField17: TFloatField;
    StringField31: TStringField;
    FloatField18: TFloatField;
    StringField32: TStringField;
    FloatField19: TFloatField;
    StringField33: TStringField;
    StringField34: TStringField;
    StringField35: TStringField;
    StringField36: TStringField;
    FloatField20: TFloatField;
    StringField37: TStringField;
    MainQueryKomplexTYP0: TStringField;
    MainQueryKomplexS0: TFloatField;
    MainQueryKomplexCURR0: TStringField;
    MainQueryKomplexD0: TDateField;
    MainQueryTYP0: TStringField;
    MainQueryS0: TFloatField;
    MainQueryCURR0: TStringField;
    MainQueryD0: TDateField;
    procedure N9Click(Sender: TObject);
    procedure N6Click(Sender: TObject);
    procedure PrintSetupMnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure FilterMnClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SortClick(Sender: TObject);
    procedure MainGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure ExplorerMnuClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure MenuItemBuroClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure KorrectMenuItemClick(Sender: TObject);
    procedure InitMenuClick(Sender: TObject);
    procedure ReservMenuClick(Sender: TObject);
    procedure AvForPolisesClick(Sender: TObject);
    procedure MagazineLossClick(Sender: TObject);
    procedure MagazineDocsClick(Sender: TObject);
    procedure MainQueryInsCompGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure AvariaListClick(Sender: TObject);
    procedure Excel1Click(Sender: TObject);
    procedure RunReportClick(Sender: TObject);
    procedure _2CClick(Sender: TObject);
    procedure RepItogClick(Sender: TObject);
    procedure AlvenaClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure OpenDialogPxCanClose(Sender: TObject; var CanClose: Boolean);
    procedure ImpParadoxClick(Sender: TObject);
    procedure InputSummClick(Sender: TObject);
    procedure N20Click(Sender: TObject);
    procedure ToGloboClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure mnuNoUnloadAvariasClick(Sender: TObject);
    procedure AgentReportClick(Sender: TObject);
    procedure MainQueryBASECARCODEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure mnuUnloadAllPolisesClick(Sender: TObject);
    procedure MainGridGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure To1CClick(Sender: TObject);
    procedure N30Click(Sender: TObject);
    procedure ExpectedPayMenuClick(Sender: TObject);
    procedure MonthlyReportClick(Sender: TObject);
    procedure StopInsInfoClick(Sender: TObject);
    procedure DivisionLabelClick(Sender: TObject);
    procedure MainQueryONE_1SUGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure K3AlwaysClick(Sender: TObject);
  private
    Mode : integer; //0 - обфзаловка 1 - private2
    W : HWnd;
    SecPayUpdated : boolean;

    lastAgCode : string;
    lastDivCode : integer;

    procedure AddIndex(name : string);
//    function CalculateSummForReport(var GoodPolises, GoodPolises2, BadPolises : integer; var SumBRB, SumEUR : double) : boolean;
//    function auxCalculateSummForReport(var GoodPolises, GoodPolises2, BadPolises : integer; var SumBRB, SumEUR : double; calcMsg, filterStr : string) : boolean;
//    procedure CalcMandPayment(F1, F2, {F3, is was valuta now BRB only}F4, filterStr : string; var Counter : integer; var SumBRB, SumEUR : double);
    procedure ImportAlvena(DBFile : string);
    procedure ImportDB(DBFile : string);
    function ExecQuery(SQLText : string; ASH : Variant; Cell : string; D1, D2 : TDateTime) : double;
    function DecodeK3(val : double) : string;
    procedure FreeWindow;

    { Private declarations }
  public
    { Public declarations }
    IsExit : boolean;
    sqlText : TStringList;
    MSWORD : Variant;
    MSEXCEL : VARIANT;

    MaxOwnerCode : integer;
    MaxSerTCCode : integer;
    MaxTCCode : integer;
    MaxAVCode : integer;

    OwnerCode_Holes : TList;
    SerTCCode_Holes : TList;
    TCCode_Holes : TList;
    AVCode_Holes : TList;

    ExpSQL : TDataSet;
    UseTable : boolean;

    procedure UpdateDupdata(filterStr : string);
//    procedure ExecTACKReport1;
    procedure OpenQuery;
    procedure FindSN(Seria : string; Number : integer);
    procedure FindDraft1Number(Number : integer);
    function  GetFilter(s : string) : string;
    function  GetFilter1 : string;
    function  GetFilter2 : string;
    function  GetOrder : string;
//    procedure CreateReportBuro(Month, Year : integer);
    function GetDivisionCode(S : string; N : integer) : string;
    function GetDivisionCode2(AgCode : string) : string;
    function ExpStr(s : string) : string;
    function GetOwnerType(N : integer) : string;
    function GetFizUrid(N : integer) : string;
    function GetDiscount2(N : double) : string;
    function GetCharacteristicAuto(s : string; Volume, Tonnage,Quantity, Power : double; IsResident : boolean) : double;
    function ExpTime(s : string) : integer;
    function ExpStateFunc(state : integer) : string;
    function ExpCurr(Curr : string) : string;
    function ExpZeroNmb(N : Integer) : string;
    function GetResident(N : integer) : string;
    function GetDecision(N : integer) : string;
    function GetPaySN(s : string) : string;
    function UpdateTabField(field, table, where : string; var Code : integer) : boolean;

    //CODE GENERATION
//    procedure InitExportBuro;
//    function GetMaxCodeFor(Table, Field : string) : integer;
//    function GetNextCodeOwner(Code : integer; S : CarOwnerDescription; var NeedUpdate : boolean) : string;
//    function GetNextCodeSerTC(Code : integer; S : SerCarDescription; var NeedUpdate : boolean) : string;
//    function GetNextCodeTC(Code : integer; S : CarDescription; var NeedUpdate : boolean) : string;
//    function GetNextCodeAv(Code : integer; var NeedUpdate : boolean) : string;
//    procedure SetMaxCodeFor(Seria : string; Number : integer; Table, Field1, Field2, Val1, Val2 : string);
  end;

var
  MainForm: TMainForm;
  lastCityKodeKey, lastCityKodeVal : string;
  lastAgCode, DecodeAgentName : string;
  lastAgName : string;
  lastAgID : string;
  AvKCount : integer = 0;
  AvK : array[0..50] of AvKType;
  lastResidentTable: string;
  lastIsResident: boolean;
  lastTableVariant: string = 'q';
  lastTableVariantCode: string = 'q';
  _1SU_GLOBO_CODES : array[0..12] of string;
  _1SU_SERIES : array[0..12] of string;
   bKomplexMode : Boolean;

const
  CodeTable : string = '0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM';
                       //backup '0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM';
                       //йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ


function DecodeAgentID(s : string) : string;
procedure WriteLn2(var F8, A : TextFile; s, s2 : string);
function DecodeCharact2(charact, s : string) : string;
function DecodeBodyShas(s : string) : string;
function DecodeInsSum(d : TDate) : string;
function DecodeSN(s : string) : string;
function DecodeFU16(s : string; pcnt : double) : string;
function DecodeOwnTypeForm(s : string; version : string) : string;
function DecodeFU12(s : string; version : string) : string;
function DecodeBlank(sect, s : string) : string;
function DecodePeriod(s : string; datefrom, dateto : TDateTIme; daysbased : boolean) : string;
function DecodeTimeX(s : string) : string;
function DecodeNB(s, paydoc : string) : string;
function KillZero(s : string) : string;
function DecodeVariant(tarifTable : string) : string;
function DecodeGlDate(dt : TDate) : string;
function DecodeGlDate2(dt, dt1, dt2, dt3 : TField) : string;
function DecodeR(s : string) : string;
function DecodeVladen(s : string; version : string) : string;
function DecodePayGraphic(s : string) : string;
function DecodePayType(s : string) : string;
function GetAvK(val : double; version : string) : string;
function GetK(val : double; version : string) : string;
function DecodeCurrency(s : string) : string;
function DecodeCharact(s, v, t, c, p : string) : string;
function DecodeNone(s, s1 : string) : string;
function DecodeAvState(s : string; DtCl, DtFree : TField) : string;
function DecodeOneDate(F1, F2, F3 : TField) : string;
procedure CheckAutoNmb(SN, AUTONUMBER : String; var Warn : TStringList);

implementation

uses Sort, FilterMandatory, MyPreview, ProgressDlg, MandatoryRpt, Export, ComObj,
  BeforeImp, Explorer, ReportDates, MainRpt6, Korrect,
  DateUtil, GetFormulaUnit, ShowInfoUnit, AboutUnit,
  KorrectPolandUnit, MandReservDlgUnit, AvarUnit, WaitFormUnit,
  GurnalUbitokUnit, ArariasFilterUnit, MagazineReport, Variants,
  GetTwoDatesUnit, rxstrutils, UnitImportPx, PolisNmbs, RepParamsFrm,
  RatesUnit, CurrRatesUnit, PcntFormUnit, RepParamsAgntFrm,
  Private2005, QClipbrd, DM, PayReportFilterUnit, ExpectedPayFilter,
  GetPeriodOnly, DataMod;

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

function GetTableName : string;
begin
        if bKomplexMode then GetTableName := 'komplex'
        else GetTableName := 'mandator'
end;

procedure TMainForm.N9Click(Sender: TObject);
begin
    About.Title.Caption := 'Отчёт'#13'по обязательному страхованию'#13'транспортных средств';
    About.ShowModal
end;

procedure TMainForm.N6Click(Sender: TObject);
begin
     Close
end;

procedure TMainForm.PrintSetupMnClick(Sender: TObject);
begin
     PrintDialog.Execute
end;

function TMainForm.GetOrder : string;
begin
     Result := SortForm.GetOrder('');
end;

function TMainForm.GetFilter(s : string) : string;
begin
     Result := MandatoryFilterDlg.GetFilter(s)
end;

function TMainForm.GetFilter1 : string;
begin
     Result := MandatoryFilterDlg.GetFilter1
end;

function TMainForm.GetFilter2 : string;
begin
     Result := MandatoryFilterDlg.GetFilter2
end;

{
procedure TMainForm.FindLikeName(s : string);
var
     F : TextFile;
     i : integer;
     fn : string;
begin
     fn := ParamStr(0);
     for i := Length(fn) downto 1 do
       if fn[i] = '\' then break;

     SYSTEM.Assign(F, Copy(fn, 1, i) + 'findres.');
     SYSTEM.ReWrite(F);

     try
        with WorkSQL, WorkSQL.SQL do begin
             Clear;
             Add('SELECT SERIA, NUMBER, NAME FROM MANDATOR');
             Add('WHERE NAME LIKE ''%' + s + '%''');
             Open;
             while not EOF do begin
                   s := FieldByName('SERIA').AsString + '/' +
                        FieldByName('NUMBER').AsString + ' ' +
                        FieldByName('NAME').AsString;

                   WriteLn(F, s);
                   Next;
             end;
        end;
     except
     end;

     SYSTEM.Close(F);

     PostMessage(FindWindow('OWLWindow31', nil), wm_user + 274, 0, 0);
end;
}
procedure TMainForm.FormCreate(Sender: TObject);
var
     R : TRect;
     buff, sPath : array[0..128] of char;
     Path, s : string;
     I, Len : integer;
     INI : TIniFile;
     divpos : integer;
begin
        if bKomplexMode then begin
                MainQuery.SQL := MainQueryKomplex.SQL;
                Caption := 'Комплексное страхование транспортных средств';
        end;

        InitMenu.Visible := not bKomplexMode;
        AvForPolises.Visible := not bKomplexMode;
        AvariaList.Visible := not bKomplexMode;
        Alvena.Visible := not bKomplexMode;
        ImpParadox.Visible := not bKomplexMode;
        MenuItemBuro.Visible := not bKomplexMode;
        N22.Visible := not bKomplexMode;
        To1C.Visible := not bKomplexMode;
        mnuNoUnloadAvarias.Visible := not bKomplexMode;
        MagazineLoss.Visible := not bKomplexMode;

     DMM.HandBookSQL.DatabaseName := DB.Name;

     INI := TIniFile.Create(BLANK_INI);
     DivisionLabel.Caption := 'ПОДРАЗДЕЛЕНИЕ: ' + INI.ReadString('DIVISION', 'ID', '?');
      mnuUnloadAllPolises.Checked := INI.ReadBool('MANDATORY', 'UnloadAllPolises', true);

     for i := 0 to 12 do
     begin
        s := INI.ReadString(MAND_SECT, '1SU_' + IntToStr(i), '');
        if s = '' then break;
        divpos := Pos(',', s);
        if divpos = 0 then break;

        _1SU_SERIES[i] := Trim(Copy(s, 1, divpos - 1));
        s := Copy(s, divpos + 1, Length(s));

        divpos := Pos(',', s);
        if divpos <> 0 then s := Copy(s, 1, divpos - 1);

        _1SU_GLOBO_CODES[i] := Trim(s);
     end;

     INI.Free;

     GetProfileString('Paradox Engine', 'NetNamePath', '', sPath, 128);
     Path := sPAth;
     DB.Open;
     Len := min(Length(DB.Directory), Length(Path));
     if (UpperCase(Copy(DB.Directory, 0, Len)) <> UpperCase(Copy(Path, 0, Len))) or (abs(Length(DB.Directory) - Length(Path)) > 1) then
        MessageDlg('Несовпадение путей для прогамм БЛАНК и ОТЧЕТ '#13 + Path + #13 + DB.Directory, mtInformation, [mbOk], 0);

     SecPayUpdated := false;
     MenuCompany.Caption := COMPAMYNAME;
     OwnerCode_Holes := TList.Create;
     SerTCCode_Holes := TList.Create;
     TCCode_Holes := TList.Create;
     AVCode_Holes := TList.Create;

     KorrectMenuItem.Visible := (GetPrivateProfileInt('MANDATORY', 'CAN_KORRECT_DATA', 0, StrPCopy(buff, BLANK_INI)) = 1) and (not bKomplexMode);

     ExpSQL := nil;

     MaxOwnerCode := -1;

     ProgressBar.Parent := StatusBar;

          Application.Title := 'Отчёт по обязятельному страхованию';
          //ExpSQLText := TStringList.Create;
          //ExpSQLText.Assign(ExportSQL.SQL);
          sqlText := TStringList.Create;
          sqlText.Assign(MainQuery.SQL);
          MainGrid.Align := alClient;
          MainGrid.Visible := true;
          Menu := MainMenu;
          //Button4.Visible := true;

     AgentTbl.Open;

     {
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
     }
     WindowState := wsMaximized;
     IsExit := false;

     if IsExit then begin
        PostQuitMessage(0);
        exit;
     end;
end;

procedure TMainForm.AddIndex(name : string);
var
     MANDATOR : TTable;
     FoundIdx : Boolean;
     I, Len : integer;
begin
    try
        FoundIdx := false;
        MANDATOR := TTable.Create(self);
        MANDATOR.DatabaseName := DB.DatabaseName;
        MANDATOR.TableName := GetTableName();
        MANDATOR.Exclusive := True;
        MANDATOR.Active := False;
        MANDATOR.IndexDefs.Update;
        for I := 0 to MANDATOR.IndexDefs.Count - 1 do
          if MANDATOR.IndexDefs.Items[I].Fields = name then
            FoundIdx := True;
        if not FoundIdx then
            MANDATOR.AddIndex('PREVIDX', 'PrevSeria;PrevNumber', []);
    except
    end;
    MANDATOR.Free;
end;

procedure TMainForm.OpenQuery;
begin
         MainQuery.Close;
         MainQuery.SQL.Assign(sqlText);
         if Length(GetFilter('')) <> 0 then begin
            MainQuery.SQL.Add('WHERE');
            MainQuery.SQL.Add(GetFilter(''));
         end;
         if Length(GetOrder) > 0 then begin
            MainQuery.SQL.Add('ORDER BY');
            MainQuery.SQL.Add(GetOrder);
         end;
         try
           Screen.Cursor := crHourGlass;
           MainQuery.Open;

           StateText.Caption := 'Записей ' + IntToStr(MainQuery.RecordCount);
         except
           on e : Exception do begin
              MessageDlg('Ошибка при выполнении запроса.'#13 + e.Message, mtInformation, [mbOK], 0);
           end
         end;
         Screen.Cursor := crDefault;
end;

procedure TMainForm.MainQueryCalcFields(DataSet: TDataSet);
begin
       if (MainQueryONE_1SU.AsString = 'N') or (MainQueryONE_1SU.AsString = '') then
       else begin
           if MainQueryONE_1SU.AsString = 'Y' then MainQuery_1SU_CALC.AsString := 'ГС'
           else
           if MainQueryONE_1SU.AsString = '0' then MainQuery_1SU_CALC.AsString := 'КВ'
           else MainQuery_1SU_CALC.AsString := _1SU_SERIES[ord(MainQueryONE_1SU.AsString[1]) - ord('A')]
       end;

     if MainQueryFROMDATE.AsString <> '' then
        MainQueryPERIOD_FT.AsString := MainQueryFROMDATE.AsString + ' - ' + MainQueryTODATE.AsString;

     MainQueryPAYWITHCURR.AsString := MainQueryPAY1.AsString + ' ' + MainQueryPAY1CURR.AsString;
     MainQueryDISCOUNTTEXT.AsString := '';
     if MainQueryDISCOUNT.AsFloat > 0.009 then
        MainQueryDISCOUNTTEXT.AsString := MainQueryDISCOUNT.AsString;

     MainQueryFee2Symbol.AsString := 'Ц';
     if MainQueryIsFee2.AsInteger = 0 then
        MainQueryFee2Symbol.AsString := '-';//'ґ';

     MainQueryPay2Symbol.AsString := 'Ц';
     if MainQueryIsPay2.AsInteger = 0 then
        MainQueryPay2Symbol.AsString := '-';//'ґ';

     MainQueryOWNEROFPOLIS.AsString := MainQueryNAME.AsString;
     if MainQueryPREVSERIA.AsString <> '' then
        MainQueryOWNEROFPOLIS.AsString := MainQueryOWNEROFPOLIS.AsString + ' ' +  MainQueryPREVSERIA.AsString + '/' + MainQueryPREVNUMBER.AsString;

     if MainQueryPREVSERIA.AsString <> '' then
        MainQueryChange.AsString := MainQueryPREVNUMBER.AsString + '/' + MainQueryPREVSERIA.AsString;

     //For month report
     if MainQueryISPAY2.AsInteger > 0 then begin //2 этапа оплачено
         MainQueryMR_PAY1_SUMVAL.AsString := MainQueryFEE2.AsString;
         MainQueryMR_PAY1_RATE.AsString := MainQueryRATE.AsString;
         MainQueryMR_PAY1_SUMBRB.AsString := MainQueryPAY1.AsString;
         if MainQueryPAY1CURR.AsString <> '' then
           if MainQueryPAY1CURR.AsString <> 'BRB' then
             MainQueryMR_PAY1_SUMBRB.AsString := MainQueryPAY1.AsString + ' ' + MainQueryPAY1CURR.AsString;
         MainQueryMR_PAY2_SUMVAL.AsString := MainQueryFEE2.AsString;
         MainQueryMR_PAY2_RATE.AsString := MainQueryPAY2RATE.AsString;
         MainQueryMR_PAY2_SUMBRB.AsString := MainQueryPay2Summa.AsString;
         if MainQueryPAY2CURR.AsString <> '' then
           if MainQueryPAY2CURR.AsString <> 'BRB' then
             MainQueryMR_PAY2_SUMBRB.AsString := MainQueryPay2Summa.AsString + ' ' + MainQueryPAY2CURR.AsString;
         MainQueryMR_PAY2_DATE.AsString := MainQueryPay2Date.AsString;
     end
     else begin
         MainQueryMR_PAY1_SUMVAL.AsString := MainQueryTARIF.AsString;
         MainQueryMR_PAY1_RATE.AsString := MainQueryRATE.AsString;
         MainQueryMR_PAY1_SUMBRB.AsString := MainQueryPAY1.AsString;
         if MainQueryPAY1CURR.AsString <> '' then
           if MainQueryPAY1CURR.AsString <> 'BRB' then
             MainQueryMR_PAY1_SUMBRB.AsString := MainQueryPAY1.AsString + ' ' + MainQueryPAY1CURR.AsString;
         MainQueryMR_PAY2_SUMVAL.AsString := '';
         MainQueryMR_PAY2_RATE.AsString := '';
         MainQueryMR_PAY2_SUMBRB.AsString := '';
         MainQueryMR_PAY2_DATE.AsString := '';

         if MainQuerySTATE.AsInteger = 0 then begin //Нормальный полис
             if MainQueryISFEE2.AsInteger = 0 then //Не нужен взнос
                 MainQueryMR_PAY2_DATE.AsString := '1 ЭТАП'
             else
             if MainQueryFEE2DATE.AsDateTime < Date then
                 MainQueryMR_PAY2_DATE.AsString := 'Просрочен';
         end;
     end;

     if MainQueryPREVSERIA.AsString <> '' then begin //Для дубликатов всё сбрасываем
         MainQueryFULLDUPNAME.AsString := '''Взамен ' + MainQueryPREVSERIA.AsString + '/' + MainQueryPREVNUMBER.AsString;
         if MainQuerySTATE.AsInteger <> 3 then begin //утерян
             MainQueryFULLDUPNAME.AsString := MainQueryFULLDUPNAME.AsString + ' 1й взнос ' +
                                              MainQueryPAY1.AsString + ' ' + MainQueryPAY1CURR.AsString + '; ';

             if MainQueryPAY2SUMMA.AsFloat > 0.001 then
                 MainQueryFULLDUPNAME.AsString := MainQueryFULLDUPNAME.AsString + ' 2й взнос ' +
                                                  MainQueryPAY2SUMMA.AsString + ' ' + MainQueryPAY2CURR.AsString;
         end;

         MainQueryFULLDUPNAME.AsString := MainQueryFULLDUPNAME.AsString + ' ''';
         MainQueryMR_PAY2_SUMVAL.AsString := '';
         MainQueryMR_PAY2_RATE.AsString := '';
         MainQueryMR_PAY2_SUMBRB.AsString := '';
         MainQueryMR_PAY2_DATE.AsString := '';
     end;

     MainQuerySTATE_TXT.AsString := '?';
     case MainQuerySTATE.AsInteger of
        0 : MainQuerySTATE_TXT.AsString := 'НОРМАЛЬНЫЙ';
        1 : MainQuerySTATE_TXT.AsString := 'ИСПОРЧЕН';
        2 : MainQuerySTATE_TXT.AsString := 'ДОСРОЧНО ЗАВЕРШЁН';
        3 : MainQuerySTATE_TXT.AsString := 'УТЕРЯН';
        4 : MainQuerySTATE_TXT.AsString := 'СРОК ДЕЙСТВИЯ ЗАКОНЧИЛСЯ';
        5 : MainQuerySTATE_TXT.AsString := 'ЗАВЕРШЁН ПО НЕУПЛАТЕ';
     end;

     if bKomplexMode then
     begin
             MainQueryTAXI.AsString := MainQuery.FieldByName('S0').AsString + ' ' + MainQuery.FieldByName('Curr0').AsString + ', ' + MainQuery.FieldByName('D0').AsString;
             if MainQuery.FieldByName('TYP0').AsString = '' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' такси';
             if MainQuery.FieldByName('TYP0').AsString = '1' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' другое';
             if MainQuery.FieldByName('TYP0').AsString = '2' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' смена типа авто';
             if MainQuery.FieldByName('TYP0').AsString = '3' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' аварийность';
     end
     else
     begin
             if not TaxiTbl.Active then
                 TaxiTbl.Open;

         if TaxiTbl.Locate('SER;N', vararrayof([MainQuerySERIA.AsString, MainQueryNUMBER.AsInteger]), []) then
         begin
             MainQueryTAXI.AsString := TaxiTbl.FieldByName('SUMMA').AsString + ' ' + TaxiTbl.FieldByName('Curr').AsString + ', ' + TaxiTbl.FieldByName('D').AsString;
             if TaxiTbl.FieldByName('TYPE').AsString = '' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' такси';
             if TaxiTbl.FieldByName('TYPE').AsString = '1' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' другое';
             if TaxiTbl.FieldByName('TYPE').AsString = '2' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' смена типа авто';
             if TaxiTbl.FieldByName('TYPE').AsString = '3' then MainQueryTAXI.AsString := MainQueryTAXI.AsString + ' аварийность';
         end
     end
end;

procedure TMainForm.FilterMnClick(Sender: TObject);
var
     s : string;
begin
     s := GetFilter('');
     MandatoryFilterDlg.ShowModal;
     if (s <> GetFilter('')) OR not MainQuery.Active then
        OpenQuery
end;

//Записывает данные последнего дубликата в первый утерянный в цепочке
procedure TMainForm.UpdateDupdata(filterStr : string);
var
    strList : TStringList;
    //strListDate : TStringList;
    strListPayCond : TStringList;
    PAY2DATE,RET2 : TDateTime;
    PAY2RATE : double;
    PAY2SUMMA : double;
    PAY2CURR : string;
    PAY2TYPE : integer;
    PAY2DRAFTNUMBER : double;

    SERIA : string;
    NUMBER : double;
    i : integer;
    strNoPay : string;
    DupCount, LostCount : integer;
    MANDATOR : TTable;
    FoundIdx : Boolean;
label
    stop;
begin
    MANDATOR := TTable.Create(self);
    MANDATOR.DatabaseName := DB.DatabaseName;
    MANDATOR.TableName := GetTableName();
    MANDATOR.Exclusive := True;
    MANDATOR.Active := False;
    MANDATOR.IndexDefs.Update;
    for I := 0 to MANDATOR.IndexDefs.Count - 1 do
      if MANDATOR.IndexDefs.Items[I].Fields = 'PrevSeria;PrevNumber' then
        FoundIdx := True;
    if not FoundIdx then
        MANDATOR.AddIndex('PREVIDX', 'PrevSeria;PrevNumber', []);
    MANDATOR.Free;

    filterstr := 'TODATE > ''' + DateToStr(Date - 100) + '''';
    if (GetAsyncKeyState(vk_shift) and $8000) <> 0 then exit;
    if SecPayUpdated then exit;
          StatusBar.Panels[0].Text := 'Синхронизация дубликатов...';
          StatusBar.Update;
          strList := TStringList.Create;
          strListPayCond := TStringList.Create;
          Screen.Cursor := crHourGlass;

          try                                      
              with MainForm.WorkSQL, MainForm.WorkSQL.SQL do begin
                   //Самые первые утерянные со вторым взносом и без 2го платежа
                   Clear;
                   Add('SELECT SERIA,NUMBER,AGENTCODE,ISFEE2 FROM ' + GetTableName());
                   Add('WHERE PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND STATE=3 AND ISFEE2<>0 AND ISPAY2=0');
                   if Length(filterStr) > 0 then begin
                      Add('AND');
                      Add(filterStr);
                   end;

                   Open;
                   while not EOF do begin
                         strList.Add(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString);
                         strListPayCond.Add(FieldByName('ISFEE2').AsString);
                         Next;
                   end;
                   Close;

                   if strList.Count > 0 then begin
                       //Проверим наличие утерянных
                       Clear;
                       Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PrevSeria <> '''' AND ISFEE2<>0');
                       Open;
                       DupCount := FieldByName('CNT').AsInteger;
                       Close;

                       Clear;
                       Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE STATE=3 AND ISFEE2<>0');
                       Open;
                       LostCount := FieldByName('CNT').AsInteger;
                       Close;

                       if DupCount <> LostCount then begin
                            MessageDlg('В БД несовпадает количество утерянных полисов и дубликатов с оплатой в ДВА этапа. Из за этого могут не совпадать суммы в отчётах т.к 2я оплата регистрируется в дубликате', mtInformation, [mbOk], 0);
                       end;
                   end;

                   for i := 0 to strList.Count - 1 do begin
                       SERIA := Copy(strList[i], 1, Pos('/', strList[i]) - 1);
                       NUMBER := StrToFloat(Copy(strList[i], Pos('/', strList[i]) + 1, 100));

                       StatusBar.Panels[0].Text := 'Синхронизация дубликатов (' + IntToStr(strList.Count - i) + ')...';
                       StatusBar.Update;

                       while true do begin
                           Clear;
                           //Ищем дубликат полиса
                           Add('SELECT RET2,STATE,SERIA,NUMBER,ISPAY2,PAY2DATE,PAY2RATE,PAY2SUMMA,PAY2CURR,PAY2TYPE,PAY2DRAFTNUMBER,AGENTCODE,FEE2DATE,ISFEE2 FROM ' + GetTableName());
                           Add('WHERE PREVSERIA=:PS AND PREVNUMBER=:PN');
                           ParamByName('PS').AsString := Copy(strList[i], 1, Pos('/', strList[i]) - 1);
                           ParamByName('PN').AsFloat := StrToInt(Copy(strList[i], Pos('/', strList[i]) + 1, 100));
                           Open;
                           if EOF then begin
                               MessageDlg('Не найден дубликат для полиса ' + strList[i] + '. Невозможно определить наличие 2й оплаты', mtInformation, [mbOK], 0);
                               WorkSQL.Close;
                               goto stop;
                           end;

                           if FieldByName('ISFEE2').AsString <> strListPayCond[i] then begin
                               MessageDlg('Обнаружена ошибка. Утерянный полис ' + strList[i] + ' и его дубликат имеют различные условия оплаты (в один и ДВА этапа).', mtInformation, [mbOK], 0);
                               WorkSQL.Close;
                               goto stop;
                           end;

                           if FieldByName('STATE').AsInteger = 3 then begin //и этот потеряли
                               strList[i] := (FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString);
                               WorkSQL.Close;
                               goto stop;
                           end
                           else begin
                               PAY2DATE := FieldByName('PAY2DATE').AsDateTime;
                               PAY2RATE := FieldByName('PAY2RATE').AsFloat;
                               PAY2SUMMA := FieldByName('PAY2SUMMA').AsFloat;
                               PAY2CURR := FieldByName('PAY2CURR').AsString;
                               PAY2TYPE := FieldByName('PAY2TYPE').AsInteger;
                               RET2 := FieldByName('RET2').AsDateTime;
                               PAY2DRAFTNUMBER := FieldByName('PAY2DRAFTNUMBER').AsFloat;
                               if FieldByName('ISPAY2').AsInteger = 1 then begin //Последний дубликат оплатили
                                  //check first polis
                                  Clear;
                                  Add('UPDATE ' + GetTableName() + ' SET RET2=:RET2,PAY2DATE=:PAY2DATE, PAY2RATE=:PAY2RATE, PAY2SUMMA=:PAY2SUMMA, PAY2CURR=:PAY2CURR, PAY2TYPE=:PAY2TYPE, PAY2DRAFTNUMBER=:PAY2DRAFTNUMBER,ISPAY2=1');
                                  Add('WHERE SERIA=:S AND NUMBER=:N');
                                  ParamByName('PAY2DATE').AsDateTime := PAY2DATE;
                                  ParamByName('RET2').AsDateTime := RET2;
                                  ParamByName('PAY2RATE').AsFloat := PAY2RATE;
                                  ParamByName('PAY2SUMMA').AsFloat := PAY2SUMMA;
                                  ParamByName('PAY2CURR').AsString := PAY2CURR;
                                  ParamByName('PAY2TYPE').AsFloat := PAY2TYPE;
                                  ParamByName('PAY2DRAFTNUMBER').AsFloat := PAY2DRAFTNUMBER;
                                  ParamByName('S').AsString := SERIA;
                                  ParamByName('N').AsFloat := NUMBER;
                                  ExecSQL;
                               end else begin
                                  if FieldByName('FEE2DATE').AsDateTime < Date then begin
                                      if (FieldByName('STATE').AsInteger in [0, 4]) then //нормальный или СРОК ДЕЙСТВИЯ ЗАКОНЧИЛСЯ
                                          strNoPay := strNoPay + #13 + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                                  end;
                               end;
                               break;
                           end;

                           Next;
                           if not EOF then begin
                               MessageDlg('Найдено 2!!! дубликата для полиса ' + strList[i], mtInformation, [mbOK], 0);
                               WorkSQL.Close;
                               break;
                           end;
                           WorkSQL.Close;
                           break;
                       end;
                       stop:;
                   end;
              end;
          except
              on E : Exception do begin
                 MessageDlg(E.Message, mtInformation, [mbOK], 0);
              end;
          end;

          strList.Free;
          strListPayCond.Free;
          Screen.Cursor := crDefault;
          StatusBar.Panels[0].Text := '';

          if strNoPay <> '' then begin
              Clipboard.AsText := strNoPay;
              MessageDlg('Просроченные полисы. Оплата должны была поступить но они не расторгнуты. См. номера в буфере обмена.'#13 + strNoPay, mtInformation, [mbOk], 0);
          end;

    SecPayUpdated := true;
end;
(*
function TMainForm.auxCalculateSummForReport(var GoodPolises, GoodPolises2, BadPolises : integer; var SumBRB, SumEUR : double; calcMsg, filterStr : string) : boolean;
begin
    GoodPolises := 0;
    GoodPolises2 := 0;
    BadPolises := 0;
    SumBRB := 0;
    SumEUR := 0;
    Result := true;
    try
          //STATE = 3 утерян
          //Расчёт сумм производится на основании базового полиса (для утерянных)
          //Отсюда проблема для 2-го платежа. Он регистрируется или в базовом
          //или в утерянном (в зависимости от даты потери)
          //или в промежуточном (что ещё хуже)
          //первый всегда есть в первом
          //второй всегда есть в последнем

          //Доп ф-ия переносит информацию о втором платеже из последнего дубликата
          //в первый утерянный полис
          //пОИСК ИДЁТ ПО ФИЛЬТРУ (ЗАПОЛНИТЬ ВСЕ УТЕРЯННЫЕ Где должен быть 2й платёж)

          UpdateDupdata(filterStr);

          ////

          Screen.Cursor := crHourGlass;
          Progress.Show;
          Progress.ProcessLabel.Caption := calcMsg;//'Засчёт сумм. Ждите...';
          Progress.Update;
          // Calculate summ
          with MainForm.WorkSQL, MainForm.WorkSQL.SQL do begin
               // Выданные
               Clear;
               Add('SELECT COUNT(* ) AS CNT FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria=''''))'); //Не учитавать дубликат
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 10;
               GoodPolises := FieldByName('CNT').AsInteger;
               Close;

               // Чужие
               Clear;
               Add('SELECT COUNT(* ) AS CNT FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NULL AND PAY2DATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria=''''))'); //Не учитавать дубликат
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 10;
               GoodPolises2 := FieldByName('CNT').AsInteger;
               Close;

               // Испорченные
               Clear;
               Add('SELECT COUNT(* ) AS CNT FROM MANDATOR WHERE (PAYDATE IS NULL)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 20;
               BadPolises := FieldByName('CNT').AsInteger;
               Close;
               // Рубли
                // Part 1  Нормальный полис и сумма в рублях
               Clear;
               Add('SELECT SUM(PAY1) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND PAY1CURR=''BRB'')');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 30;
               SumBRB := FieldByName('SUMMA').AsFloat;
               Close;
               // Part 2  Нормальный полис и сумма не в рублях (курс)
               Clear;
               Add('SELECT SUM(PAY1*RATE) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND PAY1CURR<>''BRB'' AND PAY1CURR IS NOT NULL)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 40;
               SumBRB := SumBRB + FieldByName('SUMMA').AsFloat;
               Close;
                // Part 3 pay 2 exists  Нормальный полис и есть 2й взнос и сумма в рублях
               Clear;
               Add('SELECT SUM(PAY2SUMMA) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAY2DATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND ISFEE2<>0 AND PAY2CURR=''BRB'')');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 50;
               SumBRB := SumBRB + FieldByName('SUMMA').AsFloat;
               Close;
                // Part 4 pay 2 exists  Нормальный полис и есть 2й взнос и сумма не в рублях (курс)
               Clear;
               Add('SELECT SUM(PAY2SUMMA*PAY2RATE) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAY2DATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND ISFEE2<>0 AND PAY2CURR<>''BRB'' AND PAY2CURR IS NOT NULL)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 60;
               SumBRB := SumBRB + FieldByName('SUMMA').AsFloat;
               Close;

               // ЕВРО ЕВРО ЕВРО ЕВРО ЕВРО
                // Part 1a - евро без скидок 1й платёж и единственный
               Clear;
               Add('SELECT SUM(TARIF) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND DISCOUNT=0 AND ISFEE2=0)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 70;
               SumEUR := FieldByName('SUMMA').AsFloat;
               Close;
                // Part 1b - евро со скидкой 1й платёж и единственный
               Clear;
               Add('SELECT SUM(TARIF*DISCOUNT/100) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND DISCOUNT<>0 AND ISFEE2=0)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 80;
               SumEUR := SumEUR + FieldByName('SUMMA').AsFloat;
               Close;
                // Part 2a - евро без скидок 1й платёж и не единственный
               Clear;
               Add('SELECT SUM(TARIF)/2 AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND DISCOUNT=0 AND ISFEE2<>0)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 90;
               SumEUR := SumEUR + FieldByName('SUMMA').AsFloat;
               Close;
                // Part 2b - евро со скидкой 1й платёж и не единственный
               Clear;
               Add('SELECT SUM(TARIF*DISCOUNT/100)/2 AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAYDATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND DISCOUNT<>0 AND ISFEE2<>0)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 95;
               SumEUR := SumEUR + FieldByName('SUMMA').AsFloat;
               Close;

                // Part 3 - евро 2й взнос
               Clear;
               Add('SELECT SUM(FEE2) AS SUMMA FROM MANDATOR WHERE BASECARCODE=101 AND (PAY2DATE IS NOT NULL AND (PrevSeria IS NULL OR PrevSeria = '''') AND ISFEE2<>0 AND ISPAY2 <> 0)');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Open;
               Progress.ProgressBar.Position := 100;
               SumEUR := SumEUR + FieldByName('SUMMA').AsFloat;
               Close;
               //РАЗБИВКА ПО ВАЛЮТАМ
          end;
          Progress.Hide;
          Screen.Cursor := crDefault;
          // End calculate summ
    except
          Result := false;
    end;
    Screen.Cursor := crDefault;
end;

procedure TMainForm.ExecTACKReport1;
var
    GoodPolises, GoodPolises2, BadPolises : integer;
    SumBRB, SumEUR : double;
    bigInt : int64;
begin
    Explore.Visible := false;
    try
         if CalculateSummForReport(GoodPolises, GoodPolises2, BadPolises, SumBRB, SumEUR) then begin
              MandatoryReport.PolicesCount.Caption := IntToStr(GoodPolises);
              MandatoryReport.Polices2Count.Caption := IntToStr(GoodPolises2);
              MandatoryReport.BadCount.Caption := IntToStr(BadPolises);
              bigInt := ROUND(SumBRB * 100);
              MandatoryReport.RubSumma.Caption := FloatToStr(bigInt / 100);
              bigInt := ROUND(SumEUR * 100);
              MandatoryReport.EuroSumma.Caption := FloatToStr(bigInt / 100);

              Application.CreateForm(TQRMyPreview, QRMyPreview);
              QRMyPreview.Caption := self.Caption;
              MainGrid.DataSource := nil;
              MainQuery.First;
              MandatoryReport.Preview;
          end;
    except
    end;
    Progress.Hide;
    MainGrid.DataSource := MainDataSource;
end;

*)
procedure TMainForm.TimerTimer(Sender: TObject);
begin
     if (ParamCount > 0) AND (ParamStr(1) = 'ALVENA') then begin
        Timer.Enabled := false;
        AlvenaClick(nil);
        PostQuitMessage(0);
     end;

     if (ParamCount > 0) AND (ParamStr(1) = 'IMPBASO') then begin
        Timer.Enabled := false;
        ImpParadoxClick(nil);
        PostQuitMessage(0);
     end;

     //Screen.Cursor := crDefault;
     //Timer.Enabled := false;
     //FilterMnClick(nil);
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     MainQuery.Close;
     OwnerCode_Holes.Free;
     SerTCCode_Holes.Free;
     TCCode_Holes.Free;
     AVCode_Holes.Free;
end;

procedure TMainForm.SortClick(Sender: TObject);
begin
     SortForm.m_Query := MainQuery;
     if SortForm.ShowModal = mrOk then
         OpenQuery
end;

procedure TMainForm.MainGridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
     if Field.IsNull then exit;
        
     if Field = MainQueryISFEE2 then
        if MainQueryISFEE2.AsInteger = 1 then begin
            Background := clBlue;
            AFont.Color := clWhite
        end;

     if Field = MainQueryISPAY2 then
        if (MainQueryISFEE2.AsInteger = 1) AND (MainQueryISPAY2.AsInteger = 0) AND (MainQueryFEE2DATE.AsDateTime <= Date) then begin
            if MainQuerySTATE.AsInteger in [0,4] then begin
                Background := clRed;
                AFont.Color := clWhite
            end;
        end;

     if Field = MainQueryRET1 then Background := $00E3EBB1;
     if Field = MainQueryRET2 then Background := $00E3EBB1;

end;

procedure TMainForm.FindSN(Seria : string; Number : integer);
begin
     Screen.Cursor := crHourGlass;
     if not MainQuery.Locate('Seria;Number', vararrayof([Seria, Number]), []) then
        MessageDlg('Полис ' + Seria + '/' + IntToStr(Number) + ' не найден.', mtInformation, [mbOK], 0);
     Screen.Cursor := crDefault;
end;

procedure TMAINForm.FindDraft1Number(Number : integer);
begin
     Screen.Cursor := crHourGlass;
     if not MainQuery.Locate('DraftNumber', Number, []) then
        MessageDlg('Платёжка ' + IntToStr(Number) + ' не найдена.', mtInformation, [mbOK], 0);
     Screen.Cursor := crDefault;
end;

procedure TMainForm.ExplorerMnuClick(Sender: TObject);
begin
     Explore.Visible := not Explore.Visible;
end;

procedure TMainForm.FormShow(Sender: TObject);
begin
     if Mode = 0 then begin
          Application.CreateForm(TMandatoryFilterDlg, MandatoryFilterDlg);
          Application.CreateForm(TExplore, Explore);
//          Application.CreateForm(TAvarias, Avarias);
          Application.CreateForm(TGetDates, GetDates);
          Application.CreateForm(TBeforeImport, BeforeImport);
          Application.CreateForm(TExportForm, ExportForm);
          Application.CreateForm(TMandatoryReport, MandatoryReport);
          //Avarias.Visible := true;
     end;
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
        //retCode := CodeTable[(Value MOD SystemBase) + 1] + retCode;

        Value := Value + Multer * (Pos(str[Length(str)], CodeTable) - 1);

        Multer := Multer * SystemBase;
        str := Copy(str, 1, Length(str) - 1);
    end;

    DecodeMsgId := IntToStr(Value);
end;

procedure TMainForm.MenuItemBuroClick(Sender: TObject);
begin
    lastAgCode := '????';
    lastDivCode := -1;

     if MainReport6.ShowModal = mrCancel then
        exit;
{
     StatusBar.Panels[0].Text := 'Индексирование...';
     StatusBar.Update;
     TmpTable.TableName := 'MANDATOR';
     try
       Screen.Cursor := crHourGlass;
       TmpTable.Open;
       try
          TmpTable.DeleteIndex('NAMEIDX');
       except
       end;
       try
          TmpTable.DeleteIndex('MARKAIDX');
       except
       end;
       try
          TmpTable.DeleteIndex('AUTONUMBERIDX');
       except
       end;
       TmpTable.AddIndex('NewIndex', 'Name', []);
       TmpTable.AddIndex('MARKAIDX', 'MARKA', []);
       TmpTable.AddIndex('AUTONUMBERIDX', 'AUTONUMBER', [IxNonMaintained]);
     except
        on E : Exception do begin
            TmpTable.Close;
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
            Screen.Cursor := crDefault;
            exit;
        end;
     end;
     Screen.Cursor := crDefault;
     TmpTable.Close;

     exit;
}
     UseTable := MainReport6.UseTab.Checked;
//     InitExportBuro();
//     CreateReportBuro(MainReport6.Monthes.ItemIndex + 1, ROUND(MainReport6.Year.Value));
end;

{
procedure TMainForm.CreateReportBuro(Month, Year : integer);
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

   SNList : TStringList;

   i, find : integer;

   COMPANY_CODE : string;
   DIVISION_CODE : string;
   SERIA : string[10];
   NUMBER : integer;
   CharAuto : double;

   AOwn : array of AutoOwnerStruct;
   OwnCounter : integer;
   CURR_CAR_CODE : string[10];
   CarOwnDescr : CarOwnerDescription;
   SerCarDesrc : SerCarDescription;
   CarDescr : CarDescription;

   NeedUpdate : boolean;

   StartTime, CurrTime : integer;
   TimeStr : string;

   ExceptSN : string;

   errCollectStr : string;
begin
   ExpState := true;
   errCollectStr := '';
   INI := TIniFile.Create(BLANK_INI);
   TemplateDir := INI.ReadString('MANDATORY', 'BuroDirectory', '');
   COMPANY_CODE := INI.ReadString('MANDATORY', 'BuroCompany', '');
   INI.Free;

   if (COMPANY_CODE < '1') OR (COMPANY_CODE > '9') then begin
      MessageDlg('Не определен код компании, необходимый в БЮРО', mtInformation, [mbOK], 0);
      exit;
   end;

   if Length(TemplateDir) = 0 then begin
      MessageDlg('Не указан путь к файлам-шаблонам', mtInformation, [mbOK], 0);
      exit;
   end;

   if TemplateDir[Length(TemplateDir)] <> '\' then
      TemplateDir := TemplateDir + '\';

   DestDir := MainReport6.OutDirectory.Text;
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

    if month < 10 then Appendix := '0';
    Appendix := Appendix + IntToStr(month);
    if (Year MOD 100) < 10 then Appendix := Appendix + '0';
    Appendix := Appendix + IntToStr(Year MOD 100);

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

    SNList := TStringList.Create;

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

    Table1.Exclusive := true;
    Table2.Exclusive := true;
    Table3.Exclusive := true;
    Table4.Exclusive := true;
    Table5.Exclusive := true;
    Table6.Exclusive := true;

    PeriodStr := ' UPDATEDATE >=''' + DateToStr(EncodeDate(Year, Month, 1)) + ''' AND ' +
                 ' UPDATEDATE <''' + DateToStr(IncMonth(EncodeDate(Year, Month, 1), 1)) + '''';

    if MainReport6.FlagPeriod.ItemIndex = 1 then begin //До, включительно
        PeriodStr := ' UPDATEDATE <''' + DateToStr(IncMonth(EncodeDate(Year, Month, 1), 1)) + '''';
    end;

    if MainReport6.FlagPeriod.ItemIndex = 2 then begin //После, включительно
        PeriodStr := ' UPDATEDATE >= ''' + DateToStr(EncodeDate(Year, Month, 1)) + '''';
    end;

    try
        Table1.Open;
        Table2.Open;
        Table3.Open;
        Table4.Open;
        Table5.Open;
        Table6.Open;

        WorkSQL.SQL.Clear;
        WorkSQL.SQL.Add('SELECT SERIA,NUMBER');
        WorkSQL.SQL.Add('FROM MANDATOR');
        WorkSQL.SQL.Add('WHERE ' + PeriodStr);
        WorkSQL.SQL.Add('AND');
        WorkSQL.SQL.Add('VERSION>1 AND ');
        WorkSQL.SQL.Add('PAY1 IS NOT NULL');

        WorkSQL.SQL.Add('UNION');

        WorkSQL.SQL.Add('SELECT SERIA,NUMBER');
        WorkSQL.SQL.Add('FROM MANDOWN');
        WorkSQL.SQL.Add('WHERE ' + PeriodStr);

        WorkSQL.SQL.Add('UNION');

        WorkSQL.SQL.Add('SELECT SERIA,NUMBER');
        WorkSQL.SQL.Add('FROM MANDAV2');
        WorkSQL.SQL.Add('WHERE ' + PeriodStr);
        StatusBar.Panels[0].Text := 'Поиск полисов...';
        StatusBar.Update;
        WorkSQL.Open;
        while not WorkSQL.EOF do begin
            SNList.Add(WorkSQL.FieldByName('SERIA').AsString + '/' + WorkSQL.FieldByName('NUMBER').AsString);
            WorkSQL.Next;
        end;
        WorkSQL.Close;

        if MainReport6.IsUsePrepare.Checked then begin
            StatusBar.Panels[0].Text := 'Предварительный анализ...';
            StatusBar.Update;
            WorkSQL.SQL.Clear;
            WorkSQL.SQL.Add('DELETE FROM AuxExpBuro');
            WorkSQL.ExecSQL;
            WorkSQL.SQL.Clear;
            WorkSQL.SQL.Add('INSERT INTO AuxExpBuro(SERIA,NUMBER,POR_NO)VALUES(:SERIA,:NUMBER,:N)');
            for i := 0 to SNList.Count - 1 do begin
                SERIA := Copy(SNList[i], 1, Pos('/', SNList[i]) - 1);
                NUMBER := StrToInt(Copy(SNList[i], Pos('/', SNList[i]) + 1, 100));
                WorkSQL.ParamByName('SERIA').AsString := SERIA;
                WorkSQL.ParamByName('NUMBER').AsFloat := NUMBER;
                WorkSQL.ParamByName('N').AsInteger := i;
                WorkSQL.ExecSQL;
                WorkSQL.Close;
            end;
            ExpSQLPrep.Open;
            ExpSQLPrep.First;
        end;

        StatusBar.Panels[0].Text := 'Экспортирую ' + IntToStr(SNList.Count) + ' полисов...';
        StatusBar.Update;

        SetLength(AOwn, 64);

        ProgressBar.Max := 100;
        StartTime := GetCurrentTime;
        for i := 0 to SNList.Count - 1 do begin
            if(GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 then begin
                MessageBeep(0);
                ExpState := false;
                break;
            end;

            if ExpSQLPrep.Active then begin
                if (i > 0) then
                    ExpSQLPrep.Next;

                if ExpSQLPrep.EOF then begin
                    ExpState := false;
                    MessageDlg('Ошибка в подготовленном буфере', mtInformation, [mbOK], 0);
                    break;
                end;

                if ExpSQLPrep.FieldByName('Number').IsNull then begin
                    ExpState := false;
                    MessageDlg('Найдены аварии или владельцы по полису ' + SNList[i] + ', но полиса нет!!!', mtInformation, [mbOK], 0);
                    break;
                end;
            end;

            ExceptSN := SNList[i];
            SERIA := Copy(SNList[i], 1, Pos('/', SNList[i]) - 1);
            NUMBER := StrToInt(Copy(SNList[i], Pos('/', SNList[i]) + 1, 100));
            if ExpSQLPrep.Active then
                DIVISION_CODE := GetDivisionCode2(ExpSQLPrep.FieldByName('AGENTCODE').AsString)
            else
                DIVISION_CODE := GetDivisionCode(SERIA, NUMBER);

            if DIVISION_CODE = '?' then begin
                ExpState := false;
                break;
            end;

            ProgressBar.Position := ROUND(i * 100 / SNList.Count);

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

            StatusBar.Panels[1].Text := 'N' + IntToStr(i + 1) + ', ' + TimeStr;
            StatusBar.Update;

            OwnCounter := 0;

            ExpSQL_T.Filter := 'SERIA=''' + SERIA + ''' AND NUMBER=' + IntToStr(NUMBER);
            ExpSQL_T.Filtered := true;

            //Table 3 (Данные о владельце)
            //Part 1 (other owners)
            if UseTable then begin
                ExpSQL_T.TableName := 'MANDOWN';
                //ExpSQL_T.Filter := 'SERIA=''' + SERIA + ''' AND NUMBER=' + IntToStr(NUMBER);
                ExpSQL_T.Open;
                //ExpSQL_T.Locate('Seria;Number', vararrayof([SERIA, NUMBER]), []);
                ExpSQL := ExpSQL_T;
            end
            else begin
                ExpSQL_Q.SQL.Clear;
                ExpSQL_Q.SQL.Add('SELECT NAME,CITYCODE,ADDRESS,PRAVO,FIZICH,K,WHENPRAVA,UPDATEDATE,OWNERCODE');
                ExpSQL_Q.SQL.Add('FROM MANDOWN');
                ExpSQL_Q.SQL.Add('WHERE NUMBER=:N AND SERIA=:S');
                ExpSQL_Q.ParamByName('S').AsString := SERIA;
                ExpSQL_Q.ParamByName('N').AsInteger := NUMBER;
                ExpSQL_Q.Open;
                ExpSQL := ExpSQL_Q;
            end;

            while not ExpSQL.Eof do begin
                Table3.Append;
                Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table3.FieldByName('DIVISION').AsString := DIVISION_CODE;

                CarOwnDescr.Name := ExpSQL.FieldByName('NAME').AsString;
                CarOwnDescr.Addr_code := ExpSQL.FieldByName('CITYCODE').AsString;
                CarOwnDescr.Address := ExpSQL.FieldByName('ADDRESS').AsString;
                CarOwnDescr.Owner_Type := ExpSQL.FieldByName('PRAVO').AsInteger;
                CarOwnDescr.R_N := 1;
                CarOwnDescr.F_L := ExpSQL.FieldByName('FIZICH').AsInteger;
                CarOwnDescr.DISCOUNT_2 := ExpSQL.FieldByName('K').AsFloat;
                CarOwnDescr.DVR_EXPR := -1;

                Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNERCODE').AsInteger, CarOwnDescr, NeedUpdate)); //OWNERCODE
                Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('NAME').AsString);
                Table3.FieldByName('ADDR_CODE').AsString := ExpSQL.FieldByName('CITYCODE').AsString;
                Table3.FieldByName('ADDRESS').AsString := ExpStr(ExpSQL.FieldByName('ADDRESS').AsString);
                Table3.FieldByName('OWNER_TYPE').AsString := GetOwnerType(ExpSQL.FieldByName('PRAVO').AsInteger);
                Table3.FieldByName('R_N').AsString := 'R'; //Always resident
                Table3.FieldByName('F_L').AsString := GetFizUrid(ExpSQL.FieldByName('FIZICH').AsInteger);
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('K').AsFloat);
                if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
                   errCollectStr := errCollectStr + 'Обнаружен неправильный тип скидки ' + ExpSQL.FieldByName('K').AsString + #13#10 +
                                              '(Владельцы) Полис ' + SERIA + '/' + IntToStr(NUMBER) + ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
                   ExpState := false;
                end;
                Table3.FieldByName('DRV_EXPR').Value := ExpSQL.FieldByName('WHENPRAVA').Value;
                Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
                if(NeedUpdate) then SetMaxCodeFor(Seria, Number, 'MANDOWN', 'OWNERCODE', 'NAME', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), CarOwnDescr.Name);
                Table3.Post;

                AOwn[OwnCounter].FIO := Table3.FieldByName('NAME').AsString;
                AOwn[OwnCounter].Code := Table3.FieldByName('MSG_ID').AsString;
                Inc(OwnCounter);

                //Table 2 (Данные о других владельцах)
                Table2.Append;
                Table2.FieldByName('POLICY_NUM').AsString := SERIA + ExpZeroNmb(NUMBER);
                Table2.FieldByName('OWNER_ID').AsString := Table3.FieldByName('MSG_ID').AsString;
                Table2.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table2.FieldByName('DIVISION').AsString := DIVISION_CODE;
                Table2.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
                Table2.Post;

                ExpSQL.Next;
            end;
            //END Part 1 (other owners)
            //End Table 3, Table 2 (other owners)

            ExpSQL.Close;
            //Заполнение данных о транспортном ср-ве + основной владелец + авто

            if ExpSQLPrep.Active then begin
                ExpSQL := ExpSQLPrep;
            end
            else
            if UseTable then begin
                ExpSQL_T.TableName := 'MANDATOR';
                ExpSQL_T.Open;
                ExpSQL := ExpSQL_T;
            end
            else begin
                ExpSQL_Q.SQL.Clear;
                ExpSQL_Q.SQL.Add('SELECT REGDATE,FROMDATE,FROMTIME,TODATE,STATE,STOPDATE,');
                ExpSQL_Q.SQL.Add('LETTER,DISCOUNT,CLASSAVCURR,PAYDATE,PAY1,PAY1CURR,');
                ExpSQL_Q.SQL.Add('PAY2DATE,PAY2SUMMA,PAY2CURR,UPDATEDATE,');
                //Для владельца
                ExpSQL_Q.SQL.Add('NAME,CITYCODE,ADDRESS,VLADENIE,INSURERTYPE,K,OWNERCODE,');
                //Для автомобиля 1
                ExpSQL_Q.SQL.Add('MARKA,VOLUME,TONNAGE,QUANTITY,POWER,BASECARCODE,');
                //Для автомобиля 2
                ExpSQL_Q.SQL.Add('AUTONUMBER,NUMBERBODY,CHASSIS,CARCODE,RESIDENT');
                ExpSQL_Q.SQL.Add('FROM MANDATOR');
                ExpSQL_Q.SQL.Add('WHERE NUMBER=:N AND VERSION=2 AND SERIA=:S');
                ExpSQL_Q.ParamByName('S').AsString := SERIA;
                ExpSQL_Q.ParamByName('N').AsInteger := NUMBER;
                ExpSQL_Q.Open;
                ExpSQL := ExpSQL_Q;
            end;

            if ExpSQL.Eof then begin
               MessageDlg('Не найден полис ' + SERIA + '/' + IntToStr(NUMBER), mtInformation, [mbOK], 0);
               ExpState := false;
               break;
            end;

            //Главный владелец
            Table3.Append;
            Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table3.FieldByName('DIVISION').AsString := DIVISION_CODE;

            CarOwnDescr.Name := ExpSQL.FieldByName('NAME').AsString;
            CarOwnDescr.Addr_code := ExpSQL.FieldByName('CITYCODE').AsString;
            CarOwnDescr.Address := ExpSQL.FieldByName('ADDRESS').AsString;
            CarOwnDescr.Owner_Type := ExpSQL.FieldByName('VLADENIE').AsInteger;
            CarOwnDescr.R_N := 1;
            CarOwnDescr.F_L := ExpSQL.FieldByName('INSURERTYPE').AsInteger;
            CarOwnDescr.DISCOUNT_2 := ExpSQL.FieldByName('K').AsFloat;
            CarOwnDescr.DVR_EXPR := -1;

            Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNERCODE').AsInteger, CarOwnDescr, NeedUpdate));
            Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('NAME').AsString);
            Table3.FieldByName('ADDR_CODE').AsString := ExpSQL.FieldByName('CITYCODE').AsString;
            Table3.FieldByName('ADDRESS').AsString := ExpStr(ExpSQL.FieldByName('ADDRESS').AsString);
            Table3.FieldByName('OWNER_TYPE').AsString := GetOwnerType(ExpSQL.FieldByName('VLADENIE').AsInteger);

            INI := TIniFile.Create(BLANK_INI);
            ResidentStr := INI.ReadString('MANDATORY', 'TarifTable_Is_Resident' + ExpSQL.FieldByName('RESIDENT').AsString, '');
            INI.Free;
            if (ResidentStr <> 'YES') AND (ResidentStr <> 'NO') then begin
               errCollectStr := errCollectStr + 'Не определён признак (резидент, не резидент) таблицы тарифов N' + ExpSQL.FieldByName('RESIDENT').AsString + #13#10;
               ExpState := false;
            end;

            if ResidentStr = 'YES' then
                Table3.FieldByName('R_N').AsString := GetResident(1)
            else
                Table3.FieldByName('R_N').AsString := GetResident(0);

            Table3.FieldByName('F_L').AsString := GetFizUrid(ExpSQL.FieldByName('INSURERTYPE').AsInteger);
            if ExpSQL.FieldByName('VERSION').AsInteger = 1 then
                Table3.FieldByName('DISCOUNT_2').AsString := '0'
            else
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('K').AsFloat);
            if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
               errCollectStr := errCollectStr + 'Обнаружен неправильный тип скидки ' + ExpSQL.FieldByName('K').AsString + #13#10 +
                                          'Полис ' + SERIA + '/' + IntToStr(NUMBER) + ' ' + ExpSQL.FieldByName('NAME').AsString + #13#10;
               ExpState := false;
            end;
            Table3.FieldByName('DRV_EXPR').Clear;
            Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDATOR', 'OWNERCODE', '', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), '');
            Table3.Post;

            AOwn[OwnCounter].FIO := Table3.FieldByName('NAME').AsString;
            AOwn[OwnCounter].Code := Table3.FieldByName('MSG_ID').AsString;
            Inc(OwnCounter);
            //END Главный владелец

            // Данные о серийном ТС
            Table5.Append;
            Table5.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table5.FieldByName('DIVISION').AsString := DIVISION_CODE;

            SerCarDesrc.Marka := ExpSQL.FieldByName('MARKA').AsString;
            SerCarDesrc.TypeTC := ExpSQL.FieldByName('LETTER').AsString;
            SerCarDesrc.Characteristic := GetCharacteristicAuto(SerCarDesrc.TypeTC,
                                              ExpSQL.FieldByName('VOLUME').AsFloat,
                                              ExpSQL.FieldByName('TONNAGE').AsFloat,
                                              ExpSQL.FieldByName('QUANTITY').AsFloat,
                                              ExpSQL.FieldByName('POWER').AsFloat,
                                              ResidentStr = 'YES');

            if SerCarDesrc.Characteristic < 0 then begin
                errCollectStr := errCollectStr + 'У полиса ' + SERIA + '/' + IntToStr(NUMBER) + ' не определена характеристика ТС' + #13#10;
                ExpState := false;
            end;

            Table5.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeSerTC(ExpSQL.FieldByName('BASECARCODE').AsInteger, SerCarDesrc, NeedUpdate)); //BASECARCODE
            Table5.FieldByName('VEH_MARK').AsString := ExpSQL.FieldByName('MARKA').AsString;
            Table5.FieldByName('VEH_B_TYPE').AsString := ExpSQL.FieldByName('LETTER').AsString;
            CharAuto := SerCarDesrc.Characteristic;
            if CharAuto > 0 then
                Table5.FieldByName('CHAR_VALUE').AsFloat := CharAuto;
            if CharAuto < 0 then begin
                errCollectStr := errCollectStr + 'У полиса ' + SERIA + '/' + IntToStr(NUMBER) + ' не правильная характеристика ТС ' + ExpSQL.FieldByName('LETTER').AsString + #13#10;
                ExpState := false;
            end;

            Table5.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDATOR', 'BASECARCODE', '', DecodeMsgId(Table5.FieldByName('MSG_ID').AsString), '');
            Table5.Post;
            //END Данные о серийном ТС

            // Данные о ТС
            Table4.Append;
            Table4.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table4.FieldByName('DIVISION').AsString := DIVISION_CODE;

            CarDescr.AutoNumber := ExpSQL.FieldByName('AUTONUMBER').AsString;
            CarDescr.BodyNumber := ExpSQL.FieldByName('NUMBERBODY').AsString;
            CarDescr.BaseCarCode := StrToFloat(DecodeMsgId(Table5.FieldByName('MSG_ID').AsString));

            Table4.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeTC(ExpSQL.FieldByName('CARCODE').AsInteger, CarDescr, NeedUpdate)); //CARCODE
            CURR_CAR_CODE := Table4.FieldByName('MSG_ID').AsString;

            Table4.FieldByName('NUM_SIGN').AsString := ExpSQL.FieldByName('AUTONUMBER').AsString;

            //В стороке м.б. K1=1.2, K2=3
            if(ExpSQL.FieldByName('NUMBERBODY').AsString <> '') AND
              (Pos('=', ExpSQL.FieldByName('NUMBERBODY').AsString) = 0) AND
              (Pos('.', ExpSQL.FieldByName('NUMBERBODY').AsString) = 0) AND
              (Pos(',', ExpSQL.FieldByName('NUMBERBODY').AsString) = 0) then
                 Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('NUMBERBODY').AsString;

            if(ExpSQL.FieldByName('CHASSIS').AsString <> '') AND
              (Pos('=', ExpSQL.FieldByName('CHASSIS').AsString) = 0) AND
              (Pos('.', ExpSQL.FieldByName('CHASSIS').AsString) = 0) AND
              (Pos(',', ExpSQL.FieldByName('CHASSIS').AsString) = 0) then
                 Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('CHASSIS').AsString;

            Table4.FieldByName('BASECAR_ID').AsString := Table5.FieldByName('MSG_ID').AsString;
            Table4.FieldByName('COUNTRY').Clear;

            Table4.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
            if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDATOR', 'CARCODE', '', DecodeMsgId(Table4.FieldByName('MSG_ID').AsString), '');
            Table4.Post;
            //END Данные о ТС

            //Договор
            Table1.Append;
            Table1.FieldByName('POLICY_NUM').AsString := SERIA + ExpZeroNmb(NUMBER);
            Table1.FieldByName('INS_DOC').AsString := 'S';
            Table1.FieldByName('COMPANY').AsString := COMPANY_CODE;
            Table1.FieldByName('DIVISION').AsString := DIVISION_CODE;
            Table1.FieldByName('DATE_DISTR').AsDateTime := ExpSQL.FieldByName('REGDATE').AsDateTime;
            Table1.FieldByName('DATE_BEG').AsDateTime := ExpSQL.FieldByName('FROMDATE').AsDateTime;
            Table1.FieldByName('TIME_BEG').AsFloat := ExpTime(ExpSQL.FieldByName('FROMTIME').AsString);
            Table1.FieldByName('DATE_END').AsDateTime := ExpSQL.FieldByName('TODATE').AsDateTime;
            if ExpStateFunc(ExpSQL.FieldByName('State').AsInteger) = '' then begin
                errCollectStr := errCollectStr + 'У полиса ' + SERIA + '/' + IntToStr(NUMBER) + ' не правильное состояние' + #13#10;
                ExpState := false;
            end;
            Table1.FieldByName('POL_STATE').AsString := ExpStateFunc(ExpSQL.FieldByName('State').AsInteger);
            Table1.FieldByName('DATE_TERM').Value := ExpSQL.FieldByName('STOPDATE').Value;
            Table1.FieldByName('INSUR_ID').AsString := Table3.FieldByName('MSG_ID').AsString;
            Table1.FieldByName('VEH_TYPE').AsString := ExpSQL.FieldByName('LETTER').AsString;

            if ExpSQL.FieldByName('DISCOUNT').AsInteger > 0 then
                Table1.FieldByName('DISCOUNT_1').AsString := 'H'
            else
                Table1.FieldByName('DISCOUNT_1').AsString := 'F';

            if ExpSQL.FieldByName('VERSION').AsInteger = 1 then
                Table1.FieldByName('DISCOUNT_3').AsString := 'A0'
            else
                Table1.FieldByName('DISCOUNT_3').AsString := ExpSQL.FieldByName('CLASSAVCURR').AsString;

            Table1.FieldByName('CAR_ID').AsString := Table4.FieldByName('MSG_ID').AsString;

            Table1.FieldByName('DATE_PREM1').Value := ExpSQL.FieldByName('PAYDATE').Value;
            Table1.FieldByName('PREM_SUM1').Value := ExpSQL.FieldByName('PAY1').Value;
            Table1.FieldByName('DATE_PREM2').Value := ExpSQL.FieldByName('PAY2DATE').Value;
            Table1.FieldByName('PREM_SUM2').Value := ExpSQL.FieldByName('PAY2SUMMA').Value;
            if ExpCurr(ExpSQL.FieldByName('PAY1CURR').AsString) = '' then begin
                errCollectStr := errCollectStr + 'У полиса ' + SERIA + '/' + IntToStr(NUMBER) + ' не правильная валюта оплаты' + #13#10;
                ExpState := false;
            end;
            Table1.FieldByName('CURRENCY').AsString := ExpCurr(ExpSQL.FieldByName('PAY1CURR').AsString);
            Table1.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;

            Table1.Post;
            //END Договор

            if ExpSQL <> ExpSQLPrep then begin
                ExpSQL.Next;
                if not ExpSQL.Eof then begin
                   errCollectStr := errCollectStr + 'Найдено 2 полиса ' + SERIA + '/' + IntToStr(NUMBER) + #13#10;
                   ExpState := false;
                end;
            end;

            if ExpSQL <> ExpSQLPrep then //Нельзя закрывать буфер
                ExpSQL.Close;

            //ВЫПЛАТЫ
            if UseTable then begin
                ExpSQL_T.TableName := 'MANDAV2';
                ExpSQL_T.Open;
                ExpSQL := ExpSQL_T;
            end
            else begin
                ExpSQL_Q.SQL.Clear;
                ExpSQL_Q.SQL.Add('SELECT N,AVDATE,CITYCODE,AVPLACE,NWORK,WRDATE,WHAT,SN,');
                ExpSQL_Q.SQL.Add('SUM1V,SUM1_1V,V1_DATE,');
                ExpSQL_Q.SQL.Add('SUMV2,SUM2_1V,V2_DATE,');
                ExpSQL_Q.SQL.Add('SUMV3,SUM3_1V,V3_DATE,');
                ExpSQL_Q.SQL.Add('PAYMENTCODE,UPDATEDATE,');
                //Для потерпевшего
                ExpSQL_Q.SQL.Add('FIO,CITYCODE_P,FIOADDRESS,VLADEN_P,ISRESID_P,FIOTYPE,DISCOUNT_P,OWNERCODE_P,');
                //Для автомобиля потерпевшего
                ExpSQL_Q.SQL.Add('MARKA_P,BASE_TYPE_P,CHARVAL,BASECARCODE_P,');
                //Для автомобиля 2 потерпевшего
                ExpSQL_Q.SQL.Add('AUTONMB_P,BODY_NO,CARCODE_P,');
                //Виновник
                ExpSQL_Q.SQL.Add('BADMAN');
                ExpSQL_Q.SQL.Add('FROM MANDAV2');
                ExpSQL_Q.SQL.Add('WHERE NUMBER=:N AND SERIA=:S');
                ExpSQL_Q.ParamByName('S').AsString := SERIA;
                ExpSQL_Q.ParamByName('N').AsInteger := NUMBER;
                ExpSQL_Q.Open;
                ExpSQL := ExpSQL_Q;
            end;

            while not ExpSQL.EOF do begin
                //Потерпевший
                Table3.Append;
                Table3.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table3.FieldByName('DIVISION').AsString := DIVISION_CODE;

                CarOwnDescr.Name := ExpSQL.FieldByName('FIO').AsString;
                CarOwnDescr.Addr_code := ExpSQL.FieldByName('CITYCODE_P').AsString;
                CarOwnDescr.Address := ExpSQL.FieldByName('FIOADDRESS').AsString;
                CarOwnDescr.Owner_Type := ExpSQL.FieldByName('VLADEN_P').AsInteger;
                CarOwnDescr.R_N := ExpSQL.FieldByName('ISRESID_P').AsInteger;
                CarOwnDescr.F_L := ExpSQL.FieldByName('FIOTYPE').AsInteger;
                CarOwnDescr.DISCOUNT_2 := ExpSQL.FieldByName('DISCOUNT_P').AsFloat;
                CarOwnDescr.DVR_EXPR := -1;

                Table3.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeOwner(ExpSQL.FieldByName('OWNERCODE_P').AsInteger, CarOwnDescr, NeedUpdate)); //OWNERCODE
                Table3.FieldByName('NAME').AsString := ExpStr(ExpSQL.FieldByName('FIO').AsString);
                Table3.FieldByName('ADDR_CODE').AsString := ExpSQL.FieldByName('CITYCODE_P').AsString;
                Table3.FieldByName('ADDRESS').AsString := ExpStr(ExpSQL.FieldByName('FIOADDRESS').AsString);
                Table3.FieldByName('OWNER_TYPE').AsString := GetOwnerType(ExpSQL.FieldByName('VLADEN_P').AsInteger);
                Table3.FieldByName('R_N').AsString := GetResident(ExpSQL.FieldByName('ISRESID_P').AsInteger);
                Table3.FieldByName('F_L').AsString := GetFizUrid(ExpSQL.FieldByName('FIOTYPE').AsInteger);
                Table3.FieldByName('DISCOUNT_2').AsString := GetDiscount2(ExpSQL.FieldByName('DISCOUNT_P').AsFloat);
                if Table3.FieldByName('DISCOUNT_2').AsString = '?' then begin
                   errCollectStr := errCollectStr + 'Обнаружен не правильный тип скидки ' + ExpSQL.FieldByName('DISCOUNT_P').AsString + #13#10 +
                                              'Полис ' + SERIA + '/' + IntToStr(NUMBER) + ' ' + ExpSQL.FieldByName('FIO').AsString + #13#10;
                   ExpState := false;
                end;
                Table3.FieldByName('DRV_EXPR').Clear;
                Table3.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
                if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDAV2', 'OWNERCODE_P', 'N', DecodeMsgId(Table3.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString);
                Table3.Post;

                if ExpSQL.FieldByName('MARKA_P').AsString <> '' then begin
                    //Автомобиль(Базовый) потерпевшего
                    Table5.Append;
                    Table5.FieldByName('COMPANY').AsString := COMPANY_CODE;
                    Table5.FieldByName('DIVISION').AsString := DIVISION_CODE;

                    SerCarDesrc.Marka := ExpSQL.FieldByName('MARKA_P').AsString;
                    SerCarDesrc.TypeTC := ExpSQL.FieldByName('BASE_TYPE_P').AsString;
                    SerCarDesrc.Characteristic := ExpSQL.FieldByName('CHARVAL').AsFloat;

                    Table5.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeSerTC(ExpSQL.FieldByName('BASECARCODE_P').AsInteger, SerCarDesrc, NeedUpdate)); //BASECARCODE
                    Table5.FieldByName('VEH_MARK').AsString := ExpSQL.FieldByName('MARKA_P').AsString;
                    Table5.FieldByName('VEH_B_TYPE').AsString := ExpSQL.FieldByName('BASE_TYPE_P').AsString;
                    Table5.FieldByName('CHAR_VALUE').AsFloat := ExpSQL.FieldByName('CHARVAL').AsFloat;
                    Table5.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
                    if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDAV2', 'BASECARCODE_P', 'N', DecodeMsgId(Table5.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString);
                    Table5.Post;

                   //Автомобиль потерпевшего
                    Table4.Append;
                    Table4.FieldByName('COMPANY').AsString := COMPANY_CODE;
                    Table4.FieldByName('DIVISION').AsString := DIVISION_CODE;

                    CarDescr.AutoNumber := ExpSQL.FieldByName('AUTONMB_P').AsString;
                    CarDescr.BodyNumber := ExpSQL.FieldByName('BODY_NO').AsString;
                    CarDescr.BaseCarCode := StrToFloat(DecodeMsgId(Table5.FieldByName('MSG_ID').AsString));

                    Table4.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeTC(ExpSQL.FieldByName('CARCODE_P').AsInteger, CarDescr, NeedUpdate)); //CARCODE_P
                    Table4.FieldByName('NUM_SIGN').AsString := ExpSQL.FieldByName('AUTONMB_P').AsString;
                    Table4.FieldByName('BODY_NO').AsString := ExpSQL.FieldByName('BODY_NO').AsString;
                    Table4.FieldByName('BASECAR_ID').AsString := Table5.FieldByName('MSG_ID').AsString;
                    Table4.FieldByName('COUNTRY').Clear;
                    Table4.FieldByName('UPDATE').AsDateTime := ExpSQL.FieldByName('UPDATEDATE').AsDateTime;
                    if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDAV2', 'CARCODE_P', 'N', DecodeMsgId(Table4.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString);
                    Table4.Post;
                end;

                Table6.Append;
                Table6.FieldByName('COMPANY').AsString := COMPANY_CODE;
                Table6.FieldByName('DIVISION').AsString := DIVISION_CODE;
                Table6.FieldByName('MSG_ID').AsString := CodeMsgId(GetNextCodeAv(ExpSQL.FieldByName('PAYMENTCODE').AsInteger, NeedUpdate)); //PAYMENTCODE
                Table6.FieldByName('DATE_INC').AsDateTime := ExpSQL.FieldByName('AVDATE').AsDateTime;
                Table6.FieldByName('PLACE_CODE').AsString := ExpSQL.FieldByName('CITYCODE').AsString;
                Table6.FieldByName('PLACE_INC').AsString := ExpSQL.FieldByName('AVPLACE').AsString;
                Table6.FieldByName('INS_DOC_NO').AsString := ExpSQL.FieldByName('NWORK').AsString;
                Table6.FieldByName('DATE_STATE').AsDateTime := ExpSQL.FieldByName('WRDATE').AsDateTime;
                Table6.FieldByName('DECISION').AsString := GetDecision(ExpSQL.FieldByName('WHAT').AsInteger);
                if Table6.FieldByName('DECISION').AsString = '' then begin
                   MessageDlg('Решение по результатам рассмотрения не правильно' + #13#10 + 'Полис ' + SERIA + '/' + IntToStr(NUMBER) + #13#10 + ExpSQL.FieldByName('FIO').AsString, mtInformation, [mbOK], 0);
                   ExpState := false;
                   break;
                end;
                Table6.FieldByName('POL_NUM_P').AsString := GetPaySN(ExpSQL.FieldByName('SN').AsString);
                if Table6.FieldByName('POL_NUM_P').AsString = '?' then begin
                   MessageDlg('Номер документа потерпевшего не правильный' + #13#10 + 'Полис ' + SERIA + '/' + IntToStr(NUMBER), mtInformation, [mbOK], 0);
                   ExpState := false;
                   break;
                end;

                //FK Section
                Table6.FieldByName('ID_P').AsString := Table3.FieldByName('MSG_ID').AsString;
                if ExpSQL.FieldByName('MARKA_P').AsString <> '' then
                    Table6.FieldByName('CAR_ID_P').AsString := Table4.FieldByName('MSG_ID').AsString;

                Table6.FieldByName('POL_NUM_V').AsString := SERIA + ExpZeroNmb(NUMBER);
                Table6.FieldByName('ID_V').AsString := '';
                for find := 0 to OwnCounter - 1 do
                    if ExpSQL.FieldByName('BADMAN').AsString = AOwn[find].FIO then begin
                        Table6.FieldByName('ID_V').AsString := AOwn[find].Code;
                        Table6.FieldByName('CAR_ID_V').AsString := CURR_CAR_CODE;
                        break;
                    end;

                if Table6.FieldByName('ID_V').AsString = '' then begin
                    errCollectStr := errCollectStr + 'В авариях записан виновник ' + ExpSQL.FieldByName('BADMAN').AsString + #13#10 + 'Среди владельцев такого нет!' + #13#10 + 'Полис ' + SERIA + '/' + IntToStr(NUMBER) + #13#10;
                    ExpState := false;
                end;

                Table6.FieldByName('ACT_HARM1').Value := ExpSQL.FieldByName('SUM1V').Value;
                Table6.FieldByName('PAYMENT1').Value := ExpSQL.FieldByName('SUM1_1V').Value;
                Table6.FieldByName('DATE_PAY1').Value := ExpSQL.FieldByName('V1_DATE').Value;
                Table6.FieldByName('ACT_HARM2').Value := ExpSQL.FieldByName('SUMV2').Value;
                Table6.FieldByName('PAYMENT2').Value := ExpSQL.FieldByName('SUM2_1V').Value;
                Table6.FieldByName('DATE_PAY2').Value := ExpSQL.FieldByName('V2_DATE').Value;
                Table6.FieldByName('ACT_HARM3').Value := ExpSQL.FieldByName('SUMV3').Value;
                Table6.FieldByName('PAYMENT3').Value := ExpSQL.FieldByName('SUM3_1V').Value;
                Table6.FieldByName('DATE_PAY3').Value := ExpSQL.FieldByName('V3_DATE').Value;

                if (Table6.FieldByName('PAYMENT1').AsFloat >= 0.01) OR
                   (Table6.FieldByName('PAYMENT2').AsFloat >= 0.01) OR
                   (Table6.FieldByName('PAYMENT3').AsFloat >= 0.01) OR
                   (Table6.FieldByName('ACT_HARM1').AsFloat >= 0.01) OR
                   (Table6.FieldByName('ACT_HARM2').AsFloat >= 0.01) OR
                   (Table6.FieldByName('ACT_HARM3').AsFloat >= 0.01) then begin

                    if ExpSQL.FieldByName('V1').AsString <> '' then
                        Table6.FieldByName('CURRENCY').AsString := ExpCurr(ExpSQL.FieldByName('V1').AsString)
                    else
                        Table6.FieldByName('CURRENCY').AsString := ExpCurr('BRB');
                end;
                Table6.FieldByName('UPDATE').Value := ExpSQL.FieldByName('UPDATEDATE').Value;
                if NeedUpdate then SetMaxCodeFor(Seria, Number, 'MANDAV2', 'PAYMENTCODE', 'N', DecodeMsgId(Table6.FieldByName('MSG_ID').AsString), ExpSQL.FieldByName('N').AsString);
                Table6.Post;

                ExpSQL.Next
            end;

            ExpSQL.Close;
         end;
    except
      on E : Exception do begin
         WorkSQL.Close;
         ExpSQL.Close;
         MessageDlg('Ошибка создания отчёта'#13#10 + ExceptSN + ' ' + E.Message, mtInformation, [mbOK], 0);
         ExpState := false;
      end
      else
         MessageDlg('Критическая ошибка при формировании отчёта', mtInformation, [mbOk], 0);
    end;

    ExpSQLPrep.Close;

    if not((errCollectStr <> '') OR (ExpState = false) OR (MainReport6.IsTestExport.Checked)) then begin
      Table1.Close;
      Table2.Close;
      Table3.Close;
      Table4.Close;
      Table5.Close;
      Table6.Close;

      //MessageDlg('ВНИМАНИЕ!'#13'Сделайте копию полученных файлов. Программа попатается удалить дубликаты. Если возникнет ошибка, то используйте полученную копию для ручного удаления дубликатов.', mtInformation, [mbOk], 0);

      Table1.Open;
      Table2.Open;
      Table3.Open;
      Table4.Open;
      Table5.Open;
      Table6.Open;

      //Drop dublicates
      if (not DropDublicate(Table2, DBFileName2b{'OWNER_ID')) OR
         (not DropDublicate(Table3, DBFileName3b{'MSG_ID')) OR
         (not DropDublicate(Table4, DBFileName4b{'MSG_ID')) OR
         (not DropDublicate(Table5, DBFileName5b{'MSG_ID')) then begin
         ExpState := false;
      end;
      //END Drop dublicates
    end;

    PackTable(Table1);
    //PackTable(Table2);
    //PackTable(Table3);
    //PackTable(Table4);
    //PackTable(Table5);
    PackTable(Table6);

    Table1.Close;
    Table2.Close;
    Table3.Close;
    Table4.Close;
    Table5.Close;
    Table6.Close;

    Table2.DeleteTable;
    Table3.DeleteTable;
    Table4.DeleteTable;
    Table5.DeleteTable;

    Table1.Free;
    Table2.Free;
    Table3.Free;
    Table4.Free;
    Table5.Free;
    Table6.Free;

    SNList.Free;

    StatusBar.Panels[0].Text := '';

    if (errCollectStr <> '') OR (ExpState = false) OR (MainReport6.IsTestExport.Checked) then begin
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

    if errCollectStr <> '' then begin
        //for i := 1 to Length(errCollect) do
        //  if errCollect[i] = #13 then
        //      errCollect := Copy(errCollect, 1, i) + #10 + Copy(errCollect, i + 1, Length(errCollect));

        ShowInfo.InfoPanel.Text := errCollectStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end
    else
    if ExpState then begin
        Application.MainForm.BringToFront;
        if MainReport6.IsTestExport.Checked then
            MessageDlg('Тестовый прогон прошёл успешно! Можете запускать экспорт данных', mtInformation, [mbOK], 0)
        else
            MessageDlg('Данные подготовлены для передачи в БЮРО'#13'Они расположены в каталоге ' + MainReport6.OutDirectory.Text,  mtInformation, [mbOK], 0);
    end;

    ProgressBar.Position := 0;
end;
}
procedure TMainForm.FormResize(Sender: TObject);
begin
     StatusBar.Panels[1].Width := 150;
     StatusBar.Panels[0].Width := Width - 200 - StatusBar.Panels[1].Width;
     ProgressBar.Top := 3;
     ProgressBar.Left := StatusBar.Panels.Items[0].Width + StatusBar.Panels[1].Width + 3;
     ProgressBar.Width := 200 - 12;
end;

function TMainForm.UpdateTabField(field, table, where : string; var Code : integer) : boolean;
begin
     try
{        if UseTable then begin
             TabUpdater.Close;
             TabUpdater.Filtered := true;
             TabUpdater.Filter := where;
             TabUpdater.Open;
             TabUpdater.Edit;
             TabUpdater.FieldByName('field').AsInteger := Code;
             TabUpdater.Post;
             TabUpdater.Close;
             Inc(Code);
             UpdateTabField := true;
        end
        else begin}
             HandBookSQL.SQL.Clear;
             HandBookSQL.SQL.Add('UPDATE ' + table);
             HandBookSQL.SQL.Add('SET ' + field + '=:VAL');
             HandBookSQL.SQL.Add('WHERE ' + where);
             HandBookSQL.ParamByName('VAL').AsInteger := Code;
             HandBookSQL.ExecSQL;
             Inc(Code);
             UpdateTabField := true;
         //end;
      except
         on E : Exception do begin
            MessageDlg('Ошибка обновления кода' + #13 + E.Message, mtInformation, [mbOK], 0);
            UpdateTabField := false;
         end;
      end;
end;

function TMainForm.GetDivisionCode2(AgCode : string) : string;
begin
    if lastAgCode = AgCode then begin
        GetDivisionCode2 := IntToStr(lastDivCode);
        exit;
    end;

    try
         HandBookSQL.SQL.Clear;
         HandBookSQL.SQL.Add('SELECT DivisionCode');
         HandBookSQL.SQL.Add('FROM AGENT');
         HandBookSQL.SQL.Add('WHERE AGENT_CODE=''' + AgCode + '''');
         HandBookSQL.Open;
         if HandBookSQL.EOF then begin
            MessageDlg('Не найден код подразделения для агента ' + AgCode, mtInformation, [mbOK], 0);
            GetDivisionCode2 := '?';
         end
         else begin
            if HandBookSQL.FieldByName('DivisionCode').IsNull then begin
               MessageDlg('Не определён код подразделения для агента ' + AgCode, mtInformation, [mbOK], 0);
               GetDivisionCode2 := '?';
            end
            else begin
               GetDivisionCode2 := HandBookSQL.FieldByName('DivisionCode').AsString;
               lastAgCode := AgCode;
               lastDivCode := HandBookSQL.FieldByName('DivisionCode').AsInteger;
            end;
         end;
      except
         on E : Exception do begin
            MessageDlg('Ошибка получения кода подразделения для агента ' + AgCode + #13 + E.Message, mtInformation, [mbOK], 0);
            GetDivisionCode2 := '?';
         end;
      end;
end;

function TMainForm.GetDivisionCode(S : string; N : integer) : string;
begin
     try
         HandBookSQL.SQL.Clear;
         HandBookSQL.SQL.Add('SELECT DivisionCode');
         HandBookSQL.SQL.Add('FROM AGENT A, ' + GetTableName() + ' M');
         HandBookSQL.SQL.Add('WHERE NUMBER=:N AND A.AGENT_CODE=M.AGENTCODE AND SERIA=:S');
         HandBookSQL.ParamByName('S').AsString := S;
         HandBookSQL.ParamByName('N').AsInteger := N;
         HandBookSQL.Open;
         if HandBookSQL.EOF then begin
            MessageDlg('Не найден код подразделения для ' + S + '/' + IntToStr(N), mtInformation, [mbOK], 0);
            GetDivisionCode := '?';
         end
         else begin
            if HandBookSQL.FieldByName('DivisionCode').IsNull then begin
               MessageDlg('Не определён код подразделения для ' + S + '/' + IntToStr(N), mtInformation, [mbOK], 0);
               GetDivisionCode := '?';
            end
            else
               GetDivisionCode := HandBookSQL.FieldByName('DivisionCode').AsString;
         end;
      except
         on E : Exception do begin
            MessageDlg('Ошибка получения кода подразделения для ' + S + '/' + IntToStr(N) + #13 + E.Message, mtInformation, [mbOK], 0);
            GetDivisionCode := '?';
         end;
      end;
end;

function TMainForm.ExpStr(s : string) : string;
begin
//     s := UpperCase(s);
//     s := Trim(s);
//     while(Pos('Ё', s) <> 0) do
//        s[Pos('Ё', s)] := 'E';

//     while(Pos('  ', s) <> 0) do
//        s := Copy(s, 1, Pos('  ', s)) + Copy(s, Pos('  ', s) + 2, 128);

     ExpStr := s;
end;

function TMainForm.GetOwnerType(N : integer) : string;
begin
     case N of
        1 : GetOwnerType := 'S';
        2 : GetOwnerType := 'H';
        3 : GetOwnerType := 'O';
        4 : GetOwnerType := 'A';
        5 : GetOwnerType := 'D';
        6 : GetOwnerType := 'I';
     else
         GetOwnerType := '?';
     end;
end;

function TMainForm.GetFizUrid(N : integer) : string;
begin
     if N = 1 then
        GetFizUrid := 'F'
     else
        GetFizUrid := 'L';
end;

function TMainForm.GetDiscount2(N : double) : string;
var
      rounded : integer;
begin
      rounded := ROUND(N * 100);
      if rounded = 130 then  ///11111111111
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

function TMainForm.GetCharacteristicAuto(s : string; Volume, Tonnage,Quantity, Power : double; IsResident : boolean) : double;
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
    if (s = 'C0') OR (s = 'C1') OR (s = 'C2') OR (s = 'C3') OR (s = 'C4') OR (s = 'C5') OR
    (s = 'E0') OR (s = 'E1') OR (s = 'E2') OR (s = 'E3') then
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
    if (s = 'X') OR (s = 'Х') then
       GetCharacteristicAuto := 0
    else
       GetCharacteristicAuto := -1;
end;

function TMainForm.ExpTime(s : string) : integer;
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

function TMainForm.ExpStateFunc(state : integer) : string;
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

function TMainForm.ExpCurr(Curr : string) : string;
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

function TMainForm.ExpZeroNmb(N : Integer) : string;
var
      s : string;
begin
      s := IntToStr(N);
      while Length(s) < 7 do
          s := '0' + s;
      ExpZeroNmb := s;
end;

function TMainForm.GetResident(N : integer) : string;
begin
     if N = 1 then
        GetResident := 'R'
     else
        GetResident := 'N'
end;

function TMainForm.GetDecision(N : integer) : string;
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

function TMainForm.GetPaySN(s : string) : string;
var
     divIdx : integer;
     resStr : string;
begin
     s := Trim(s);
     if s = '' then
         resStr := ''
     else begin
         divIdx := Pos('/', s);
         if divIdx < 2 then begin
            resStr := '?'
         end
         else begin
            resStr := Copy(s, 1, divIdx - 1);
            try
                resStr := resStr + ExpZeroNmb(StrToInt(Copy(s, divIdx + 1, 64)));
            except
                resStr := '?';
            end
         end;
     end;
     GetPaySN := resStr;
end;

procedure TMainForm.KorrectMenuItemClick(Sender: TObject);
begin
     if KorrectData = nil then begin
        Application.CreateForm(TKorrectData, KorrectData);
        //KorrectData.Parent := self;
     end;
     KorrectData.BringToFront
end;

function FTS(V : double) : string;
var
    s : string;
begin
    s := FloatToStr(V);
//    if Pos(',', s) > 0 then
//        s[Pos(',', s)] := '.';
    FTS := '''' + s + '''';
end;

         {
function TMainForm.GetNextCodeOwner(Code : integer; S : CarOwnerDescription; var NeedUpdate : boolean) : string;
begin
     if MainReport6.IsTestExport.Checked then begin
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
         //Поиск по трём таблицам
         //Без учёта резидентов и года получения прав
         with HandBookSQL, HandBookSQL.SQL do begin
              Clear;
              Add('SELECT OWNERCODE');
              Add('FROM MANDATOR');
              Add('WHERE OWNERCODE IS NOT NULL AND NAME=:N');
              Add('AND ADDRESS=:A');
              Add('AND CITYCODE=:C');
              Add('AND VLADENIE=:V');
              Add('AND INSURERTYPE=:I');
              Add('AND K=:K');

              ParamByName('N').AsString := S.Name;
              ParamByName('A').AsString := S.Address;
              ParamByName('C').AsString := S.Addr_code;
              ParamByName('V').AsFloat := S.Owner_Type;
              ParamByName('I').AsFloat := S.F_L;
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if FieldByName('OWNERCODE').AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(FieldByName('OWNERCODE').AsInteger);
                       Close;
                       exit;
                    end;
                    Next;
              end;

              //Поиск в других владельцах
              Clear;
              Add('SELECT OWNERCODE');
              Add('FROM MANDOWN');
              Add('WHERE OWNERCODE IS NOT NULL AND NAME=:N');
              Add('AND ADDRESS=:A');
              Add('AND CITYCODE=:C');
              Add('AND PRAVO=:V');
              Add('AND FIZICH=:I');
              Add('AND K=:K');

              ParamByName('N').AsString := S.Name;
              ParamByName('A').AsString := S.Address;
              ParamByName('C').AsString := S.Addr_code;
              ParamByName('V').AsFloat := S.Owner_Type;
              ParamByName('I').AsFloat := S.F_L;
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if FieldByName('OWNERCODE').AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(FieldByName('OWNERCODE').AsInteger);
                       Close;
                       exit;
                    end;
                    Next;
              end;

              //Поиск в выплатах
              Clear;
              Add('SELECT OWNERCODE_P AS OWNERCODE');
              Add('FROM MANDAV2');
              Add('WHERE OWNERCODE_P IS NOT NULL AND FIO=:N');
              Add('AND FIOADDRESS=:A');
              Add('AND CITYCODE_P=:C');
              Add('AND VLADEN_P=:V');
              Add('AND FIOTYPE=:I');
              Add('AND DISCOUNT_P=:K');

              ParamByName('N').AsString := S.Name;
              ParamByName('A').AsString := S.Address;
              ParamByName('C').AsString := S.Addr_code;
              ParamByName('V').AsFloat := S.Owner_Type;
              ParamByName('I').AsFloat := S.F_L;
              ParamByName('K').AsFloat := S.DISCOUNT_2;

              Open;

              while not EOF do begin
                    if FieldByName('OWNERCODE').AsInteger > 0 then begin
                       GetNextCodeOwner := IntToStr(FieldByName('OWNERCODE').AsInteger);
                       Close;
                       exit;
                    end;
                    Next;
              end;
         end; // with
//     end;

     if OwnerCode_Holes.Count > 0 then begin
         GetNextCodeOwner := IntToStr(Integer(OwnerCode_Holes[0]));
         OwnerCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxOwnerCode);
     GetNextCodeOwner := IntToStr(MaxOwnerCode);
end;

function TMainForm.GetMaxCodeFor(Table, Field : string) : integer;
begin
     with HandBookSQL, HandBookSQL.SQL do begin
          Clear;
          Add('SELECT MAX(' + Field + ') AS CODE');
          Add('FROM ' + Table);
          Add('WHERE ' + Field + ' IS NOT NULL');
          Open;
          GetMaxCodeFor := FieldByName('CODE').AsInteger;
          Close;
     end;
end;

procedure TMainForm.SetMaxCodeFor(Seria : string; Number : integer; Table, Field1, Field2, Val1, Val2 : string);
begin
    if UseTable then begin
         TabUpdater.Close;
         TabUpdater.Filtered := true;
         TabUpdater.Filter := 'NUMBER=' + IntToStr(Number) +
                              ' AND SERIA=''' + Seria + '''';
         if Length(Field2) > 0 then begin
            try //Проверка строка или число
               StrToInt(Val2);
               TabUpdater.Filter := TabUpdater.Filter + ' AND ' + Field2 + '=' + Val2;
            except
               TabUpdater.Filter := TabUpdater.Filter + ' AND ' + Field2 + '=''' + Val2 + '''';
            end;
         end;

         TabUpdater.TableName := Table;
         TabUpdater.Open;
         TabUpdater.Edit;
         TabUpdater.FieldByName(Field1).AsInteger := StrToInt(Val1);
         TabUpdater.Post;
         TabUpdater.Close;
     end
     else begin
         with HandBookSQL, HandBookSQL.SQL do begin
              Clear;
              Add('UPDATE ' + Table);
              Add('SET '  + Field1  + '=:VAL');
              Add('WHERE NUMBER=:N AND SERIA=:S');
              if Length(Field2) > 0 then
                 Add('AND ' + Field2 + '=:VAL2');

              ParamByName('S').AsString := Seria;
              ParamByName('N').AsFloat := Number;
              ParamByName('VAL').AsFloat := StrToInt(Val1);
              if Length(Field2) > 0 then begin
                  try //Проверка строка или число
                     StrToInt(Val2);
                     ParamByName('VAL2').AsFloat := StrToInt(Val2);
                  except
                     ParamByName('VAL2').AsString := Val2;
                  end;
              end;
              ExecSQL;
         end;
     end;
end;

function TMainForm.GetNextCodeSerTC(Code : integer; S : SerCarDescription; var NeedUpdate : boolean) : string;
begin
     if MainReport6.IsTestExport.Checked then begin
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

     with HandBookSQL, HandBookSQL.SQL do begin
          Clear;
          Add('SELECT BASECARCODE');
          Add('FROM MANDATOR');
          Add('WHERE BASECARCODE IS NOT NULL');
          Add('AND MARKA=:M');
          Add('AND LETTER=:BT');
          Add('AND ((Tonnage = :CH) OR (Quantity = :CH) OR (Power = :CH) OR (Volume = :CH))');

          ParamByName('M').AsString := S.Marka;
          ParamByName('BT').AsString := S.TypeTC;
          ParamByName('CH').AsFloat := S.Characteristic;
          Open;

          while not EOF do begin
                if FieldByName('BASECARCODE').AsInteger > 0 then begin
                   GetNextCodeSerTC := IntToStr(FieldByName('BASECARCODE').AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;

          Clear;
          Add('SELECT BASECARCODE_P');
          Add('FROM MANDAV2');
          Add('WHERE BASECARCODE_P IS NOT NULL');
          Add('AND MARKA_P=:M');
          Add('AND BASE_TYPE_P=:BT');
          Add('AND CHARVAL=:CH');

          ParamByName('M').AsString := S.Marka;
          ParamByName('BT').AsString := S.TypeTC;
          ParamByName('CH').AsFloat := S.Characteristic;
          Open;

          while not EOF do begin
                if FieldByName('BASECARCODE_P').AsInteger > 0 then begin
                   GetNextCodeSerTC := IntToStr(FieldByName('BASECARCODE_P').AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;

     end; // with

     if SerTCCode_Holes.Count > 0 then begin
         GetNextCodeSerTC := IntToStr(Integer(SerTCCode_Holes[0]));
         SerTCCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxSerTCCode);
     GetNextCodeSerTC := IntToStr(MaxSerTCCode);
end;

function TMainForm.GetNextCodeTC(Code : integer; S : CarDescription; var NeedUpdate : boolean) : string;
begin
     if MainReport6.IsTestExport.Checked then begin
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

     //check with SerTCCode
     with HandBookSQL, HandBookSQL.SQL do begin
          Clear;
          Add('SELECT CARCODE');
          Add('FROM MANDATOR');
          Add('WHERE CARCODE IS NOT NULL');
          Add('AND AUTONUMBER=:NMB');
          Add('AND Numberbody=:NB');
          Add('AND BASECARCODE=:CODE');

          ParamByName('NMB').AsString := S.AutoNumber;
          ParamByName('NB').AsString := S.BodyNumber;
          ParamByName('CODE').AsFloat := S.BaseCarCode;
          Open;

          while not EOF do begin
                if FieldByName('CARCODE').AsInteger > 0 then begin
                   GetNextCodeTC := IntToStr(FieldByName('CARCODE').AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;

          Clear;
          Add('SELECT CARCODE_P');
          Add('FROM MANDAV2');
          Add('WHERE CARCODE_P IS NOT NULL');
          Add('AND AUTONMB_P=:NMB');
          Add('AND body_no=:NB');
          Add('AND BASECARCODE_P=:CODE');

          ParamByName('NMB').AsString := S.AutoNumber;
          ParamByName('NB').AsString := S.BodyNumber;
          ParamByName('CODE').AsFloat := S.BaseCarCode;
          Open;

          while not EOF do begin
                if FieldByName('CARCODE_P').AsInteger > 0 then begin
                   GetNextCodeTC := IntToStr(FieldByName('CARCODE_P').AsInteger);
                   Close;
                   exit;
                end;
                Next;
          end;
     end; // with

     if TCCode_Holes.Count > 0 then begin
         GetNextCodeTC := IntToStr(Integer(TCCode_Holes[0]));
         TCCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxTCCode);
     GetNextCodeTC := IntToStr(MaxTCCode);
end;

function TMainForm.GetNextCodeAv(Code : integer; var NeedUpdate : boolean) : string;
begin
     if MainReport6.IsTestExport.Checked then begin
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

     if AVCode_Holes.Count > 0 then begin
         GetNextCodeAv := IntToStr(Integer(AVCode_Holes[0]));
         AVCode_Holes.Delete(0);
         exit;
     end;

     Inc(MaxAVCode);
     GetNextCodeAv := IntToStr(MaxAVCode);
end;
procedure TMainForm.InitExportBuro;
var
    CurrCode : integer;
begin
    try
      Screen.Cursor := crHourGlass;
      
      OwnerCode_Holes.Clear;
      SerTCCode_Holes.Clear;
      TCCode_Holes.Clear;
      AVCode_Holes.Clear;

      //OWNER CODE
      StatusBar.SimpleText := 'Поиск кода владельца...';
      StatusBar.Update;

      MaxOwnerCode := GetMaxCodeFor('MANDATOR', 'OWNERCODE');
      CurrCode := GetMaxCodeFor('MANDOWN', 'OWNERCODE');
      if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;
      CurrCode := GetMaxCodeFor('MANDAV2', 'OWNERCODE_P');
      if CurrCode > MaxOwnerCode then MaxOwnerCode := CurrCode;

      // Ser TC
      StatusBar.SimpleText := 'Поиск кода характеристики авто...';
      StatusBar.Update;
      MaxSerTCCode := GetMaxCodeFor('MANDATOR', 'BaseCarCode');
      CurrCode := GetMaxCodeFor('MANDAV2', 'BaseCarCode_P');
      if CurrCode > MaxSerTCCode then MaxSerTCCode := CurrCode;

      // TC
      StatusBar.SimpleText := 'Поиск кода авто...';
      StatusBar.Update;
      MaxTCCode := GetMaxCodeFor('MANDATOR', 'CarCode');
      CurrCode := GetMaxCodeFor('MANDAV2', 'CarCode_P');
      if CurrCode > MaxTCCode then MaxTCCode := CurrCode;

      //AVCode
      StatusBar.SimpleText := 'Поиск кода аварии...';
      StatusBar.Update;
      MaxAVCode := GetMaxCodeFor('MANDAV2', 'PaymentCode');

      //Find Holes
      with WorkSQL, WorkSQL.SQL do begin
           StatusBar.SimpleText := 'Поиск пустых кодов (1)...';
           StatusBar.Update;

           Clear;
           Add('SELECT OWNERCODE FROM MANDATOR WHERE OWNERCODE > 0');
           Add('UNION');
           Add('SELECT OWNERCODE FROM MANDOWN WHERE OWNERCODE > 0');
           Add('UNION');
           Add('SELECT OWNERCODE_P FROM MANDAV2 WHERE OWNERCODE_P > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName('OWNERCODE').AsInteger do begin
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
           Add('SELECT BaseCarCode FROM MANDATOR WHERE BaseCarCode > 0');
           Add('UNION');
           Add('SELECT BaseCarCode_P FROM MANDAV2 WHERE BaseCarCode_P > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName('BaseCarCode').AsInteger do begin
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
           Add('SELECT CarCode FROM MANDATOR WHERE CarCode > 0');
           Add('UNION');
           Add('SELECT CarCode_P FROM MANDAV2 WHERE CarCode_P > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName('CarCode').AsInteger do begin
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
           Add('SELECT DISTINCT PaymentCode FROM MANDAV2 WHERE PaymentCode > 0');
           Add('ORDER BY 1');
           Open;
           CurrCode := 1;
           while not EOF do begin
                 while CurrCode < FieldByName('PaymentCode').AsInteger do begin
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

    StatusBar.SimpleText := '';
    StatusBar.Update;
    Screen.Cursor := crDefault;
end;
}
{
procedure TMainForm.BASO_Month(Sender: TObject);
var
    GoodPolises, GoodPolises2, BadPolises : integer;
    SumBRB, SumEUR : double;
    bigInt : int64;
begin
    if BasoReportsForm <> nil then exit;

    with MandatoryFilterDlg do
        if not IsPutDate.Checked AND
           not IsPayDate.Checked AND
           not IsWasActive.Checked AND
           not IsInsStart.Checked AND
           not IsInsEnd.Checked AND
           not IsFee2Date.Checked then
               if MessageDlg('В фильтре не указан ни один период. Продолжить?', mtInformation, [mbYes, mbNo], 0) = mrNo then exit;

    Explore.Visible := false;

    if CalculateSummForReport(GoodPolises, GoodPolises2, BadPolises, SumBRB, SumEUR) then begin
        if MessageDlg('Сумма в рублях ' + FloatToStr(SumBRB) + '. Показать отчёт?', mtInformation, [mbYes, mbNo], 0) = mrNo then exit;
        if BasoReportsForm = nil then begin
            Application.CreateForm(TBasoReportsForm, BasoReportsForm);
        end;
        BasoReportsForm.frReportMonth.Dictionary.Variables['Summary1'] := '''' + IntToStr(GoodPolises) + ', чужих ' + IntToStr(GoodPolises2) + '''';
        BasoReportsForm.frReportMonth.Dictionary.Variables['Summary2'] := '''' + IntToStr(BadPolises) + '''';
        bigInt := ROUND(SumBRB * 100);
        BasoReportsForm.frReportMonth.Dictionary.Variables['Summary3'] := '''' + FloatToStr(bigInt / 100) + '''';
        bigInt := ROUND(SumEUR * 100);
        BasoReportsForm.frReportMonth.Dictionary.Variables['Summary4'] := '''' + FloatToStr(bigInt / 100) + '''';

        BasoReportsForm.Show;
        BasoReportsForm.BringToFront;

        MainGrid.DataSource := nil;
        MainQuery.First;
        BasoReportsForm.frReportMonth.ShowReport;
        MainGrid.DataSource := MainDataSource;
    end
end;

procedure TMainForm.BASO_Quarter(Sender: TObject);
var
    GoodPolises, GoodPolises2, BadPolises : integer;
    SumBRB, SumEUR : double;
    bigInt : int64;
    filterStr : string;

    AllCounter, Counter : integer;

    AllSumBRB, AllSumEUR : double;
begin
    if BasoReportsForm <> nil then exit;

    if MandatoryFilterDlg.IsWasActive.Checked = false then begin
       MessageDlg('Вы не указали в фильтре период, за который необходимо подготовить отчёт', mtInformation, [mbOK], 0);
       exit;
    end;

    if GetFormula.ShowModal <> mrOk then exit;

    Explore.Visible := false;

    MandatoryFilterDlg.IsWasActive.Checked := false;
    filterStr := GetFilter('');
    MandatoryFilterDlg.IsWasActive.Checked := true;

    if CalculateSummForReport(GoodPolises, GoodPolises2, BadPolises, SumBRB, SumEUR) then begin
        if BasoReportsForm = nil then begin
            Application.CreateForm(TBasoReportsForm, BasoReportsForm);
        end;
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary0'] := '''C ' + DateToStr(MandatoryFilterDlg.WasActiveFrom.Date) + ' - ' + DateToStr(MandatoryFilterDlg.WasActiveTo.Date) + '''';
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary1'] := '''' + IntToStr(GoodPolises) + '''';
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary1x'] := '''' + IntToStr(GoodPolises2) + '''';
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary4'] := '''' + IntToStr(BadPolises) + '''';
        bigInt := ROUND(SumBRB * 100);
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary2'] := '''' + FloatToStr(bigInt / 100) + '''';
        bigInt := ROUND(SumEUR * 100);
        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary3'] := '''' + FloatToStr(bigInt / 100) + '''';

        BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary1_0'] := '''На ' + DateToStr(MandatoryFilterDlg.WasActiveTo.Date + 1) + '''';

        ProgressBar.Position := 0;

        try
            //количество действующих на дату
            with WorkSQL, WorkSQL.SQL do begin
                Clear;
                Add('SELECT COUNT(* ) AS S FROM MANDATOR M');
                Add('WHERE FROMDATE <= :DT AND');

                if filterStr <> '' then begin
                    Add('(' + filterStr + ') AND');
                end;

                //0 ""; 1 этап 2 этапа (опл и неопл)
                //1 "ИСПОРЧЕН"; не учитывать
                //2 "ДОСРОЧНО ЗАВЕРШЁН"; дата завершения как дата окончания
                //3 "УТЕРЯН"; //не учитывается
                //4 "СРОК ДЕЙСТВИЯ ЗАКОНЧИЛСЯ"; как нормальный
                //5 "ЗАВЕРШЁН ПО НЕУПЛАТЕ"; тоже что и нормальный без оплаты 2го взноса


                Add('(');
                //Для нормальных и завершённых
                //1 этап //2 этап (оплачено)
                Add(' (STATE IN (0, 4) AND (IsFee2 = 0 OR (IsFee2 > 0 AND ISPAY2 > 0)) AND TODATE >= :DT)');
                Add(' OR ');
                // 2 этапа (не оплачено)
                Add(' (STATE IN (0, 4, 5) AND IsFee2 > 0 AND ISPAY2 = 0 AND FEE2DATE >= :DT)');

                //2 "ДОСРОЧНО ЗАВЕРШЁН";
                Add(' OR ');
                Add(' (STATE IN (2) AND STOPDATE > :DT)');

                Add(')');
                ParamByName('DT').AsDateTime := CutTime(MandatoryFilterDlg.WasActiveTo.Date);

                Open;
                BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary1_1'] := '''' + FieldByName('S').AsString + '''';
                Close;

               ProgressBar.Position := 20;

               //Переоформлено
               Clear;
               Add('SELECT COUNT(* ) AS CNT FROM MANDATOR WHERE (PAYDATE IS NOT NULL AND STATE=3)'); //Утерянные (т.к дубликат)
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Add(' AND (FROMDATE BETWEEN ''' + DateToStr(MandatoryFilterDlg.WasActiveFrom.Date) + ''' AND ''' + DateToStr(MandatoryFilterDlg.WasActiveTo.Date) + ''')');
               Open;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary5'] := '''' + FieldByName('CNT').AsString + '''';
               Close;

               ProgressBar.Position := 40;

               Clear;
               Add('SELECT COUNT(* ) AS CNT,SUM(RetBRB) AS B, SUM(RetEUR) AS E FROM MANDATOR '); //Утерянные (т.к дубликат)
               Add('WHERE (PAYDATE IS NOT NULL AND RETBRB IS NOT NULL AND STATE<>3)');
               Add(' AND (' + GetFilter('') + ')');
               Open;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['RetCOUNT'] := '''' + FieldByName('CNT').AsString + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['RetBRB'] := '''' + FieldByName('B').AsString + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['RetEUR'] := '''' + FieldByName('E').AsString + '''';
               Close;

               BasoReportsForm.frReportQuarter.Dictionary.Variables['Summary1_2'] := '[Summary1_1]' + GetFormula.Formula.Text;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Otvetstv'] := '[Summary1]' + GetFormula.Formula.Text;

               //Суммы выплат (Sum1_1, Sum1_1v, V1_1, V1_Date)
               //Суммы выплат (Sum2_1, Sum2_1v, V2_1, V2_Date)
               //Суммы выплат (Sum3_1, Sum3_1v, V3_1, V3_Date)

               AllCounter := 0;
               AllSumBRB := 0;
               AllSumEUR := 0;
               ProgressBar.Position := 55;
               CalcMandPayment('Sum1_1', 'Sum1_1v', 'V1_Date', filterStr, Counter, SumBRB, SumEUR);
               AllCounter := AllCounter + Counter;
               AllSumBRB := AllSumBRB + SumBRB;
               AllSumEUR := AllSumEUR + SumEUR;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutCount'] := '''' + IntToStr(Counter) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutBRB'] := '''' + FloatToStr(SumBRB) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutEUR'] := '''' + FloatToStr(SumEUR) + '''';
               ProgressBar.Position := 75;
               CalcMandPayment('Sum2_1', 'Sum2_1v',  'V2_Date', filterStr, Counter, SumBRB, SumEUR);
               AllCounter := AllCounter + Counter;
               AllSumBRB := AllSumBRB + SumBRB;
               AllSumEUR := AllSumEUR + SumEUR;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out2Count'] := '''' + IntToStr(Counter) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out2BRB'] := '''' + FloatToStr(SumBRB) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out2EUR'] := '''' + FloatToStr(SumEUR) + '''';
               ProgressBar.Position := 90;
               CalcMandPayment('Sum3_1', 'Sum3_1v', 'V3_Date', filterStr, Counter, SumBRB, SumEUR);
               AllCounter := AllCounter + Counter;
               AllSumBRB := AllSumBRB + SumBRB;
               AllSumEUR := AllSumEUR + SumEUR;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out3Count'] := '''' + IntToStr(Counter) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out3BRB'] := '''' + FloatToStr(SumBRB) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Out3EUR'] := '''' + FloatToStr(SumEUR) + '''';
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutCount+'] := IntToStr(AllCounter);
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutBRB+'] := FloatToStr(AllSumBRB);
               BasoReportsForm.frReportQuarter.Dictionary.Variables['OutEUR+'] := FloatToStr(AllSumEUR);
               ProgressBar.Position := 100;

               //Отказы в выплате
               Clear;
               Add('SELECT COUNT(* ) AS CNT');
               Add('FROM MANDAV2 M2, MANDATOR M');
               Add('WHERE WHAT = 1 AND M2.SERIA=M.SERIA AND M2.NUMBER=M.NUMBER ');
               if Length(filterStr) > 0 then begin
                  Add('AND');
                  Add(filterStr);
               end;
               Add(' AND (WRDATE BETWEEN ''' + DateToStr(MandatoryFilterDlg.WasActiveFrom.Date) + ''' AND ''' + DateToStr(MandatoryFilterDlg.WasActiveTo.Date) + ''')');
               Open;
               BasoReportsForm.frReportQuarter.Dictionary.Variables['Otkaz'] := FieldByName('CNT').AsString;
               Close;
            end;
        except
            on E : Exception do begin
               MessageDlg('Ошибка подготовки квартального отчёта'#13 + E.Message, mtInformation, [mbOK], 0);
            end;
        end;

        BasoReportsForm.Show;
        BasoReportsForm.BringToFront;
        BasoReportsForm.frReportQuarter.ShowReport
    end;

    //Progress.ProcessLabel.Caption := '';
    ProgressBar.Position := 0;
end;

procedure TMainForm.TASK_Report1(Sender: TObject);
begin
     ExecTACKReport1
end;
 }
function getPayTypeName(Flag : integer) : string;
begin
    getPayTypeName := 'Нал';
    if Flag = 0 then
        getPayTypeName := 'Безнал';
end;

function getPayTypeName2(Flag : String) : string;
begin
    getPayTypeName2 := 'Нал';
    if Flag <> 'Y' then
        getPayTypeName2 := 'Безнал';
end;
(*
procedure TMainForm.MICalcSummClick(Sender: TObject);
var
     BTWDATE : string;
     msgStr  : string;
     msgStr1 : string;
     msgStr2 : string;
     msgStr3 : string;
     msgStr4 : string;
     //flt : string;

     AllSumma : double;
begin
     AllSumma := 0;
     GetDates.Label1.Visible := false;
     if GetDates.ShowModal <> mrYes then exit;

     UpdateDupData(GetFilter(''));

     BTWDATE := '''' + DateToStr(GetDates.FromDate.Date) + ''' AND ''' + DateToStr(GetDates.ToDate.Date) + '''';
     //flt := GetFilter('');

     if GetFilter('') <> '' then
         MessageDlg('У вас установлен фильтр. Отчёт будет создан на основании указанного фильтра', mtInformation, [mbOk], 0);

     try
       Screen.Cursor := crHourGlass;
       with WorkSQL, WorkSQL.SQL do begin
            StatusBar.Panels[0].Text := 'Расчёт 1й части взносов...';
            StatusBar.Update;
            Clear;
            Add('SELECT SUM(PAY1) AS S, PAY1CURR, PAYTYPE');
            Add('FROM MANDATOR');
            Add('WHERE PAYDATE IS NOT NULL');
            Add('AND (PREVSERIA IS NULL or PREVSERIA='''')');
            Add('AND RET1 BETWEEN ' + BTWDATE);
            Add('AND BASECARCODE=101');
            if GetFilter('AgentCode') <> '' then
                Add(' AND ' + GetFilter('AgentCode'));
            Add('GROUP BY PAY1CURR, PAYTYPE');
            Add('ORDER BY PAYTYPE');
            Open;
            while not EOF do begin
                if not FieldByName('S').IsNull then
                    msgStr1 := msgStr1 +
                               FieldByName('S').AsString +
                               ' ' + FieldByName('PAY1CURR').AsString +
                               ' ' + getPayTypeName(FieldByName('PAYTYPE').AsInteger) + #13;

                    if FieldByName('PAY1CURR').AsString = 'BRB' then
                        AllSumma := AllSumma + FieldByName('S').AsFloat;
                Next;
            end;
            Close;

            StatusBar.Panels[0].Text := 'Расчёт 2й части взносов...';
            StatusBar.Update;
            Clear;
            Add('SELECT SUM(PAY2SUMMA) AS S, PAY2CURR, PAY2TYPE');
            Add('FROM MANDATOR');
            Add('WHERE ISPAY2 <> 0 AND PAYDATE IS NOT NULL');
            Add('AND BASECARCODE=101');
            Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
            Add('AND RET2 BETWEEN ' + BTWDATE);
            if GetFilter('Fee2Text') <> '' then
                Add(' AND ' + GetFilter('Fee2Text'));
            Add('GROUP BY PAY2CURR, PAY2TYPE');
            Add('ORDER BY PAY2TYPE');
            Open;
            while not EOF do begin
                if not FieldByName('S').IsNull then
                    msgStr2 := msgStr2 +
                               FieldByName('S').AsString +
                               ' ' + FieldByName('PAY2CURR').AsString +
                               ' ' + getPayTypeName(FieldByName('PAY2TYPE').AsInteger) + #13;
                    if FieldByName('PAY2CURR').AsString = 'BRB' then
                        AllSumma := AllSumma + FieldByName('S').AsFloat;
                Next;
            end;
            Close;

            StatusBar.Panels[0].Text := 'Расчёт 2й части взносов (чужие полисы)...';
            StatusBar.Update;
            Clear;
            Add('SELECT SUM(PAY2SUMMA) AS S, PAY2CURR, PAY2TYPE');
            Add('FROM MANDATOR');
            Add('WHERE ISPAY2 <> 0 AND AGENTCODE IS NULL');
            Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
            Add('AND BASECARCODE=101');
            Add('AND RET2 BETWEEN ' + BTWDATE);
            if GetFilter('Fee2Text') <> '' then
                Add(' AND ' + GetFilter('Fee2Text'));
            Add('GROUP BY PAY2CURR, PAY2TYPE');
            Add('ORDER BY PAY2TYPE');
            Open;
            while not EOF do begin
                if not FieldByName('S').IsNull then
                    msgStr4 := msgStr4 +
                               FieldByName('S').AsString +
                               ' ' + FieldByName('PAY2CURR').AsString +
                               ' ' + getPayTypeName(FieldByName('PAY2TYPE').AsInteger) + #13;
                    if FieldByName('PAY2CURR').AsString = 'BRB' then
                        AllSumma := AllSumma + FieldByName('S').AsFloat;
                Next;
            end;
            Close;

            StatusBar.Panels[0].Text := 'Расчёт доплат...';
            StatusBar.Update;
            Clear;
            Add('SELECT SUM(SUMMA) AS S, CURR PAY2CURR, NAL');
            Add('FROM MANDATOR M, TAXI T');
            Add('WHERE SERIA=SER AND NUMBER=N');
            Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
            Add('AND BASECARCODE=101');
            Add('AND D BETWEEN ' + BTWDATE);
            if GetFilter('') <> '' then
                Add(' AND ' + GetFilter(''));
            Add('GROUP BY CURR, NAL');
            Add('ORDER BY NAL');
            Open;
            while not EOF do begin
                if not FieldByName('S').IsNull then
                    msgStr3 := msgStr3 +
                               FieldByName('S').AsString +
                               ' ' + FieldByName('PAY2CURR').AsString +
                               ' ' + getPayTypeName2(FieldByName('NAL').AsString) + #13;
                    if FieldByName('PAY2CURR').AsString = 'BRB' then
                        AllSumma := AllSumma + FieldByName('S').AsFloat;
                Next;
            end;
            Close;

            msgStr := 'Данные за период C ' + DateToStr(GetDates.FromDate.Date) + ' ПО ' + DateToStr(GetDates.ToDate.Date) + #13'с учётом фильтра'#13#13;
            if msgStr1 <> '' then
               msgStr := msgStr + 'Первые (и единственные) платежи '#13 + msgStr1
            else
               msgStr := msgStr + 'Первых (и единственных) платежей небыло'#13;

            if msgStr2 <> '' then
               msgStr := msgStr + #13'Вторые платежи '#13 + msgStr2
            else
               msgStr := msgStr + #13'Вторых платежей небыло'#13;

            if msgStr3 <> '' then
               msgStr := msgStr + #13'Доплаты за такси '#13 + msgStr3
            else
               msgStr := msgStr + #13'Доплат за такси небыло'#13;

            if msgStr4 <> '' then
               msgStr := msgStr + #13'Чужие полисы '#13 + msgStr4;
            //else
            //   msgStr := msgStr + #13'Доплат за такси небыло'#13;

            StatusBar.Panels[0].Text := '';
            if MessageDlg(msgStr + #13 + 'Сумма всех рублей ' + FloatToStr(AllSumma) + #13#13 + 'Вы хотите получить этот отчёт в Excel?', mtInformation, [mbYes, mbNo], 0) = mrYes then begin
                with WorkSQL, WorkSQL.SQL do begin
                    Clear;
                    Add('SELECT ''1я оплата'' Оплата, SERIA Cерия, NUMBER Номер, NAME Страхователь, FROMDATE C, TODATE ПО, PERIOD Период, TARIFGLASS Тариф, K1 Местность, K2 Аварийность, TARIF ТарифПолный');
                    Add(',RET1 Дата_оплаты, PAY1 Сумма_взноса, PAY1CURR Валюта, AGENTCODE Код_агента');
                    Add('FROM MANDATOR');
                    Add('WHERE PAYDATE IS NOT NULL');
                    Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
                    Add('AND BASECARCODE=101');
                    Add('AND RET1 BETWEEN ' + BTWDATE);

                    Add('UNION ALL ');
                    Add('SELECT ''2я оплата'' Оплата, SERIA Cерия, NUMBER Номер, NAME Страхователь, FROMDATE C, TODATE ПО, PERIOD Период, TARIFGLASS Тариф, K1 Местность, K2 Аварийность, TARIF ТарифПолный');
                    Add(',RET2 Дата_оплаты, PAY2SUMMA Сумма_1го_взноса, PAY2CURR Валюта, AGENTCODE Код_агента');
                    Add('FROM MANDATOR');
                    Add('WHERE PAYDATE IS NOT NULL');
                    Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
                    Add('AND BASECARCODE=101');
                    Add('AND RET2 BETWEEN ' + BTWDATE);

                    Add('UNION ALL ');
                    Add('SELECT ''только 2я оплата'' Оплата, SERIA Cерия, NUMBER Номер, NAME Страхователь, FROMDATE C, TODATE ПО, PERIOD Период, TARIFGLASS Тариф, K1 Местность, K2 Аварийность, TARIF ТарифПолный');
                    Add(',RET2 Дата_оплаты, PAY2SUMMA Сумма_1го_взноса, PAY2CURR Валюта, AGENTCODE Код_агента');
                    Add('FROM MANDATOR');
                    Add('WHERE AGENTCODE IS NULL AND ISPAY2=1');
                    Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
                    Add('AND BASECARCODE=101');
                    Add('AND RET2 BETWEEN ' + BTWDATE);

                    Add('UNION ALL ');
                    Add('SELECT ''Доплата за такси'' Оплата, SERIA Cерия, NUMBER Номер, NAME Страхователь, FROMDATE C, TODATE ПО, PERIOD Период, TARIFGLASS Тариф, K1 Местность, K2 Аварийность, TARIF ТарифПолный');
                    Add(',D Дата_оплаты, SUMMA Сумма_1го_взноса, CURR Валюта, AGENTCODE Код_агента');
                    Add('FROM MANDATOR M, TAXI T');
                    Add('WHERE SERIA=SER AND NUMBER=N');
                    Add('AND (PREVSERIA IS NULL OR PREVSERIA='''')');
                    Add('AND BASECARCODE=101');
                    Add('AND D BETWEEN ' + BTWDATE);
                    Add('ORDER BY 2, 3, 1');
                    Open;
                    DatSetToExcel(WorkSQL);
                    Close;
                end;
            end;
       end;
     except
       on E : Exception do begin
          MessageDlg('Ошибка расчёта'#13 + E.Message, mtInformation, [mbOK], 0);
       end;
     end;
     Screen.Cursor := crDefault;
     StatusBar.Panels[0].Text := '';
end;
*)
procedure TMainForm.InitMenuClick(Sender: TObject);
begin
     if KorrectPoland = nil then begin
         Application.CreateForm(TKorrectPoland, KorrectPoland);
     end;

     KorrectPoland.DbName := DB.DatabaseName;
     KorrectPoland.TableName := GetTableName();
     KorrectPoland.AgentCodeFld := 'AgentCode';
     KorrectPoland.AgentPercentFld := 'AgPercent';
     KorrectPoland.UridichFizichFld := 'InsComp';
     KorrectPoland.IsUridichFldStr := true;
     KorrectPoland.UR_FIZ_VLA := 'UF';
     KorrectPoland.NULLFld := 'PayDate';
     KorrectPoland.RegDateFld := 'RegDate';
     KorrectPoland.ShowModal;

     if MainQuery.Active then
         MainQuery.Refresh
end;

procedure TMainForm.ReservMenuClick(Sender: TObject);
var
    stopdate, todt, frdt : TField;
    getPolises : TQuery;
    PolisCounter : integer;
    PAYCURR : string;
    rs : array[0..9] of ReportStructure;
    f_i : integer;
    dataptr : ^ReportStructure;
    BruttoPremium : double;
    Otchislen, BRBSum, AgentSumma, bp : double;
    AllPeriod : integer;
    AgType : string;
    WorkPeriod : integer;
    PercentPay : double;
    rnp : double;
begin
  {  PolisCounter := 0;
    InitRezerv.RestoreName('ОСТС');
    if InitRezerv.ShowModal <> mrOk then exit;

    getPolises := WorkSQL;
    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT * FROM MANDATOR WHERE PAY1DT IS NOT NULL AND (PSER IS NULL OR PSER='') AND (PAY1DT <= :REPDT AND TODT >= :REPDT))');
    WorkSQL.ParamByName('REPDT').AsDate := InitRezerv.RepData.Date;
    WorkSQL.Open;
    while(not WorkSQL.Eof) do begin
        stopdate := getPolises.FieldByName('STOPDATE');
        if(not stopdate.IsNull) then
            if(stopdate.AsDateTime <= InitRezerv.RepData.Date) then begin
                WorkSQL.Next;
                continue;
            end;

        //if(f)
        //    fprintf(f, "%s/%lu ", (LPCTSTR)getPolises["SER'), (int)getPolises["NMB'));

        Inc(PolisCounter);

        if((PolisCounter MOD 100) = 0) then begin
            StatusBar.SimpleText := 'Обработано ' + IntToStr(PolisCounter);
            StatusBar.Update;
        end;

        PAYCURR := getPolises.FieldByName('PAY1CURR').AsString;
        dataptr := nil;
        for f_i := 0 to 10 do begin
            if(rs[f_i].Curr = '') OR (PAYCURR = rs[f_i].Curr) then begin
                dataptr := @rs[f_i];
                dataptr.Curr := PAYCURR;
                break;
            end
        end;

        todt := getPolises.FieldByName('TODT');
        frdt := getPolises.FieldByName('FRDT');
        //sqldate pay1dt = getPolises["PAY1DT');
        AllPeriod := round(todt.AsDateTime - frdt.AsDateTime + 1);

        //if(f)
        //    fprintf(f, "до %s ", (const char* )todt.to_string());

        if(PAYCURR = '') then
            raise Exception.Create('Не определена валюта оплаты ' + getPolises.FieldByName('SER').AsString + '/' + IntToStr(getPolises.FieldByName('NMB').AsInteger));

        if(PAYCURR <> 'BRB') then begin
            //Валюта
            //валютное отчисление в фонд
            BruttoPremium := getPolises.FieldByName('PAY1').AsFloat;
            Otchislen := BruttoPremium * GetOneTax(InitRezerv.List, getPolises.FieldByName('PAY1DT').AsDateTime) / 100;
            dataptr.FOND_PREMUIM := dataptr.FOND_PREMUIM + Otchislen;

            //if(f)
            //    fprintf(f, "1пл. %0.2f %s, ", BruttoPremium, PAYCURR);

            //% продажи
            PercentPay := dlg.GetSellPercent(getPolises.FieldByName('PAY1DT'), (LPCTSTR)getPolises.FieldByName('SER'), getPolises.FieldByName('NMB')) / 100;

            //if(f)
            //    fprintf(f, "Продажа %0.2f%c, ", PercentPay * 100, '%');

            BRBSum := BruttoPremium * PercentPay;
            AgentSumma := BruttoPremium * getPolises.FieldByName('AGPCNT').AsFloat / 100;
            AgType := getPolises.FieldByName('AGTYPE').AsString;
            if(AgType.GetLength() != 1 || CBlankString("FU").Find(AgType[0]) == -1)
                throw xsql("Неизвестный тип агента " + CBlankString(getPolises.FieldByName('SER')) + '/' + LongToString(getPolises.FieldByName('NMB')));

            if(AgType[1] == 'U')
                AgentSumma = AgentSumma * (1 + dlg.m_UrTax / 100);
            else
                AgentSumma = AgentSumma * (1 + dlg.m_FizTax / 100);

            //if(f)
            //    fprintf(f, "Агент %0.2f, ", AgentSumma);

            BruttoPremium *= (1 - PercentPay);
            dataptr.BRUTTO_PREMUIM := dataptr.BRUTTO_PREMUIM + BruttoPremium;

            bp := BruttoPremium - Otchislen;
            dataptr.BASE_PREMUIM := dataptr.BASE_PREMUIM + bp;

            //if(f)
            //    fprintf(f, "БП %0.2f, ", bp);

            if(true {AllPeriod > 30) begin
                WorkPeriod := todt.AsDateTime - dlg.m_RepDate.m_dt;
                if WorkPeriod < 0 then rnp := bp
                else rnp := bp * max(0, WorkPeriod) / AllPeriod;
                dataptr.RNP2_PREMUIM := dataptr.RNP2_PREMUIM + rnp;

                //if(f)
                //    fprintf(f, "РНП %0.2f, ", rnp);
            end

            Rate := GetRateCurrency(getPolises.FieldByName('PAY1DT'), PAYCURR);
            //if(f)
            //    fprintf(f, "КУРС %0.2f, ", Rate);

            PAYCURR := 'BRB';
            dataptr := nil;
            for f_i := 0 to 10 do begin
                if(rs[f_i].Curr = '') OR (PAYCURR = rs[f_i].Curr) begin
                    dataptr = @rs[f_i];
                    dataptr.Curr := PAYCURR;
                    break;
                end
            end
            //ASSERT(dataptr);

            if(getPolises.FieldByName('ISPAY2').AsString == 'Y' && getPolises.FieldByName('PAY2CURR').AsString != 'BRB')
                raise Exception.Create('Валютная оплата в 2 этапа " + getPolises.FieldByName('SER').AsString + '/' + IntToStr(getPolises.FieldByName('NMB').AsInteger));

            bp := 0.;
            if(getPolises.FieldByName('ISPAY2').AsString == 'Y' begin
                double lBruttoPremium = (double)getPolises.FieldByName('PAY2');
                double lOtchislen = lBruttoPremium * dlg.GetFondPercent(getPolises.FieldByName('PAY2DT'), (LPCTSTR)getPolises.FieldByName('SER'), getPolises.FieldByName('NMB')) / 100.;
                double lAgSumma = lBruttoPremium * (double)getPolises.FieldByName('AGPCNT') / 100.;
                if(AgType[0] == 'U')
                    lAgSumma = lAgSumma * (1 + dlg.m_UrTax / 100);
                else
                    lAgSumma = lAgSumma * (1 + dlg.m_FizTax / 100);
                dataptr.BRUTTO_PREMUIM += lBruttoPremium; //полученные рубли
                dataptr.FOND_PREMUIM += lOtchislen; //отчисление в фонд
                dataptr.AGENT_PREMUIM += lAgSumma; //процент агенту
                bp = lBruttoPremium - lOtchislen - lAgSumma;
                dataptr.BASE_PREMUIM += bp; //базовая

                //if(f)
                //    fprintf(f, "2й пл. %0.2f, ", lBruttoPremium);
                //if(f)
                //    fprintf(f, "Фонды %0.2f, ", lOtchislen);
                //if(f)
                //    fprintf(f, "Агент %0.2f, ", lAgSumma);
                //if(f)
                //    fprintf(f, "БП %0.2f, ", bp);
            end

            dataptr.BRUTTO_PREMUIM := dataptr.BRUTTO_PREMUIM + BRBSum * Rate;
            dataptr.AGENT_PREMUIM := dataptr.AGENT_PREMUIM + AgentSumma * Rate;
            bp := bp + (BRBSum - AgentSumma) * Rate;
            dataptr.BASE_PREMUIM := dataptr.BASE_PREMUIM + bp;

            if(true {AllPeriod > 30) begin
                WorkPeriod := getDate(todt).m_dt - dlg.m_RepDate.m_dt;
                double rnp := (dlg.m_RepDate.m_dt < getDate(frdt).m_dt) ? bp : (bp * max(0, WorkPeriod) / AllPeriod);
                dataptr.RNP2_PREMUIM := dataptr.RNP2_PREMUIM + rnp;
                //if(f)
                //    fprintf(f, "РНП %0.2f, ", rnp);
            end
        end

        if(PAYCURR == 'BRB') begin
            double BruttoPremium = getPolises.FieldByName('PAY1');
            double BruttoPremium1 = getPolises.FieldByName('PAY1');
            double BruttoPremium2 = 0;
            double FondPercent2 = 0;

            if(f)
                fprintf(f, "1пл. %0.2f, ", BruttoPremium);

            double PAY2DUP = 0;

            if(int(getPolises.FieldByName('STATE')) == STATE_POLIS_LOST) //Утерян
                if(CBlankString(getPolises.FieldByName('ISFEE2')) == "Y") //Нужна оплата
                    if(CBlankString(getPolises.FieldByName('ISPAY2')) == "N") begin //Нет оплаты
                        sqldate fee2date = getPolises.FieldByName('FEE2DT');
                        if(fee2date.is_null())
                            throw xsql("Нет даты 2й оплаты " + CBlankString(getPolises.FieldByName('SER')) + '/' + LongToString(getPolises.FieldByName('NMB')));
                        if(getDate(fee2date).m_dt < dlg.m_RepDate.m_dt) begin //Должна быть!!!
                            SQLQuery getDupPolis(RateCur, "SELECT * FROM SYSADM.MANDATOR WHERE PSER=:1 AND PNMB=:2");
                            getDupPolis << CBlankString(getPolises.FieldByName('SER'))
                                        << int(getPolises.FieldByName('NMB'));
                            if(getDupPolis.StartFetch()) begin //Есть дубликат
                                if(int(getDupPolis["STATE')) != STATE_POLIS_STOPPAY) begin
                                    if(CBlankString(getDupPolis["ISPAY2')) == "Y") begin
                                        PAY2DUP = getDupPolis["PAY2');
                                        FondPercent2 = dlg.GetFondPercent(getDupPolis["PAY2DT'), (LPCTSTR)getDupPolis["SER'), getDupPolis["NMB')) / 100;
                                        if(PAY2DUP < 0.01)
                                            throw xsql("Нет 2й оплаты " + CBlankString(getDupPolis["SER')) + '/' + LongToString(getDupPolis["NMB')));
                                    end
                                    else
                                    if(int(getDupPolis["STATE')) == STATE_POLIS_LOST)
                                        NotSupportPolis++;
                                end
                            end
                            else
                                NotFoundDup++;
                        end
                    end

            //check 2 pay
            if(CBlankString(getPolises.FieldByName('ISPAY2')) == "Y") begin
                if(dlg.m_RepDate.m_dt > getDate(getPolises.FieldByName('PAY2DT'))) begin
                    if(CBlankString(getPolises.FieldByName('PAY2CURR')) != "BRB")
                        throw xsql("Проблема со 2й оплатой (она валютная) " + CBlankString(getPolises.FieldByName('SER')) + '/' + LongToString(getPolises.FieldByName('NMB')));
                    BruttoPremium += (double)getPolises.FieldByName('PAY2');
                    BruttoPremium2 = (double)getPolises.FieldByName('PAY2');
                    FondPercent2 = dlg.GetFondPercent(getPolises.FieldByName('PAY2DT'), (LPCTSTR)getPolises.FieldByName('SER'), getPolises.FieldByName('NMB')) / 100;
                end
            end

            if(BruttoPremium2 < 0.01) begin
                BruttoPremium2 = PAY2DUP;
                BruttoPremium := BruttoPremium + PAY2DUP;
            end

            //if(f && BruttoPremium2 > 1)
            //    fprintf(f, "2пл. %0.2f, ", BruttoPremium2);

            dataptr.BRUTTO_PREMUIM := dataptr.BRUTTO_PREMUIM + BruttoPremium;

            AgentSumma := BruttoPremium * (double)getPolises.FieldByName('AGPCNT') / 100;
            AgType := getPolises.FieldByName('AGTYPE');
            if(AgType.GetLength() != 1 || CBlankString("FU").Find(AgType[0]) == -1)
                throw xsql("Неизвестный тип агента " + CBlankString(getPolises.FieldByName('SER')) + '/' + LongToString(getPolises.FieldByName('NMB')));

            if(AgType[0] == 'U')
                AgentSumma = AgentSumma * (1 + dlg.m_UrTax / 100);
            else
                AgentSumma = AgentSumma * (1 + dlg.m_FizTax / 100);

            dataptr.AGENT_PREMUIM += AgentSumma;

            if(f && AgentSumma > 1)
                fprintf(f, "агент %0.2f, ", AgentSumma);

            double Otchislen = BruttoPremium1 * dlg.GetFondPercent(getPolises.FieldByName('PAY1DT'), (LPCTSTR)getPolises.FieldByName('SER'), getPolises.FieldByName('NMB')) / 100;
            if(BruttoPremium2 > 0)
                Otchislen += BruttoPremium2 * FondPercent2;// dlg.GetFondPercent(getPolises.FieldByName('PAY2DT'), (LPCTSTR)getPolises.FieldByName('SER'), getPolises.FieldByName('NMB')) / 100;
            dataptr.FOND_PREMUIM += Otchislen;

            if(f && Otchislen > 1)
                fprintf(f, "фонды %0.2f, ", Otchislen);

            bp := (BruttoPremium - AgentSumma - Otchislen);
            dataptr.BASE_PREMUIM := dataptr.BASE_PREMUIM + bp;

            //if(f)
            //    fprintf(f, "БП %0.2f, ", bp);

            if(true {AllPeriod > 30) begin
                WorkPeriod := getDate(todt).m_dt - dlg.m_RepDate.m_dt + 1;
                if (dlg.m_RepDate.m_dt < getDate(frdt).m_dt) then rnp := bp
                else rnp := (bp * max(0, WorkPeriod) / AllPeriod);
                dataptr.RNP2_PREMUIM := dataptr.RNP2_PREMUIM + rnp;
                //if(f)
                //    fprintf(f, "РНП2 %lu/%lu = %0.2f\r", WorkPeriod, AllPeriod, rnp);
            end
        end
    end;
     Screen.Cursor := crDefault;
     }
end;

procedure TMainForm.AvForPolisesClick(Sender: TObject);
begin
     if AvForm = nil then begin
        Application.CreateForm(TAvForm, AvForm);
        AvForm.ShowModal;
     end;
end;

procedure TMainForm.MagazineLossClick(Sender: TObject);
var
     Line : integer;
     LineStr : String;
     OutFileName : String;
     Val : VARIANT;
     INI : TIniFile;
     AvState : string;
     MaxDate : TDateTime;

     DocsPay, DocsNoPay : integer;
     ASH : VARIANT;
begin
     DocsPay := 0;
     DocsNoPay := 0;

     FilterForUbitok.DateFrom.Enabled := true;
     FilterForUbitok.DateTo.Enabled := true;
     FilterForUbitok.Caption := 'Журнал убытков';
     FilterForUbitok.CheckedDate1 := true;
     FilterForUbitok.CheckedDate2 := true;
     if FilterForUbitok.ShowModal = mrCancel then exit;

     Line := 9;

     INI := TIniFile.Create(BLANK_INI);
     try
       Screen.Cursor := crHourGlass;
       with WorkSQL, WorkSQL.SQL do begin
            Clear;
            Add('SELECT M.SERIA, M.NUMBER, M.REGDATE, M.NAME, M.MARKA,');
            Add('MA.NWORK, MA.WRDATE, MA.DtFreeRz, MA.DtClLoss, MA.LossCode, MA.WHAT, ');
            Add('MA.DeclTC, MA.DeclHL, MA.DeclIM, ');
            Add('MA.SumPay, V1_Date, V2_Date, V3_Date, V1');
            Add('FROM MANDAV2 MA, MANDATOR M');
            Add('WHERE M.SERIA=MA.SERIA');
            Add('AND BASECARCODE=101');
            Add('AND M.NUMBER=MA.NUMBER');
            if FilterForUbitok.DateFrom.Checked then
                Add('AND MA.WRDATE>=''' + DateToStr(FilterForUbitok.DateFrom.Date) + '''');
            if FilterForUbitok.DateTo.Checked then
                Add('AND MA.WRDATE<=''' + DateToStr(FilterForUbitok.DateTo.Date) + '''');

            Add('ORDER BY V1,MA.WRDATE');

            Open;

            if EOF then begin
               MessageDlg('Не данных', mtInformation, [mbOk], 0);
               exit;
            end;

            Application.CreateForm(TWaitForm, WaitForm);
            WaitForm.Update;

            if not InitExcel(MSEXCEL) then exit;
            MSEXCEL.Visible := true;

            MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\ЖУРНАЛ УБЫТКОВ.XLS');
            ASH := MSEXCEL.ActiveSheet;

            while not EOF do begin
                AvState := INI.ReadString(MAND_SECT, 'AVDesision' + FieldByName('WHAT').AsString, '');
                AvState := AnsiUpperCase(AvState);

                if (Pos('ОТКАЗ', AvState) = 0) AND
                   (Pos('ОПЛАЧ', AvState) = 0) AND
                   (Pos('ЗАЯВЛ', AvState) = 0) then begin
                   Application.BringToFront;
                   if MessageDlg('Неизвестное состояние аварии ' + AvState + ' в полисе ' + FieldByName('NUMBER').AsString + '. Пропустить?', mtInformation, [mbYes, mbNo], 0) = mrNo then
                      break;
                end;

                if Pos('ЗАЯВЛ', AvState) <> 0 then begin
                   Next;
                   continue;
                end;

                LineStr := IntToStr(Line);
                ASH.Range['A' + LineStr, 'A' + LineStr].Value := FieldByName('NWORK').AsString;
                ASH.Range['B' + LineStr, 'B' + LineStr].Value := FieldByName('WRDATE').AsString;
                ASH.Range['C' + LineStr, 'C' + LineStr].Value := 'Заявление';
                ASH.Range['D' + LineStr, 'D' + LineStr].Value := FieldByName('WRDATE').AsString;
                ASH.Range['E' + LineStr, 'E' + LineStr].Value := '?';
                ASH.Range['F' + LineStr, 'F' + LineStr].Value := FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                ASH.Range['G' + LineStr, 'G' + LineStr].Value := FieldByName('REGDATE').AsString;
                ASH.Range['H' + LineStr, 'H' + LineStr].Value := FieldByName('NAME').AsString;
                ASH.Range['I' + LineStr, 'I' + LineStr].Value := FieldByName('MARKA').AsString;
                ASH.Range['J' + LineStr, 'J' + LineStr].Value := FilterForUbitok.ValCode.Text;
                ASH.Range['K' + LineStr, 'K' + LineStr].Value := FilterForUbitok.InsSumma.Text;
                Val := FieldByName('DeclTC').AsFloat +
                       FieldByName('DeclHL').AsFloat +
                       FieldByName('DeclIM').AsFloat;
                if FieldByName('V1').AsString <> '' then
                    ASH.Range['L' + LineStr, 'L' + LineStr].Value := FieldByName('V1').AsString
                else
                    ASH.Range['L' + LineStr, 'L' + LineStr].Value := 'BRB';

                ASH.Range['M' + LineStr, 'M' + LineStr].Value := Val;

                ASH.Range['N' + LineStr, 'N' + LineStr].Value := FieldByName('LOSSCODE').AsString;

                //AvState := INI.ReadString('MANDATORY', 'AvariaState' + FieldByName('What').AsString, '');
                //AvState := AnsiUpperCase(AvState);

                //if (Pos('ОТКАЗ', AvState) == 0) AND
                //   (Pos('ОПЛАЧ', AvState) == 0) AND
                //   (Pos('ЗАЯВЛ', AvState) == 0) then begin
                //end;

                //AvariaState0=ЗАЯВЛЕНО
                //AvariaState1=ОТКАЗАНО
                //AvariaState2=ОПЛАЧЕНО

                //Оплата
                if Pos('ОПЛАЧ', AvState) <> 0 then begin
                    Inc(DocsPay);
                    MaxDate := FieldByName('V1_Date').AsDateTime;
                    if MaxDate < FieldByName('V2_Date').AsDateTime then
                        MaxDate := FieldByName('V2_Date').AsDateTime;
                    if MaxDate < FieldByName('V3_Date').AsDateTime then
                        MaxDate := FieldByName('V3_Date').AsDateTime;

                    if MaxDate > EncodeDate(1999, 1, 1) then
                        MSEXCEL.ActiveSheet.Range['O' + LineStr, 'O' + LineStr].Value := DateToStr(MaxDate);

                    if FieldByName('V1').AsString <> '' then
                        ASH.Range['P' + LineStr, 'P' + LineStr].Value := FieldByName('V1').AsString
                    else
                        ASH.Range['P' + LineStr, 'P' + LineStr].Value := 'BRB';
                    Val := FieldByName('SumPay').AsFloat;
                    ASH.Range['Q' + LineStr, 'Q' + LineStr].Value := Val;
                    ASH.Range['R' + LineStr, 'R' + LineStr].Value := Val;
                end;

                //Отказ
                if Pos('ОТКАЗ', AvState) <> 0 then begin
                    Inc(DocsNoPay);
                    ASH.Range['S' + LineStr, 'S' + LineStr].Value := FieldByName('DtClLoss').AsString;;
                    if FieldByName('V1').AsString <> '' then
                        ASH.Range['T' + LineStr, 'T' + LineStr].Value := FieldByName('V1').AsString
                    else
                        ASH.Range['T' + LineStr, 'T' + LineStr].Value := 'BRB';
                    ASH.Range['U' + LineStr, 'U' + LineStr].Value := Val;
                    ASH.Range['V' + LineStr, 'V' + LineStr].Value := Val;
                end;

                ASH.Range['W' + LineStr, 'W' + LineStr].Value := FieldByName('DtClLoss').AsString;

                Line := Line + 1;

                Next;
            end;

            LineStr := IntToStr(Line);
            //Требует разбивки по суммам
            //MSEXCEL.ActiveSheet.Range['M' + LineStr, 'M' + LineStr].Value := '=SUM(M9:M' + IntToStr(Line - 1) + ')';
            //MSEXCEL.ActiveSheet.Range['Q' + LineStr, 'Q' + LineStr].Value := '=SUM(Q9:Q' + IntToStr(Line - 1) + ')';
            //MSEXCEL.ActiveSheet.Range['R' + LineStr, 'R' + LineStr].Value := '=SUM(R9:R' + IntToStr(Line - 1) + ')';
            //MSEXCEL.ActiveSheet.Range['U' + LineStr, 'U' + LineStr].Value := '=SUM(U9:U' + IntToStr(Line - 1) + ')';
            //MSEXCEL.ActiveSheet.Range['V' + LineStr, 'V' + LineStr].Value := '=SUM(V9:V' + IntToStr(Line - 1) + ')';

            LineStr := IntToStr(Line + 2);
            MSEXCEL.ActiveSheet.Range['A' + LineStr, 'A' + LineStr].Value := 'ОПЛАЧЕННО ' + IntToStr(DocsPay);
            LineStr := IntToStr(Line + 3);
            MSEXCEL.ActiveSheet.Range['A' + LineStr, 'A' + LineStr].Value := 'ОТКАЗАНО ' + IntToStr(DocsNoPay);

            OutFileName := GetAppDir + 'DOCS\_ЖУРНАЛ УБЫТКОВ.XLS';
            DeleteFile(OutFileName);
            MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
       end;
     except
       on E : Exception do begin
            Application.BringToFront;
            if WaitForm <> nil then WaitForm.Close;
            MessageDlg('Возникла ошибка при формировании отчёта'#13 + e.Message, mtInformation, [mbOk], 0);
       end
     end;

     Application.BringToFront;
     INI.Free;
     Screen.Cursor := crDefault;
     if WaitForm <> nil then WaitForm.Close;
end;

procedure TMainForm.MagazineDocsClick(Sender: TObject);
var
     Line : integer;
     LineStr : string;
     TempStr : string;
     OutFileName : string;
     ASH : VARIANT;
     LastLine : integer;
     ArrayData : VARIANT;
begin
     if MagazineForm.ShowModal <> mrOk then exit;

     if not InitExcel(MSEXCEL) then exit;
     MSEXCEL.Visible := true;
     MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\ЖУРНАЛ УЧЁТА ЗАКЛЮЧЁННЫХ ДОГОВОРОВ.XLS');
     ASH := MSEXCEL.ActiveSheet;
     ASH.Range['C4','C4'].Value := 'Обязательное страхование тр. средств';
     try
           WorkSQL.SQL.Clear;
           WorkSQL.SQL.Add('SELECT * FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND (PREVSERIA IS NULL OR PREVSERIA='''') AND (FROMDATE <= :D1 AND TODATE >= :D1 OR (TODATE - FROMDATE + 1) <= 30 AND PAYDATE BETWEEN (:D2) AND :D1) ORDER BY SERIA,NUMBER');
           WorkSQL.ParamByName('D1').AsdateTime := MagazineForm.RepDate.Date;
           WorkSQL.ParamByName('D2').AsdateTime := IncMonth(MagazineForm.RepDate.Date, -1);
           WorkSQL.Open;

           //MainGrid.DataSource := nil;
           //MainQuery.First;
           Line := 22;
           ArrayData := VarArrayCreate([1, 1, 1, 18], varVariant);
           while not WorkSQL.EOF do begin
                 //Не испорчен и не дубликат
                 with WorkSQL do begin
                       ArrayData[1, 1] := FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString;
                       ArrayData[1, 2] := FieldByName('FROMDATE').AsString;
                       ArrayData[1, 3] := FieldByName('NAME').AsString;
                       ArrayData[1, 4] := FieldByName('MARKA').AsString;
                       ArrayData[1, 5] := MagazineForm.OtvSumma.Text;
                       ArrayData[1, 6] := MagazineForm.OtvCurr.Text;
                       ArrayData[1, 7] := FieldByName('FROMDATE').AsString;
                       ArrayData[1, 8] := FieldByName('TODATE').AsString;

                       TempStr := FieldByName('PAYDATE').AsString;
                       if FieldByName('Draftnumber').AsFloat > 0.5 then
                           TempStr := TempStr + ', ' + FieldByName('Draftnumber').AsString;
                       ArrayData[1, 9] := TempStr;
                       ArrayData[1, 10] := FieldByName('Pay1Curr').AsString;
                       ArrayData[1, 11] := FieldByName('Pay1').AsFloat;
                       if FieldByName('Pay1Curr').AsString = 'BRB' then
                           ArrayData[1, 12] := FieldByName('Pay1').AsFloat
                       else
                           ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Pay1Curr').AsString, FieldByName('PAYDATE').AsDateTime) * FieldByName('Pay1').AsFloat;

                       if (FieldByName('IsPay2').AsInteger = 1) AND (FieldByName('PAY2DATE').AsDateTime <= MagazineForm.RepDate.Date) then begin
                           //перекинуть в итог
                           ArrayData[1, 13] := FieldByName('Pay1').AsFloat + FieldByName('Pay2Summa').AsFloat;
                           if FieldByName('Pay2Curr').AsString = 'BRB' then
                               ArrayData[1, 14] := FieldByName('Pay1').AsFloat + FieldByName('Pay2Summa').AsFloat
                           else
                               ArrayData[1, 14] := double(ArrayData[1, 12]) + GetRateOfCurrency(FieldByName('Pay2Curr').AsString, FieldByName('PAY2DATE').AsDateTime) * FieldByName('Pay2Summa').AsFloat;
                           //Заполнить последний
                           TempStr := FieldByName('PAY2DATE').AsString;
                           if FieldByName('Pay2Draftnumber').AsFloat > 0.5 then
                               TempStr := TempStr + ', ' + FieldByName('pay2Draftnumber').AsString;
                           ArrayData[1, 9] := TempStr;
                           ArrayData[1, 10] := FieldByName('Pay2Curr').AsString;
                           ArrayData[1, 11] := FieldByName('Pay2Summa').AsFloat;
                           if FieldByName('Pay2Curr').AsString = 'BRB' then
                               ArrayData[1, 12] := FieldByName('Pay2Summa').AsFloat
                           else
                               ArrayData[1, 12] := GetRateOfCurrency(FieldByName('Pay2Curr').AsString, FieldByName('PAY2DATE').AsDateTime) * FieldByName('Pay2Summa').AsFloat;
                       end
                       else begin
                           ArrayData[1, 13] := ArrayData[1, 11];
                           ArrayData[1, 14] := ArrayData[1, 12];
                       end;

                       ArrayData[1, 15] := double(ArrayData[1, 11]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 16] := double(ArrayData[1, 12]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 17] := double(ArrayData[1, 13]) * FieldByName('AgPercent').AsFloat / 100;
                       ArrayData[1, 18] := double(ArrayData[1, 14]) * FieldByName('AgPercent').AsFloat / 100;
                       LastLine := Line;
                 end;
                 ASH.Range['A' + InttoStr(Line), 'R' + InttoStr(Line)] := ArrayData;
                 Inc(Line);
                 WorkSQL.Next;
           end;

           Inc(Line);
           LineStr := IntToStr(Line);
           ASH.Range['L' + LineStr, 'L' + LineStr].Value := '=SUM(L22:L' + IntToStr(LastLine) + ')';
           ASH.Range['N' + LineStr, 'N' + LineStr].Value := '=SUM(N22:N' + IntToStr(LastLine) + ')';
           ASH.Range['P' + LineStr, 'P' + LineStr].Value := '=SUM(P22:P' + IntToStr(LastLine) + ')';
           ASH.Range['R' + LineStr, 'R' + LineStr].Value := '=SUM(R22:R' + IntToStr(LastLine) + ')';

           VarClear(ArrayData);
           OutFileName := GetAppDir + 'DOCS\_ЖУРНАЛ УЗД.XLS';
           DeleteFile(OutFileName);
           MSEXCEL.ActiveWorkbook.SaveAs(OutFileName, 1);
     except
           on E : Exception do
              MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
    //MainGrid.DataSource := MainDataSource;
end;

procedure TMainForm.MainQueryInsCompGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
     if Sender.AsString = 'U' then
         Text := 'Юридич.'
     else
     if Sender.AsString = 'F' then
         Text := 'Физич.'
     else
         Text := '?'
end;

procedure TMainForm.AvariaListClick(Sender: TObject);
begin
     if ListAvarias = nil then begin
        Application.CreateForm(TListAvarias, ListAvarias);
        ListAvarias.ShowModal;
     end;
end;

procedure TMainForm.Excel1Click(Sender: TObject);
begin
    if not MainQuery.Active then
        FilterMnClick(nil);
    DatSetToExcel(MainQuery)
end;

{procedure TMainForm.N14Click(Sender: TObject);
begin
  if MessageDlg('Вы сделали копию данных??', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
      exit;
  try
    Screen.Cursor := crHourglass;
    with WorkSQL, WorkSQL.SQL do begin
        Clear;
        Add('ALTER TABLE MANDATOR');
        Add('ADD RET1 DATE');
        ExecSQL;
        Clear;
        Add('ALTER TABLE MANDATOR');
        Add('ADD RET2 DATE');
        ExecSQL;
        Clear;
        Add('UPDATE MANDATOR');
        Add('SET RET1=PAYDATE, RET2=PAY2DATE');
        ExecSQL;
        MessageDlg('Всё хорошо! Выйдите из программы, откройте каталог с Базой Данных и запустите программу PX2CYRR.EXE', mtInformation, [mbOk], 0);
    end;
  except
    on E : Exception do begin
      MessageDlg('Ошибка!!!' + E.Message, mtError, [mbOk], 0);
    end;
  end;
  Screen.Cursor := crDefault;
end;
 }
procedure TMainForm.RunReportClick(Sender: TObject);
var
    IsFound : boolean;
    i : integer;
begin
    if not PrepareHiLo(RptTbl, WorkSQL, false, ProgressBar) then exit;

    OpenExcel2C(MSEXCEL);
{    //Запуск экселя открытие файла
    if not InitExcel(MSEXCEL) then begin
        MessageDlg('Excel ен найден!!!', mtInformation, [mbOk], 0);
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
        MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\2с.xls');  }
    RunReport(MSEXCEL, TMenuItem(Sender).Tag, WorkSQL);
    Application.BringToFront;
end;

procedure TMainForm._2CClick(Sender: TObject);
var
    Name : string;
    i, n : integer;
    Item : TMenuItem;
    INI : TINIFile;
begin
    INI := TINIFile.Create('reports.ini');
    n := 0;
    _2C.Clear;
    for i := 1 to 100 do begin
        Name := INI.ReadString('REPORT' + IntToStr(i), 'Name', '');
        if Name <> '' then begin
            Item := TMenuItem.Create(Self);
            Item.Tag := i;
            Item.OnClick := RunReportClick;
            Item.Caption := Name;
            Item.Tag := i;
            _2C.Add(Item);
            n := n + 1;
        end;
    end;

    INI.Free;

    if _2C.Count = 0 then begin
        Item := TMenuItem.Create(Self);
        Item.Caption := 'Не заполнен файл reports.ini';
        _2C.Add(Item);
    end;
end;

function TMainForm.ExecQuery(SQLText : string; ASH : Variant; Cell : string; D1, D2 : TDateTime) : double;
var
   Value : double;
   DT : TDateTime;
begin
    with WorkSQL, WorkSQL.SQL do begin
        Close;
        Clear;
        Add(SQLText);
        if GetFilter('') <> '' then
            Add(' AND ' + GetFilter(''));
        if D1 <> 0 then
            ParamByName('D1').AsDateTime := D1;
        ParamByName('D2').AsDateTime := D2;
        Open;
        Value := 0;
        while not Eof do begin
            if Fields.Count > 1 then begin
                Dt := 0;
                if not Fields[2].IsNull then Dt := Fields[2].AsDateTime;
                if not Fields[3].IsNull then Dt := Fields[3].AsDateTime;
                if not Fields[4].IsNull then Dt := Fields[4].AsDateTime;
                if Dt = 0 then
                    MessageDlg('Не определена дата выплаты', mtInformation, [mbOk], 0);
                Value := Value + FieldByName('VAL').AsFloat * GetRateOfCurrency(Fields[1].AsString, Dt)
            end    
            else
                Value := Value + FieldByName('VAL').AsFloat;
            Next;
        end;
        if Cell <> '' then
            ASH.Range[Cell, Cell].Value := VARIANT(Value);
        ExecQuery := Value;
        Close;
    end;
end;

procedure TMainForm.RepItogClick(Sender: TObject);
var
    ASH : VARIANT;
    StartYear : TDateTime;
    D, M, Y : WORD;
    flt : string;
    D1, D2 : TDateTime;
    Summa : double;
begin
    flt := GetFilter('');

    if flt <> '' then
        MessageDlg('У вас установлен фильтр. отчёт будет создан в соответствии с указанным фильтром!', mtInformation, [mbOk], 0);

    if GetTwoDates.ShowModal <> mrOk then exit;

    UpdateDupdata('');

    MessageDlg('Расчёт будет выполнен на основе даты отчёта', mtInformation, [mbOk], 0);

    if not InitExcel(MSEXCEL) then exit;
    MSEXCEL.Visible := true;

    MSEXCEL.WorkBooks.Open(GetAppDir + 'TEMPLATES\ИТОГОВЫЙОСТС.XLS');
    ASH := MSEXCEL.ActiveSheet;

    DecodeDate(GetTwoDates.StartDate.DateTime, Y, M, D);
    StartYear := EncodeDate(Y, 1, 1);
    D1 := GetTwoDates.StartDate.DateTime;
    D2 := GetTwoDates.EndDate.DateTime;

    ASH.Range['C4', 'C4'].Value := DateToStr(GetTwoDates.StartDate.DateTime);
    ASH.Range['C5', 'C5'].Value := DateToStr(GetTwoDates.EndDate.DateTime);

    //Страховая сумма по заключенным договорам
    //Сумма ответственности по действующим договорам

    try
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND RET1 BETWEEN :D1 AND :D2 '+
                  'AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'D8',
                  D1, D2
                  );
        ExecQuery('SELECT 10000*RATE AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND RET1 BETWEEN :D1 AND :D2 '+
                  'AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'D19',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND RET1 BETWEEN :D1 AND :D2 '+
                  'AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'E8',
                  StartYear, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND RET1 BETWEEN :D1 AND :D2 '+
                  'AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)'+
                  'AND (STATE IN (0,3,4) AND TODATE > :D2 '+
                  'OR STATE IN (5) AND FEE2DATE > :D2 '+
                  'OR STATE IN (2) AND STOPDATE > :D2 )',
                  ASH, 'D9',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND RET1 BETWEEN :D1 AND :D2 '+
                  'AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)'+
                  'AND (STATE IN (0,3,4) AND TODATE > :D2 '+
                  'OR STATE IN (5) AND FEE2DATE > :D2 '+
                  'OR STATE IN (2) AND STOPDATE > :D2)',
                  ASH, 'E9',
                  StartYear, D2
                  );
        //ВЫПЛАТЫ ЗА ПЕРИОД
        Summa :=
        ExecQuery('SELECT SUM(PAY1) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR=''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR=''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY1*RATE) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR<>''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*PAY2RATE) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR<>''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  D1, D2
                  );
        if not bKomplexMode then
        begin
              Summa := Summa +
              ExecQuery('SELECT SUM(SUMMA) AS VAL FROM MANDATOR, TAXI WHERE SERIA=SER AND NUMBER=N '+
                        ' AND BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR=''BRB'' AND D BETWEEN :D1 AND :D2',
                        ASH, '',
                        D1, D2
                        );
        end
        else
        begin
              Summa := Summa +
              ExecQuery('SELECT SUM(S0) AS VAL FROM KOMPLEX WHERE ' +
                        ' BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR0=''BRB'' AND D0 BETWEEN :D1 AND :D2',
                        ASH, '',
                        D1, D2
                        );
        end;
        ASH.Range['D10', 'D10'].Value := FloatToStr(round(Summa) / 1000);

        if not bKomplexMode then
        begin
            Summa :=
            ExecQuery('SELECT SUM(SUMMA) AS VAL FROM MANDATOR, TAXI WHERE SERIA=SER AND NUMBER=N '+
                      ' AND BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR<>''BRB'' AND D BETWEEN :D1 AND :D2',
                      ASH, '',
                      D1, D2
                      );
            if Summa > 0.01 then
                ASH.Range['F10', 'F10'].Value := 'За период есть валютные доплаты!!!';
        end
        else
        begin
            Summa :=
            ExecQuery('SELECT SUM(S0) AS VAL FROM KOMPLEX WHERE ' +
                      ' BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR0<>''BRB'' AND D0 BETWEEN :D1 AND :D2',
                      ASH, '',
                      D1, D2
                      );
            if Summa > 0.01 then
                ASH.Range['F10', 'F10'].Value := 'За период есть валютные доплаты!!!';
        end;

        //ВЫПЛАТЫ С НАЧАЛА ГОДА
        Summa :=
        ExecQuery('SELECT SUM(PAY1) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR=''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR=''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY1*RATE) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR<>''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*PAY2RATE) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR<>''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  StartYear, D2
                  );
        if bKomplexMode then
        begin
        Summa := Summa +
        ExecQuery('SELECT SUM(S0) AS VAL FROM KOMPLEX WHERE '+
                  ' BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR0=''BRB'' AND D0 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        end
        else
        begin
        Summa := Summa +
        ExecQuery('SELECT SUM(SUMMA) AS VAL FROM MANDATOR, TAXI WHERE SERIA=SER AND NUMBER=N '+
                  ' AND BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR=''BRB'' AND D BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        end;
        ASH.Range['E10', 'E10'].Value := FloatToStr(round(Summa) / 1000);

        if bKomplexMode then
        begin
        Summa :=
        ExecQuery('SELECT SUM(S0) AS VAL FROM KOMPLEX WHERE '+
                  ' BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR0<>''BRB'' AND D0 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        end
        else
        begin
        Summa :=
        ExecQuery('SELECT SUM(SUMMA) AS VAL FROM MANDATOR, TAXI WHERE SERIA=SER AND NUMBER=N '+
                  ' AND BASECARCODE=101 AND PAYDATE IS NOT NULL AND CURR<>''BRB'' AND D BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        end;
        if Summa > 0.01 then
            ASH.Range['G10', 'G10'].Value := 'С Начала года есть валютные доплаты!!!';

        if not bKomplexMode then
        begin
        //ВЫПЛАТЫ
        ExecQuery('SELECT COUNT(* ) AS VAL FROM MANDAV2 A,MANDATOR M '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND (V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, 'D11',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND  (V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, 'E11',
                  StartYear, D2
                  );
        //СУММЫ ВЫПЛАТ
        Summa :=
        ExecQuery('SELECT SUM(SUMPAY) AS VAL FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND V1=''BRB'' AND'+
                  '(V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUMPAY AS VAL, V1, V1_Date, V2_Date, V3_Date FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND  V1<>''BRB'' AND'+
                  '(V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, '',
                  D1, D2
                  );
        ASH.Range['D12', 'D12'].Value := VARIANT(Summa/1000);

        Summa :=
        ExecQuery('SELECT SUM(SUMPAY) AS VAL FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND  V1=''BRB'' AND'+
                  '(V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUMPAY AS VAL, V1, V1_Date, V2_Date, V3_Date FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE SUMPAY IS NOT NULL AND BASECARCODE=101 AND A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND  V1<>''BRB'' AND'+
                  '(V1_Date BETWEEN :D1 AND :D2 '+
                  'OR V2_Date BETWEEN :D1 AND :D2 '+
                  'OR V3_Date BETWEEN :D1 AND :D2)'+
                  'AND WHAT=0', //ОПЛАЧЕНО
                  ASH, '',
                  StartYear, D2
                  );
        ASH.Range['E12', 'E12'].Value := VARIANT(Summa/1000);

         //Не ЗАКРЫТЫЕ ДЕЛА
        ExecQuery('SELECT COUNT(* ) AS VAL FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND  (WRDate BETWEEN :D1 AND :D2 '+
                  'OR WRDate BETWEEN :D1 AND :D2 '+
                  'OR WRDate BETWEEN :D1 AND :D2)'+
                  'AND WHAT=2', //ЗАЯВЛЕНО
                  ASH, 'D13',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM MANDAV2 A,MANDATOR M  '+
                  'WHERE A.SERIA=M.SERIA AND A.NUMBER=M.NUMBER AND   (WRDate BETWEEN :D1 AND :D2 '+
                  'OR WRDate BETWEEN :D1 AND :D2 '+
                  'OR WRDate BETWEEN :D1 AND :D2)'+
                  'AND WHAT=2', //ЗАЯВЛЕНО
                  ASH, 'E13',
                  StartYear, D2
                  );
        end; //komplex mode for avarias
        //РАСТОРГНУТО ДОГОВОРОВ
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE STOPDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'D14',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE STOPDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'E14',
                  StartYear, D2
                  );
        //СУММЫ РАСТОРЖЕНИЯ
        ExecQuery('SELECT SUM(RETBRB)/1000 AS VAL FROM ' + GetTableName() + ' WHERE STOPDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND RETBRB>0 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'D15',
                  D1, D2
                  );
        ExecQuery('SELECT SUM(RETBRB)/1000 AS VAL FROM ' + GetTableName() + ' WHERE STOPDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND RETBRB>0 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1',
                  ASH, 'E15',
                  StartYear, D2
                  );

        //ИСПОРЧЕНО
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE REGDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND STATE=1',
                  ASH, 'D16',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE REGDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND STATE=1',
                  ASH, 'E16',
                  StartYear, D2
                  );
        //ДУБЛИКАТЫ
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE REGDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND (PREVNUMBER>0)AND STATE<>1',
                  ASH, 'D17',
                  D1, D2
                  );
        ExecQuery('SELECT COUNT(* ) AS VAL FROM ' + GetTableName() + ' WHERE REGDATE BETWEEN :D1 AND :D2 '+
                  'AND BASECARCODE=101 AND (PREVNUMBER>0)AND STATE<>1',
                  ASH, 'E17',
                  StartYear, D2
                  );

        //ВЫПЛАТЫ АГЕНТАМ =======================================
        //ЗА ПЕРИОД
        Summa :=
        ExecQuery('SELECT SUM(PAY1*AgPercent/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR=''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*carcode/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR=''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY1*RATE*AgPercent/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR<>''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  D1, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*PAY2RATE*carcode/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR<>''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  D1, D2
                  );
        ASH.Range['D18', 'D18'].Value := FloatToStr(Round(Summa) / 1000);

        //ВЫПЛАТЫ С НАЧАЛА ГОДА
        Summa :=
        ExecQuery('SELECT SUM(PAY1*AgPercent/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR=''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*carcode/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR=''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY1*RATE*AgPercent/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL AND PREVSERIA IS NULL AND PAY1CURR<>''BRB'' '+
                  'AND RET1 BETWEEN :D1 AND :D2',
                  ASH, '',
                  StartYear, D2
                  );
        Summa := Summa +
        ExecQuery('SELECT SUM(PAY2SUMMA*PAY2RATE*carcode/100) AS VAL FROM ' + GetTableName() + ' '+
                  'WHERE BASECARCODE=101 AND PREVSERIA IS NULL AND PAY2CURR<>''BRB'' '+
                  'AND RET2 BETWEEN :D1 AND :D2 '+
                  'AND ISPAY2 <> 0',
                  ASH, '',
                  StartYear, D2
                  );
        ASH.Range['E18', 'E18'].Value := FloatToStr(Round(Summa) / 1000);

        ASH.Range['D22', 'D22'].Value := VARIANT(GetRateOfCurrency('EUR', D2));
    except
      on E : Exception do begin
          MessageDlg(E.Message, mtInformation, [mbOk], 0);
      end
    end;
end;

procedure TMainForm.AlvenaClick(Sender: TObject);
begin
    if not OpenDialogPx.Execute then exit;
    ImportAlvena(OpenDialogPx.FileName);
end;

function DivCurr(S1, S2 : double) : double;
begin
    S1 := S1 / S2;
    DivCurr := (ROUND(S1 * 100)) / 100;
end;

function GetCityCode(Reg, Distr, City : string) : string;
label
    Stop;
var
    i, j : integer;
    INI : TIniFile;
    s : string;
    RegStr, CityCode : string;
    reg_found : boolean;
begin
    if lastCityKodeKey = Reg + '|' + Distr + '|' + City then begin
        GetCityCode := lastCityKodeVal;
        exit;
    end;

    lastCityKodeKey := Reg + '|' + Distr + '|' + City;

    CityCode := '';
    Reg := AnsiUpperCase(Reg);
    Distr := AnsiUpperCase(Distr);
    City := AnsiUpperCase(City);

    if City = '' then
        raise Exception.Create('Не укзан город ' + Reg + ' ' + Distr);

    reg_found := false;
    INI := TIniFile.Create(BLANK_INI);
    for i := 0 to 10 do begin
        for j := 0 to 50 do begin
            s := INI.ReadString(MAND_SECT, 'CityCode' + IntToStr(i) + '_' + IntToStr(j), '');
            s := AnsiUpperCase(s);
            if s = '' then begin
                if j = 0 then goto Stop;
            end;
            if j = 0 then begin
                if (Reg <> '') AND (Pos(Reg, s) <> 0) then
                    reg_found := true;
                continue;
            end;

            //Поиск города
            if (Copy(s, 1, Length(City) + 1) = (City + ' ')) OR
               (Copy(s, Length(s) - Length(City), 100) = (' ' + City)) OR
               (City = s) then begin
               CityCode := Copy(s, 1, Pos(',', s) - 1);
               goto Stop;
            end
            else
            if Distr <> '' then begin
                if (Pos(Distr + ' ', s) <> 0) OR
                   (Pos(' ' + Distr, s) <> 0) OR
                   (Distr = s) then begin
                   CityCode := Copy(s, 1, Pos(',', s) - 1);
                end
            end
        end;

        if reg_found then break;
    end;
    Stop:;
    if CityCode = '' then
        CityCode := INI.ReadString(MAND_SECT, 'CityCodeNF', '');
    INI.Free;

    if CityCode = '' then
        raise Exception.Create('Не могу определить код местности ' + Reg + ' ' + Distr + ' ' + City);

    lastCityKodeVal := CityCode;
    GetCityCode := CityCode;
end;

function GetPlacementK(Vid : string) : double;
var
    INI : TIniFile;
    s : string;
begin
    INI := TIniFile.Create(BLANK_INI);
    s := INI.ReadString(MAND_SECT, Vid, '');
    INI.Free;

    if s = '' then
        raise Exception.Create('Не найдено значение коэффициента местности ' + Vid);

    GetPlacementK := StrToFloat(s);
end;

function CalcSuper(Letter : string; var _TableTarif : integer; var IsResident : boolean; ShowExcept : boolean) : integer;
label
    Stop;
var
    INI : TIniFile;
    Section, s : string;
    SuperCode : integer;
    i, j : integer;
    TableTarif : integer;
begin
    INI := TIniFile.Create(BLANK_INI);

//    if Letter = 'C0' then
  //     INI:=INI;

    for TableTarif := 10 downto 2 do begin
        S := INI.ReadString(MAND_SECT, 'TarifTable_Is_Resident' + IntToStr(TableTarif), '');
        if (S = 'YES') AND (not IsResident) then continue;
        if (S = 'NO') AND (IsResident) then continue;

        S := INI.ReadString(MAND_SECT, 'TarifTable' + IntToStr(TableTarif), '');
        Section := ExtractWord(2, s, ['|']);
        SuperCode := -1;
        _TableTarif := TableTarif;

        for i := 0 to 100 do begin
            s := IntToStr(i);
            if i < 10 then s := '0' + s;
            S := INI.ReadString(MAND_SECT, Section + s + '_0', '');
            if s = '' then break;
            if s[1] = Letter[1] then begin
                //Символ из Базы длиной 1 то это надо проверить наша ли таблица
                //s := IntToStr(i);
                //if i < 10 then s := '0' + s;
                if (Length(Letter) = 1){ AND (INI.ReadString('MANDATORY', Section + s + '_1', '') = '') }then begin
                    SuperCode := (Ord(Letter[1]) SHL 24) +
                                 (i SHL 16) +
                                 0;
                    goto Stop;
                end;
                for j := 1 to 100 do begin
                    s := IntToStr(i);
                    if i < 10 then s := '0' + s;
                    S := INI.ReadString('MANDATORY', Section + s + '_' + IntToStr(j), '');
                    if S = '' then break;
                    if Pos(',', s) <> 0 then s := Copy(s, 1, Pos(',', s) - 1);
                    if Copy(Letter, 2, 100) = s then begin
                        SuperCode := (Ord(Letter[2]) SHL 24) +
                                     (i SHL 16) +
                                     j;
                        goto Stop;
                    end;
                end;
                //break;
            end
        end;
    end;

    Stop:;
    INI.Free;

    if SuperCode = -1 then begin
        s := 'Не резидент';
        if IsResident then s := 'Резидент';
        if ShowExcept then
            raise Exception.Create('Не найден тип авто ' + Letter + ' в таблице тарифов ' + IntToStr(_TableTarif) + '. проверь признак Резидента в данных. ' + s);
        IsResident := not IsResident;
        CalcSuper := CalcSuper(Letter, _TableTarif, IsResident, true);
    end;

    CalcSuper := SuperCode;
end;

function CalcCountAv(OldCl, NewCl : string) : integer;
var
    INI : TIniFile;
    s : string;
    i : integer;
begin
    CalcCountAv := -1;
    
    INI := TIniFile.Create(BLANK_INI);
    s := INI.ReadString('MANDATORY', 'ClassAvMove_' + OldCl, '');
    INI.Free;

    if s = '' then
        raise Exception.Create('Нету класса аварийности ' + OldCl);

    for i := 1 to 10 do begin
        //if i = 10 then
        //    raise Exception.Create('Нету класса аварийности ' + NewCl);

        if ExtractWord(i, s, [',']) = NewCl then begin
            CalcCountAv := i - 1;
            break;
        end;
    end
end;

function GetMobileTel(s: string) : string;
begin
    GetMobileTel := '';
    if WordCount(s, [' ']) > 0 then
    begin
        GetMobileTel := ExtractWord(1, s, [' ']);
    end;
    if WordCount(s, [' ']) > 1 then
    begin
        GetMobileTel := ExtractWord(2, s, [' ']);
    end
end;

function GetHomeTel(s: string) : string;
begin
    GetHomeTel := '';
    if WordCount(s, [' ']) > 1 then
    begin
        GetHomeTel := ExtractWord(1, s, [' ']);
    end
end;

procedure TMainForm.ImportAlvena(DBFile : string);
var
    TableMain, TablePay2{, TableOwn} : TTable;
    Mandator : TTable;
    i, n : integer;
    Seria, Number : string;
    NSeria, NNumber : string;
    dv : integer;
    Field : TField;
    IsNewRecord : boolean;
    PolisId, RateCurrency : string;
    TableTarif : integer;
    NAll, NPos : integer;
    errStr : string;
    SuccessCount : integer;
    IsResident : boolean;
    IsRussian, IsEnglish, IsDigit, IsBad : boolean;
    fld : integer;
    Is2PaysFlag : boolean;
    ErrCount : integer;
    s : string;
    digitDotSrc, digitDotDest : string;
begin
    MessageDlg('Текущая версия не распознает ТИП СТРАХОВАТЕЛЯ - ИП без образования Юр. лица', mtInformation, [mbOk], 0);
    digitDotDest := FloatToStr(5.5)[2];
    digitDotSrc := '.';
    if digitDotSrc = digitDotDest then digitDotSrc := ',';
    errStr := '';
    ErrCount := 0;
    TableMain := TTable.Create(self);
    TablePay2 := TTable.Create(self);
    //TableOwn := TTable.Create(self);
    Mandator := TTable.Create(self);
    Mandator.TableName := 'mandator.db';
    Mandator.DatabaseName := 'DB';

    for i := Length(DBFile) downto 1 do
        if DBFile[i] = '\' then break;

    if PcntForm.ShowModal <> mrOk then exit;

    Is2PaysFlag := (GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0;
    if Is2PaysFlag then
        if MessageDlg('Импортировать только вторые платежи?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
            exit;

    try
        TableMain.TableName := DBFile{Copy(DBFile, 1, i) + 'GORB.DB'};
        n := Pos('.DB', AnsiUpperCase(DBFile)) - 1;
        TablePay2.TableName := Copy(DBFile, 1, n) + '$.DB';
        //TablePay2.TableName := Copy(DBFile, 1, n) + 'TS_$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '_$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '__$.DB';
        if not FileExists(TablePay2.TableName) then
            TablePay2.TableName := Copy(DBFile, 1, n) + '___$.DB';


        //TableOwn.TableName := Copy(DBFile, 1, i) + 'GORB___@.DB';
        TableMain.Open;
        //TablePay2.Exclusive := true;
        TablePay2.Open;
        //TablePay2.AddIndex('NmbIdx', 'Номер полиса', [ixPrimary]);
//        TableOwn.Open;
        Mandator.Open;

        Screen.Cursor := crHourGlass;
        NAll := TableMain.RecordCount;
        NPos := 1;
        SuccessCount := 0;
        while not TableMain.Eof do begin
            StatusBar.Panels[0].Text := InttoStr(NPos) + ' из ' + IntToStr(NAll) + ' Ошибок ' + IntToStr(ErrCount) + ', ' + IntToStr(round(ErrCount * 100 / NAll)) + '%';
            StatusBar.Update;
            NPos := NPos + 1;


           if Is2PaysFlag then
                if TableMain.FieldByName('Форма с_платежа').AsString <> 'С' then begin
                    TableMain.Next;
                    continue;
                end;

            try
                if not TableMain.FieldByName('Состояние').AsInteger in [1, 2, 3] then
                    raise Exception.Create('Полис имеет непонятное состояние.');
                if (TableMain.FieldByName('Состояние').AsInteger = 2) AND
                   (TableMain.FieldByName('Расторжение действия').IsNull) then
                    raise Exception.Create('Полис расторгнут но не указана дата расторжения');
                //if (not TableMain.FieldByName('Расторжение действия').IsNull) AND
                //   (TableMain.FieldByName('Состояние').AsInteger <> 2) then
                //    raise Exception.Create('Полис имеет дату расторжения но по состоянию ДЕЙСТВУЕТ');
                if TableMain.FieldByName('Нетто тариф').AsFloat <> 0. then
                    raise Exception.Create('Полис имеет нетто тариф.');
                //if TableMain.FieldByName('Адрес_Страна').AsString <> 'Беларусь' then
                //    raise Exception.Create('Полис зарегистрирован не на резидента.');
                //if TableMain.FieldByName('Скидка тарифа').AsFloat <> 0 then
                //    raise Exception.Create('Импорт полиса  со скидкой пока не поддержимается.');
                Field := TableMain.FieldByName('Номер полиса');
                PolisId := TableMain.FieldByName('Номер полиса').AsString;
                dv := Pos('/', Field.AsString);
                if dv = 0 then
                    raise Exception.Create('Не правильно указана номер/серия ' + Field.AsString);
                Seria := Copy(Field.AsString, 1, dv - 1);
                Number := Copy(Field.AsString, dv + 1, 1024);
                 if StrToInt(Number) = 180800 then
                    messagebeep(0);

//                IgnoreImportState := false;
                IsNewRecord := false;
                if not Mandator.Locate('Seria;Number', vararrayof([Seria, StrToInt(Number)]), []) then begin
                    Mandator.Append;
                    //Mandator.ClearFields;
                    IsNewRecord := true;
                end
                else begin
                    //Сравнить агента!!!
                    if Mandator.FieldByName('AgentCode').AsString <> 'АЛВН' then
                        raise Exception.Create('Полис ' + Field.AsString + ' уже зарегистрирован на другую компанию');
                    //if Mandator.FieldByName('State').AsInteger = 2 then begin
                    //    if TableMain.FieldByName('Состояние').AsInteger <> 2 then
                    //        IgnoreImportState := true;
                        //    raise Exception.Create('Полис в БД ' + Field.AsString + ' досрочно завершён или расторгнут.');
                    //end;
                    if (TableMain.FieldByName('Состояние').AsInteger = 3) AND (Mandator.FieldByName('State').AsInteger <> 1) OR
                        (TableMain.FieldByName('Состояние').AsInteger <> 3) AND (Mandator.FieldByName('State').AsInteger = 1) then
                        raise Exception.Create('Полис ' + Field.AsString + ' досрочно конфликт состояний.');

                    //не будем переимпортировать данные т.к. один платёж и данные уже есть в Базе
                    if TableMain.FieldByName('Форма с_платежа').AsString <> 'С' then begin
                        Mandator.Edit;
                        Mandator.FieldByName('TarifGlass').AsFloat := TableMain.FieldByName('Полный страховой тариф').AsFloat;
                        Mandator.FieldByName('Tarif').AsFloat := TableMain.FieldByName('Страховой тариф').AsFloat;
                        Mandator.Post;
                        TableMain.Next;
                        Inc(SuccessCount);
                        continue;
                    end;

                    if TableMain.FieldByName('Форма с_платежа').AsString = 'С' then
                        if Mandator.FieldByName('IsPay2').AsInteger = 1 then begin
                            Mandator.Edit;
                            Mandator.FieldByName('TarifGlass').AsFloat := TableMain.FieldByName('Полный страховой тариф').AsFloat;
                            Mandator.FieldByName('Tarif').AsFloat := TableMain.FieldByName('Страховой тариф').AsFloat;
                            Mandator.Post;
                            Inc(SuccessCount);
                            TableMain.Next;
                            continue;
                        end;

                    Mandator.Edit;
                end;

                if {not IgnoreImportState or }IsNewRecord then
                    Mandator.ClearFields;

                Mandator.FieldByName('Seria').AsString := Seria;
                Mandator.FieldByName('Number').AsInteger := StrToInt(Number);

                Mandator.FieldByName('AgentCode').AsString := 'АЛВН';
                Mandator.FieldByName('Fee2Text').AsString := 'АЛВН';

                if TableMain.FieldByName('Состояние').AsInteger = 3 then begin
                    Mandator.FieldByName('Name').AsString := 'ИСПОРЧЕН';
                    Mandator.FieldByName('UpdateDate').AsDateTime := Date;
                    Mandator.FieldByName('State').AsInteger := 1;
                    Mandator.FieldByName('Super').AsInteger := -1;
                    Mandator.FieldByName('RegDate').AsDateTime := TableMain.FieldByName('Дата выдачи').AsDateTime;
                    Mandator.FieldByName('Version').AsInteger := 2;
                    Mandator.FieldByName('RET1').Value := TableMain.FieldByName('Дата выдачи').AsDateTime;
                    TableMain.Next;
                    continue;
                end;

                //Поля прописываются на основании 'Дубликат полиса'
                try
                    if Length(TableMain.FieldByName('ПредПолис_Номер').AsString) > 5 then
                    begin
                        s := TableMain.FieldByName('ПредПолис_Номер').AsString;
                        Mandator.FieldByName('PrevSeria').AsString := Copy(s, 1, Pos('/', s) - 1);
                        Mandator.FieldByName('PrevNumber').AsInteger := -StrToInt(Trim(Copy(s, Pos('/', s) + 1, 100)));
                    end;
                except
                end;

                if Length(TableMain.FieldByName('К3 Коэф').AsString) > 0 then
                begin
                    try
                        Mandator.FieldByName('Placement').AsString := Chr(Trunc(TableMain.FieldByName('К3 Коэф').AsFloat * 10) + 33);
                    except
                        raise Exception.Create('K3 некорректный ' + TableMain.FieldByName('К3 Коэф').AsString);
                    end
                end
                else
                if Mandator.FieldByName('RegDate').AsDateTime >= EncodeDate(2006, 9, 1) then
                begin
                    raise Exception.Create('K3 не указан. Полис выдан ' + Mandator.FieldByName('RegDate').AsString);
                end;

                //Признак юр лица агента
                Mandator.FieldByName('InsComp').AsString := 'U';
                Mandator.FieldByName('Country').AsString := 'U';
                Mandator.FieldByName('Name').AsString := TableMain.FieldByName('Страхователь').AsString;

                //DataMod.ExpStr(Mandator.FieldByName('Name').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                //if IsBad then
                //    raise Exception.Create('Полис ' + Field.AsString + ' недопустимый символ в имени ' + Mandator.FieldByName('Name').AsString);

                Mandator.FieldByName('TarifGlass').AsFloat := TableMain.FieldByName('Полный страховой тариф').AsFloat;
                Mandator.FieldByName('Tarif').AsFloat := TableMain.FieldByName('Страховой тариф').AsFloat;

                if IsNewRecord then begin
                    Mandator.FieldByName('Address').AsString := PrepareStr(
                                                                TableMain.FieldByName('Адрес_Город').AsString,
                                                                TableMain.FieldByName('Адрес_Улица').AsString,
                                                                TableMain.FieldByName('Адрес_Дом').AsString
                                                                //TableMain.FieldByName('Адрес_Телефон').AsString
                                                                );
                    //if Length(Mandator.FieldByName('Address').AsString) < 10 then
                    //    raise Exception.Create('Полис ' + Field.AsString + ' адрес слишком мал ' + Mandator.FieldByName('Address').AsString);
                    Mandator.FieldByName('Fromdate').AsDateTime := TableMain.FieldByName('Начало действия').AsDateTime;
                    Mandator.FieldByName('Fromtime').AsString := CalcAlvenaTime(TableMain.FieldByName('Время начала действия').AsInteger);
                    Mandator.FieldByName('Todate').AsDateTime := TableMain.FieldByName('Окончание действия').AsDateTime;
                    Mandator.FieldByName('Period').AsString := CalcPeriod(Mandator.FieldByName('Fromdate').AsDateTime,
                                                                          Mandator.FieldByName('Todate').AsDateTime);
                    Mandator.FieldByName('Letter').AsString := TableMain.FieldByName('Категория машины').AsString;
                    IsResident := (TableMain.FieldByName('Характеристика клиента').AsInteger AND 2) = 0;
                    Mandator.FieldByName('Super').AsInteger := CalcSuper(Mandator.FieldByName('Letter').AsString, TableTarif, IsResident, false);
                    Mandator.FieldByName('Marka').AsString := TableMain.FieldByName('Марка машины').AsString;
                    Mandator.FieldByName('Autonumber').AsString := TableMain.FieldByName( 'Номер машины').AsString;

                    Mandator.FieldByName('MobTel').AsString := GetMobileTel(TableMain.FieldByName('Адрес_Телефон').AsString);
                    Mandator.FieldByName('HomeTel').AsString := GetHomeTel(TableMain.FieldByName('Адрес_Телефон').AsString);
                    //Mandator.FieldByName('Contact').AsString := TableMain.FieldByName('Адрес_Телефон').AsString;
                    //Mandator.FieldByName('SMS').AsString := TableMain.FieldByName('Адрес_Телефон').AsString;
                    //DataMod.ExpStr(Mandator.FieldByName('Marka').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    //if IsBad then
                    //    raise Exception.Create('Полис ' + Field.AsString + ' недопустимый символ в Марке ' + Mandator.FieldByName('Marka').AsString);

                    try
                        TableMain.FieldByName('Тех. характеристика').AsFloat;
                        case ROUND(GetCharacteristicAuto(Mandator.FieldByName('Letter').AsString, 1, 2, 3, 4, TRUE)) of
                          1 : Mandator.FieldByName('Volume').AsFloat := TableMain.FieldByName('Тех. характеристика').AsFloat;
                          2 : Mandator.FieldByName('Tonnage').AsFloat := TableMain.FieldByName('Тех. характеристика').AsFloat;
                          3 : Mandator.FieldByName('Quantity').AsFloat := TableMain.FieldByName('Тех. характеристика').AsFloat;
                          4 : Mandator.FieldByName('Power').AsFloat := TableMain.FieldByName('Тех. характеристика').AsFloat;
                        end;
                    except
                    end;

                    Mandator.FieldByName('Numberbody').AsString := TableMain.FieldByName('Номер кузова').AsString;
                    Mandator.FieldByName('Chassis').AsString := TableMain.FieldByName('Номер шасси').AsString;
                    Mandator.FieldByName('TarifGlass').AsFloat := TableMain.FieldByName('Полный страховой тариф').AsFloat;
                    Mandator.FieldByName('Tarif').AsFloat := TableMain.FieldByName('Страховой тариф').AsFloat;
                    //Mandator.FieldByName('Tarif').AsFloat := TableMain.FieldByName('Cтраховой тариф').AsFloat;
                                                             //Mandator.FieldByName('TarifGlass').AsFloat *
                                                             //TableMain.FieldByName('Бонус').AsFloat *
                                                             //TableMain.FieldByName('Малус').AsFloat *
                                                             //TableMain.FieldByName('Коэф K2').AsFloat *
                                                             //(100 - Mandator.FieldByName('Discount').AsFloat) / 100;

                    Mandator.FieldByName('Discount').AsFloat := TableMain.FieldByName('Скидка тарифа').AsFloat;

                    Mandator.FieldByName('Allfeetext').AsString := 'N'; //1СУ
                    Mandator.FieldByName('Pay1').AsFloat := TableMain.FieldByName('Страховой платеж').AsFloat;
                    Mandator.FieldByName('Pay1Curr').AsString := TableMain.FieldByName('Валюта с_платежа').AsString;
                    Mandator.FieldByName('PayType').AsFloat := 1;
                    if TableMain.FieldByName('Валюта с_платежа').AsString = 'BYB' then
                        Mandator.FieldByName('Pay1Curr').AsString := 'BRB';
                    //не используется
                    //Mandator.FieldByName('Pay1Text').AsString := '';

                    if TableMain.FieldByName('Форма с_платежа').AsString = 'Б' then
                        Mandator.FieldByName('PayType').AsInteger := 0;
                    if TableMain.FieldByName('Форма с_платежа').AsString = '*' then
                        Mandator.FieldByName('PayType').AsInteger := 1;

                    try
                        Mandator.FieldByName('Draftnumber').AsFloat := TableMain.FieldByName('Номер плат_документа').AsFloat;
                    except
                    end;
                    Mandator.FieldByName('Paydate').AsDateTime := TableMain.FieldByName('Дата с_платежа').AsDateTime;
                    //Mandator.FieldByName('PayType').AsInteger := 0;
                    //if TablePay2.FieldByName('Форма платежа').AsString = '*' then
                    //    Mandator.FieldByName('PayType').AsInteger := 1; //Наличка

                    RateCurrency := 'EUR';
                    if Mandator.FieldByName('Pay1Curr').AsString <> 'BRB' then
                        RateCurrency := Mandator.FieldByName('Pay1Curr').AsString;

                    if RateCurrency = '' then
                        raise Exception.Create('Не заполнено поле Валюта с_платежа');

                    Mandator.FieldByName('Rate').AsFloat := -1000000;
                    //if TableMain.FieldByName('Дата с_платежа').IsNull then
                    //    Mandator.FieldByName('Rate').AsFloat := GetRateOfCurrency(RateCurrency, TableMain.FieldByName('Дата выдачи').AsDateTime, 'дата выдачи')
                    //else
                    //    Mandator.FieldByName('Rate').AsFloat := GetRateOfCurrency(RateCurrency, TableMain.FieldByName('Дата с_платежа').AsDateTime, '1й платёж');

                    //Сколько всего надо заплатить - в валюте первого платежа
                    Mandator.FieldByName('Allfee').AsFloat := -1000000;
                    //if Mandator.FieldByName('Pay1Curr').AsString <> 'BRB' then
                    //    Mandator.FieldByName('Allfee').AsFloat := DivCurr(TableMain.FieldByName('Страховой тариф').AsFloat, GetRateOfCurrency('EUR', max(TableMain.FieldByName('Дата выдачи').AsDateTime, TableMain.FieldByName('Дата с_платежа').AsDateTime), '1й платёж') / Mandator.FieldByName('Rate').AsFloat)
                    //else
                    //    Mandator.FieldByName('Allfee').AsFloat := DivCurr(TableMain.FieldByName('Страховой тариф').AsFloat * Mandator.FieldByName('Rate').AsFloat, 1);

                    //Mandator.FieldByName('Allfee').AsFloat := Mandator.FieldByName('Allfee').AsFloat * (100 - Mandator.FieldByName('Discount').AsFloat) / 100;
                end;

                Mandator.FieldByName('IsFee2').AsInteger := 0;
                Mandator.FieldByName('IsPay2').AsInteger := 0;
                if TableMain.FieldByName('Форма с_платежа').AsString = 'С' then begin
                    Mandator.FieldByName('IsFee2').AsInteger := 1;

                    Mandator.FieldByName('Fee2').AsFloat := Mandator.FieldByName('Tarif').AsFloat / 2.;
                    //Mandator.FieldByName('Fee2text').AsString;
                    Mandator.FieldByName('Fee2date').AsdateTime := (Mandator.FieldByName('Fromdate').AsDateTime + Mandator.FieldByName('Todate').AsDateTime) / 2.;

                    //TablePay2.Close;
                    TablePay2.Filter := '[Номер полиса]=''' + PolisId + ''' AND [Вид платежа]=1';
                    TablePay2.Filtered := True;
                    TablePay2.Open;

                    if TablePay2.Eof then
                        if TableMain.FieldByName('Страховой платеж').AsFloat = TableMain.FieldByName('Страховой тариф').AsFloat then
                            raise Exception.Create('Не обнаружены платежи хотя оплата полная');

                    if TablePay2.RecordCount < 3 then
                    while not TablePay2.Eof do begin
                        if IsNewRecord AND (TablePay2.RecordCount = 1) AND ((TablePay2.FieldByName('Номер платежа').AsString = '') or (TablePay2.FieldByName('Номер платежа').AsInteger = 1)) or (TablePay2.FieldByName('Дата платежа').AsDateTime <= Mandator.FieldByName('Fromdate').AsDateTime) then begin
                            Mandator.FieldByName('Pay1').AsFloat := TablePay2.FieldByName('Сумма платежа').AsFloat;
                            Mandator.FieldByName('Pay1Curr').AsString := TablePay2.FieldByName('Валюта платежа').AsString;

                            if Mandator.FieldByName('Pay1Curr').AsString = 'BYB' then
                                Mandator.FieldByName('Pay1Curr').AsString := 'BRB';
                            Mandator.FieldByName('PayDate').AsDateTime := TablePay2.FieldByName('Дата платежа').AsDateTime;
                            Mandator.FieldByName('PayType').AsInteger := 0;
                            if TablePay2.FieldByName('Форма платежа').AsString = '*' then
                                Mandator.FieldByName('PayType').AsInteger := 1; //Наличка
                        end
                        else begin
                            Mandator.FieldByName('IsPay2').AsInteger := 1;
                            Mandator.FieldByName('Pay2Date').AsDateTime := TablePay2.FieldByName('Дата платежа').AsDateTime;
                            if Mandator.FieldByName('Pay2Date').AsDateTime < TableMain.FieldByName('Дата выдачи').AsDateTime then
                                raise Exception.Create('Дата 2й оплаты слишком рано ' + Mandator.FieldByName('Pay2Date').AsString + ' ' + TableMain.FieldByName('Дата выдачи').AsString + '-' + Mandator.FieldByName('Todate').AsString);
                            Mandator.FieldByName('Pay2Summa').AsFloat := TablePay2.FieldByName('Сумма платежа').AsFloat;
                            Mandator.FieldByName('Pay2Curr').AsString := TablePay2.FieldByName('Валюта платежа').AsString;

                            if Mandator.FieldByName('Pay2Curr').AsString = 'BYB' then
                                Mandator.FieldByName('Pay2Curr').AsString := 'BRB';

                            Mandator.FieldByName('Pay2Type').AsInteger := 0;
                            if TablePay2.FieldByName('Форма платежа').AsString = '*' then
                                Mandator.FieldByName('Pay2Type').AsInteger := 1; //Наличка
                            Mandator.FieldByName('Pay2DraftNumber').Clear;
                            RateCurrency := 'EUR';
                            if Mandator.FieldByName('Pay2Curr').AsString <> 'BRB' then
                                RateCurrency := Mandator.FieldByName('Pay2Curr').AsString;

                            if RateCurrency = '' then
                                raise Exception.Create('Не заполнено поле Валюта платежа в платежах');

                            Mandator.FieldByName('Pay2Rate').AsFloat := -1000000;    
                            //Mandator.FieldByName('Pay2Rate').AsFloat := GetRateOfCurrency(RateCurrency, TablePay2.FieldByName('Дата платежа').AsDateTime, '2й платёж');
                        end;
                        TablePay2.Next
                    end;

                    if (Mandator.FieldByName('Pay1Curr').AsString = '') AND (Mandator.FieldByName('Pay2Curr').AsString <> '') then
                        raise Exception.Create('Платежи не обработаны');
                end;

                if Mandator.FieldByName('PayDate').AsDateTime > Mandator.FieldByName('Fromdate').AsDateTime then
                    if (TableMain.FieldByName('Дубликат полиса').AsString = '') {and (TableMain.FieldByName('Форма с_платежа').AsString <> 'С')} then
                        raise Exception.Create('Дата оплаты после начала или не указана дата первой оплаты в полученном файле. Оплата ' + Mandator.FieldByName('PayDate').AsString + ' Начало ' + Mandator.FieldByName('Fromdate').AsString);

                Mandator.FieldByName('RegDate').AsDateTime := TableMain.FieldByName('Дата выдачи').AsDateTime;

                Mandator.FieldByName('Resident').AsInteger := TableTarif;
                Mandator.FieldByName('CityCode').AsString := GetCityCode(TableMain.FieldByName('Адрес_Область').AsString,
                                                                         TableMain.FieldByName('Адрес_Район').AsString,
                                                                         TableMain.FieldByName('Адрес_Город').AsString);
                //не используется
                // Mandator.FieldByName('Placement').AsFloat
                Mandator.FieldByName('K').AsFloat := GetPlacementK(Mandator.FieldByName('CityCode').AsString);
                Mandator.FieldByName('ClassOldAv').AsString := TableMain.FieldByName('Разряд бонуса пред').AsString;
                Mandator.FieldByName('ClassAvCurr').AsString := TableMain.FieldByName('Разряд бонуса тек').AsString;
                Mandator.FieldByName('CountAv').AsInteger := CalcCountAv(Mandator.FieldByName('ClassOldAv').AsString, Mandator.FieldByName('ClassAvCurr').AsString);
                if TableMain.FieldByName('К2 Коэф').AsString = '' then
                begin
                    raise Exception.Create('K2 не указан');
                end;
                
                try
                    Mandator.FieldByName('K1').AsFloat := StrToFloat(ReplaceStr(Trim(Mandator.FieldByName('K').AsString), digitDotSrc, digitDotDest));
                except
                    on E : EConvertError do begin
                        raise Exception.Create('Не могу конвертировать K1 строку (' + Mandator.FieldByName('K').AsString + ') в число. Возможно не правильный символ-разделитель целой и дробной части в Windows');
                    end
                end;

                try
                    Mandator.FieldByName('K2').AsFloat := 1;    
                    Mandator.FieldByName('K2').AsFloat := StrToFloat(ReplaceStr(Trim(TableMain.FieldByName('К2 Коэф').AsString), digitDotSrc, digitDotDest));
                except
                    on E : EConvertError do begin
                        if (Mandator.FieldByName('Todate').AsDateTime - Mandator.FieldByName('Fromdate').AsDateTime) > 15 then
                            raise Exception.Create('Не могу конвертировать K2 строку (' + TableMain.FieldByName('Коэф K2').AsString + ') в число. Возможно не правильный символ-разделитель целой и дробной части в Windows');
                    end
                end;

                //Состояние обновляем тока в новых полисах
                if Mandator.State = dsInsert then begin
                    if TableMain.FieldByName('Состояние').AsInteger = 2 then
                        Mandator.FieldByName('StopDate').Value := TableMain.FieldByName('Расторжение действия').Value;
                    Mandator.FieldByName('State').AsInteger := 0;
                    if not TableMain.FieldByName('Расторжение действия').IsNull then
                        Mandator.FieldByName('State').AsInteger := 2;
                end;

                Mandator.FieldByName('Vladenie').AsInteger := 6;
                case TableMain.FieldByName('Основание владения').Asinteger of
                 1 : Mandator.FieldByName('Vladenie').AsInteger := 1;
                 2 : Mandator.FieldByName('Vladenie').AsInteger := 5;
                 3 : Mandator.FieldByName('Vladenie').AsInteger := 4;
                 4 : Mandator.FieldByName('Vladenie').AsInteger := 4;
                 5 : Mandator.FieldByName('Vladenie').AsInteger := 1;
                end;

                //Физич = 1 Юрид = 0 2 ИП
                Mandator.FieldByName('InsurerType').AsInteger := 1;
                if (TableMain.FieldByName('Характеристика клиента').AsInteger AND 1) = 1 then
                    Mandator.FieldByName('InsurerType').AsInteger := 0;
                Mandator.FieldByName('Version').AsInteger := 2;
                if TableMain.FieldByName('Дубликат полиса').AsString <> '' then
                    Mandator.FieldByName('State').AsInteger := 3;
                Mandator.FieldByName('RetBRB').Clear;
                Mandator.FieldByName('RetEUR').Clear;
                Mandator.FieldByName('UpdateDate').AsDateTime := Date;
//                Mandator.FieldByName('Country').Clear;
                Mandator.FieldByName('OwnerCode').AsInteger := 0;
                Mandator.FieldByName('BaseCarCode').AsInteger := 0;
//                Mandator.FieldByName('CarCode').AsInteger := 0;


                Mandator.FieldByName('AgPercent').Value := PcntForm.PcntBefore.Value;
                Mandator.FieldByName('carcode').Value := PcntForm.PcntBefore.Value;
                if Mandator.FieldByName('Paydate').Value >= PcntForm.Date.Date then Mandator.FieldByName('AgPercent').Value := PcntForm.PcntAfter.Value;

                Mandator.FieldByName('RET1').Value := Mandator.FieldByName('PayDate').Value;
                Mandator.FieldByName('RET2').Value := Mandator.FieldByName('Pay2Date').Value;

                for fld := 0 to Mandator.FieldCount - 1 do
                    if Mandator.Fields[i].DataType = ftDate then
                        if not Mandator.Fields[i].IsNull then
                            if (Mandator.Fields[i].AsDateTime < (Date - 400)) OR (Mandator.Fields[i].AsDateTime > (Date + 400)) then
                                raise Exception.Create('Полис ' + Field.AsString + ' дата не в диапазоне ' + Mandator.Fields[i].FieldName + '=' + Mandator.Fields[i].AsString);

                Mandator.Post;
                SuccessCount := SuccessCount + 1;

                if TableMain.FieldByName('Дубликат полиса').AsString <> '' then begin
                    Field := TableMain.FieldByName('Дубликат полиса');
                    dv := Pos('/', Field.AsString);
                    if dv = 0 then
                        raise Exception.Create('Не правильно указана номер/серия дубликата' + Field.AsString);
                    NSeria := Copy(Field.AsString, 1, dv);
                    NNumber := Copy(Field.AsString, dv + 1, 1024);
                    if not Mandator.Locate('Seria;Number', vararrayof([NSeria, StrToInt(NNumber)]), []) then
                    begin
                        errStr := errStr + PolisId + ' Полис импортирован хотя дубликат не найден ' + Field.AsString + #13#10
                    end;
                    
                    Mandator.Edit;
                    Mandator.FieldByName('PrevSeria').AsString := NSeria;
                    Mandator.FieldByName('PrevNumber').AsInteger := StrToInt(NNumber);
                    Mandator.Post;
                end;
            except
                on  E : Exception do begin
                    if IsNewRecord then
                        Mandator.Delete
                    else
                    if Mandator.State in [dsEdit] then
                        Mandator.Cancel;

                    errStr := errStr + PolisId + ' ' + e.Message + #13#10;
                    Inc(ErrCount);
                end;
            end;

            TableMain.Next;
        end;
    except
        on  E : Exception do begin
            MessageDlg('Ошибка импорта. ' + e.Message, mtInformation, [mbOk], 0);
        end
    end;

    Screen.Cursor := crDefault;
    Mandator.Free;
    TableMain.Free;
    TablePay2.Free;
//    TableOwn.Free;

    MessageDlg(Format('Импортировано %d из %d. %d пропущено', [SuccessCount, NAll, NAll - SuccessCount]), mtInformation, [mbOk], 0);

    if errStr <> '' then begin
        ShowInfo.Caption := 'Результат импорта данных Альвены';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
    FreeWindow;
end;

procedure TMainForm.FreeWindow;
begin
     MainQuery.Close;
     if W <> 0 then begin
        ShowWindow(W, SW_SHOW);
        EnableWindow(W, true);
        SetActiveWindow(W);
        W := 0;
     end;
end;

procedure TMainForm.OpenDialogPxCanClose(Sender: TObject;
  var CanClose: Boolean);
var
    Table : TTable;
begin
    CanClose := true;
    {if (GetAsyncKeyState(VK_SHIFT) AND $8000) > 0 then exit;
    if Pos('$', OpenDialogPx.FileName) <> 0 then CanClose := false;
    if Pos('_', OpenDialogPx.FileName) <> 0 then CanClose := false;
    Table := TTable.Create(self);
    Table.TableName := OpenDialogPx.FileName;
    try
        Table.Open;
        if Table.FieldCount < 60 then
            CanClose := false;
    except
        CanClose := false;
    end;
    Table.Close;}
end;

procedure TMainForm.ImpParadoxClick(Sender: TObject);
begin
     if not OpenDialogPx2.Execute then exit;
     ImportDB(OpenDialogPx2.FileName)
end;

procedure TMainForm.ImportDB(DBFile : string);
var
    TableMain : TTable;
    Mandator : TTable;
    i, n : integer;
    Seria : string;
    Number : integer;
    errStr : string;
    NAll, NPos : integer;
    SuccessCount : integer;
    IsNewRecord : boolean;
    IsRussian, IsEnglish, IsDigit, IsBad : boolean;
    {NSeria, NNumber : string;
    dv : integer;
    Field : TField;
    PolisId, RateCurrency : string;
    TableTarif : integer;
    IsResident : boolean;}
    fld : integer;
begin
    errStr := '';
    TableMain := TTable.Create(self);
    Mandator := TTable.Create(self);
    Mandator.TableName := 'mandator.db';
    Mandator.DatabaseName := 'DB';

    //for i := Length(DBFile) downto 1 do
    //    if DBFile[i] = '\' then break;

    if formImportPx.ShowModal <> mrOk then exit;

    try
        TableMain.TableName := DBFile;
        TableMain.Open;
        Mandator.Open;

        Screen.Cursor := crHourGlass;
        NAll := TableMain.RecordCount;
        NPos := 1;
        SuccessCount := 0;
        while not TableMain.Eof do begin
            StatusBar.Panels[0].Text := InttoStr(NPos) + ' из ' + IntToStr(NAll);
            StatusBar.Update;
            NPos := NPos + 1;

            try
                Seria := TableMain.FieldByName('Seria').AsString;
                Number := TableMain.FieldByName('Number').AsInteger;

                IsNewRecord := false;
                if not Mandator.Locate('Seria;Number', vararrayof([Seria, Number]), []) then begin
                    Mandator.Append;
                    Mandator.ClearFields;
                    IsNewRecord := true;
                    Mandator.FieldByName('Seria').AsString := Seria;
                    Mandator.FieldByName('Number').AsInteger := Number;
                end
                else begin
                    //Сравнить агента!!!
                    if Mandator.FieldByName('AgentCode').AsString <> formImportPx.Agent.Value then
                        raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' уже зарегистрирован на агента ' + Mandator.FieldByName('AgentCode').AsString);
                    if Mandator.FieldByName('AgPercent').AsFloat <> formImportPx.AgPercent.Value then
                        raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' Процент агенту другой');
                    Mandator.Edit;
                end;

                Mandator.ClearFields;
                for i := 0 to TableMain.FieldCount - 1 do begin
                    Mandator.FieldByName(TableMain.Fields[i].FieldName).Value := TableMain.Fields[i].Value
                end;

                Mandator.FieldByName('AgentCode').AsString := formImportPx.Agent.Value;
                Mandator.FieldByName('fee2text').AsString := formImportPx.Agent.Value;

                if Mandator.FieldByName('State').AsInteger <> 1 then begin //не испорчен
                    Mandator.FieldByName('AgPercent').AsFloat := formImportPx.AgPercent.Value;
                    Mandator.FieldByName('carcode').AsFloat := formImportPx.AgPercent.Value;

                    //DataMod.ExpStr(Mandator.FieldByName('Name').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    //if IsBad then
                    //    raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' недопустимый символ в имени ' + Mandator.FieldByName('Name').AsString);

                    //if Length(Mandator.FieldByName('Address').AsString) < 10 then
                    //    raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' адрес слишком мал ' + Mandator.FieldByName('Address').AsString);
                    //DataMod.ExpStr(Mandator.FieldByName('Marka').AsString, IsRussian, IsEnglish, IsDigit, IsBad, true);
                    //if IsBad then
                    //    raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' недопустимый символ в Марке ' + Mandator.FieldByName('Marka').AsString);

                    if Mandator.FieldByName('PayDate').AsDateTime > Mandator.FieldByName('Fromdate').AsDateTime then
                        raise Exception.Create('Дата оплаты после начала');
                end;

                for fld := 0 to Mandator.FieldCount - 1 do
                    if Mandator.Fields[i].DataType = ftDate then
                        if not Mandator.Fields[i].IsNull then
                            if (Mandator.Fields[i].AsDateTime < (Date - 400)) OR (Mandator.Fields[i].AsDateTime > (Date + 400)) then
                                raise Exception.Create('Полис ' + Seria + '/' + IntToStr(Number) + ' дата не в диапазоне ' + Mandator.Fields[i].FieldName + '=' + Mandator.Fields[i].AsString);

                Mandator.Post;
                SuccessCount := SuccessCount + 1;
            except
                on  E : Exception do begin
                    if IsNewRecord then
                        Mandator.Delete
                    else
                    if Mandator.State in [dsEdit] then
                        Mandator.Cancel;

                    errStr := errStr + Seria + '/' + IntToStr(Number) + ' ' + e.Message + #13#10;
                end;
            end;

            TableMain.Next;
        end;
    except
        on  E : Exception do begin
            MessageDlg('Ошибка импорта. ' + e.Message, mtInformation, [mbOk], 0);
        end
    end;

    Screen.Cursor := crDefault;
    Mandator.Free;
    TableMain.Free;

    MessageDlg(Format('Импортировано %d из %d', [SuccessCount, NAll]), mtInformation, [mbOk], 0);

    if errStr <> '' then begin
        ShowInfo.Caption := 'Результат импорта данных';
        ShowInfo.InfoPanel.Text := errStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end;
end;


procedure TMainForm.InputSummClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText : string;
begin
    if RepParams.ShowModal <> mrOk then exit;

    MessageDlg('Расчёт будет выполнен на основе даты оплаты', mtInformation, [mbOk], 0);

    DtFrom := min(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    DtTo := max(RepParams.DtFrom.Date, RepParams.DtTo.Date);
    sqlText := 'SELECT PAY1,PAY1CURR,PAYDATE,NUMBER,RET1 FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND (PREVSERIA='''' OR PREVSERIA IS NULL) AND PAYDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' UNION ALL ' +
               'SELECT PAY2SUMMA,PAY2CURR,PAY2DATE,NUMBER,RET1 FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND STATE IN (0,2,4) AND ISPAY2=1 AND PAY2DATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';
    UpdateDupdata('');
    BuildReportInputSum(DtFrom, DtTo, sqlText, '', WorkSQL);
end;

procedure TMainForm.N20Click(Sender: TObject);
begin
    CurrencyRates.ShowModal
end;

function DecodeRN(s : string) : string;
begin
    if s = '1' then DecodeRN := '1'
    else DecodeRN := '2';
end;

function NotExistsPay2(s, n : string; dt : TDateTime) : boolean;
begin
    NotExistsPay2 := true;
    MainForm.HandBookSQL.Close;
    MainForm.HandBookSQL.SQL.Clear;
    MainForm.HandBookSQL.SQL.Add('SELECT ISPAY2 FROM ' + GetTableName() + ' WHERE SERIA=''' + s + ''' AND NUMBER=' + n + ' AND PAY2DATE<''' + DateToStr(dt) + '''');
    MainForm.HandBookSQL.Open;
    if not MainForm.HandBookSQL.EOF then
        NotExistsPay2 := MainForm.HandBookSQL.FieldByName('ISPAY2').AsInteger = 0;
end;

function DecodeAgentNamePx(s : string) : string;
var
    code : array[0..16] of char;
    val : double;
begin
    if lastAgCode = s then begin
        DecodeAgentNamePx := lastAgName;
        exit;
    end;
    lastAgCode := s;
    DMM.HandBookSQL.Close;
    DMM.HandBookSQL.SQL.Clear;
    DMM.HandBookSQL.SQL.Add('SELECT NAME,SOVAG,POLANDPCNT FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    DMM.HandBookSQL.Open;
    if DMM.HandBookSQL.Eof then lastAgName := s
    else begin
        lastAgName := DMM.HandBookSQL.FieldByName('NAME').AsString;
        val := DMM.HandBookSQL.FieldByName('POLANDPCNT').AsFloat;
        Move(val, code[0], 8);
        val := DMM.HandBookSQL.FieldByName('SOVAG').AsFloat;
        Move(val, code[8], 8);
        code[16] := #0;
        lastAgID := code;
    end;
    DecodeAgentNamePx := lastAgName;
end;

function DecodeAgentID(s : string) : string;
var
    code : array[0..16] of char;
    val : double;
begin
    if lastAgCode = s then begin
        DecodeAgentID := lastAgID;
        exit;
    end;
    lastAgCode := s;
    DMM.HandBookSQL.Close;
    DMM.HandBookSQL.SQL.Clear;
    DMM.HandBookSQL.SQL.Add('SELECT NAME, SOVAG,POLANDPCNT FROM AGENT WHERE AGENT_CODE = ''' + s + '''');
    DMM.HandBookSQL.Open;
    if DMM.HandBookSQL.Eof then lastAgName := s
    else begin
        lastAgName := DMM.HandBookSQL.FieldByName('NAME').AsString;
        val := DMM.HandBookSQL.FieldByName('POLANDPCNT').AsFloat;
        Move(val, code[0], 8);
        val := DMM.HandBookSQL.FieldByName('SOVAG').AsFloat;
        Move(val, code[8], 8);
        code[16] := #0;
        lastAgID := code;
    end;
    DecodeAgentID := lastAgID;
end;

function IsOnlySecondPay(RegDt, DtSecPay : TField) : boolean;
begin
    IsOnlySecondPay := false;

    //no filter by period - ret false
    if not MandatoryFilterDlg.IsPayDate.Checked then exit;

    //Выдан в периоде - ret false
    if (RegDt.AsDateTime >= MandatoryFilterDlg.Pay1Start_.Date) AND
       (RegDt.AsDateTime <= MandatoryFilterDlg.Pay1End_.Date) then exit;

    //2й платеж не попал в период - ret false
    if not DtSecPay.IsNull then
    if (DtSecPay.AsDateTime < MandatoryFilterDlg.Pay1Start_.Date) OR
       (DtSecPay.AsDateTime > MandatoryFilterDlg.Pay1End_.Date) then exit;

    IsOnlySecondPay := not DtSecPay.IsNull;
end;

function DecodeZatr(s : string) : string;
begin
    if Length(s) < 2 then s := '0' + s;
    DecodeZatr := s;
end;

function FindInList(List : TStringList; s : string; add, del : boolean) : boolean;
var
    Index : integer;
begin
    if List.Find(s, Index) then begin
        if del then List.Delete(Index);
        FindInList := true;
        exit;
    end;
    if add then begin
        List.Add(s);
        List.Sort;
    end;
    FindInList := false;
end;

procedure WriteLn2(var F8, A : TextFile; s, s2 : string);
begin
    WriteLn(F8, s);
    WriteLn(A, s + '|' + s2);
end;

function DecodeSeria1SU(s : string) : string;
begin
    DecodeSeria1SU := '?';
    if s = 'Y' then DecodeSeria1SU := '12'
    else if s = '0' then DecodeSeria1SU := '13'
    else DecodeSeria1SU := _1SU_GLOBO_CODES[ord(s[1]) - ord('A')];
end;

function GetSeria1SU(s : string) : string;
begin
    GetSeria1SU := '?';
    if s = 'Y' then GetSeria1SU := 'ГС'
    else if s = '0' then GetSeria1SU := 'КВ'
    else GetSeria1SU := _1SU_SERIES[ord(s[1]) - ord('A')];
end;

function RemoveTel(addr, agent : string) : string;
var
    tel_pos : integer;
begin
    RemoveTel := addr;
    if agent = 'АЛВН' then
    begin
        tel_pos := Pos('ТЕЛ ', addr);
        if tel_pos > 0 then RemoveTel := Copy(addr, 1, tel_pos - 1);
    end
end;



function _33Code(s : string) : string;
const
  symb : string = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var
        val : integer;
        res : string;
begin
        val := StrToInt(s);
        while val > 33 do
        begin
                res := res + symb[(val mod 33) + 1];
                val := Trunc( val / 33);
        end;
        res := res + symb[val + 1];
       _33Code := ReverseString(res);
end;

function ToENG(s : string) : string;
const
  RU : string = 'ЦУКЕНГШЗХФЫВАПРОЛДЖЭЯЧСМИТБЮ';
  EN : string = 'CUKENGSZXFIWAPROLDZEJQCMITBU';
var
  i, idx : integer;
begin
  for i := 1 to Length(s) do
    for idx := 1 to Length(RU) do
            if s[i] = RU[idx]   then
            begin
                    s[i] := EN[idx];
            break;
            end;

ToENG := s;
end;

procedure TMainForm.TOGloboClick(Sender: TObject);
var
    searchResult : TSearchRec;
    D : TextFile;
    X : TextFile;
    F3 : TextFile;
    F6 : TextFile;
    P0, P, A, A_ : TextFile; //P0 old file
    F10 : TextFile;
    F11 : TextFile;
    E : TextFile;
    //F12 : TextFile;
    F13 : TextFile;
    AVAR : TextFile;
    PrevKod, PrevSeria, PrevNumber : string;

    N, Q, V, U, Log : TextFile;

    All, Curr : integer;
    IsError : boolean;
    pbuff : array[0..128] of char;
    UnloadN, WarnStr : string;
    UnloadDt : string;
    Division, sDivision : string;
    DD, MM, YY : WORD;
    errStr, Prefix : string;
    INI : TIniFile;
    path : string;
    Ltr : char;
    OnlyAvarFlag, bUnloadAllPolises : boolean;
    ImgCount, PolisCount, PayCount, BadCount, AvCount : integer;
    TaxiFound : boolean;
    Letter : string;
    RequestFind : boolean;
    MinSumma : integer;
    Hour, Min, Sec, MSec: Word;
    n_fld : integer;
    zatr_list : array[1..50] of integer;
    s : string;
    i : integer;
    prevWork : string;
    imgDir : string;
    ErrorCount : integer;

    finder : integer;
    Numbers : TStringList;
    UnloadedList, Warn : TStringList;
    StrNumbers : string;
    F : TextFile;
    sline : string;
    MinNmb, MaxNmb : string;
    div_nmb_pos : integer;
    fileSource : string;
    _SERIA_, _NUMBER_, _DIVISION_ : string;
    _1SULen : integer;
    Dt : TDateTime;
    korrectAutoNmb : string;
    MsgCode : string;
    serEq : string;
    unloadCycles : integer;
    firstFlag : boolean;
    delta : integer;
    InsTypeCode : string;
label
    ExpNext, RestartExport, skipK3;
begin

    firstFlag := true;
    unloadCycles := 5;

    if bKomplexMode then
            InsTypeCode := '40'
    else
            InsTypeCode := '30';

    if TMenuItem(Sender).Tag = 0 then
    begin
        if (GetFilter('') = '') or (not MandatoryFilterDlg.IsPayDate.Checked) then
        begin
            if MessageDlg('Не указан период выгрузки данных. Программа не сможет определить полисы выгружаемые по вторым платежам и будет экспортировать все данные. Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
        end
        else
        begin
            if mnuUnloadAllPolises.Checked then
            begin
                if MessageDlg('Установлен режим выгрузки всех данных по полисам. Программа выгрузит данные по некоторым полисам повторно, хотя они уже передавались в Дирекцию. Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
            end
        end;
    end;

    AddIndex('PrevSeria;PrevNumber');
    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT PREVSERIA, PREVNUMBER, COUNT(* ) AS CNT, MIN(SERIA || ''/'' || CAST(NUMBER AS CHAR(10))) AS MNN, MAX(SERIA || ''/'' || CAST(NUMBER AS CHAR(10))) AS MXN FROM ' + GetTableName() + ' WHERE PREVSERIA IS NOT NULL GROUP BY PREVSERIA, PREVNUMBER HAVING COUNT(*)>1');
    WorkSQL.Open;
    while not WorkSQL.Eof do
    begin
        if s <> '' then s := s + #13;
        s := s + WorkSQL.FieldByName('PREVSERIA').AsString + '/' + WorkSQL.FieldByName('PREVNUMBER').AsString + ' ' + WorkSQL.FieldByName('CNT').AsString + ' шт.' + ' Дубликаты: ' + WorkSQL.FieldByName('MNN').AsString + ' и ' + WorkSQL.FieldByName('MXN').AsString;
        WorkSQL.Next;
    end;
    WorkSQL.Close;

    if s <> '' then
    begin
        Clipboard.AsText := s;
        MessageDlg(s + #13'Указанные полисы имеют несколько дубликатов. Выгрузка невозможна. Исправляте ошибки. Список полисов находится в буфере обмена', mtInformation, [mbOk], 0);
        exit;
    end;

    s := '';
    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT * FROM AGENT WHERE AGENT_CODE<>''АЛВН'' AND TYPE=''F''');
    WorkSQL.Open;
    while not WorkSQL.Eof do
    begin
        if DecodeAgentID(WorkSQL.FieldByName('AGENT_CODE').AsString) = '' then
        begin
            if s <> '' then s := s + #13;
            s := s + WorkSQL.FieldByName('NAME').AsString;
        end;
        WorkSQL.Next;
        if Length(s) > 300 then break;
    end;
    WorkSQL.Close;

    if s <> '' then
    begin
        MessageDlg('ПОСЛЕДНЕЕ Китайское предупреждение!!!'#13'Не указан личный номер агента'#13 + s + #13'Следующая версия программы не позволит выгружать полисы у которых агент не имеет личного номера.', mtInformation, [mbOk], 0);
    end;

    Warn := TStringList.Create;
    UnloadedList := TStringList.Create;
    Numbers := nil;
    UnloadedList.CaseSensitive := true;

    OnlyAvarFlag := TMenuItem(Sender).Tag = 1;

    if FloatToStr(5.5)[2] <> '.' then begin
        MessageDlg('Установлен не правильный разделитель целой и дробной части. Нужна точка.', mtInformation, [mbOk], 0);
        ShellExecute(0, 'open', 'control.exe', 0, 0, 0);
        exit;
    end;

    if GetFilter('') <> '' then
    begin
        if (TMenuItem(Sender).Tag <> 2) AND (not OnlyAvarFlag) then
            if not MandatoryFilterDlg.IsStopped.Checked then
                if MessageDlg('Вы не указали выгрузку досрочно расторгнутых (В Фильтре не установлена пометка Прекращён). Продолжить экспорт?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
    end;

    if TMenuItem(Sender).Tag = 2 then
    begin
        Numbers := TStringList.Create;
        Numbers.CaseSensitive := true;
        if not OpenDialogNMB.Execute then exit;
        AssignFile(F, OpenDialogNMB.FileName);
        Reset(F);
        while not Eof(F) do
        begin
            ReadLn(F, sline);
            Numbers.Add(Trim(sline));
        end;
        Numbers.Sort;
        CloseFile(F);
        unloadCycles := 0;
        firstFlag := false;
        //restart := false;
    end;

    bUnloadAllPolises := mnuUnloadAllPolises.Checked;
    PolisCount := 0;
    ErrorCount := 0;
    PayCount := 0;
    BadCount := 0;
    AvCount := 0;

    INI := TIniFile.Create(BLANK_INI);
    SaveDialog.InitialDir := INI.ReadString('DIVISION', 'OutDir', 'C:\');
    Division := INI.ReadString('DIVISION', 'ID', '');
    MinSumma := INI.ReadInteger('MANDATORY', 'MinSumma', 1000);
    _1SULen := INI.ReadInteger('MANDATORY', '1SULen', 7);

    for i:= 1 to 100 do begin
        s := INI.ReadString('MANDATORY', 'Zatr' + IntToStr(i - 1), '');
        zatr_list[i] := 0;
        if Pos(',', s) <> 0 then begin
            s := Trim(Copy(s, 1, Pos(',', s) - 1));
            try
                zatr_list[i] := StrToInt(s);
            except
            end;
        end;
    end;

    INI.Free;
    if Length(Division) < 1 then begin
        MessageDlg('Не указан код вашего подразделения', mtInformation, [mbOk], 0);
        if Numbers <> nil then Numbers.Free;
        if UnloadedList <> nil then UnloadedList.Free;
        if Warn <> nil then Warn.Free;
        exit;
    end;

    if OnlyAvarFlag then
        mnuNoUnloadAvarias.Checked := false;

    SaveDialog.FileName := UnloadN + '.unload';
    if not SaveDialog.Execute then exit;
    path := Copy(SaveDialog.FileName, 1, Length(SaveDialog.FileName) - Pos('\', ReverseString(SaveDialog.FileName)) + 1);
    INI := TIniFile.Create(BLANK_INI);
    INI.WriteString('DIVISION', 'OutDir', path);
    INI.Free;

    Prefix := '';
    errStr := '';
    UnloadN := '00';
    //Division := '00';
    DecodeDate(Date, YY, MM, DD);
    UnloadDt := IntToStr(MM);
    if MM < 10 then UnloadDt := '0' + UnloadDt;
    if DD < 10 then UnloadDt := UnloadDt + '0';
    UnloadDt := UnloadDt + IntToStr(DD);
    imgDir := 'IMG-' + UnloadDt;
    IsError := false;

    if TMenuItem(Sender).Tag = 1 then begin
        if GetTwoDates.ShowModal <> mrOK then exit;
    end;

    sDivision := Division;
    if Length(sDivision) < 2 then sDivision := '0' + sDivision;

    DecodeTime(Now, Hour, Min, Sec, MSec);
    //UnloadN := IntToStr(Hour);
    //if Hour < 10 then UnloadN := '0' + UnloadN;


    UnloadN := Stringreplace(Format('%02d', [Hour]) + Format('%02d', [Min]) + Format('%02d', [Sec]), ' ', '0', [rfReplaceAll]);
    imgDir := imgDir + '-' + UnloadN;

    AssignFile(Log,  path + sDivision + 'LOG' + UnloadDt + UnloadN + '.Log0' + InsTypeCode);
    ReWrite(Log);
    AssignFile(D,  path + sDivision + 'D' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(D);
    AssignFile(X, path + sDivision + 'X' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(X);
    AssignFile(F3, path + sDivision + 'T' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F3);
    AssignFile(F6, path + sDivision + 'C' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F6);
    AssignFile(P0, path + sDivision + 'P' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(P0);
    AssignFile(A, path + sDivision + 'A' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(A);
    AssignFile(A_, path + sDivision + 'A-' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(A_);
    AssignFile(P, path + sDivision + 'P-' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(P);
    AssignFile(F10, path + sDivision + 'H' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F10);
    AssignFile(F11, path + sDivision + 'Z' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F11);
    //AssignFile(F12, path + sDivision + 'G' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    //ReWrite(F12);
    if bKomplexMode then
    begin
            AssignFile(E, path + sDivision + 'E' + UnloadDt + UnloadN + '.0' + InsTypeCode);
            ReWrite(E);
    end;
    AssignFile(F13, path + sDivision + 'B' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(F13);
    AssignFile(AVAR, path + sDivision + 'M' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(AVAR);
    AssignFile(N, path + sDivision + 'N' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(N);
    AssignFile(Q, path + sDivision + 'Q' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(Q);
    AssignFile(V, path + sDivision + 'V' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(V);
    AssignFile(U, path + sDivision + 'U' + UnloadDt + UnloadN + '.0' + InsTypeCode);
    ReWrite(U);

    WriteLn(Log, 'Версия прогаммы от ' + DateToStr(FileDateToDateTime(FileAge(Application.ExeName))));
    WriteLn(Log, 'Лимит повторов догрузки ' + IntToStr(unloadCycles) + ' раз');

    if Numbers <> nil then
        WriteLn(Log, 'Выгрузка по номерам. Указано ' + IntToStr(Numbers.Count) + ' штук')
    else
    if OnlyAvarFlag then
        WriteLn(Log, 'Выгрузка только аварий')
    else
        WriteLn(Log, 'Выгрузка полисов и аварий');

    if bUnloadAllPolises then
        WriteLn(Log, 'Выгрузка всех данных')
    else
        WriteLn(Log, 'Выгрузка остаточных данных');

    if mnuNoUnloadAvarias.Checked then
        WriteLn(Log, 'Не выгружать аварии')
    else
        WriteLn(Log, 'Выгрузка с авариями');


    RestartExport:;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT ''' + Division + ''' AS DIVISION, M.* FROM ' + GetTableName() + ' M ');
    if (not OnlyAvarFlag) and firstFlag then
        if Length(GetFilter('')) <> 0 then
            WorkSQL.SQL.Add(' WHERE ' + GetFilter(''));

    if not firstFlag then
    begin
        serEq := '';
        for i := 0 to Numbers.Count - 1 do
        begin
            sline := Numbers[i];
            WriteLn(Log, 'Догружаем ' + sline);
            div_nmb_pos := Pos(',', sline);
            if(div_nmb_pos <> 0) then begin
            if serEq <> '?' then
                begin
                    if serEq = '' then
                        serEq := Copy(sline, 1, div_nmb_pos - 1)
                    else
                    if serEq <> Copy(sline, 1, div_nmb_pos - 1) then
                        serEq := '?'
                end;

                if MinNmb = '' then begin
                    MinNmb := Trim(Copy(sline, div_nmb_pos + 1, 100));
                    MaxNmb := Trim(Copy(sline, div_nmb_pos + 1, 100));
                end;
                if StrToInt(Copy(sline, div_nmb_pos + 1, 100)) < StrToInt(MinNmb) then
                    MinNmb := Copy(sline, div_nmb_pos + 1, 100);
                if StrToInt(Copy(sline, div_nmb_pos + 1, 100)) > StrToInt(MaxNmb) then
                    MaxNmb := Copy(sline, div_nmb_pos + 1, 100);
            end;
        end;

        if (MinNmb = '') or (MaxNmb = '') then begin
            MessageDlg('Обнаружены некорректные данные в дубликатах или утерянных полисах', mtInformation, [mbOk], 0);
        end;
        WorkSQL.SQL.Add(' WHERE NUMBER BETWEEN 0' + MinNmb + ' AND 0' + MaxNmb);
        if (serEq <> '') AND (serEq <> '?') then
        begin
            WorkSQL.SQL.Add(' AND SERIA=''' + serEq + '''');
        end;
    end;

    if not bKomplexMode and OnlyAvarFlag and (firstFlag) then begin
        WorkSQL.SQL.Add(',MANDAV2 AV WHERE AV.SERIA=M.SERIA AND AV.NUMBER=M.NUMBER');
        WorkSQL.SQL.Add('AND ((V1_DATE >= ''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND V1_DATE <= ''' + DateToStr(GetTwoDates.EndDate.Date) + ''')');
        WorkSQL.SQL.Add('OR (V2_DATE >= ''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND V2_DATE <= ''' + DateToStr(GetTwoDates.EndDate.Date) + ''')');
        WorkSQL.SQL.Add('OR (V3_DATE >= ''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND V3_DATE <= ''' + DateToStr(GetTwoDates.EndDate.Date) + ''')');
        WorkSQL.SQL.Add('OR (UPDATEDATE >= ''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND UPDATEDATE <= ''' + DateToStr(GetTwoDates.EndDate.Date) + '''))');

        WorkSQL.SQL.Add('UNION');
        WorkSQL.SQL.Add('SELECT ''' + Division + ''' AS DIVISION, M.* FROM MANDATOR M ');
        WorkSQL.SQL.Add(',MANDAV2 AV, AVPAYS AP WHERE AV.SERIA=M.SERIA AND AV.NUMBER=M.NUMBER');
        WorkSQL.SQL.Add('AND AP.SER=AV.SERIA AND AP.NMB=AV.NUMBER AND AP.N=AV.N');
        WorkSQL.SQL.Add('AND (AP.D >=''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND AP.D <=''' + DateToStr(GetTwoDates.EndDate.Date) + '''');
        WorkSQL.SQL.Add('OR UPDATEDATE >= ''' + DateToStr(GetTwoDates.StartDate.Date) + '''');
        WorkSQL.SQL.Add('AND UPDATEDATE <= ''' + DateToStr(GetTwoDates.EndDate.Date) + ''')');
    end;

    WriteLn(Log, '=== QUERY ===');
    WriteLn(Log, WorkSQL.SQL.Text);
    WriteLn(Log, '=== QUERY ===');

    WorkSQL.Open;
    All := WorkSQL.RecordCount;

    WriteLn(Log, IntToStr(All) + ' records selected');

    Curr := 1;

    if Numbers = nil then
    begin
        Numbers := TStringList.Create;
        Numbers.CaseSensitive := true;
    end;

    with WorkSQL do begin
        while not Eof do begin
          try
            Curr := Curr + 1;
            StatusBar.Panels[0].Text := IntToStr(Curr) + ' шт., ' + IntToStr(ErrorCount) + ' ошибок.';
            StatusBar.Update;
            ProgressBar.Position := round(Curr * 100 / All);

            _SERIA_ := FieldByName('SERIA').AsString;
            _NUMBER_ := FieldByName('NUMBER').AsString;
            _DIVISION_ := FieldByName('DIVISION').AsString;

            if (not firstFlag) and (Numbers <> nil) then begin
                //РЕЖИМ ДОГРУЗКИ
                if FindInList(Numbers, _SERIA_ + ',' + _NUMBER_, false, true) then
                    goto ExpNext; //Если есть в логрузке то догружаем

                Next;
                continue;
            end;

            if unloadCycles > 0 then
            begin
                //РЕЖИМ ВЫГРУЗКИ - ДОБАВЛЯЕМ ДОГРУЗКУ
                if (FieldByName('PREVNUMBER').AsFloat > 0) then begin
                    FindInList(Numbers, FieldByName('PREVSERIA').AsString + ',' + FieldByName('PREVNUMBER').AsString, true, false);

                    if FieldByName('PREVNUMBER').AsString = '' then
                        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' некорректный номер утерянного полиса');
                end;
                if FieldByName('STATE').AsString = '3' then
                begin //lost
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('select * from ' + GetTableName() + ' where prevseria=''' + _SERIA_ + ''' and prevnumber=0' + _NUMBER_);
                    HandBookSQL.Open;
                    if not HandBookSQL.Eof then
                    begin
                        FindInList(Numbers, HandBookSQL.FieldByName('SERIA').AsString + ',' + HandBookSQL.FieldByName('NUMBER').AsString, true, false);
                    end
                    else
                    begin
                        WriteLn(Log, 'Не найден дубликат для догрузки утерянного ' + _SERIA_ + '/' + _NUMBER_);
                    end;
                    HandBookSQL.Close;
                end;
            end;

            ExpNext:;

            //Если уже выгружался то уходим и удалим из догрузки
            if FindInList(UnloadedList, _SERIA_ + ',' + _NUMBER_, true, false) then
            begin
                FindInList(Numbers, _SERIA_ + ',' + _NUMBER_, false, true);
                Next;
                continue;
            end;

            if _DIVISION_ = '' then
                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' нет кода подразделения');
            if FieldByName('STATE').AsInteger = 1 then begin //ИСПОРЧЕН
                if not OnlyAvarFlag then begin
                WriteLn(F13, _DIVISION_ + '|8|' +
                            DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                            _SERIA_ + '|' +
                            _NUMBER_ + '|' +
                            _NUMBER_ + '|' +
                            DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                );

                Inc(BadCount);
                end;
            end
            else begin
                    if Trim(FieldByName('NAME').AsString) = '' then
                        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' имя страхователя не указано');
                    if FieldByName('AGENTCODE').AsString <> '' then
                    begin //Нет агента - только 2я оплата
                        if (FieldByName('INSCOMP').AsString = '') OR not (FieldByName('INSCOMP').AsString[1] in ['F', 'U', 'I']) then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' тип агента не определён "' + FieldByName('INSCOMP').AsString + '"');
                        if DecodePeriod(FieldByName('PERIOD').AsString, 0, 0, false) = '?' then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' период не определён "' + FieldByName('PERIOD').AsString + '"');
                        if (GetK(FieldByName('K').AsFloat, FieldByName('version').AsString) = '?') then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' коэффициент местности не правильный "' + FieldByName('K').AsString + '"');
                        if GetAVK(FieldByName('K2').AsFloat, FieldByName('version').AsString) = '?' then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' коэффициент аварийности не правильный "' + FieldByName('K2').AsString + '"');
                        if FieldByName('TONNAGE').AsString <> ''then
                            if FieldByName('TONNAGE').AsFloat > 100 then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' слишком большая грузоподъёмность ' + FieldByName('TONNAGE').AsString + ' тонн');
                        if not FieldByName('STOPDATE').IsNull then
                            if FieldByName('STOPDATE').AsDateTime > FieldByName('TODATE').AsDateTime then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' расторжение после окончания');
                        if FieldByName('STATE').AsString = '5' then
                            if FieldByName('ISPAY2').AsString = 'Y' then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' указана 2я оплата');
                        if DecodeVladen(FieldByName('VLADENIE').AsString, FieldByName('version').AsString) = '?' then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' основание на вледение "' + FieldByName('VLADENIE').AsString + '"');
                        if Length(FieldByName('AUTONUMBER').AsString) > 10 then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' гос. номер очень длиный');
                        korrectAutoNmb := ReplaceStr(ReplaceStr(ReplaceStr(FieldByName('AUTONUMBER').AsString, '-', ''), ' ', ''), ',', '');
                        if korrectAutoNmb <> FieldByName('AUTONUMBER').AsString then
                        begin
                            HandBookSQL.Close;
                            HandBookSQL.SQL.Text := 'update ' + GetTableName() + ' set AUTONUMBER=''' + korrectAutoNmb + ''' where seria=''' + FieldByName('seria').AsString + ''' and number=' + FieldByName('number').AsString;
                            HandBookSQL.ExecSQL;
                        end;

                        CheckAutoNmb(_SERIA_ + '/' + _NUMBER_, korrectAutoNmb, Warn);

                        delta := 31;
                        //Ur
                        //if FieldByName('InsurerType').AsInteger in [0] then delta := 365;
                        //Beznal
                        if FieldByName('PayType').AsInteger in [0] then delta := 365;

                        if (FieldByName('PAYDATE').AsDateTime < FieldByName('FROMDATE').AsDateTime - delta) or (FieldByName('PAYDATE').AsDateTime > FieldByName('FROMDATE').AsDateTime) then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' несоответствие оплаты и периода действия ' + FieldByName('PAYDATE').AsString + ' ' + FieldByName('FROMDATE').AsString + '-' + FieldByName('TODATE').AsString + ' ' + FieldByName('INSCOMP').AsString);

                        //if FieldByName('PREVNUMBER').AsInteger < 1 then //не дубликат
                        //if not FieldByName('FROMDATE').IsNull then
                        //    if FieldByName('FROMDATE').AsDateTime < FieldByName('REGDATE').AsDateTime then
                        //        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' дата выдачи после окончания');

                        if not FieldByName('TODATE').IsNull then
                            if FieldByName('TODATE').AsDateTime < FieldByName('REGDATE').AsDateTime then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' дата выдачи ' + FieldByName('REGDATE').AsString + ' после!!! срока окончания действия ' + FieldByName('TODATE').AsString);

                        if FieldByName('regdate').AsDateTime >= EncodeDAte(2000, 1, 1) then
                        begin
                            //if FieldByName('RATE').AsFloat > 5000 then
                            //        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' курс валюты не правильный');
                            //if not FieldByName('PAY2RATE').IsNull then
                            //if FieldByName('PAY2RATE').AsFloat > 5000 then
                            //        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' курс валюты 2го платежа не правильный');
                        end;
                    end;
                    if FieldByName('STATE').AsString = '2' then
                        if FieldByName('STOPDATE').IsNull then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' не указана дата расторжения');

                    if(not FieldByName('STOPDATE').IsNull) and (not FieldByName('REGDATE').IsNull) then
                        if FieldByName('STOPDATE').AsDateTime < FieldByName('REGDATE').AsDateTime then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' дата расторжения до выдачи полиса');

                    if not bKomplexMode and not mnuNoUnloadAvarias.Checked then begin
                        ExportAvariasGLOBO.Close;
                        ExportAvariasGLOBO.MacroByName('WHERE').AsString := 'A.SERIA=''' + _SERIA_ + ''' AND A.NUMBER=' + _NUMBER_;
                        ExportAvariasGLOBO.Open;
                        while not ExportAvariasGLOBO.Eof do begin
                            Inc(AvCount);

                            if not ExportAvariasGLOBO.FieldByName(Prefix + 'DtFreeRz').IsNull then
                                if ExportAvariasGLOBO.FieldByName(Prefix + 'DtFreeRz').AsDateTime < ExportAvariasGLOBO.FieldByName(Prefix + 'AVDATE').AsDateTime then
                                    raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, Дата освобождения резерва до аварии');

                            if not ExportAvariasGLOBO.FieldByName(Prefix + 'DTCLLOSS').IsNull then
                                if ExportAvariasGLOBO.FieldByName(Prefix + 'DTCLLOSS').AsDateTime < ExportAvariasGLOBO.FieldByName(Prefix + 'AVDATE').AsDateTime then
                                    raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, Дата закрытия резерва до аварии');

                            if ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString <> '' then
                                if GetCharacteristicAuto(ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString,
                                                  100, 100, 100, 100, ExportAvariasGLOBO.FieldByName('Doc1Type').AsInteger <> 2) < 0 then
                                    raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, Тип ТС не правильный');

                            if Length(ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString) < 1 then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, не указан номер дела');

                            Write(Log, 'MANDAV ' + _SERIA_ + '/' + _NUMBER_ + ' ');
                            for n_fld := 0 to ExportAvariasGLOBO.FieldCount - 1 do
                                Write(Log, ExportAvariasGLOBO.Fields[n_fld].FieldName + ':"' + ExportAvariasGLOBO.Fields[n_fld].AsString + '" ');
                            WriteLn(Log);

                            if prevWork = ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' дублирование номера дела ' + prevWork);

                            prevWork := ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString;

                            //Старые и новые аварии
                            RequestFind := false;
                            if ExportAvariasGLOBO.FieldByName(Prefix + 'SUMPAY').IsNull then begin

                            ExportAvariasGLOBO2.Close;
                            ExportAvariasGLOBO2.MacroByName('WHERE').AsString := 'SER=''' + _SERIA_ + ''' AND NMB=' + _NUMBER_ + ' AND N=' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString;
                            ExportAvariasGLOBO2.Open;

                            if ExportAvariasGLOBO2.Eof then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, нет выплат по аварии N' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString);

                            WriteLn(N, Division + '|30|' +
                                          DecodeBlank(MAND_SECT, ExportAvariasGLOBO.FieldByName(Prefix + 'SERIA').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'SERIA').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'NUMBER').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString + '|' +
                                          '004' + '|' + //Заявление
                                          IntToStr(ExportAvariasGLOBO2.FieldByName(Prefix + 'RISK').AsInteger+3)  + '|' +
                                          // '1' + '|' + //ДТП

                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'AVDATE').AsDateTime) + '|' +
                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'WRDATE').AsDateTime) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'AVPLACE').AsString + '|' +
                                          'BYS' + '|' + //Код страны (место страхового случая)
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'CITYCODE').AsString + '|' +
                                          '6' + '|' + //Потерпевший

                                          ExportAvariasGLOBO.FieldByName(Prefix + 'FIO').AsString + '|' +
                                          DecodeFU12(ExportAvariasGLOBO.FieldByName(Prefix + 'FIOTYPE').AsString, '9') + '|' +
                                          DecodeRN(ExportAvariasGLOBO.FieldByName(Prefix + 'ISRESID_P').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'CITYCODE_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'FIOADDRESS').AsString + '|' +
                                          DecodeSN(ExportAvariasGLOBO.FieldByName(Prefix + 'SN').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'MARKA_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'AUTONMB_P').AsString + '|' +
                                          '' + '|' +  //год выпуска
                                          DecodeBodyShas(ExportAvariasGLOBO.FieldByName(Prefix + 'BODY_NO').AsString) + '|' +
                                          DecodeCharact2(ExportAvariasGLOBO.FieldByName(Prefix + 'CHARVAL').AsString, ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString) + '|'); //Дата платежного документа

                                if DecodeAvState(ExportAvariasGLOBO.FieldByName(Prefix + 'WHAT').AsString, ExportAvariasGLOBO.FieldByName('DTCLLOSS'), ExportAvariasGLOBO.FieldByName('DtFreeRz')) <> '' then
                                WriteLn(U, Division + '|30|' +
                                              DecodeBlank(MAND_SECT, ExportAvariasGLOBO.FieldByName('SERIA').AsString) + '|' +
                                              ExportAvariasGLOBO.FieldByName('SERIA').AsString + '|' +
                                              ExportAvariasGLOBO.FieldByName('NUMBER').AsString + '|' +
                                              ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString + '|' +
                                              DecodeAvState(ExportAvariasGLOBO.FieldByName(Prefix + 'WHAT').AsString, ExportAvariasGLOBO.FieldByName('DTCLLOSS'), ExportAvariasGLOBO.FieldByName('DtFreeRz')) + '|' +
                                              DecodeGlDate2(ExportAvariasGLOBO.FieldByName('DTCLLOSS'), ExportAvariasGLOBO.FieldByName('DTCLLOSS'), ExportAvariasGLOBO.FieldByName('DTCLLOSS'), ExportAvariasGLOBO.FieldByName('DTCLLOSS'))
                                );

                                //MOVED UP
                                //ExportAvariasGLOBO2.Close;
                                //ExportAvariasGLOBO2.MacroByName('WHERE').AsString := 'SER=''' + _SERIA_ + ''' AND NMB=' + _NUMBER_ + ' AND N=' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString;
                                //ExportAvariasGLOBO2.Open;

                                WriteLn(Log, 'AVPAYS ' + IntToStr(ExportAvariasGLOBO2.RecordCount));

                                while not ExportAvariasGLOBO2.Eof do begin

                                    try
                                        for n_fld := 0 to ExportAvariasGLOBO2.FieldCount - 1 do
                                            Write(Log, ExportAvariasGLOBO2.Fields[n_fld].FieldName + ':"' + ExportAvariasGLOBO2.Fields[n_fld].AsString + '" ');
                                        WriteLn(Log);
                                    except
                                        WriteLn(Log, 'AV DUMP BUG');
                                    end;

                                    for i := 1 to 95 do begin
                                        if zatr_list[i] = ExportAvariasGLOBO2.FieldByName('ZATR').AsInteger then begin
                                            break;
                                        end;
                                    end;

                                    if i >= 90 then
                                        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария N' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString + ', неопределеный вид затраты. ' + ' ' + ExportAvariasGLOBO2.FieldByName('SUM').AsString + ' ' + ExportAvariasGLOBO2.FieldByName('CURR').AsString);

                                    if ExportAvariasGLOBO2.FieldByName('TYPE').AsString = '0' then begin //Заявлено
                                    WriteLn(Q, Division + '|30|' +
                                                  DecodeBlank(MAND_SECT, ExportAvariasGLOBO2.FieldByName('SER').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('SER').AsString + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('NMB').AsString + '|' +
                                                  ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString + '|' +
                                                  DecodeGlDate(ExportAvariasGLOBO2.FieldByName('D').AsDateTime) + '|' +
                                                  '|' + //Doc Code
                                                  DecodeZatr(ExportAvariasGLOBO2.FieldByName('ZATR').AsString) + '|' +
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('RISK').AsInteger + 1) + '|' +
                                                  DecodeCurrency(ExportAvariasGLOBO2.FieldByName('CURR').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('SUM').AsString + '|' +
                                                  Format('%0.2f', [ExportAvariasGLOBO2.FieldByName('SUM').AsFloat *
                                                                   GetRateOfCurrency(ExportAvariasGLOBO2.FieldByName('CURR').AsString,
                                                                                      ExportAvariasGLOBO2.FieldByName('D').AsDateTime) /
                                                                   GetRateOfCurrency('EUR', ExportAvariasGLOBO2.FieldByName('D').AsDateTime, 'Заявлено')])  + '|' + //in polis currency
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('RCVR').AsInteger + 1) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('NAME').AsString
                                    );
                                    RequestFind := true;
                                    end
                                    else
                                    if ExportAvariasGLOBO2.FieldByName('TYPE').AsString = '1' then //Оплата
                                    WriteLn(V, Division + '|30|' +
                                                  DecodeBlank(MAND_SECT, ExportAvariasGLOBO2.FieldByName('SER').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('SER').AsString + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('NMB').AsString + '|' +
                                                  ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString + '|' +
                                                  DecodeCurrency(ExportAvariasGLOBO2.FieldByName('CURR').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('SUM').AsString + '|' +
                                                  Format('%0.2f', [ExportAvariasGLOBO2.FieldByName('SUM').AsFloat *
                                                                       GetRateOfCurrency(ExportAvariasGLOBO2.FieldByName('CURR').AsString,
                                                                                         ExportAvariasGLOBO2.FieldByName('D').AsDateTime) /
                                                                       GetRateOfCurrency('EUR', ExportAvariasGLOBO2.FieldByName('D').AsDateTime, 'Оплата')])  + '|' + //in polis currency
                                                  DecodePayType(ExportAvariasGLOBO2.FieldByName('NAL').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('DOC').AsString + '|' +
                                                  DecodeGlDate(ExportAvariasGLOBO2.FieldByName('D').AsDateTime) + '|' +

                                                  DecodeZatr(ExportAvariasGLOBO2.FieldByName('ZATR').AsString) + '|' +
                                                  //13
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('RISK').AsInteger + 1) + '|' +
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('SUBTYPE2').AsInteger) + '|' +
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('RCVR').AsInteger + 1) + '|' +
                                                  //DecodeZatr(ExportAvariasGLOBO2.FieldByName('ZATR').AsString) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('NAME').AsString

                                                  {old version
                                                  DecodeZatr(ExportAvariasGLOBO2.FieldByName('ZATR').AsString) + '|' +
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('RISK').AsInteger + 1) + '|' +
                                                  IntToStr(ExportAvariasGLOBO2.FieldByName('SUBTYPE2').AsInteger + 1) + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('RCVR').AsString + '|' +
                                                  ExportAvariasGLOBO2.FieldByName('NAME').AsString
                                                  }
                                    )
                                    else
                                        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, неопределенный признак по аварии N' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString);
                                    ExportAvariasGLOBO2.Next;
                                end;
                                if RequestFind = false then
                                    raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' авария, нет ЗАЯВЛЕННЫХ СУММ по аварии N' + ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString);
                            end
                            else begin
                            WriteLn(AVAR, Division + '|30|' +
                                          DecodeBlank(MAND_SECT, ExportAvariasGLOBO.FieldByName(Prefix + 'SERIA').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'SERIA').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'NUMBER').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'N').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'NWORK').AsString + '|' +
                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'AVDATE').AsDateTime) + '|' +
                                          DecodeGlDate(ExportAvariasGLOBO.FieldByName(Prefix + 'WRDATE').AsDateTime) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'AVPLACE').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'CITYCODE').AsString + '|' +
                                          DecodeAvState(ExportAvariasGLOBO.FieldByName(Prefix + 'WHAT').AsString, nil, nil) + '|' +

                                          DecodeGlDate2(ExportAvariasGLOBO.FieldByName(Prefix + 'DTCLLOSS'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'V1_DATE'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'V2_DATE'),
                                                       ExportAvariasGLOBO.FieldByName(Prefix + 'V3_DATE')
                                                       ) + '|' +

                                          ExportAvariasGLOBO.FieldByName(Prefix + 'FIO').AsString + '|' +
                                          DecodeFU12(ExportAvariasGLOBO.FieldByName(Prefix + 'FIOTYPE').AsString, '2') + '|' +
                                          DecodeRN(ExportAvariasGLOBO.FieldByName(Prefix + 'ISRESID_P').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'CITYCODE_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'FIOADDRESS').AsString + '|' +
                                          DecodeSN(ExportAvariasGLOBO.FieldByName(Prefix + 'SN').AsString) + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'MARKA_P').AsString + '|' +
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'AUTONMB_P').AsString + '|' +
                                          '' + '|' +  //год выпуска
                                          DecodeBodyShas(ExportAvariasGLOBO.FieldByName(Prefix + 'BODY_NO').AsString) + '|' +
                                          DecodeCharact2(ExportAvariasGLOBO.FieldByName(Prefix + 'CHARVAL').AsString, ExportAvariasGLOBO.FieldByName(Prefix + 'BASE_TYPE_P').AsString) + '|' +
                                          DecodeCurrency(ExportAvariasGLOBO.FieldByName(Prefix + 'V1').AsString) + '|' + //Код валюты (заявленная)
                                          Format('%0.2f', [ExportAvariasGLOBO.FieldByName(Prefix + 'DECLTC').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DECLHL').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DECLIM').AsFloat]) + '|' + //Заявленная сумма
                                          DecodeCurrency(ExportAvariasGLOBO.FieldByName(Prefix + 'V1').AsString) + '|' + //Код валюты выплаты
                                          ExportAvariasGLOBO.FieldByName(Prefix + 'SUMPAY').AsString + '|' + //Выплачено
                                          '2' + '|' + //Форма оплаты
                                          '' + '|' + //Номер платежного документа
                                          DecodeOneDate(ExportAvariasGLOBO.FieldByName(Prefix + 'V1_DATE'), ExportAvariasGLOBO.FieldByName(Prefix + 'V2_DATE'), ExportAvariasGLOBO.FieldByName(Prefix + 'V3_DATE')) + '|'); //Дата платежного документа

                            if (ExportAvariasGLOBO.FieldByName(Prefix + 'DECLTC').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DECLHL').AsFloat + ExportAvariasGLOBO.FieldByName(Prefix + 'DECLIM').AsFloat) = 0 then
                                if FieldByName('version').AsString > '1' then
                                    raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' нет требований по суммам в аварии');
                            end;

                            ExportAvariasGLOBO.Next;
                        end;
                    end;

                    //or (FieldByName('BASECARCODE').AsInteger = 100)

                    if not bKomplexMode and (OnlyAvarFlag) and (ExportAvariasGLOBO.RecordCount = 0) then begin
                        //WriteLn(Log, 'Не подотчетный');
                        Next;
                        continue;
                    end;

                    if (FieldByName('ISPAY2').AsInteger > 0) then begin
                        if (FieldByName('AllFeeText').AsString <> '') AND (FieldByName('AllFeeText').AsString[1] in ['Y','0','A','B','C','D','E','F','G','H','I','J','K','L','M']) then begin //1СУ
                            WriteLn(F13, _DIVISION_ + '|27|' +
                                        DecodeSeria1SU(FieldByName('AllFeeText').AsString) + '|' + //12
                                        GetSeria1SU(FieldByName('AllFeeText').AsString) + '|' +  // 'ГС' {_SERIA_} + '|' +
                                        FieldByName('PAY2DRAFTNUMBER').AsString + '|' +
                                        FieldByName('PAY2DRAFTNUMBER').AsString + '|' +
                                        DecodeGlDate(FieldByName('PAY2DATE').AsDateTime)
                            );
                           if Length(FieldByName('PAY2DRAFTNUMBER').AsString)<_1SULen then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' некорректный номер 1СУ. ' + FieldByName('PAY2DRAFTNUMBER').AsString + '. Длина меньше ' + IntToStr(_1SULen));
                        end
                        else
                        if FieldByName('AllFeeText').AsString <> 'N' then begin
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' некорректный тип 1СУ. ' + FieldByName('AllFeeText').AsString);
                        end
                    end;

                    //if(FieldByName('AllFeeText').AsString = 'Y') then //1СУ
                    //    if FieldByName('ISPAY2').AsInteger < 1 then
                    //        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' признак 1-СУ и нет 2й оплаты');


                    if FieldByName('AGENTCODE').AsString <> '' then
                    if (not bKomplexMode and (ExportAvariasGLOBO.RecordCount <> 0)) or bUnloadAllPolises or (not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2'))) then begin
                    WriteLn(D, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                DecodeFU16(FieldByName('INSCOMP').AsString, FieldByName('AGPERCENT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('FROMDATE').AsDateTime, FieldByName('TODATE').AsDateTime, false) + '|' +
                                DecodeGlDate(FieldByName('FROMDATE').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('FROMTIME').AsString) + '|' +
                                DecodeGlDate(FieldByName('TODATE').AsDateTime) + '|' +
                                DecodeVariant(FieldByName('RESIDENT').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('NAME').AsString + '|' +
                                DecodeFU12(FieldByName('INSURERTYPE').AsString, FieldByName('version').AsString) + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                FieldByName('CITYCODE').AsString + '|' +
                                FieldByName('ADDRESS').AsString + '|' +
                                DecodeOwnTypeForm(FieldByName('INSURERTYPE').AsString, FieldByName('version').AsString) + '|' + //Форма собственности
                                DecodePayGraphic(FieldByName('ISFEE2').AsString) + '|' + //График оплаты
                                DecodePayType(FieldByName('PAYTYPE').AsString) + '|' +
                                DecodeAgentNamePx(FieldByName('AGENTCODE').AsString) + '|' +
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма по полису (на весь период)
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('EUR') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('TARIF').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TARIFGLASS').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('EUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' //Для заметок
                    );
                    PrevKod := '';
                    PrevSeria := '';
                    PrevNumber := '';
                    if FieldByName('PREVNUMBER').AsInteger < 0 then begin
                        PrevKod := DecodeBlank(MAND_SECT, FieldByName('PREVSERIA').AsString);
                        PrevSeria := FieldByName('PREVSERIA').AsString;
                        PrevNumber := IntToStr(-FieldByName('PREVNUMBER').AsInteger);
                    end;
                    WriteLn(X,  _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                DecodeFU16(FieldByName('INSCOMP').AsString, FieldByName('AGPERCENT').AsFloat) + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime) + '|' +
                                '' + '|' + //time
                                DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('FROMDATE').AsDateTime, FieldByName('TODATE').AsDateTime, false) + '|' +
                                DecodeGlDate(FieldByName('FROMDATE').AsDateTime) + '|' +
                                DecodeTimeX(FieldByName('FROMTIME').AsString) + '|' +
                                DecodeGlDate(FieldByName('TODATE').AsDateTime) + '|' +
                                DecodeVariant(FieldByName('RESIDENT').AsString) + '|' + //Код варианта страхования
                                '1' + '|' + //Количество застрахованных
                                FieldByName('NAME').AsString + '|' +
                                DecodeFU12(FieldByName('INSURERTYPE').AsString, FieldByName('version').AsString) + '|' +
                                DecodeR(FieldByName('RESIDENT').AsString) + '|' +
                                FieldByName('CITYCODE').AsString + '|' +
                                RemoveTel(FieldByName('ADDRESS').AsString, FieldByName('AGENTCODE').AsString) + '|' +
                                DecodeOwnTypeForm(FieldByName('INSURERTYPE').AsString, FieldByName('version').AsString) + '|' + //Форма собственности
                                DecodePayGraphic(FieldByName('ISFEE2').AsString) + '|' + //График оплаты
                                DecodePayType(FieldByName('PAYTYPE').AsString) + '|' +
                                DecodeAgentNamePx(FieldByName('AGENTCODE').AsString) + '|' +
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма по полису (на весь период)
                                DecodeInsSum(FieldByName('REGDATE').AsDateTime) + '|' + //Страховая сумма на один случай по полису
                                DecodeCurrency('EUR') + '|' + //Код валюты (страховой суммы)
                                Format('%f', [FieldByName('TARIF').AsFloat]) + '|' + //Страховая премия
                                Format('%f', [FieldByName('TARIFGLASS').AsFloat]) + '|' + //Страховой тариф (сумма)
                                DecodeCurrency('EUR') + '|' + //Код валюты (премии)
                                '' + '|' + //БСT  (процент)
                                '' + '|' + //Серия спецзнака (карточки)
                                '' + '|' + //Номер спецзнака (карточки)
                                '' + '|' + //Для заметок
                                FieldByName('MOBTEL').AsString + '|' +
                                FieldByName('HOMETEL').AsString + '|' +
                                DecodeAgentID(FieldByName('AGENTCODE').AsString) + '|' + //Agent Doverennost
                                FieldByName('SMS').AsString + '|' +
                                FieldByName('CONTACT').AsString + '|' +
                                PrevKod + '|' +
                                PrevSeria + '|' +
                                PrevNumber
                    );
                    Inc(PolisCount);
                    end;
                    {WriteLn(F2, _DIVISION_ + '|30|' +
                                DecodeBlank(_SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '1' + '|' + //Номер застрахованного
                                DecodeCurrency('EUR') + '|' + //Код валюты
                                '10000' + '|' + //Страховая сумма
                                FieldByName('NAME').AsString + '|' +
                                '' + '|' + //Дата рождения
                                '' + '|' + //Паспорт
                                FieldByName('ADDR').AsString
                    );}

            //Images
            if bKomplexMode then
            begin
                 if SysUtils.FindFirst(DB.Directory + '\IMG\' + ToENG(_SERIA_) + '\' + _33Code(_NUMBER_) + '\*.*', faAnyFile, searchResult) = 0 then
                  begin
                    repeat

                    if (searchResult.Attr AND faDirectory) <> 0 then continue;
                    Inc(ImgCount);
                    WriteLn(E, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '01' + '|' + //фотография объекта страхования (  01 - фотография объекта страхования
                                                //  02 - паспорт
                                                //  03 - водительское удостоверение )
                                _SERIA_ + '-' + _NUMBER_ + '-' + searchResult.Name + '|' + //FileName
                                ''); //Comments

                        fileSource := DB.Directory + '\IMG\' + ToENG(_SERIA_) + '\' + _33Code(_NUMBER_) + '\' + searchResult.Name;

                        try
                        CopyFile(PChar(fileSource), PChar(path + imgDir + '\' + _SERIA_ + '-' + _NUMBER_ + '-' + searchResult.Name), nil);
                        except
                        on e : Exception do
                        begin
                            FindClose(searchResult);
                          raise Exception.Create('не могу скопировать файл ' + searchResult.Name + '. Ошибка ' + IntToStr(GetLastError()) + ' ' + e.Message);
                        end;
                        end;
                    until SysUtils.FindNext(searchResult) <> 0;

                    // Должен освободить ресурсы, используемые этими успешными, поисками
                    FindClose(searchResult);
                  end;

            end;

                    if not bKomplexMode and not TaxiTbl.Active then
                         TaxiTbl.Open;

                    TaxiFound := not bKomplexMode and TaxiTbl.Locate('SER;N', vararrayof([_SERIA_, FieldByName('NUMBER').AsInteger]), []);

                    if TaxiFound then WriteLn(Log, 'TAXI ' + _SERIA_ + '/' + _NUMBER_);

                    if FieldByName('AGENTCODE').AsString <> '' then
                    if bUnloadAllPolises or (not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2'))) then begin
                        Letter := FieldByName('LETTER').AsString;
                        if TaxiFound and (TaxiTbl.FieldByName('SUMMA').AsFloat > 0) and (TaxiTbl.FieldByName('TYPE').AsString = '') then Letter := 'A6';
                        WriteLn(F3, _DIVISION_ + '|' + InsTypeCode + '|' +
                                    DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                    _SERIA_ + '|' +
                                    _NUMBER_ + '|' +
                                    '1' + '|' + //Номер застрахованного
                                    Letter + '|' + //Тип транспортного средства
                                    FieldByName('MARKA').AsString + '|' +
                                    DecodeNone(korrectAutoNmb, 'НЕ УКАЗАНО') + '|' +
                                    '' + '|' + //Год выпуска
                                    FieldByName('NUMBERBODY').AsString + '|' +
                                    FieldByName('CHASSIS').AsString + '|' +
                                    '' + '|' + //Серия техпаспорта
                                    '' + '|' + //Номер техпаспорта
                                    DecodeCharact(FieldByName('LETTER').AsString,
                                                  FieldByName('VOLUME').AsString,
                                                  FieldByName('TONNAGE').AsString,
                                                  FieldByName('QUANTITY').AsString,
                                                  FieldByName('POWER').AsString) + '|' +
                                    '' + '|' + //Стоимость
                                    '' + '|' + //Код валюты (стоимости и страховой суммы)
                                    '' + '|' + //Страховая сумма по объекту
                                    DecodeVladen(FieldByName('VLADENIE').AsString, FieldByName('version').AsString) + '|' +
                                    '' + '|'//Символ страны прибытия
                        );
                    end;
                    if FieldByName('AGENTCODE').AsString <> '' then
                    if FieldByName('DISCOUNT').AsFloat <> 0 then
                    if bUnloadAllPolises or not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    WriteLn(F6, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '|' + //Номер застрахованного не указан
                                'H|' + //Код коэффициента
                                '3|' + //Единица измерения Коэффициент
                                FieldByName('DISCOUNT').AsString  + '|' //Значение коэффициента
                    );

                    if FieldByName('REGDATE').AsDateTime >= EncodeDate(2006, 9, 1) then
                    if FieldByName('AGENTCODE').AsString <> '' then
                    //if FieldByName('Placement').AsString <> '' then
                    if bUnloadAllPolises or not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    if (K3Always.Checked) or (FieldByName('InsurerType').AsInteger <> 0) then //не юр лицо
                    begin
                        if not(FieldByName('InsurerType').AsInteger <> 0) then
                        if FieldByName('Placement').AsString = '' then goto skipK3;

                        if DecodeK3((Ord(FieldByName('Placement').AsString[1]) - 33) / 10) = '?' then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' K3 некорректный');

                        WriteLn(F6, _DIVISION_ + '|' + InsTypeCode + '|' +
                                    DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                    _SERIA_ + '|' +
                                    _NUMBER_ + '|' +
                                    '|' + //Номер застрахованного не указан
                                    DecodeK3((Ord(FieldByName('Placement').AsString[1]) - 33) / 10) + '|' + //Код коэффициента
                                    '4|' + //Единица измерения Коэффициент !!!!!!!!!
                                    FloatToStr((Ord(FieldByName('Placement').AsString[1]) - 33) / 10)  + '|' //Значение коэффициента
                        );
                    end;

                    skipK3:;

                    if FieldByName('AGENTCODE').AsString <> '' then
                    if DecodeR(FieldByName('RESIDENT').AsString) = '1' then //резидент
                    if bUnloadAllPolises or not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    WriteLn(F6, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '|' + //Номер застрахованного не указан
                                 GetK(FieldByName('K').AsFloat, FieldByName('version').AsString) + '|' + //Код коэффициента
                                '4|' + //Единица измерения Коэффициент
                                FieldByName('K').AsString  + '|' //Значение коэффициента
                    );
                    if FieldByName('AGENTCODE').AsString <> '' then
                    if bUnloadAllPolises or not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    WriteLn(F6, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '|' + //Номер застрахованного не указан
                                 GetAvK(FieldByName('K2').AsFloat, FieldByName('version').AsString) + '|' + //Код коэффициента
                                '4|' + //Единица измерения Коэффициент
                                FieldByName('K2').AsString + '|'//Значение коэффициента
                    );
                    if FieldByName('AGENTCODE').AsString <> '' then
                    if FieldByName('PREVNUMBER').AsInteger <= 0 then //не дубликат
                    if not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    //Расторгнутые полисы не нужен платеж ()
                    if FieldByName('STOPDATE').IsNull or not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('STOPDATE')) then begin //true если попал в период
                        WriteLn2(P0, A, _DIVISION_ + '|' + InsTypeCode + '|' +
                                    DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                    _SERIA_ + '|' +
                                    _NUMBER_ + '|' +
                                    '5' + '|' + //Код события
                                    DecodeNB(FieldByName('PAYTYPE').AsString, FieldByName('DRAFTNUMBER').AsString) + '|' + //Форма оплаты
                                    KillZero(FieldByName('DRAFTNUMBER').AsString) + '|' +
                                    DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' + //Дата платежного документа
                                    DecodeGlDate(FieldByName('PAYDATE').AsDateTime) + '|' +
                                    DecodeGlDate(FieldByName('RET1').AsDateTime) + '|' +
                                    FieldByName('PAY1').AsString + '|' + //Сумма
                                    DecodeCurrency(FieldByName('PAY1CURR').AsString) + '|' + //Код валюты
                                    FieldByName('AGPERCENT').AsString + '|' + //Комиссия
                                    DecodeAgentNamePx(FieldByName('AGENTCODE').AsString) + '|' +
                                    DecodeFU16(FieldByName('INSCOMP').AsString, FieldByName('AGPERCENT').AsFloat),
                                    DecodeAgentID(FieldByName('AGENTCODE').AsString)
                        );
                        Inc(PayCount);
                    end;
                    //2й платеж выгружаем если он у первого полиса который не утерян или утерян после оплаты
                    //или если у утерянного нету платежа 2го платежа и дата дубликата до оплаты

                    WriteLn(Log);
                    WriteLn(Log, 'N' + IntToStr(Curr) + ': ' + _SERIA_ + '/' + _NUMBER_ + ' ISPAY2:' + FieldByName('ISPAY2').AsString + ' STATE:' + FieldByName('STATE').AsString + ' PREVNUMBER:' + FieldByName('PREVNUMBER').AsString + ' REGDATE:' + FieldByName('REGDATE').AsString + ' PAY2DATE:' + FieldByName('PAY2DATE').AsString + ' STOPDATE:' + FieldByName('STOPDATE').AsString);

                    if FieldByName('ISPAY2').AsInteger = 1 then begin
                        if((FieldByName('STATE').AsInteger = 3) AND FieldByName('STOPDATE').IsNull) then
                            raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' утерян, но не указана дата утери');
                    end;

                    //3 "УТЕРЯН";
                    if FieldByName('ISPAY2').AsInteger = 1 then
                    //не дубликат не утерян или утерян после оплаты
                    if (FieldByName('PREVNUMBER').AsInteger <= 0) AND ((FieldByName('STATE').AsInteger <> 3) or (FieldByName('STATE').AsInteger = 3) AND (FieldByName('STOPDATE').AsDateTime > FieldByName('PAY2DATE').AsDateTime)) or
                    //дубликат + выдан до оплаты (не утерян или утерян после оплаты)
                        ((FieldByName('PREVSERIA').AsString <> '') AND (FieldByName('REGDATE').AsDateTime <= FieldByName('PAY2DATE').AsDateTime) AND ((FieldByName('STATE').AsInteger <> 3) OR ((FieldByName('STATE').AsInteger = 3) AND (FieldByName('STOPDATE').AsDateTime > FieldByName('PAY2DATE').AsDateTime)))) then begin
                        //((FieldByName('PREVSERIA').AsString <> '') AND NotExistsPay2(FieldByName('PREVSERIA').AsString, FieldByName('PREVNUMBER').AsString, FieldByName('REGDATE').AsDateTime) AND (FieldByName('REGDATE').AsDateTime <= FieldByName('PAY2DATE').AsDateTime)) then begin
                    WriteLn2(P0, A,  _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '5' + '|' + //Код события
                                DecodeNB(FieldByName('PAY2TYPE').AsString, FieldByName('PAY2DRAFTNUMBER').AsString) + '|' + //Форма оплаты
                                KillZero(FieldByName('PAY2DRAFTNUMBER').AsString) + '|' +
                                DecodeGlDate(FieldByName('PAY2DATE').AsDateTime) + '|' + //Дата платежного документа
                                DecodeGlDate(FieldByName('PAY2DATE').AsDateTime) + '|' +
                                DecodeGlDate(FieldByName('RET2').AsDateTime) + '|' +
                                FieldByName('PAY2SUMMA').AsString + '|' + //Сумма
                                DecodeCurrency(FieldByName('PAY2CURR').AsString) + '|' + //Код валюты
                                FieldByName('CarCode').AsString + '|' + //Комиссия
                                DecodeAgentNamePx(FieldByName('Fee2Text').AsString) + '|' +
                                DecodeFU16(FieldByName('country').AsString, FieldByName('CarCode').AsFloat),
                                DecodeAgentID(FieldByName('Fee2Text').AsString)
                    );
                    if (FieldByName('PAY2CURR').AsString = 'BRB') AND (FieldByName('PAY2SUMMA').AsFloat < MinSumma) then
                        raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' 2я оплата слишком мала. ini [MANDATORY] MinSumma');
                    Inc(PayCount);
                    end;

                   if TaxiFound then begin
                          if (TaxiTbl.FieldByName('SUMMA').AsFloat > 0) then
                            MsgCode := '5' //Код события
                          else
                            MsgCode := '51'; //Код события

                          WriteLn2(P, A_, _DIVISION_ + '|' + InsTypeCode + '|' +
                                      DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                      _SERIA_ + '|' +
                                      _NUMBER_ + '|' +
                                      MsgCode + '|' +
                                      DecodeNB(TaxiTbl.FieldByName('NAL').AsString, TaxiTbl.FieldByName('PAYDOC').AsString) + '|' + //Форма оплаты
                                      KillZero(TaxiTbl.FieldByName('PAYDOC').AsString) + '|' +
                                      DecodeGlDate(TaxiTbl.FieldByName('D').AsDateTime) + '|' + //Дата платежного документа
                                      DecodeGlDate(TaxiTbl.FieldByName('D').AsDateTime) + '|' +
                                      DecodeGlDate(TaxiTbl.FieldByName('D').AsDateTime) + '|' +
                                      TaxiTbl.FieldByName('SUMMA').AsString + '|' + //Сумма
                                      DecodeCurrency(TaxiTbl.FieldByName('CURR').AsString) + '|' + //Код валюты
                                      FieldByName('CARCODE').AsString + '|' + //Комиссия
                                      DecodeAgentNamePx(FieldByName('FEE2TEXT').AsString) + '|' +
                                      DecodeFU16(FieldByName('COUNTRY').AsString, FieldByName('CARCODE').AsFloat),
                                      DecodeAgentID(FieldByName('AGENTCODE').AsString)
                          );
                          if TaxiTbl.FieldByName('SU').AsString = 'Y' then begin //1СУ
                          WriteLn(F13, _DIVISION_ + '|27|' +
                                      '12' {DecodeBlank(_SERIA_)} + '|' +
                                      'ГС' {_SERIA_} + '|' +
                                      TaxiTbl.FieldByName('PAYDOC').AsString + '|' +
                                      TaxiTbl.FieldByName('PAYDOC').AsString + '|' +
                                      DecodeGlDate(TaxiTbl.FieldByName('D').AsDateTime)
                          );
                            if Length(TaxiTbl.FieldByName('PAYDOC').AsString)<_1SULen then
                                raise Exception.Create(_SERIA_ + '/' + _NUMBER_ + ' некрректный номер 1СУ в доплате. ' + TaxiTbl.FieldByName('PAYDOC').AsString);
                          end;
                    end;

                    if (FieldByName('STATE').AsString = '2') AND (FieldByName('RETEUR').AsFloat >= 0.01) then begin
                        Dt := FieldByName('STOPDATE').AsDateTime;
                        if(not FieldByName('OWNERCODE').IsNull AND (FieldByName('OWNERCODE').AsInteger > 5000)) then
                            Dt := FieldByName('OWNERCODE').AsInteger - (731953 - EncodeDate(2005, 1, 7));
                        WriteLn2(P0, A, _DIVISION_ + '|' + InsTypeCode + '|' +
                                    DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                    _SERIA_ + '|' +
                                    _NUMBER_ + '|' +
                                    '51' + '|' + //Код события
                                    DecodeNB('N', '') + '|' + //Форма оплаты
                                    '|' +
                                    DecodeGlDate(Dt) + '|' + //Дата платежного документа
                                    DecodeGlDate(Dt) + '|' +
                                    DecodeGlDate(Dt) + '|' +
                                    FieldByName('RETBRB').AsString + '|' + //Сумма
                                    DecodeCurrency('BRB') + '|' + //Код валюты
                                    '|' + //Комиссия
                                    '|' +
                                    DecodeFU16('', 0),
                                    ''
                        );
                    end;
                    {if (FieldByName('ISFEE2').AsString = 'Y' ) AND (FieldByName('ISPAY2').AsString = 'N') then
                    WriteLn(F12, _DIVISION_ + '|30|' +
                                DecodeBlank(_SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                DecodeGlDate(FieldByName('FEE2DT').AsDateTime) + '|' +
                                Format('%0.2f', [FieldByName('FEE2').AsFloat]) + '|' + //Сумма
                                DecodeCurrency('EUR') //Код валюты
                    );}
                //end
                //else begin
                if FieldByName('PREVNUMBER').AsFloat > 0 then begin
                    if not IsOnlySecondPay(FieldByName('REGDATE'), FieldByName('RET2')) then
                    WriteLn(F10, _DIVISION_ + '|' + InsTypeCode + '|' +
                                DecodeBlank(MAND_SECT, FieldByName('PREVSERIA').AsString) + '|' +
                                FieldByName('PREVSERIA').AsString + '|' +
                                FieldByName('PREVNUMBER').AsString + '|' +
                                DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                                _SERIA_ + '|' +
                                _NUMBER_ + '|' +
                                '9' + '|' +
                                DecodeGlDate(FieldByName('REGDATE').AsDateTime)
                    );
                end;
                if FieldByName('STATE').AsString = '5' then
                WriteLn(F11, _DIVISION_ + '|' + InsTypeCode + '|' +
                            DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                            _SERIA_ + '|' +
                            _NUMBER_ + '|' +
                            '|' +
                            DecodeGlDate(FieldByName('STOPDATE').AsDateTime) + '|' +
                            '2' + '|' +
                            '2' + '|' + //расторгнут по неуплате
                            '2'
                );
                if FieldByName('STATE').AsString = '2' then
                WriteLn(F11, _DIVISION_ + '|' + InsTypeCode + '|' +
                            DecodeBlank(MAND_SECT, _SERIA_) + '|' +
                            _SERIA_ + '|' +
                            _NUMBER_ + '|' +
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
                clipboard.AsText := WorkSQL.SQL.Text;
                Inc(ErrorCount);
                errStr := errStr + #13#10 + e.Message + ' ' + DecodeAgentNamePx(FieldByName('AGENTCODE').AsString);
                errStr := errStr + ' (' + _NUMBER_ + ')';
                IsError := true;
                if Length(errStr) > 10240 then break;
                Next;
            end
        end; //while
    end; //width}
  end;

  firstFlag := false;
  if unloadCycles > 0 then begin
    Dec(unloadCycles);
    UnloadedList.Sort;
    if Numbers.Count > 0 then begin
        WriteLn(Log, 'Догрузка данных по дубликатам/утерянным. Указано ' + IntToStr(Numbers.Count) + ' штук');
        errStr := errStr + 'Догрузка данных: ' + IntToStr(Numbers.Count) + ' полисов';
        goto RestartExport;
    end
  end;

    if (Numbers.Count > 0) then
    begin
        WriteLn(Log, 'Догрузка данных превысила лимит повторов. Осталось невыгружено ' + IntToStr(Numbers.Count) + ' штук');
    end;

  StatusBar.Panels[0].Text := '';
   if bKomplexMode then     CloseFile(E);
    CloseFile(D);
    CloseFile(X);
    CloseFile(A_);
    CloseFile(F3);
    CloseFile(F6);
    CloseFile(P0);
    CloseFile(A);
    CloseFile(P);
    CloseFile(F10);
    CloseFile(F11);
    CloseFile(F13);
    CloseFile(AVAR);
    CloseFile(N);
    CloseFile(Q);
    CloseFile(V);
    CloseFile(U);
    ProgressBar.Position := 0;
    if Numbers <> nil then begin
        if Numbers.Count > 0 then begin
            for i := 0 to Numbers.Count - 1 do begin
                StrNumbers := StrNumbers + #13 + Numbers[i];
                if i > 20 then begin
                    StrNumbers := StrNumbers + #13 + '...';
                    break;
                end;
            end;
            MessageDlg('Из указанного списка полисов не обнаружено ' + IntToStr(Numbers.Count) + ' штук'#13 + StrNumbers, mtInformation, [mbOk], 0)
        end;
        Numbers.Free;
    end;

    for i := 0 to Warn.Count - 1 do begin
        WarnStr := WarnStr + Warn[i] + #13;
    end;
    if Warn <> nil then Warn.Free;

    if not IsError then begin
        WriteLn(Log, 'Экспорт выполнен успешно'#13 + Format('Выгружено %d полисов, %d платежей, %d испорченных, %d аварий', [PolisCount, PayCount, BadCount, AvCount]));

        if bKomplexMode then
        begin
                if MessageDlg('Экспорт выполнен успешно'#13 + Format('Выгружено %d полисов, %d платежей, %d испорченных, %d изображений'#13'Путь куда скопированы изображения ' + path + imgDir + #13'Открыть папку с изображениями?', [PolisCount, PayCount, BadCount, ImgCount]), mtInformation, [mbYes ,  mbNo], 0) = mrYes then
                        ShellExecute(0, 'open', StrPCopy(pbuff, path + imgDir), 0, 0, SW_SHOWNORMAL);
        end
        else
        begin
                MessageDlg('Экспорт выполнен успешно'#13 + Format('Выгружено %d полисов, %d платежей, %d испорченных, %d аварий', [PolisCount, PayCount, BadCount, AvCount]), mtInformation, [mbOk], 0);
        end;
        ShowInfo.Caption := 'ПРЕДУПРЕЖДЕНИЯ!!! Возможно ошибки';
        ShowInfo.InfoPanel.Text := WarnStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        WriteLn(Log, WarnStr);
        if WarnStr <> '' then  ShowInfo.ShowModal;
    end
    else begin
        if MessageDlg('Обнаружены ошибки. Удалить полученные файлы?', mtInformation, [mbYes, mbNo], 0) = mrYes then
        begin
            for Ltr := 'A' to 'Z' do
                DeleteFile(path + sDivision + Ltr + UnloadDt + UnloadN + '.0' + InsTypeCode);
            DeleteFile(path + sDivision + 'P-' + UnloadDt + UnloadN + '.0' + InsTypeCode);
        end;
        ShowInfo.Caption := 'Ошибки в данных';
        WriteLn(Log, 'Ошибки в данных');
        WriteLn(Log, errStr);
        ShowInfo.InfoPanel.Text := errStr + #13 + WarnStr;
        SetActiveWindow(Handle);
        Application.MainForm.BringToFront;
        ShowInfo.ShowModal;
    end;
    CloseFile(Log);
end;

procedure TMainForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     FreeWindow;
     MainForm.WindowState := wsMinimized;
     try
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('SELECT 1 FROM ' + GetTableName() + ' WHERE BASECARCODE IS NULL OR BASECARCODE NOT IN (101, 100)');
     WorkSQL.Open;
     if WorkSQL.Eof then exit;

     SetThreadPriority(GetCurrentThread, THREAD_PRIORITY_IDLE);

     WorkSQL.Close;
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('UPDATE ' + GetTableName() + ' SET BASECARCODE=101 WHERE BASECARCODE IS NULL OR (BASECARCODE <> 100)');
     WorkSQL.ExecSQL;
     WorkSQL.Close;
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('UPDATE ' + GetTableName() + ' SET FEE2TEXT=AGENTCODE,CARCODE=AGPERCENT,COUNTRY=INSCOMP WHERE ISPAY2=1 AND COUNTRY=''''');
     WorkSQL.ExecSQL;
     WorkSQL.Close;
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('UPDATE ' + GetTableName() + ' SET ISPAY2=0 WHERE ISPAY2 IS NULL');
     WorkSQL.ExecSQL;
     WorkSQL.Close;
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('UPDATE ' + GetTableName() + ' SET RET2=PAY2DATE WHERE RET2 IS NULL AND PAY2DATE IS NOT NULL');
     WorkSQL.ExecSQL;
     WorkSQL.Close;
     WorkSQL.SQL.Clear;
     WorkSQL.SQL.Add('UPDATE ' + GetTableName() + ' SET FEE2TEXT='''' WHERE ISPAY2=0');
     WorkSQL.ExecSQL;
     except
        on e : exception do begin
         MainForm.WindowState := wsMaximized;
         MessageDlg(e.Message, mtInformation, [mbOk], 0);
        end;
     end;
end;

procedure TMainForm.mnuNoUnloadAvariasClick(Sender: TObject);
begin
    mnuNoUnloadAvarias.Checked := not mnuNoUnloadAvarias.Checked;
end;

procedure InitAgRep(Data : PAgReportStruct; PAYTYPE : Integer; PAYDATE : TDate; PAY : double; CURR : String; RATE : double; PCNT : double);
var
    i : integer;
    Ptr : ^AgReportPayListStruct;
begin
    if PAYTYPE = 1 then
        Ptr := @Data^.NalSumm
    else
        Ptr := @Data^.BezNalSumm;

    Data^.LastRubSumm := -1;
        
    for i := 1 to 10 do
    begin
        if (Data^.Curr[i] = CURR) or (Data^.Curr[i] = '') then
        begin
            Data^.Curr[i] := CURR;
            Ptr^[i].Summ := Ptr^[i].Summ + PAY;
            if CURR = 'BRB' then
            begin
                Data^.LastRubSumm := PAY;
                Ptr^[i].RubSumm := Ptr^[i].RubSumm + PAY
            end
            else
            begin
                Data^.LastRubSumm := Round(PAY * GetRateOfCurrency(CURR, PAYDATE) / 5) * 5;
                Ptr^[i].RubSumm := Ptr^[i].RubSumm + Data^.LastRubSumm; //в бд курс валюты
            end;
            Data^.Curr[i + 1] := '';
            break;
        end;
    end;
end;

procedure TMainForm.AgentReportClick(Sender: TObject);
var
    DtFrom, DtTo : TDAteTime;
    sqlText, AG_FILTER : string;
    DataList : TList;
    Data : PAgReportStruct;
    i, j, Line, StartLine, SumLine : integer;
    ASH : VARIANT;
    AllNal, AllBezNal : double;
    CurrSheet : integer;
    AllRecords, N : Integer;
    ArrayData : VARIANT;
begin
    AddIndex('PrevSeria;PrevNumber');

    AllNal := 0;
    AllBezNal := 0;
    CurrSheet := 0;

    //RepParamsAgnt.filtParam := 'MANDPCNT > 0';
    if RepParamsAgnt.ShowModal <> mrOk then exit;

    ArrayData := VarArrayCreate([1, 1, 1, 14], varVariant);

    if RepParamsAgnt.agResultCodeList.Count = 0 then
    begin
        AG_FILTER := ' IS NOT NULL';
    end
    else
    begin
        AG_FILTER := ' IN ' + RepParamsAgnt.agSQL;
    end;
    DtFrom := Trunc(min(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date));
    DtTo := Trunc(max(RepParamsAgnt.DtFrom.Date, RepParamsAgnt.DtTo.Date));
    sqlText := 'SELECT 1 AS N, SERIA, NUMBER, NAME, REGDATE, PAY1 AS PAY, PAYDATE,  RATE,     PAY1CURR AS CURR, PAYTYPE,  RET1 AS DT, AGPERCENT AS PCNT, AGENTCODE, DRAFTNUMBER, INSURERTYPE, '''' AS ALLFEETEXT FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND AGENTCODE ' + AG_FILTER + ' AND PAYDATE IS NOT NULL  AND RET1 BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' AND (PREVSERIA IS NULL OR PREVNUMBER<=0)' +
               ' UNION ALL ' +
               'SELECT 2,      SERIA, NUMBER, NAME, REGDATE, PAY2SUMMA,   PAY2DATE, PAY2RATE, PAY2CURR,         PAY2TYPE, RET2,       CARCODE,           FEE2TEXT,  PAY2DRAFTNUMBER, INSURERTYPE, ALLFEETEXT  FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND FEE2TEXT  ' + AG_FILTER + ' AND PAY2DATE IS NOT NULL AND RET2 BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''' +
               ' ORDER BY AGENTCODE, SERIA, NUMBER';

    DataList := TList.Create;
    Data := nil;
    with WorkSQL, WorkSQL.SQL do
    begin
        try
            if not InitExcel(MSEXCEL) then exit;
            MSEXCEL.WorkBooks.Add;

            Screen.Cursor := crHourGlass;
            Close;
            Clear;
            Add(sqlText);
            Open;

            AllRecords := RecordCount;
            N := 0;
            while not Eof do
            begin
                Inc(N);
                StatusBar.Panels[0].Text := IntToStr((N * 100) DIV AllRecords) + '%';
                StatusBar.Update;
                //Для 2й оплаты проверим дубликат
                if (FieldByName('N').AsInteger = 2) then
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                    HandBookSQL.Open;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                    begin
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                    end;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                    begin
                        HandBookSQL.Close;
                        Next;
                        continue;
                    end;
                    HandBookSQL.Close;
                end;

                //Show Bad Polises
                if RepParamsAgnt.IsFullRep.Checked and (Data <> nil) and (Data^.AgCode <> FieldByName('AGENTCODE').AsString) then
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT * FROM ' + GetTableName() + ' WHERE STATE=1 AND REGDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' AND AGENTCODE =''' + Data^.AgCode + '''');
                    HandBookSQL.Open;

                    SumLine := Line;
                    Inc(Line);

                    if not HandBookSQL.Eof then
                    begin
                        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := 'ИСПОРЧЕННЫЕ';
                    end;

                    while not HandBookSQL.Eof do
                    begin
                        Inc(Line);
                        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := HandBookSQL.FieldByName('Seria').AsString;
                        ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)].Value := HandBookSQL.FieldByName('Number').AsString;
                        ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := HandBookSQL.FieldByName('RegDate').AsString;
                        HandBookSQL.Next;
                    end;
                end;

                if (Data = nil) or (Data^.AgCode <> FieldByName('AGENTCODE').AsString) then
                begin
                    New(Data);
                    DataList.Add(Data);
                    Data^.AgCode := FieldByName('AGENTCODE').AsString;
                    Data^.Curr[1] := '';
                    Data^.Curr[1] := '';
                    for i := 1 to 10 do
                    begin
                        Data^.NalSumm[i].Summ := 0;
                        Data^.BezNalSumm[i].Summ := 0;
                        Data^.NalSumm[i].RubSumm := 0;
                        Data^.BezNalSumm[i].RubSumm := 0;
                    end;

                    if RepParamsAgnt.IsFullRep.Checked then
                    begin
                        if CurrSheet > 0 then
                        begin
                            ASH.Range['J' + IntToStr(SumLine) + ':J' + IntToStr(SumLine)].Value := '=SUM(J8:J' + IntToStr(SumLine - 1) + ')';
                            ASH.Range['J6:J6'].Value := '=SUM(J8:J' + IntToStr(SumLine - 1) + ')';
                            ASH.Range['J6:J6'].Font.Bold := True;
                            ASH.Columns['A:M'].EntireColumn.AutoFit;
                        end;

                        //Init Sheet
                        Inc(CurrSheet);
                        if CurrSheet > MSEXCEL.Sheets.Count then
                        begin
                            MSEXCEL.Sheets[1].Select;
                            MSEXCEL.Sheets.Add;
                            ASH := MSEXCEL.Sheets[1];
                        end
                        else
                        begin
                            ASH := MSEXCEL.Sheets[CurrSheet];
                        end;
                        ASH.Select;
                        try
                            ASH.Name := RepParamsAgnt.GetAgentName(Data^.AgCode);
                        except
                        end;
                        Data^.PageName := ASH.Name;
                        if Pos(' ', Data^.PageName) <> 0 then Data^.PageName := '''' + Data^.PageName + '''';
                        ASH.Range['A1:M1'].MergeCells := True;
                        ASH.Range['A1:A1'].Value := 'Отчёт о поступивших страховых взносах за ' + DateToStr(DtFrom) + ' - ' + DateToStr(DtTo);
                        ASH.Range['A2:M2'].MergeCells := True;
                        ASH.Range['A2:A2'].Value := 'от агента ' + RepParamsAgnt.GetAgentName(Data^.AgCode);
                        ASH.Range['A4:M4'].MergeCells := True;
                        ASH.Range['A4:A4'].Value := 'Вид страхования  "30.Обязательное ГО авто владельцев"';
                        ASH.Range['A7:A7'].Value := 'Серия';
                        ASH.Range['B7:B7'].Value := 'Номер';
                        ASH.Range['C7:C7'].Value := 'Выдан';
                        ASH.Range['D7:D7'].Value := 'Страх-ль';
                        ASH.Range['E7:E7'].Value := 'Ф/Ю';
                        ASH.Range['F7:F7'].Value := 'Дата';
                        ASH.Range['G7:G7'].Value := 'Форма';
                        ASH.Range['H7:H7'].Value := 'В валюте';
                        ASH.Range['I7:I7'].Value := 'Валюта';
                        ASH.Range['J7:J7'].Value := 'Эквивалент';
                        ASH.Range['K7:K7'].Value := 'Проведено по бухг.';
                        ASH.Range['L7:L7'].Value := 'Квитанция';
                        ASH.Range['N7:N7'].Value := 'Номер платежа';
                        ASH.Range['A7:N7'].Font.Bold := True;
                        ASH.Range['A5:A5'].Value := 'ИТОГИ';
                        ASH.Hyperlinks.Add(ASH.Range['A5:A5'], '', 'ИТОГИ!A' + IntToStr(6 + DataList.Count * 4), 'Перейти в ИТОГИ');
                    end;
                    Line := 8;
                end;

                InitAgRep(Data,
                          FieldByName('PAYTYPE').AsInteger,
                          FieldByName('PAYDATE').AsDateTime,
                          FieldByName('PAY').AsFloat,
                          FieldByName('CURR').AsString,
                          FieldByName('RATE').AsFloat,
                          FieldByName('PCNT').AsFloat
                         );

                if RepParamsAgnt.IsFullRep.Checked then
                begin
                    ArrayData[1, 1] := FieldByName('SERIA').AsString;
                    ArrayData[1, 2] := FieldByName('NUMBER').AsInteger;
                    ArrayData[1, 3] := FieldByName('REGDATE').AsString;
                    ArrayData[1, 4] := FieldByName('NAME').AsString;
                    if FieldByName('INSURERTYPE').AsInteger = 1 then
                    begin
                        ArrayData[1, 5] := 'Ф';
                    end;
                    if FieldByName('INSURERTYPE').AsInteger = 0 then
                    begin
                        ArrayData[1, 5] := 'Ю';
                    end;
                    if FieldByName('INSURERTYPE').AsInteger = 2 then
                    begin
                        ArrayData[1, 5] := 'ИП';
                    end;
                    ArrayData[1, 6] := FieldByName('PAYDATE').AsString;
                    if FieldByName('PAYTYPE').AsInteger = 1 then
                    begin
                        ArrayData[1, 7] := 'нал';
                    end
                    else
                    begin
                        ArrayData[1, 7] := 'безнал';
                    end;
                    ArrayData[1, 8] := FieldByName('PAY').AsFloat;
                    ArrayData[1, 9] := FieldByName('CURR').AsString;
                    ArrayData[1, 10] := Data^.LastRubSumm;
                    ArrayData[1, 11] := FieldByName('DT').AsString;
                    ArrayData[1, 12] := '';
                    ArrayData[1, 13] := '';
                    if FieldByName('N').AsInteger = 2 then
                    if FieldByName('DRAFTNUMBER').AsInteger > 0 then
                    begin
                        if FieldByName('AllFeeText').AsString = '0' then ArrayData[1, 12] := 'КВ';
                        if FieldByName('AllFeeText').AsString = 'Y' then ArrayData[1, 12] := 'ГС';
                        ArrayData[1, 13] := FieldByName('DRAFTNUMBER').AsInteger;
                    end;
                    ArrayData[1, 14] := FieldByName('N').AsInteger;
                    ASH.Range['A' + IntToStr(Line) + ':N' + IntToStr(Line)] := ArrayData;
                    Inc(Line);
                end;
                Next;
            end;

            Close;
        except
            on E : Exception do
            begin
                MessageDlg(e.Message + #13'Создание отчета остановлено', mtInformation, [mbOk], 0);
            end;
        end;
    end;

    VarClear(ArrayData);

    if CurrSheet > 0 then
    begin
        ASH.Range['J' + IntToStr(Line) + ':J' + IntToStr(Line)].Value := '=SUM(J8:J' + IntToStr(Line - 1) + ')';
        ASH.Range['J6:J6'].Value := '=SUM(J8:J' + IntToStr(Line - 1) + ')';
        ASH.Range['J6:J6'].Font.Bold := True;
        ASH.Columns['A:M'].EntireColumn.AutoFit;
    end;

    Inc(CurrSheet);
    if CurrSheet > MSEXCEL.Sheets.Count then
    begin
        MSEXCEL.Sheets[1].Select;
        MSEXCEL.Sheets.Add;
        ASH := MSEXCEL.Sheets[1];
    end
    else
    begin
        ASH := MSEXCEL.Sheets[CurrSheet];
    end;
    ASH.Select;

    MSEXCEL.Visible := true;

    ASH.Name := 'ИТОГИ';
    ASH.Range['A1:G1'].MergeCells := True;
    ASH.Range['A1:A1'].Value := 'ИТОГИ по платежам за ' + DateToStr(DtFrom) + ' - ' + DateToStr(DtTo);
    ASH.Range['A5:A5'].Value := 'Валюта';
    ASH.Range['B5:B5'].Value := 'В валюте';
    ASH.Range['C5:C5'].Value := 'Эквивалент';
    ASH.Range['D5:D5'].Value := 'В валюте';
    ASH.Range['E5:E5'].Value := 'Эквивалент';
    ASH.Range['F5:F5'].Value := 'В валюте';
    ASH.Range['G5:G5'].Value := 'Эквивалент';
    ASH.Range['B4:B4'].Value := 'Всего';
    ASH.Range['B4:C4'].MergeCells := True;
    ASH.Range['D4:D4'].Value := 'Безналичными';
    ASH.Range['D4:E4'].MergeCells := True;
    ASH.Range['F4:F4'].Value := 'Наличными';
    ASH.Range['F4:G4'].MergeCells := True;
    ASH.Range['H6:H6'].Select;
    MSEXCEL.ActiveWindow.FreezePanes := True;
    Line := 6;
    for i := 0 to DataList.Count - 1 do
    begin
        ASH.Range['A' + IntToStr(Line) + ':G' + IntToStr(Line)].MergeCells := True;
        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := 'Агент ' + RepParamsAgnt.GetAgentName(Data^.AgCode);
        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Font.Bold := True;
        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Select;
        if RepParamsAgnt.IsFullRep.Checked then
        begin
            ASH.Hyperlinks.Add(ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)], '', Data^.PageName + '!A1', ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value);
        end;
        //ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= "АЛЬВЕНА!A1", TextToDisplay:="Агент АЛЬВЕНА"

        Inc(Line);
        Data := DataList[i];
        StartLine := Line;
        for j := 1 to 10 do
        begin
            if Data^.Curr[j] = '' then break;
            ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := Data^.Curr[j];
            ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)].Value := Data^.NalSumm[j].Summ + Data^.BezNalSumm[j].Summ;
            ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := Data^.NalSumm[j].RubSumm + Data^.BezNalSumm[j].RubSumm;
            ASH.Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)].Value := Data^.BezNalSumm[j].Summ;
            ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)].Value := Data^.BezNalSumm[j].RubSumm;
            ASH.Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)].Value := Data^.NalSumm[j].Summ;
            ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)].Value := Data^.NalSumm[j].RubSumm;
            AllNal := AllNal + Data^.NalSumm[j].RubSumm;
            AllBezNal := AllBezNal + Data^.BezNalSumm[j].RubSumm;

            Inc(Line);
        end;
        ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := 'ИТОГО';
        ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := '=SUM(C' + IntToStr(StartLine) + ':C' + IntToStr(Line-1) + ')';
        ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)].Value := '=SUM(E' + IntToStr(StartLine) + ':E' + IntToStr(Line-1) + ')';
        ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)].Value := '=SUM(G' + IntToStr(StartLine) + ':G' + IntToStr(Line-1) + ')';
        Inc(Line, 2);
    end;

    ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := 'ВСЕГО';
    ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := AllNal + AllBezNal;
    ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)].Value := AllBezNal;
    ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)].Value := AllNal;
    ASH.Columns['A:G'].EntireColumn.AutoFit;

    for i := 0 to DataList.Count - 1 do
    begin
        Dispose(DataList[i]);
    end;
    DataList.Free;
    Screen.Cursor := crDefault;
    //BuildReportInputSum(DtFrom, DtTo, sqlText, RepParamsAgnt.listAgnt.Text, WorkSQL);
end;

procedure TMainForm.MainQueryBASECARCODEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
    if Sender.AsInteger = 101 then Text := 'ДА';
    if Sender.AsInteger = 100 then Text := 'НЕТ';
end;

procedure TMainForm.mnuUnloadAllPolisesClick(Sender: TObject);
var
     INI : TIniFile;
begin
     INI := TIniFile.Create(BLANK_INI);
      mnuUnloadAllPolises.Checked := not mnuUnloadAllPolises.Checked;
      INI.WriteBool('MANDATORY', 'UnloadAllPolises', mnuUnloadAllPolises.Checked);
      INI.Free
end;

procedure TMainForm.MainGridGetCellParams(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
    if (MainQueryAGENTCODE.AsString = '') then begin
        Background := clPurple;
        AFont.Color := clWhite;
    end;
    if MainQueryBASECARCODE.AsInteger = 100 then Background := clRed;
end;

procedure TMainForm.To1CClick(Sender: TObject);
label
    NextLabel;
const
    InsTypeCode : string = '30';
var
    PolCount, PayCount : integer;
    F : TextFile;
begin
    if not SaveDialog1C.Execute then exit;
    //Create1CTable(Table1C, SaveDialog1C.FileName);
    PolCount := 0;
    PayCount := 0;

    try
        //Table1C.Open;
        DeleteFile(SaveDialog1C.FileName);
        AssignFile(F, SaveDialog1C.FileName);
        Rewrite(F);
//        Truncate(F);

        MainQuery.First;

        Screen.Cursor := crHourGlass;
        MainGrid.DisableScroll;

        while not MainQuery.Eof do begin
            if MainQueryPrevSeria.AsString <> '' then goto NextLabel;
            if (MainQueryState.AsString = '1') then goto NextLabel;

            Inc(PolCount);
            Inc(PayCount);
            {Table1C.Append;
            Table1C.Fields[0].AsString := InsTypeCode;
            Table1C.Fields[1].AsDateTime := MainQueryRET1.AsDateTime;
            Table1C.Fields[2].AsString := MainQueryNumber.AsString;
            Table1C.Fields[3].AsString := MainQueryNAME.AsString;
            Table1C.Fields[4].AsFloat := MainQueryPAY1.AsFloat;
            Table1C.Fields[5].AsString := MainQueryPAY1CURR.AsString;
            Table1C.Fields[6].AsDateTime := MainQueryFROMDATE.AsDateTime;
            Table1C.Fields[7].AsDateTime := MainQueryTODATE.AsDateTime;
            Table1C.Fields[8].AsString := MainQueryAGNAME.AsString;
            Table1C.Post;
            }
            WriteLn(F, InsTypeCode + ';' +
                    MainQueryRET1.AsString + ';' +
                    MainQueryNumber.AsString + ';' +
                    MainQueryNAME.AsString + ';' +
                    MainQueryPAY1.AsString + ';' +
                    MainQueryPAY1CURR.AsString + ';' +
                    MainQueryFROMDATE.AsString + ';' +
                    MainQueryTODATE.AsString + ';' +
                    MainQueryAGNAME.AsString);

            if MainQueryPAY2CURR.AsString <> '' then begin
              //Table1C.Append;
              Inc(PayCount);
              {
              Table1C.Fields[0].AsString := InsTypeCode;
              Table1C.Fields[1].AsDateTime := MainQueryRET1.AsDateTime;
              Table1C.Fields[2].AsString := MainQueryNumber.AsString;
              Table1C.Fields[3].AsString := MainQueryNAME.AsString;
              Table1C.Fields[4].AsFloat := MainQueryPAY2SUMMA.AsFloat;
              Table1C.Fields[5].AsString := MainQueryPAY2CURR.AsString;
              Table1C.Fields[6].AsDateTime := MainQueryFROMDATE.AsDateTime;
              Table1C.Fields[7].AsDateTime := MainQueryTODATE.AsDateTime;
              Table1C.Fields[8].AsString := MainQueryAGNAME.AsString;
              Table1C.Post;
              }
              Writeln(F, InsTypeCode + ';' +
                         MainQueryRET1.AsString + ';' +
                         MainQueryNumber.AsString + ';' +
                         MainQueryNAME.AsString + ';' +
                         MainQueryPAY2SUMMA.AsString + ';' +
                         MainQueryPAY2CURR.AsString + ';' +
                         MainQueryFROMDATE.AsString + ';' +
                         MainQueryTODATE.AsString + ';' +
                         MainQueryAGNAME.AsString);
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

function DecodeFU16(s : string; pcnt : double) : string;
begin
    DecodeFU16 := '?';
    if pcnt < 0.000001 then DecodeFU16 := '3'
    else
    if s = 'F' then DecodeFU16 := '1'
    else
    if s = 'U' then DecodeFU16 := '6';
end;

function DecodeOwnTypeForm(s, version : string) : string;
begin
    DecodeOwnTypeForm := '?';
    if version = '1' then s := 'F';
    if s = 'F' then DecodeOwnTypeForm := '1';
    if s = 'U' then DecodeOwnTypeForm := ''; //'3'
    if s = 'I' then DecodeOwnTypeForm := '3'; 
    if s = '1' then DecodeOwnTypeForm := '1';
    if s = '0' then DecodeOwnTypeForm := ''; //'3'
    if s = '2' then DecodeOwnTypeForm := '3';
end;

function DecodeFU12(s, version : string) : string;
begin
    DecodeFU12 := '?';
    if version = '1' then s := 'F';
    if s = 'F' then DecodeFU12 := '1';
    if s = 'U' then DecodeFU12 := '2';
    if s = 'I' then DecodeFU12 := '3';
    if s = '1' then DecodeFU12 := '1';
    if s = '0' then DecodeFU12 := '2';
    if s = '2' then DecodeFU12 := '3';
end;

function DecodeBlank(sect, s : string) : string;
var
    INI : TIniFile;
    iniName, kod : string;
begin
    DecodeBlank := '?';

    if s = 'АБ' then kod := '7';
    if s = 'АГ' then kod := '28';
    if s = 'TA' then kod := '19';
    if s = 'АД' then kod := '43';
    if s = 'BY02' then kod := '29';

    INI := TIniFile.Create(BLANK_INI);

    if bKomplexMode then
    begin
        iniName := INI.ReadString(KOMPLEX_SECT, 'INI', '');
        INI.Free;
        INI := TIniFile.Create(iniName);
    end;

    kod := INI.ReadString(sect, 'Seria_' + s + '_KOD', kod);
    INI.Free;
    if kod = '?' then
        raise Exception.Create('Не указан код серии бланка ' + s + ' в ГЛОБО ');
    DecodeBlank := kod;
end;

function DecodePeriod(s : string; datefrom, dateto : TDateTIme; daysbased : boolean) : string;
var
    s1 : string;
    days : double;
    D, m, y : word;
begin
    D := 0;
    m := 0;
    y := 0;
    days := dateto - datefrom + 1;
    s := Trim(s);
    DecodePeriod := '?';
    if Pos(' ', s) <> 0 then begin
        s1 := AnsiLowerCase(Trim(Copy(s, 1, Pos(' ', s) - 1)));
        DecodePeriod := s1;
        s := AnsiLowerCase(s);
        if Pos('год', s) <> 0 then DecodePeriod := s1 + 'Г'
        else
        if Pos('мес', s) <> 0 then DecodePeriod := s1 + 'М'
        else
        if(days >= 31) then
            raise Exception.Create('Некорректный период ' + s);
    end
    else begin
        if daysbased then
        begin
            DateDiff(dateto+1, datefrom, D, m, y);
            if (d > 0) and (m > 0) or (d > 0) and (y > 0) or (y > 0) and (m > 0) then
                raise Exception.Create('Некорректный период ' + s);
            if Y > 0 then DecodePeriod := IntToStr(y) + 'Г'
            else if M > 0 then DecodePeriod := IntToStr(m) + 'М'
            else DecodePeriod := IntToStr(d);
        end;
        try
            if(days >= 31) then
                raise Exception.Create('Некорректный период ' + s);
            DecodePeriod := IntToStr(StrToInt(s));
        except
        end;
    end;
end;

function DecodeTimeX(s : string) : string;
var
    s1, s2 : string;
begin
    s1 := ExtractWord(1, s, [',','.',':','-']);
    s2 := ExtractWord(2, s, [',','.',':','-']);
    while (Length(s1) < 2) do s1 := '0' + s1;
    while (Length(s2) < 2) do s2 := '0' + s2;

    if (StrToInt(s2) > 59) or (StrToInt(s1) > 23) then
        raise Exception.Create('Неправильный формат времени ' + s);

    DecodeTimeX := s1 + s2;
end;

function DecodeNB(s, paydoc : string) : string;
var
    rs : string;
begin
    rs := '';
    if (s = 'N') or (s = 'Y') then rs := '1';
    if (s = 'B') or (s = 'N') then rs := '2';
    if s = '1' then rs := '1';
    if s = '0' then rs := '2';
    if s = '2' then rs := '2'; //raise Exception.Create('Неизвестный тип платежа КАРТА');

    if (rs = '') and (Length(paydoc) < 2) then rs := '1';
    if (rs = '') and (Length(paydoc) >= 2) then rs := '2';

    DecodeNB := rs;
end;

function KillZero(s : string) : string;
begin
    if s = '0' then s := '';
    KillZero := s;
end;

function DecodeVariant(tariftable : string) : string;
var
    INI : TIniFile;
    iniName : string;
begin
    if (lastTableVariant = (tarifTable + BoolToStr(bKomplexMode))) then
    begin
        DecodeVariant := lastTableVariantCode;
        exit;
    end;

    lastTableVariant := (tarifTable + BoolToStr(bKomplexMode));
    lastTableVariantCode := '?';

    //if not bKomplexMode and (s = '2') then lastTableVariantCode := '2'
    //else if not bKomplexMode and (s = '3') then lastTableVariantCode := '1'
    //else if not bKomplexMode and (s = '4') then lastTableVariantCode := '3'
    //else
    begin
        iniName := BLANK_INI;
        INI := TIniFile.Create(iniName);
        if bKomplexMode then
        begin
                iniName := INI.ReadString(KOMPLEX_SECT, 'INI', '');
                INI.Free;
                INI := TIniFile.Create(iniName);
        end;
        lastTableVariantCode := INI.ReadString('MANDATORY', 'TarifTableVariant' + tarifTable, '');
        if lastTableVariantCode = '' then
        begin
            INI.Free;
            raise Exception.Create('TarifTableVariant' + tarifTable + ' не определён в ' + iniName);
        end;
        DecodeVariant := lastTableVariantCode;
        INI.Free;
    end;

    DecodeVariant := lastTableVariantCode;
end;

function DecodeGlDate(dt : TDate) : string;
var
    d, m, y : Word;
    sd, sm : string;
begin
    DecodeDate(dt, y, m, d);
    sd := IntToStr(d);
    if Length(sd) < 2 then sd := '0' + sd;
    sm := IntToStr(m);
    if Length(sm) < 2 then sm := '0' + sm;
    DecodeGlDate := IntToStr(y) + sm + sd;
end;

function DecodeGlDate2(dt, dt1, dt2, dt3 : TField) : string;
begin
    if not dt.IsNull then DecodeGlDate2 := DecodeGlDate(dt.AsDateTime)
    else if not dt1.IsNull then DecodeGlDate2 := DecodeGlDate(dt1.AsDateTime)
    else if not dt2.IsNull then DecodeGlDate2 := DecodeGlDate(dt2.AsDateTime)
    else if not dt3.IsNull then DecodeGlDate2 := DecodeGlDate(dt3.AsDateTime)
    else DecodeGlDate2 := '';                                      
end;

function DecodeR(s : string) : string;
var
    INI : TIniFile;
    ss : string;
begin
    if lastResidentTable <> s then
    begin
        INI := TIniFile.Create(BLANK_INI);
        if bKomplexMode then
        begin
                ss := INI.ReadString(KOMPLEX_SECT, 'INI', '');
                INI.Free;
        INI := TIniFile.Create(ss);
        end;
        lastIsResident := INI.ReadString('MANDATORY', 'TarifTable_Is_Resident' + s, '') = 'YES';
        INI.Free;
    end;
    lastResidentTable := s;
    DecodeR := '2';
    if lastIsResident then DecodeR := '1';
end;

function DecodeVladen(s : string; version : string) : string;
begin
    DecodeVladen := '?';
    if version = '1' then s := '1';
    if s = '1' then DecodeVladen := 'S';
    if s = '2' then DecodeVladen := 'H';
    if s = '3' then DecodeVladen := 'O';
    if s = '4' then DecodeVladen := 'A';
    if s = '5' then DecodeVladen := 'D';
    if s = '6' then DecodeVladen := 'I';
end;

function DecodePayGraphic(s : string) : string;
begin
    DecodePayGraphic := '?';
    if s = 'N' then DecodePayGraphic := '1';
    if s = '0' then DecodePayGraphic := '1';
    if s = 'Y' then DecodePayGraphic := '2';
    if s = '1' then DecodePayGraphic := '2';
end;

function DecodePayType(s : string) : string;
begin
    DecodePayType := '1';
    if s = 'N' then DecodePayType := '1';
    if s = 'B' then DecodePayType := '2';
    if s = '1' then DecodePayType := '1';
    if s = '0' then DecodePayType := '2';
    if (s = 'F') OR (s = 'I') then DecodePayType := '1';
    if s = 'U' then DecodePayType := '2';
end;

function GetAvK(val : double; version : string) : string;
var
    i : integer;
    s : string;
    INI : TIniFile;
begin
    if version = '1' then
    begin
        GetAvK := '1';
        exit;
    end;
    if AvKCount = 0 then begin
        INI := TIniFile.Create(BLANK_INI);
        while true do begin
            s := INI.ReadString('MANDATORY', 'ClassAV' + IntToStr(AvKCount), '');
            if s = '' then break;
            i := Pos(',', s);
            if i = 0 then break;
            AvK[AvKCount].Val := round(StrToFloat(Copy(s, 1, i - 1)) * 10);
            AvK[AvKCount].K := Trim(Copy(s, i + 1, 100));
            i := Pos(' ', AvK[AvKCount].K);
            if i > 0 then AvK[AvKCount].K := Trim(Copy(AvK[AvKCount].K, i + 1, 100));

            AvKCount := AvKCount + 1;
        end;
        INI.Free;
    end;

    if AvKCount < 3 then
    begin
        raise Exception.Create('Не найдены ClassAV в blank.ini');
    end;

    val := round(val * 10 + 0.001);
    GetAvK := '?';
    for i := 0 to AvKCount - 1 do
        if abs(AvK[i].Val - val) < 1e-5 then
            GetAvK := AvK[i].K;
end;

function GetK(val : double; version : string) : string;
begin
    if version = '1' then GetK := '1'
    else
    begin
        val := round(val * 10 + 0.001);        
        //GetK := IntToStr(round(val * 10 + 0.001) MOD 10);
        if val = 13 then GetK := '3'
        else if val = 15 then GetK := '3'
        else if val = 12 then GetK := '2'
        else if val = 10 then GetK := '0'
        else if val = 8 then GetK := '8'
        else GetK := '?';
    end;
end;

function DecodeCurrency(s : string) : string;
begin
    DecodeCurrency := '?';
    if s = 'USD' then DecodeCurrency := '840';
    if s = 'EUR' then DecodeCurrency := '978';
    if s = 'RUR' then DecodeCurrency := '643';
    if s = 'BRB' then DecodeCurrency := '974';
    if s = 'UAH' then DecodeCurrency := '980';
end;

function DecodeCharact(s, v, t, c, p : string) : string;
begin
    DecodeCharact := '|';
    //if (v <> '0') and (v <> '') then DecodeCharact := '1|' + v;
    //if (t <> '0') and (t <> '') then DecodeCharact := '3|' + t;
    //if (c <> '0') and (c <> '') then DecodeCharact := '4|' + c;
    //if (p <> '0') and (p <> '') then DecodeCharact := '2|' + p;

//    try
//        if v <> '' then
 //       begin
  //          v := FloatToStr(StrToFloat(v) / VolumeK);
   //     end
    //except
    //end;

    if (s = 'N1') OR (s = 'N2') OR (s = 'N3') OR (s = 'N4') OR (s = 'N5') OR
       (s = 'A1') OR (s = 'A2') OR (s = 'A3') OR (s = 'A4') OR (s = 'A5') OR
       (s = 'F1') OR (s = 'F2') OR (s = 'F3') then
        if (v <> '0') and (v <> '') then DecodeCharact := '1|' + v;

    if (s = 'C0') OR (s = 'C1') OR (s = 'C2') OR (s = 'C3') OR (s = 'C4') OR
       (s = 'C5') OR (s = 'E1') OR (s = 'E2') OR (s = 'E3') then
        if (t <> '0') and (t <> '') then DecodeCharact := '3|' + t;

    if (s = 'L1') OR (s = 'L2') OR (s = 'L3') then
        if (c <> '0') and (c <> '') then DecodeCharact := '4|' + c;

    if (s = 'V1') OR (s = 'V2') OR (s = 'V3') then
        if (p <> '0') and (p <> '') then DecodeCharact := '2|' + p;
end;

function DecodeNone(s, s1 : string) : string;
begin
    if Trim(s) = '' then DecodeNone := s1
    else DecodeNone := s;
end;

function DecodeAvState(s : string; DtCl, DtFree : TField) : string;
begin
    DecodeAvState := '';
    if (s = '1') or ((DtFree <> nil) and (not DtFree.IsNull)) then DecodeAvState := '3'; //ОТКАЗ
    if s = '0' then DecodeAvState := '2'; //ОПЛАЧЕНО
end;

function DecodeOneDate(F1, F2, F3 : TField) : string;
begin
    DecodeOneDate := '';
    if not f1.IsNull then DecodeOneDate := DecodeGlDate(f1.AsDateTime);
    if not f2.IsNull then DecodeOneDate := DecodeGlDate(f2.AsDateTime);
    if not f3.IsNull then DecodeOneDate := DecodeGlDate(f3.AsDateTime);
end;

{
function DecodeRN(s : string) : string;
begin
    DecodeRN := '?';
    if s = 'Y' then DecodeRN := '1';
    if s = 'N' then DecodeRN := '2';
end;
 }
function DecodeInsSum(d : TDate) : string;
begin
    if d < EncodeDate(2005, 4, 1) then DecodeInsSum := '10000'
    else DecodeInsSum := '20000';
end;

function DecodeSN(s : string) : string;
var
    p : integer;
begin
    s := Trim(s);
    p := Pos('/', s);
    if p <> 0 then
        s[p] := '|'
    else
        s := '|' + s;

    //if Length(Trim(s)) <= 2 then DecodeSN := '|'
    //else
    DecodeSN := s;
end;

function DecodeBodyShas(s : string) : string;
begin
    DecodeBodyShas := s + '|';
end;

function DecodeCharact2(charact, s : string) : string;
var
    res : string;
begin
    if (s = 'N1') OR (s = 'N2') OR (s = 'N3') OR (s = 'N4') OR (s = 'N5') OR
       (s = 'A1') OR (s = 'A2') OR (s = 'A3') OR (s = 'A4') OR (s = 'A5') OR
       (s = 'F1') OR (s = 'F2') OR (s = 'F3') then
    begin
        //try
        //    if charact <> '' then
        //    begin
        //        charact := FloatToStr(StrToFloat(charact) / VolumeK);
        //    end
        //except
        //end;
        res := '1'
    end
    else
    if (s = 'C0') OR (s = 'C1') OR (s = 'C2') OR (s = 'C3') OR (s = 'C4') OR
       (s = 'C5') OR (s = 'E1') OR (s = 'E2') OR (s = 'E3') then
       res := '3'
    else
    if (s = 'V1') OR (s = 'V2') OR (s = 'V3') then
       res := '2'
    else
    if (s = 'L1') OR (s = 'L2') OR (s = 'L3') then
       res := '4'
    else
    if (s = 'A6') OR (s = 'B1') OR (s = 'B2') OR (s = 'D') OR (s = 'M') OR
       (s = 'L4') OR (s = 'W') then
       res := ''
    else
       res := '';

    if res = '' then res := res + '|'
    else res := res + '|' + charact;

    DecodeCharact2 := res
end;

procedure CheckAutoNmb(SN, AUTONUMBER : String; var Warn : TStringList);
begin
  if Copy(AUTONUMBER, 1, 2) = 'AO' then
      Warn.Add(SN + ' гос. номер. Возможно A0 вместо AO');
  if Copy(AUTONUMBER, 1, 2) = 'CO' then
      Warn.Add(SN + ' гос. номер. Возможно C0 вместо CO');
  if Copy(AUTONUMBER, 1, 2) = 'EO' then
      Warn.Add(SN + ' гос. номер. Возможно E0 вместо EO');
end;

procedure TMainForm.N30Click(Sender: TObject);
var
    DtFrom, DtTo : TDateTime;
    i, j, Line : integer;
    sqlText, filterField1, filterField2 : string;
    Data : AgReportStruct;
    ASH : Variant;
    AllRecords, N : Integer;
    ArrayData : VARIANT;
    delta : double;
begin
    delta := 732312 - EncodeDate(2006, 1, 1);
    
    if PayReportFilter.ShowModal <> mrOk then
        exit;

    AddIndex('PrevSeria;PrevNumber');

    ArrayData := VarArrayCreate([1, 1, 1, 16], varVariant);
    if PayReportFilter.listConds.ItemIndex = 0 then //ДАТЕ ОТЧЕТА
    begin
        filterField1 := 'RET1';
        filterField2 := 'RET2';
    end
    else
    begin
        filterField1 := 'PAYDATE';
        filterField2 := 'PAY2DATE';
    end;

    DtFrom := min(PayReportFilter.DtFrom.Date, PayReportFilter.DtTo.Date);
    DtTo := max(PayReportFilter.DtFrom.Date, PayReportFilter.DtTo.Date);
    sqlText := 'SELECT 1 AS N, SERIA, FROMDATE, TODATE, NUMBER, NAME, InsurerType, REGDATE, PAY1 AS PAY, PAYDATE,  RATE,     PAY1CURR AS CURR, PAYTYPE,  RET1 AS DT, AGPERCENT AS PCNT, AGENTCODE, DRAFTNUMBER FROM ' + GetTableName() + ' WHERE BASECARCODE=101 ' + 'AND PAYDATE IS NOT NULL  AND ' + filterField1 + ' BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' AND PREVSERIA IS NULL ' +
               ' UNION ALL ' +
               'SELECT 2,      SERIA, FROMDATE, TODATE, NUMBER, NAME, InsurerType, REGDATE, PAY2SUMMA,   PAY2DATE, PAY2RATE, PAY2CURR,         PAY2TYPE, RET2,       CARCODE,           FEE2TEXT,  PAY2DRAFTNUMBER  FROM ' + '' + GetTableName() + ' WHERE BASECARCODE=101 AND PAY2DATE IS NOT NULL AND ' + filterField2 + ' BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''' +
               ' ORDER BY AGENTCODE, SERIA, NUMBER';

    Clipboard.AsText := sqlText;
    with WorkSQL, WorkSQL.SQL do
    begin
        try
            if not InitExcel(MSEXCEL) then exit;
            MSEXCEL.WorkBooks.Add;
            while MSEXCEL.Sheets.Count < 3 do
            begin
                MSEXCEL.Sheets.Add;
            end;
            ASH := MSEXCEL.Sheets[1];
            ASH.Name := 'Полисы';

            Screen.Cursor := crHourGlass;
            Close;
            Clear;
            Add(sqlText);
            Open;
            Line := 5;
            ASH.Range['A1:N1'].MergeCells := True;
            ASH.Range['A3:N3'].MergeCells := True;
            ASH.Range['A1:A1'].Value := 'Справка о платежах за период ' + DateToStr(DtFrom) + ' - ' + DateToStr(DtTo);
            ASH.Range['A3:A3'].Value := 'Вид страхования  "30.Обязательное ГО авто владельцев"';

            ASH.Range['A4:A4'].Value := 'Полис';
            ASH.Range['B4:B4'].Value := 'Выдан';
            ASH.Range['C4:C4'].Value := 'Начало действия полиса';
            ASH.Range['D4:D4'].Value := 'Окончание действия полиса';
            ASH.Range['E4:E4'].Value := 'Форма оплаты';
            ASH.Range['F4:F4'].Value := 'Дата платежа';
            ASH.Range['G4:G4'].Value := 'Валюта';
            ASH.Range['H4:H4'].Value := 'Сумма';
            ASH.Range['I4:I4'].Value := 'Эквивалент';
            ASH.Range['J4:J4'].Value := '№ документа';
            ASH.Range['K4:K4'].Value := 'Квитанция';
            ASH.Range['L4:L4'].Value := 'Дата отчета';
            ASH.Range['M4:M4'].Value := 'Плательщик';
            ASH.Range['N4:N4'].Value := 'Ф/Ю';
            ASH.Range['O4:O4'].Value := 'Получатель (АГЕНТ)';
            ASH.Range['P4:P4'].Value := 'Номер платежа';
            ASH.Range['A4:P4'].Font.Bold := True;

            AllRecords := RecordCount;
            N := 0;
            while not Eof do
            begin
                Inc(N);
                StatusBar.Panels[0].Text := IntToStr((N * 100) DIV AllRecords) + '%';
                StatusBar.Update;
                //Для 2й оплаты проверим дубликат
                if (FieldByName('N').AsInteger = 2) then
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                    HandBookSQL.Open;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                    begin
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                    end;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                    begin
                        HandBookSQL.Close;
                        Next;
                        continue;
                    end;
                    HandBookSQL.Close;
                end;

                Data.Curr[1] := '';
                Data.Curr[1] := '';
                Data.NalSumm[1].Summ := 0;
                Data.BezNalSumm[1].Summ := 0;
                Data.NalSumm[1].RubSumm := 0;
                Data.BezNalSumm[1].RubSumm := 0;

                InitAgRep(@Data,
                          FieldByName('PAYTYPE').AsInteger,
                          FieldByName('PAYDATE').AsDateTime,
                          FieldByName('PAY').AsFloat,
                          FieldByName('CURR').AsString,
                          FieldByName('RATE').AsFloat,
                          FieldByName('PCNT').AsFloat
                         );

                ArrayData[1, 1] := FieldByName('SERIA').AsString+ '/' + FieldByName('NUMBER').AsString;
                ArrayData[1, 2] := FieldByName('REGDATE').AsString;
                ArrayData[1, 3] := FieldByName('FROMDATE').AsString;
                ArrayData[1, 4] := FieldByName('TODATE').AsString;
                if FieldByName('PAYTYPE').AsInteger = 1 then
                begin
                    ArrayData[1, 5] := 'нал';
                end
                else
                begin
                    ArrayData[1, 5] := 'безнал';
                end;
                ArrayData[1, 6] := FieldByName('PAYDATE').AsString;
                ArrayData[1, 7] := FieldByName('CURR').AsString;
                ArrayData[1, 8] := FieldByName('PAY').AsFloat;
                ArrayData[1, 9] := Data.LastRubSumm;
                ArrayData[1, 10] := '';
                ArrayData[1, 11] := '';
                if FieldByName('DRAFTNUMBER').AsInteger > 0 then
                begin
                    ArrayData[1, 10] := FieldByName('DRAFTNUMBER').AsInteger;
                    ArrayData[1, 11] := FieldByName('DRAFTNUMBER').AsInteger;
                end;
                ArrayData[1, 12] := FieldByName('DT').AsString;
                ArrayData[1, 13] := FieldByName('NAME').AsString;
                ArrayData[1, 14] := '';
                if FieldByName('InsurerType').AsString = '1' then
                begin
                    ArrayData[1, 14] := 'Ф';
                end;
                if FieldByName('InsurerType').AsString = '0' then
                begin
                    ArrayData[1, 14] := 'Ю';
                end;
                if FieldByName('InsurerType').AsString = '2' then
                begin
                    ArrayData[1, 14] := 'ИП';
                end;
                ArrayData[1, 15] := DecodeAgentNamePx(FieldByName('AGENTCODE').AsString);
                ArrayData[1, 16] := FieldByName('N').AsInteger;
                ASH.Range['A' + IntToStr(Line) + ':P' + IntToStr(Line)] := ArrayData;
                Inc(Line);

                Next;
            end;

            if Line > 5 then
            begin
                ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)].Value := '=SUM(I5:I' + IntToStr(Line - 1) + ')';
            end;

            /////////////////////////////////////////////////////////////////////
            Close;
            Clear;
            ASH := MSEXCEL.Sheets[2];
            ASH.Name := 'Доплаты';
            if not bKomplexMode then
                  Add('SELECT M.*, T.*, A.NAME AS ANAME FROM TAXI T,MANDATOR M,AGENT A WHERE M.AGENTCODE=A.AGENT_CODE AND NUMBER=N AND SERIA=SER AND D BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''')
            else
                  Add('SELECT M.*, A.NAME AS ANAME, nal0 as nal, su0 as su, typ0 as type, s0 as summa, curr0 as curr, d0 as d, doc0 as paydoc FROM MANDATOR M,AGENT A WHERE M.AGENTCODE=A.AGENT_CODE AND D0 BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''')
            ;
            Open;
            Line := 1;
            ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)] := 'Полис';
            ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)] := 'Страхователь';
            ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)] := 'Сумма';
            ASH.Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)] := 'Валюта';
            ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)] := 'Дата платежа';
            ASH.Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)] := 'Документ';
            ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)] := 'Наличные';
            ASH.Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)] := '1 СУ';
            ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'Вид';
            ASH.Range['J' + IntToStr(Line) + ':J' + IntToStr(Line)] := 'Агент';
            Line := 2;
            while not Eof do
            begin
                ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)] := FieldByName('Ser').AsString + '/' + FieldByName('N').AsString;
                ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)] := FieldByName('NAME').AsString;
                ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)] := FieldByName('summa').AsString;
                ASH.Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)] := FieldByName('curr').AsString;
                ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)] := FieldByName('d').AsString;
                ASH.Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)] := FieldByName('paydoc').AsString;
                if FieldByName('Nal').AsString = 'Y' then
                begin
                    ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)] := 'нал';
                end;
                if FieldByName('Nal').AsString = 'N' then
                begin
                    ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)] := 'безнал';
                end;
                if FieldByName('SU').AsString = 'Y' then
                begin
                    ASH.Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)] := '1СУ';
                end;
                if FieldByName('TYPE').AsString = '' then
                begin
                    ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'За такси';
                end;
                if FieldByName('TYPE').AsString = '1' then
                begin
                    ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'Другое';
                end;
                if FieldByName('TYPE').AsString = '2' then
                begin
                    ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'За тип авто';
                end;
                if FieldByName('TYPE').AsString = '3' then
                begin
                    ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'Коэфф. аварийности';
                end;
                ASH.Range['J' + IntToStr(Line) + ':J' + IntToStr(Line)] := FieldByName('AName').AsString;
                Inc(Line);
                Next;
            end;

            Close;

            /////////////////////////////////////////////////////////////////////
            Close;
            Clear;
            ASH := MSEXCEL.Sheets[3];
            ASH.Name := 'Возвраты';
            Add('SELECT M.*, A.NAME AS ANAME FROM ' + GetTableName() + ' M, AGENT A WHERE M.AGENTCODE=A.AGENT_CODE AND (OwnerCode BETWEEN ' + FloatToStr(DtFrom+delta) + ' AND ' + FloatToStr(DtTo+delta) + ' OR (OwnerCode is null or OwnerCode < 1000) AND STOPDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''') AND RETBRB>0');
            Open;
            Line := 1;
            ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)] := 'Полис';
            ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)] := 'Страхователь';
            ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)] := 'С';
            ASH.Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)] := 'ПО';
            ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)] := 'Сумма';
            ASH.Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)] := 'Валюта';
            ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)] := 'Дата платежа';
            ASH.Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)] := 'Дата расторжения';
            ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'Вид';
            ASH.Range['J' + IntToStr(Line) + ':J' + IntToStr(Line)] := 'Агент';
            Line := 2;
            while not Eof do
            begin
                ASH.Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)] := FieldByName('Seria').AsString + '/' + FieldByName('Number').AsString;
                ASH.Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)] := FieldByName('NAME').AsString;
                ASH.Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)] := FieldByName('FROMDATE').AsString;
                ASH.Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)] := FieldByName('TODATE').AsString;
                ASH.Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)] := -FieldByName('RETBRB').AsFloat;
                ASH.Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)] := 'BRB';
                ASH.Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)] := DateToStr(FieldByName('OwnerCode').AsInteger - delta);
                ASH.Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)] := FieldByName('StopDate').AsString;
                ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)] := 'ВОЗВРАТ';
                ASH.Range['J' + IntToStr(Line) + ':J' + IntToStr(Line)] := FieldByName('AName').AsString;
                Inc(Line);
                Next;
            end;

            Close;
        except
            on E : Exception do
            begin
                MessageDlg(e.Message + #13'Создание отчета остановлено', mtInformation, [mbOk], 0);
            end;
        end;
    end;

    VarClear(ArrayData);
    ASH := MSEXCEL.Sheets[1];
    ASH.Columns['A:O'].EntireColumn.AutoFit;
    ASH := MSEXCEL.Sheets[2];
    ASH.Columns['A:J'].EntireColumn.AutoFit;
    ASH.Range['A1:J1'].Font.Bold := True;
    ASH := MSEXCEL.Sheets[3];
    ASH.Columns['A:J'].EntireColumn.AutoFit;
    ASH.Range['A1:J1'].Font.Bold := True;
    MSEXCEL.Visible := true;
    Screen.Cursor := crDefault;
end;

function GetPhone(s : string) : string;
var
    i : integer;
begin
    Result := '';
    for i := Length(s) downto 1 do
    begin
        if (s[i] < '0') or (s[i] > '9') then break;
    end;
    if Length(s) - i > 5 then
    begin
        GetPhone := Copy(s, i, 100);
    end
end;

procedure TMainForm.ExpectedPayMenuClick(Sender: TObject);
var
    ArrayData, ASH : VARIANT;
    N, AllRecords, Line : integer;
    sqlText : string;
    Data : AgReportStruct;
    DtFrom, DtTo : TDAteTime;
begin
    ExpectedPaysForm.Caption := 'Ожидаемые платежи';
    if ExpectedPaysForm.ShowModal <> mrOk then
        exit;

    AddIndex('PrevSeria;PrevNumber');

    ArrayData := VarArrayCreate([1, 1, 1, 18], varVariant);

    DtFrom := min(ExpectedPaysForm.DtFrom.Date, ExpectedPaysForm.DtTo.Date);
    DtTo := max(ExpectedPaysForm.DtFrom.Date, ExpectedPaysForm.DtTo.Date);
    sqlText := 'SELECT * FROM ' + GetTableName() + ' M LEFT OUTER JOIN AGENT A ON A.AGENT_CODE=M.AGENTCODE WHERE BASECARCODE=101 AND ISPAY2=0 AND ISFEE2=1 AND STOPDATE IS NULL AND FEE2DATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';
    if Trim(ExpectedPaysForm.Name.Text) <> '' then
    begin
        sqlText := sqlText + ' AND M.NAME LIKE ''%' + ExpectedPaysForm.Name.Text + '%'''
    end;
    if Trim(ExpectedPaysForm.Agent.Text) <> '' then
    begin
        sqlText := sqlText + ' AND AGENTCODE IN (SELECT AGENT_CODE FROM AGENT WHERE UPPER(A.NAME) LIKE ''%' + ExpectedPaysForm.Agent.Text + '%'')'
    end;

    with WorkSQL, WorkSQL.SQL do
    begin
        try
            if not InitExcel(MSEXCEL) then exit;
            MSEXCEL.WorkBooks.Add;
            ASH := MSEXCEL.ActiveSheet;

            Screen.Cursor := crHourGlass;
            Close;
            Clear;
            Add(sqlText);
            Open;
            Line := 5;
            ASH.Range['A1:R1'].MergeCells := True;
            ASH.Range['A2:R2'].MergeCells := True;
            ASH.Range['A1:A1'].Value := 'Справка о сроках оплаты за период ' + DateToStr(DtFrom) + ' - ' + DateToStr(DtTo);
            ASH.Range['A2:A2'].Value := 'Вид страхования  "30.Обязательное ГО авто владельцев"';
            ASH.Range['A4:A4'].Value := 'Дата ' + DateToStr(Date + Time);

            ASH.Range['A3:R4'].Font.Bold := True;
            ASH.Range['A4:A4'].Value := 'Полис';
            ASH.Range['B4:B4'].Value := 'Выдан';
            ASH.Range['C4:C4'].Value := 'Страхователь';
            ASH.Range['D4:D4'].Value := 'Ф/Ю';
            ASH.Range['E4:E4'].Value := 'Адрес';
            ASH.Range['F4:F4'].Value := 'Телефон';
            ASH.Range['G4:G4'].Value := 'Валюта';

            ASH.Range['H3:H3'].Value := 'Оплачено';
            ASH.Range['H3:I3'].MergeCells := True;
            ASH.Range['H4:H4'].Value := 'В валюте полиса';
            ASH.Range['I4:I4'].Value := 'В BYR';
            ASH.Range['J3:K3'].MergeCells := True;
            ASH.Range['J3:J3'].Value := 'Подлежит оплате';
            ASH.Range['J4:J4'].Value := 'В валюте полиса';
            ASH.Range['K4:K4'].Value := 'В BYR';
            ASH.Range['L4:L4'].Value := 'Срок платежа';
            ASH.Range['M4:M4'].Value := 'Тип';
            ASH.Range['N4:N4'].Value := 'Марка';
            ASH.Range['O4:O4'].Value := 'Гос. номер';
            ASH.Range['P4:P4'].Value := 'Страховщик';
            ASH.Range['Q4:Q4'].Value := 'Филиал';
            ASH.Range['R4:R4'].Value := 'Вид';

            AllRecords := RecordCount;
            N := 0;
            while not Eof do
            begin
                Inc(N);
                StatusBar.Panels[0].Text := IntToStr((N * 100) DIV AllRecords) + '%';
                StatusBar.Update;
                //Для 2й оплаты проверим дубликат
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                    HandBookSQL.Open;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                    begin
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                    end;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                    begin
                        HandBookSQL.Close;
                        Next;
                        continue;
                    end;
                    HandBookSQL.Close;
                end;

                Data.Curr[1] := '';
                Data.Curr[1] := '';
                Data.NalSumm[1].Summ := 0;
                Data.BezNalSumm[1].Summ := 0;
                Data.NalSumm[1].RubSumm := 0;
                Data.BezNalSumm[1].RubSumm := 0;


                ArrayData[1, 1] := FieldByName('SERIA').AsString+ '/' + FieldByName('NUMBER').AsString;
                ArrayData[1, 2] := FieldByName('REGDATE').AsString;
                ArrayData[1, 3] := FieldByName('NAME').AsString;
                if FieldByName('InsurerType').AsString = '1' then
                begin
                    ArrayData[1, 4] := 'Ф';
                end;
                if FieldByName('InsurerType').AsString = '0' then
                begin
                    ArrayData[1, 4] := 'Ю';
                end;
                if FieldByName('InsurerType').AsString = '2' then
                begin
                    ArrayData[1, 4] := 'ИП';
                end;
                ArrayData[1, 5] := FieldByName('ADDRESS').AsString;
                ArrayData[1, 6] := GetPhone(FieldByName('ADDRESS').AsString);
                ArrayData[1, 7] := 'EUR';
                ArrayData[1, 8] := FieldByName('ALLFEE').AsFloat-FieldByName('FEE2').AsFloat;

                if FieldByName('PAY1CURR').AsString = 'BRB' then
                    ArrayData[1, 9] := FieldByName('PAY1').AsFloat
                else
                begin
                    InitAgRep(@Data,
                              FieldByName('PAYTYPE').AsInteger,
                              FieldByName('PAYDATE').AsDateTime,
                              FieldByName('ALLFEE').AsFloat-FieldByName('FEE2').AsFloat,
                              'EUR',
                              FieldByName('RATE').AsFloat,
                              -1
                             );
                    ArrayData[1, 9] := Data.LastRubSumm;
                end;



                InitAgRep(@Data,
                          FieldByName('PAYTYPE').AsInteger,
                          Date,
                          FieldByName('FEE2').AsFloat,
                          'EUR',
                          FieldByName('RATE').AsFloat,
                          -1
                         );
                ArrayData[1, 10] := FieldByName('FEE2').AsFloat;
                ArrayData[1, 11] := Data.LastRubSumm;
                ArrayData[1, 12] := FieldByName('FEE2DATE').AsString;
                ArrayData[1, 13] := FieldByName('LETTER').AsString;
                ArrayData[1, 14] := FieldByName('MARKA').AsString;
                ArrayData[1, 15] := FieldByName('AUTONUMBER').AsString;
                ArrayData[1, 16] := DecodeAgentNamePx(FieldByName('AGENTCODE').AsString);
                ArrayData[1, 17] := FieldByName('DIVISIONCODE').AsString;
                ArrayData[1, 18] := '30';

                ASH.Range['A' + IntToStr(Line) + ':R' + IntToStr(Line)] := ArrayData;
                Inc(Line);

                Next;
            end;

            Close;
        except
            on E : Exception do
            begin
                MessageDlg(e.Message + #13'Создание отчета остановлено', mtInformation, [mbOk], 0);
            end;
        end;
    end;

    VarClear(ArrayData);
    ASH.Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)].Value := '=SUM(I5:I' + IntToStr(Line - 1) + ')';
    ASH.Range['K' + IntToStr(Line) + ':K' + IntToStr(Line)].Value := '=SUM(K5:K' + IntToStr(Line - 1) + ')';
    ASH.Columns['A:R'].EntireColumn.AutoFit;
    MSEXCEL.Visible := true;
    Screen.Cursor := crDefault;
end;

procedure TMainForm.MonthlyReportClick(Sender: TObject);
var
    DtFrom, DtTo : TDateTime;
    RegDocs : integer;
    CurrDocs, CntZajavl, DupCount, ActiveWorks, CntClosed : integer;
    IncomeSumm, RetSumm, delta : double;
    AvPaySumm : double;
    Data : AgReportStruct;

    Curr, All : integer;
    errStr, s1, s2 : string;
begin
    AddIndex('PrevSeria;PrevNumber');

    delta := 732312 - EncodeDate(2006, 1, 1);
    //GetPeriodForm.DtFrom.Date := EncodeDate(2005, 1, 1);
    //GetPeriodForm.DtTo.Date := EncodeDate(2005, 1, 31);
    if GetPeriodForm.ShowModal <> mrOK then exit;
    DtFrom := Trunc(min(GetPeriodForm.DtFrom.Date, GetPeriodForm.DtTo.Date));
    DtTo := Trunc(max(GetPeriodForm.DtFrom.Date, GetPeriodForm.DtTo.Date));

    with WorkSQL, WorkSQL.SQL do
    begin
        try
            Screen.Cursor := crHourGlass;
            //Количество заключенных
            Clear;
            Close;
            Add('SELECT COUNT(*) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND REGDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' ' +
                  'AND  PAYDATE IS NOT NULL  AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)AND STATE<>1');
            StatusBar.Panels[0].Text := 'Количество заключенных...';
            StatusBar.Update;
            Open;
            RegDocs := FieldByName('VAL').AsInteger;

            //Количество дубликатов
            Clear;
            Close;
            Add('SELECT COUNT(*) AS VAL FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND REGDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' ' +
                  'AND PAYDATE IS NOT NULL  AND (PREVNUMBER>0)AND STATE<>1');
            StatusBar.Panels[0].Text := 'Количество дубликатов...';
            StatusBar.Update;
            Open;
            DupCount := FieldByName('VAL').AsInteger;

            //Количество действующих договоров (на конец отчетного периода).
            CurrDocs := 0;
            Clear;
            Close;
            Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PAYDATE IS NOT NULL  AND BASECARCODE=101 AND FROMDATE <= :D2 AND TODATE > :D2 AND STATE<>1 AND (STOPDATE IS NULL OR STOPDATE<:D2)');
            //ParamByName('D1').AsDateTime := DtFrom;
            ParamByName('D2').AsDateTime := DtTo;
            Open;
            CurrDocs := FieldByName('CNT').AsINteger;

            Clear;
            Close;
            Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' M, ' + GetTableName() + ' M2 WHERE M.SERIA=M2.PREVSERIA AND M2.PREVNUMBER=M.NUMBER AND M.PAYDATE IS NOT NULL  AND M.BASECARCODE=101 AND M.FROMDATE <= :D2 AND M.TODATE > :D2 AND M.STATE<>1');
            //ParamByName('D1').AsDateTime := DtFrom;
            ParamByName('D2').AsDateTime := DtTo;
            Open;
            CurrDocs := CurrDocs - FieldByName('CNT').AsINteger;
        {
            Curr := 0;
            All := RecordCount;

            while (not Eof) do
            begin
                if (Curr MOD 50) = 0 then
                begin
                    StatusBar.Panels[0].Text := 'Количество действующих ' + IntToStr((Curr * 100) DIV All) + '%...';
                    StatusBar.Update;
                end;
                Curr := Curr + 1;
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM MANDATOR WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                    HandBookSQL.Open;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                    begin
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                    end;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                    begin
                        HandBookSQL.Close;
                        Next;
                        continue;
                    end
                    else
                    begin
                        if FieldByName('STATE').AsInteger = 3 then
                        begin
                            errStr := errStr + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' утерян но нету дубликата'#10;
                        end;
                    end;
                    HandBookSQL.Close;
                end;

                CurrDocs := CurrDocs + 1;
                Next;
            end;
     }
            //Сумма полученных взносов
            IncomeSumm := 0;
            Clear;
            Close;
            Add('SELECT 1 AS N, SERIA, NUMBER, PAY1 AS PAY, RATE, PAY1CURR AS CURR, RET1 AS DT, STATE FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND PAYDATE IS NOT NULL  AND RET1 BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' AND PREVSERIA IS NULL ' +
                ' UNION ALL ' +
                'SELECT 2,      SERIA, NUMBER, PAY2SUMMA,   PAY2RATE, PAY2CURR,      RET2     , STATE FROM ' + GetTableName() + ' WHERE BASECARCODE=101 AND PAY2DATE IS NOT NULL AND RET2 BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''');
            Clipboard.AsText := SQL.Text;
            Open;
            Curr := 0;
            All := RecordCount;
            while not Eof do
            begin
                if (Curr MOD 50) = 0 then
                begin
                    StatusBar.Panels[0].Text := 'Сумма полученных взносов... ' + IntToStr((Curr * 100) DIV All) + '%...';
                    StatusBar.Update;
                end;
                Curr := Curr + 1;

                if FieldByName('N').AsInteger = 2 then
                begin
                    HandBookSQL.Close;
                    HandBookSQL.SQL.Clear;
                    HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                    HandBookSQL.Open;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                    begin
                        raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                    end;
                    if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                    begin
                        HandBookSQL.Close;
                        Next;
                        continue;
                    end
                    else
                    begin
                        if FieldByName('STATE').AsInteger = 3 then
                        begin
                            errStr := errStr + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' утерян но нету дубликата'#10;
                        end;
                    end;
                    HandBookSQL.Close;
                end;

                InitAgRep(@Data,
                          1,
                          FieldByName('DT').AsDateTime,
                          FieldByName('PAY').AsFloat,
                          FieldByName('CURR').AsString,
                          FieldByName('RATE').AsFloat,
                          0
                         );

                IncomeSumm := IncomeSumm + Data.LastRubSumm;
                Next;
            end;

            if not bKomplexMode then
            begin
            //Количество заявленных страховых сумм
            CntZajavl := 0;
            Clear;
            Close;
            Add('SELECT DISTINCT AV.SERIA, AV.NUMBER, AV.N FROM MANDAV2 AV, AVPAYS P WHERE AV.SERIA=SER AND AV.NUMBER=NMB AND AV.N=P.N AND TYPE=''0'' AND D BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' ');
            StatusBar.Panels[0].Text := 'Количество заявленных страховых сумм...';
            StatusBar.Update;
            Open;
            while not Eof do begin
                CntZajavl := CntZajavl + 1;
                Next;
            end;

            //Сумма выплат по внутреннему страхованию
            AvPaySumm := 0;
            Clear;
            Close;
            Add('SELECT P."SUM", CURR, D FROM AVPAYS P WHERE TYPE=''1'' AND D BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' ');
            StatusBar.Panels[0].Text := 'Сумма выплат по внутреннему страхованию...';
            StatusBar.Update;
            Open;
            while not Eof do
            begin
                if FieldByName('CURR').AsString = 'BRB' then
                begin
                    AvPaySumm := AvPaySumm + FieldByName('SUM').AsFloat
                end
                else
                begin
                    AvPaySumm := AvPaySumm + RoundTo(FieldByName('SUM').AsFloat * GetRateOfCurrency(FieldByName('CURR').AsString, FieldByName('D').AsDateTime), -2);
                end;

                Next;
            end;

            //Количество дел, находящихся в работе
            Clear;
            Close;
            Add('SELECT COUNT(*) AS CNT FROM MANDAV2 P WHERE (DTCLLOSS IS NULL OR DTCLLOSS > ''' + DateToStr(DtTo) + ''') AND WRDATE < ''' + DateToStr(DtTo) + ''' ');
            StatusBar.Panels[0].Text := 'Количество дел, находящихся в работе...';
            StatusBar.Update;
            Open;
            ActiveWorks := FieldByName('CNT').AsInteger;

            //Сумма возвратов
            Clear;
            Close;
            Add('SELECT SUM(RETBRB) AS SUMM FROM ' + GetTableName() + ' WHERE OwnerCode BETWEEN ' + FloatToStr(DtFrom+delta) + ' AND ' + FloatToStr(DtTo+delta));
            StatusBar.Panels[0].Text := 'Сумма возвратов...';
            StatusBar.Update;
            Open;
            RetSumm := -FieldByName('SUMM').AsFloat / 1000;

            CntClosed := 0;
            Clear;
            Close;
            Add('SELECT * FROM ' + GetTableName() + ' WHERE STOPDATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + ''' AND STOPDATE<TODATE OR FEE2DATE BETWEEN ''' + DateToStr(DtFrom) + '''  AND ''' + DateToStr(DtTo) + ''' AND ISPAY2=0');
            Open;
            All := RecordCount;
            Curr := 0;
            while not Eof do
            begin
                if (Curr MOD 50) = 0 then
                begin
                    StatusBar.Panels[0].Text := 'Количество досрочно закрытых ' + IntToStr((Curr * 100) DIV All) + '%...';
                    StatusBar.Update;
                end;
                Curr := Curr + 1;
                HandBookSQL.Close;
                HandBookSQL.SQL.Clear;
                HandBookSQL.SQL.Add('SELECT COUNT(*) AS CNT FROM ' + GetTableName() + ' WHERE PREVSERIA=''' + FieldByName('SERIA').AsString + ''' AND PREVNUMBER=' + FieldByName('NUMBER').AsString);
                HandBookSQL.Open;
                if HandBookSQL.FieldByName('CNT').AsInteger > 1 then
                begin
                    raise Exception.Create(FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' имеет ' + HandBookSQL.FieldByName('CNT').AsString + ' дубликатов');
                end;
                if HandBookSQL.FieldByName('CNT').AsInteger > 0 then
                begin
                    HandBookSQL.Close;
                    Next;
                    continue;
                end;
                HandBookSQL.Close;

                Inc(CntClosed);
                Next;
            end;
            end; //avarias
            //END
            Close;
        except
            on E : Exception do
            begin
                MessageDlg(e.Message, mtInformation, [mbOk], 0);
                Screen.Cursor := crDefault;
                StatusBar.Panels[0].Text := '';
                exit;
            end;
        end;
    end;

    StatusBar.Panels[0].Text := '';

    if not InitExcel(MSEXCEL) then exit;
    MSEXCEL.WorkBooks.Add;
    if MSEXCEL.Sheets.Count = 1 then
    begin
        MSEXCEL.Sheets.Add;
        MSEXCEL.Sheets[1].Select;
    end;
    MSEXCEL.ActiveSheet.Range['A1:A1'].Value := 'Отчет за период с ' + DateToStr(DtFrom) + ' по ' +  DateToStr(DtTo);

    MSEXCEL.ActiveSheet.Range['A3:A3'].Value := 'Показатель';
    MSEXCEL.ActiveSheet.Range['B3:B3'].Value := 'Количество';
    MSEXCEL.ActiveSheet.Range['C3:C3'].Value := 'Сумма в тыс. руб.';

    MSEXCEL.ActiveSheet.Range['A4:A4'].Value := 'Количество заключенных договоров';
    MSEXCEL.ActiveSheet.Range['A5:A5'].Value := 'Количество действующих договоров';
    MSEXCEL.ActiveSheet.Range['A6:A6'].Value := 'Сумма взносов по внутреннему страхованию';
    MSEXCEL.ActiveSheet.Range['A7:A7'].Value := 'Количество заявленных страховых случаев';
    MSEXCEL.ActiveSheet.Range['A8:A8'].Value := 'Сумма выплат по внутреннему страхованию';
    MSEXCEL.ActiveSheet.Range['A9:A9'].Value := 'Количество дел, находящихся в работе';
    MSEXCEL.ActiveSheet.Range['A10:A10'].Value := 'Количество досрочно прекращенных договоров';
    MSEXCEL.ActiveSheet.Range['A11:A11'].Value := 'Количество дубликатов и переоформленных договоров';
    MSEXCEL.ActiveSheet.Range['A12:A12'].Value := 'Сумма страховых взносов, которые были возвращены страхователям в связи с переоформлением или досрочным прекращением договора.';
    MSEXCEL.ActiveSheet.Rows['12:12'].RowHeight := MSEXCEL.ActiveSheet.Rows['12:12'].RowHeight * 2;
    MSEXCEL.ActiveSheet.Range['A12:A12'].WrapText := True;

    MSEXCEL.ActiveSheet.Range['B4:B4'].Value := RegDocs;
    MSEXCEL.ActiveSheet.Range['B5:B5'].Value := CurrDocs;
    MSEXCEL.ActiveSheet.Range['C6:C6'].Value := IncomeSumm / 1000;
    MSEXCEL.ActiveSheet.Range['B7:B7'].Value := CntZajavl;
    MSEXCEL.ActiveSheet.Range['C8:C8'].Value := AvPaySumm / 1000;
    MSEXCEL.ActiveSheet.Range['B9:B9'].Value := ActiveWorks;
    MSEXCEL.ActiveSheet.Range['B10:B10'].Value := CntClosed;
    MSEXCEL.ActiveSheet.Range['B11:B11'].Value := DupCount;
    MSEXCEL.ActiveSheet.Range['C12:C12'].Value := RetSumm;

    MSEXCEL.ActiveSheet.Range['A3:C3'].Font.Bold := True;
    MSEXCEL.ActiveSheet.Columns['A:C'].EntireColumn.AutoFit;

    if errStr <> '' then
    begin
        MSEXCEL.ActiveSheet.Range['A14:A14'].WrapText := True;
        MSEXCEL.ActiveSheet.Range['A14:A14'].VerticalAlignment := xlTop;
        MSEXCEL.ActiveSheet.Rows['14:14'].RowHeight := MSEXCEL.ActiveSheet.Rows['14:14'].RowHeight * 10;
        MSEXCEL.ActiveSheet.Range['A14:A14'].Value :=  'Ошибки' + #10 + errStr;
    end;

    MSEXCEL.ActiveSheet.Name := 'Отчет';
    MSEXCEL.Visible := True;

    with WorkSQL, WorkSQL.SQL do
    begin
            //Количество дел, находящихся в работе
            Clear;
            Close;
            Add('SELECT * FROM MANDAV2 P WHERE (DTCLLOSS IS NULL OR DTCLLOSS > ''' + DateToStr(DtTo) + ''') AND WRDATE < ''' + DateToStr(DtTo) + ''' ');
            StatusBar.Panels[0].Text := 'Количество дел, находящихся в работе...';
            StatusBar.Update;
            Open;
            MSEXCEL.Sheets[2].Name := 'Дела в работе';
            Line := 1;
            MSEXCEL.Sheets[2].Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := 'Полис';
            MSEXCEL.Sheets[2].Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)].Value := 'Номер случая';
            MSEXCEL.Sheets[2].Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := 'Номер дела';
            MSEXCEL.Sheets[2].Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)].Value := 'Дата заявления';
            MSEXCEL.Sheets[2].Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)].Value := 'Дата закрытия';
            MSEXCEL.Sheets[2].Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)].Value := 'Заявлено';
            MSEXCEL.Sheets[2].Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)].Value := 'Выплачено';
            MSEXCEL.Sheets[2].Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)].Value := 'Потерпевший';
            MSEXCEL.Sheets[2].Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)].Value := 'Виновник';
            Line := 2;
            while not Eof do
            begin
                MSEXCEL.Sheets[2].Range['A' + IntToStr(Line) + ':A' + IntToStr(Line)].Value := FieldByName('Seria').AsString + '/' + FieldByName('Number').AsString;
                MSEXCEL.Sheets[2].Range['B' + IntToStr(Line) + ':B' + IntToStr(Line)].Value := FieldByName('N').AsString;
                MSEXCEL.Sheets[2].Range['C' + IntToStr(Line) + ':C' + IntToStr(Line)].Value := FieldByName('NWork').AsString;
                MSEXCEL.Sheets[2].Range['D' + IntToStr(Line) + ':D' + IntToStr(Line)].Value := FieldByName('WrDate').AsString;
                MSEXCEL.Sheets[2].Range['E' + IntToStr(Line) + ':E' + IntToStr(Line)].Value := FieldByName('DTCLLOSS').AsString;


                HandBookSQL.SQL.Clear;
                HandBookSQL.SQL.Add('SELECT SUM(P."SUM") AS S, CURR, TYPE FROM MANDAV2 AV, AVPAYS P WHERE AV.SERIA=SER AND AV.NUMBER=NMB AND AV.N=P.N AND SER=''' + FieldByName('Seria').AsString + ''' AND NMB = ''' + FieldByName('Number').AsString + ''' GROUP BY CURR, TYPE ORDER BY TYPE');
                HandBookSQL.Open;
                s1 := '';
                s2 := '';                
                while not HandBookSQL.Eof do begin
                    if HandBookSQL.FieldByName('TYPE').AsString = '0' then
                    begin
                        s1 := s1 + HandBookSQL.FieldByName('S').AsString + ' ' +  HandBookSQL.FieldByName('CURR').AsString + ' ';
                    end;
                    if HandBookSQL.FieldByName('TYPE').AsString = '1' then
                    begin
                        s2 := s2 + HandBookSQL.FieldByName('S').AsString + ' ' +  HandBookSQL.FieldByName('CURR').AsString + ' ';
                    end;
                    HandBookSQL.Next;
                end;
                MSEXCEL.Sheets[2].Range['F' + IntToStr(Line) + ':F' + IntToStr(Line)].Value := s1;
                MSEXCEL.Sheets[2].Range['G' + IntToStr(Line) + ':G' + IntToStr(Line)].Value := s2;
                MSEXCEL.Sheets[2].Range['H' + IntToStr(Line) + ':H' + IntToStr(Line)].Value := FieldByName('FIO').AsString;
                MSEXCEL.Sheets[2].Range['I' + IntToStr(Line) + ':I' + IntToStr(Line)].Value := FieldByName('BadMan').AsString;
                INc(Line);
                Next;
            end;
            Close;
    end;
    MSEXCEL.Sheets[2].Range['A1:I1'].Font.Bold := True;
    MSEXCEL.Sheets[2].Columns['A:I'].EntireColumn.AutoFit;

    Screen.Cursor := crDefault;

end;

procedure TMainForm.StopInsInfoClick(Sender: TObject);
var
    ArrayData, ASH : VARIANT;
    N, AllRecords, Line : integer;
    sqlText : string;
    Data : AgReportStruct;
    DtFrom, DtTo : TDAteTime;
    INI : TIniFile;
    ss : string;
begin
    ExpectedPaysForm.Caption := 'Окончание страхования';
    if ExpectedPaysForm.ShowModal <> mrOk then
        exit;

    ArrayData := VarArrayCreate([1, 1, 1, 28], varVariant);

    DtFrom := min(ExpectedPaysForm.DtFrom.Date, ExpectedPaysForm.DtTo.Date);
    DtTo := max(ExpectedPaysForm.DtFrom.Date, ExpectedPaysForm.DtTo.Date);
    sqlText := 'SELECT M.*, A.NAME AS ANAME FROM ' + GetTableName() + ' M LEFT OUTER JOIN AGENT A ON A.AGENT_CODE=M.AGENTCODE WHERE BASECARCODE=101 AND STOPDATE IS NULL AND PAYDATE IS NOT NULL AND STATE=0 AND TODATE BETWEEN ''' + DateToStr(DtFrom) + ''' AND ''' + DateToStr(DtTo) + '''';
    if Trim(ExpectedPaysForm.Name.Text) <> '' then
    begin
        sqlText := sqlText + ' AND M.NAME LIKE ''%' + ExpectedPaysForm.Name.Text + '%'''
    end;
    if Trim(ExpectedPaysForm.Agent.Text) <> '' then
    begin
        sqlText := sqlText + ' AND AGENTCODE IN (SELECT AGENT_CODE FROM AGENT WHERE UPPER(A.NAME) LIKE ''%' + ExpectedPaysForm.Agent.Text + '%'')'
    end;

    INI := TIniFile.Create(BLANK_INI);
    if bKomplexMode then
    begin
        ss := INI.ReadString(KOMPLEX_SECT, 'INI', '');
        INI.Free;
        INI := TIniFile.Create(ss);
    end;

    with WorkSQL, WorkSQL.SQL do
    begin
        try
            if not InitExcel(MSEXCEL) then exit;
            MSEXCEL.WorkBooks.Add;
            ASH := MSEXCEL.ActiveSheet;

            Screen.Cursor := crHourGlass;
            Close;
            Clear;
            Add(sqlText);
            Open;
            Line := 6;
            ASH.Range['A1:A1'].Value := 'Состояние по полисам за период ' + DateToStr(DtFrom) + ' - ' + DateToStr(DtTo);
            ASH.Range['A2:A2'].Value := 'Вид страхования  "30.Обязательное ГО авто владельцев"';
            ASH.Range['A3:A3'].Value := 'Состояние полисов. Полисы, по котором "истекает" срок страхования';
            ASH.Range['A1:D1'].MergeCells := True;
            ASH.Range['A2:D2'].MergeCells := True;
            ASH.Range['A3:D3'].MergeCells := True;

            ASH.Range['A5:A5'].Value := 'Страхователь';
            ASH.Range['B5:B5'].Value := 'Полис';
            ASH.Range['C5:C5'].Value := 'Дата выдачи';
            ASH.Range['D5:D5'].Value := 'Страховая сумма';
            ASH.Range['E5:E5'].Value := 'Валюта';
            ASH.Range['F5:F5'].Value := 'Премия';

            ASH.Range['G5:G5'].Value := 'Валюта';
            ASH.Range['H5:H5'].Value := 'Срок';
            ASH.Range['I5:I5'].Value := 'Дата начала страх.';
            ASH.Range['J5:J5'].Value := 'Дата окончания страх.';
            ASH.Range['K5:K5'].Value := 'Количество дней';

            ASH.Range['L5:L5'].Value := 'Дата закрытия';
            ASH.Range['M5:M5'].Value := 'Лицо';
            ASH.Range['N5:N5'].Value := 'Резидент';
            ASH.Range['O5:O5'].Value := 'Агент';
            ASH.Range['P5:P5'].Value := 'Район';
            ASH.Range['Q5:Q5'].Value := 'Адрес';

            ASH.Range['R5:R5'].Value := 'К.застр.';
            ASH.Range['S5:S5'].Value := 'Вариант';
            ASH.Range['T5:T5'].Value := 'Кол-во случаев';
            ASH.Range['U5:U5'].Value := 'Тип';
            ASH.Range['V5:V5'].Value := 'Марка';

            ASH.Range['W5:W5'].Value := 'Гос.номер';

            ASH.Range['X5:X5'].Value := 'Год';
            ASH.Range['Y5:Y5'].Value := 'Кузов,шасси';
            ASH.Range['Z5:Z5'].Value := 'К1';
            ASH.Range['AA5:AA5'].Value := 'К2';
            ASH.Range['AB5:AB5'].Value := 'К3';

            AllRecords := RecordCount;
            N := 0;
            while not Eof do
            begin
                Inc(N);
                StatusBar.Panels[0].Text := IntToStr((N * 100) DIV AllRecords) + '%';
                StatusBar.Update;

                Data.Curr[1] := '';
                Data.Curr[1] := '';
                Data.NalSumm[1].Summ := 0;
                Data.BezNalSumm[1].Summ := 0;
                Data.NalSumm[1].RubSumm := 0;
                Data.BezNalSumm[1].RubSumm := 0;


                ArrayData[1, 1] := FieldByName('NAME').AsString;
                ArrayData[1, 2] := FieldByName('SERIA').AsString+ '/' + FieldByName('NUMBER').AsString;
                ArrayData[1, 3] := FieldByName('REGDATE').AsString;
                if FieldByName('InsurerType').AsString = '1' then
                begin
                    ArrayData[1, 13] := 'ФИЗИЧЕСКОЕ';
                end;
                if FieldByName('InsurerType').AsString = '0' then
                begin
                    ArrayData[1, 13] := 'ЮРИДИЧЕСКОЕ';
                end;
                if FieldByName('InsurerType').AsString = '2' then
                begin
                    ArrayData[1, 13] := 'ИП';
                end;
                ArrayData[1, 15] := FieldByName('ANAME').AsString;
                ArrayData[1, 9] := FieldByName('FROMDATE').AsString;
                ArrayData[1, 10] := FieldByName('TODATE').AsString;
                ArrayData[1, 11] := FieldByName('TODATE').AsFloat - FieldByName('FROMDATE').AsFloat + 1;
                ArrayData[1, 12] := FieldByName('STOPDATE').AsString;
                ArrayData[1, 17] := FieldByName('ADDRESS').AsString;
                ArrayData[1, 18] := 1;
                ArrayData[1, 21] := FieldByName('LETTER').AsString;
                ArrayData[1, 22] := FieldByName('MARKA').AsString;
                ArrayData[1, 23] := FieldByName('AUTONUMBER').AsString;
                ArrayData[1, 25] := FieldByName('NUMBERBODY').AsString + ' ' + FieldByName('CHASSIS').AsString;
                ArrayData[1, 8] := DecodePeriod(FieldByName('PERIOD').AsString, FieldByName('FROMDATE').AsFloat, FieldByName('TODATE').AsFloat, false);

                ArrayData[1, 26] := FieldByName('K1').AsString;
                ArrayData[1, 27] := FieldByName('K2').AsString;
                ArrayData[1, 28] := '';
                if FieldByName('RegDate').AsDateTime >= EncodeDate(2006, 9, 1) then
                if FieldByName('InsurerType').AsInteger <> 0 then
                begin
                    if FieldByName('Placement').AsString = '' then
                    begin
                        raise Exception.Create('Не указан K3 для ' + FieldByName('SERIA').AsString+ '/' + FieldByName('NUMBER').AsString);
                    end;
                    ArrayData[1, 28] := (Ord(FieldByName('Placement').AsString[1]) - 33) / 10;
                end;
                ArrayData[1, 14] := INI.ReadString('MANDATORY', 'TarifTable_Is_Resident' + FieldByName('Resident').AsString, '');
                ArrayData[1, 16] := FieldByName('CityCode').AsString;

                ASH.Range['A' + IntToStr(Line) + ':AB' + IntToStr(Line)] := ArrayData;
                Inc(Line);

                Next;
            end;

            Close;
        except
            on E : Exception do
            begin
                MessageDlg(e.Message + #13'Создание отчета остановлено.', mtInformation, [mbOk], 0);
            end;
        end;
    end;

    INI.Free;
    VarClear(ArrayData);
    ASH.Range['A5:AB5'].Font.Bold := True;
    ASH.Columns['A:AB'].EntireColumn.AutoFit;
    MSEXCEL.Visible := true;
    Screen.Cursor := crDefault;
end;

function TMainForm.DecodeK3(val : double) : string;
var
    s : string;
begin
    //DecodeK3 := '?';
    //if abs(val - 1.3) < 1e-5 then DecodeK3 := '11';
    //if abs(val - 1.2) < 1e-5 then DecodeK3 := '13';
    //if abs(val - 1.1) < 1e-5 then DecodeK3 := '12';
    //s := Format('%1.1f', [val * 10]);
    s := IntToStr(round(val * 10));
    DecodeK3 := s;
end;

procedure TMainForm.DivisionLabelClick(Sender: TObject);
begin
    MessageDlg(COMPAMYNAME + ' ' + DivisionLabel.Caption, mtInformation, [mbOk], 0);
end;

procedure TMainForm.MainQueryONE_1SUGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
   if (Sender.AsString = 'N') or (Sender.AsString = '') then exit;
   if Sender.AsString = 'Y' then Text := 'ГС'
   else
   if Sender.AsString = '0' then Text := 'КВ'
   else Text := _1SU_SERIES[ord(Sender.AsString[1]) - ord('A')];
end;

procedure TMainForm.K3AlwaysClick(Sender: TObject);
begin
    K3Always.Checked := not K3Always.Checked 
end;

end.

