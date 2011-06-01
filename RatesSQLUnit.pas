unit RatesSQLUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, StdCtrls, ExtCtrls, Grids, DBGrids, RXDBCtrl, RXSpin,
  ComCtrls, Mask;

type
  TRatesSQL = class(TForm)
    Panel1: TPanel;
    RatesTbl: TTable;
    RatesSource: TDataSource;
    RatesGrid: TRxDBGrid;
    RatesTblRateDate: TDateField;
    RatesTblEUR_NAL: TFloatField;
    RatesTblEUR_BNL: TFloatField;
    RatesTblUSD_NAL: TFloatField;
    RatesTblUSD_BNL: TFloatField;
    RatesTblDM_NAL: TFloatField;
    RatesTblDM_BNL: TFloatField;
    RatesTblRUR_NAL: TFloatField;
    RatesTblRUR_BNL: TFloatField;
    Panel2: TPanel;
    AddBtn: TButton;
    DelBtn: TButton;
    RATESDB: TDatabase;
    Panel3: TPanel;
    Button1: TButton;
    DeltaDate: TRxSpinEdit;
    FindDate: TDateTimePicker;
    RatesTblUAH: TFloatField;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AddBtnClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure RatesTblBeforeDelete(DataSet: TDataSet);
    procedure DelBtnClick(Sender: TObject);
    procedure RatesGridGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure RatesTblEUR_NALChange(Sender: TField);
    procedure RatesTblUSD_NALChange(Sender: TField);
    procedure RatesTblDM_NALChange(Sender: TField);
    procedure RatesTblRUR_NALChange(Sender: TField);
    procedure FindDateChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RatesSQL: TRatesSQL;

implementation

uses inifiles;

{$R *.DFM}

procedure TRatesSQL.Button1Click(Sender: TObject);
begin
     if RatesTbl.State in [dsEdit, dsInsert] then
         RatesTbl.Post;
     Close
end;

procedure TRatesSQL.FormCreate(Sender: TObject);
var
     INI : TIniFile;
begin
     INI := TIniFile.Create(BLANK_INI);
     RATESDB.AliasName := INI.ReadString('RATES', 'ODBCALIAS', '');
     RatesTbl.IndexName := INI.ReadString('RATES', 'ORDER', '');
     INI.Free;

     try
        RATESDB.Open;
     except
        //PostMessage(Handle, WM_CLOSE, 0, 0);
        //exit;
        Halt
     end;

     RatesTbl.Open;
     RatesTbl.Last;
     FindDate.Date := Date;
end;

procedure TRatesSQL.AddBtnClick(Sender: TObject);
var
     D : TDate;
begin
     if RatesTbl.State in [dsEdit, dsInsert] then
         RatesTbl.Post;

     RatesTbl.Last;
     D := RatesTblRateDate.AsDateTime;
     RatesTbl.Append;
     RatesTblRateDate.AsDateTime := D + DeltaDate.Value;
end;

procedure TRatesSQL.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     CanClose := false;
     if RatesTbl.State in [dsEdit, dsInsert] then
         RatesTbl.Post;
     CanClose := true;
end;

procedure TRatesSQL.RatesTblBeforeDelete(DataSet: TDataSet);
begin
     if not DelBtn.Focused then
        Abort;

     if MessageDlg('Вы хотите удалить курсы валют за ' + RatesTblRateDate.AsString, mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
        Abort;
end;

procedure TRatesSQL.DelBtnClick(Sender: TObject);
begin
     if RatesTbl.State = dsInsert then begin
        RatesTbl.Cancel;
        exit;
     end;

     RatesTbl.Delete
end;

procedure TRatesSQL.RatesGridGetCellParams(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor; Highlight: Boolean);
var
     d, m, y : WORD;
begin
     DecodeDate(RatesTblRateDate.AsDateTime, y, m, d);
     if (m MOD 2) = 1 then
         Background := $00FFFFFF
     else
         Background := $00DDDDDD;

     if RatesTblRateDate = Field then begin
         AFont.Color := $00;
         AFont.Style := [fsBold];
     end;

     if Highlight then begin
         AFont.Color := $00;
         Background := $0000FF00;
     end;
end;

procedure TRatesSQL.RatesTblEUR_NALChange(Sender: TField);
begin
     if RatesTblEUR_BNL.IsNull then
         RatesTblEUR_BNL.Value := RatesTblEUR_NAL.Value;
end;

procedure TRatesSQL.RatesTblUSD_NALChange(Sender: TField);
begin
     if RatesTblUSD_BNL.IsNull then
         RatesTblUSD_BNL.Value := RatesTblUSD_NAL.Value;
end;

procedure TRatesSQL.RatesTblDM_NALChange(Sender: TField);
begin
     if RatesTblDM_BNL.IsNull then
         RatesTblDM_BNL.Value := RatesTblDM_NAL.Value;
end;

procedure TRatesSQL.RatesTblRUR_NALChange(Sender: TField);
begin
     if RatesTblRUR_BNL.IsNull then
         RatesTblRUR_BNL.Value := RatesTblRUR_NAL.Value;
end;

procedure TRatesSQL.FindDateChange(Sender: TObject);
var
     Counter : integer;
begin
     Counter := 0;
     while not RatesTbl.Locate('RateDate', FindDate.Date - Counter, []) do begin
           Inc(Counter);
           if(Counter > 30) then break;
     end;
end;

end.
