unit AvarUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, RxQuery, StdCtrls, ExtCtrls, Grids, DBGrids, RXDBCtrl,
  ComCtrls;

type
  TAvForm = class(TForm)
    AVQuery: TRxQuery;
    AVQuerySeria: TStringField;
    AVQueryNumber: TFloatField;
    AVQueryN: TFloatField;
    AVQueryAvDate: TDateField;
    AVQueryAvPlace: TStringField;
    AVQueryCityCode: TStringField;
    AVQueryNWork: TStringField;
    AVQueryWrDate: TDateField;
    AVQueryWhat: TFloatField;
    AVQueryDoc1Type: TFloatField;
    AVQuerySN: TStringField;
    AVQueryFIO: TStringField;
    AVQueryFIOType: TFloatField;
    AVQueryFIOAddress: TStringField;
    AVQueryBadMan: TStringField;
    AVQueryCityCode_p: TStringField;
    AVQueryDiscount_p: TFloatField;
    AVQueryIsResid_p: TFloatField;
    AVQueryVladen_p: TFloatField;
    AVQueryBMStaj: TFloatField;
    AVQueryBMVTR: TFloatField;
    AVQuerySum1: TFloatField;
    AVQuerySum1V: TFloatField;
    AVQueryV1: TStringField;
    AVQuerySum1_1: TFloatField;
    AVQuerySum1_1v: TFloatField;
    AVQueryV1_1: TStringField;
    AVQueryV1_Date: TDateField;
    AVQuerySum2: TFloatField;
    AVQuerySumV2: TFloatField;
    AVQueryV2: TStringField;
    AVQuerySum2_1: TFloatField;
    AVQuerySum2_1V: TFloatField;
    AVQueryV2_1: TStringField;
    AVQueryV2_Date: TDateField;
    AVQuerySum3: TFloatField;
    AVQuerySumV3: TFloatField;
    AVQueryV3: TStringField;
    AVQuerySum3_1: TFloatField;
    AVQuerySum3_1V: TFloatField;
    AVQueryV3_1: TStringField;
    AVQueryV3_Date: TDateField;
    AVQueryUpdateDate: TDateField;
    AVQueryCountry_p: TStringField;
    AVQueryMarka_p: TStringField;
    AVQueryBase_Type_p: TStringField;
    AVQueryCharval: TFloatField;
    AVQueryAutoNmb_p: TStringField;
    AVQueryBody_no: TStringField;
    AVQueryOwnerCode_p: TFloatField;
    AVQueryBaseCarCode_p: TFloatField;
    AVQueryCarCode_p: TFloatField;
    AVQueryPaymentCode: TFloatField;
    AVQueryOutSumma: TStringField;
    RxDBGrid1: TRxDBGrid;
    DataSourceAV: TDataSource;
    Panel1: TPanel;
    Label1: TLabel;
    AvCount: TLabel;
    Panel2: TPanel;
    Button1: TButton;
    Button2: TButton;
    AvTable: TTable;
    WorkSQL: TRxQuery;
    AVQueryDeclSumma: TFloatField;
    AVQueryPayInfo: TStringField;
    Label2: TLabel;
    PeriodFrom: TDateTimePicker;
    PeriodTo: TDateTimePicker;
    btnRefresh: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure RxDBGrid1GetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure Button2Click(Sender: TObject);
    procedure AVQueryCalcFields(DataSet: TDataSet);
    procedure btnRefreshClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AvForm: TAvForm;

implementation

uses MF, datamod;

{$R *.DFM}

procedure TAvForm.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TAvForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     Action := caFree
end;

procedure TAvForm.FormDestroy(Sender: TObject);
begin
     AvForm := nil
end;

procedure TAvForm.FormCreate(Sender: TObject);
begin
     try
         AvTable.Open;
         AVQueryNWork.Size := AvTable.FieldByName('NWork').Size;
         AVQuerySN.Size := AvTable.FieldByName('SN').Size;
         AvTable.Close;

         Screen.Cursor := crHourGlass;
         AVQuery.MacroByName('WHERE').AsString := '1=1';
         if MainForm.GetFilter('') <> '' then
             AVQuery.MacroByName('WHERE').AsString := MainForm.GetFilter('');
         AVQuery.Open;
         AvCount.Caption := IntToStr(AVQuery.RecordCount);
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;

     Screen.Cursor := crDefault; 
     Left := MAINFORM.Left;
     Top := MAINFORM.Top;
     Width := MAINFORM.Width;
     Height := MAINFORM.Height - GetSystemMetrics(SM_CYCAPTION) - 5;
end;

procedure TAvForm.RxDBGrid1GetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
     if AvQuery.FieldByName('N').AsInteger > 1 then
         Background := clRed;
end;
procedure TAvForm.Button2Click(Sender: TObject);
begin
    DatSetToExcel(AVQuery);
end;

procedure TAvForm.AVQueryCalcFields(DataSet: TDataSet);
var
    s : string;
    period : string;
begin
    if PeriodFrom.Checked then period := 'D >= ''' + DateToStr(PeriodFrom.Date) + '''';
    if PeriodTo.Checked then begin
        if period <> '' then period := period + ' AND ';
        period := period + 'D <= ''' + DateToStr(PeriodTo.Date) + '''';
    end;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT SUM(T."SUM"), CURR FROM AVPAYS T');
    WorkSQL.SQL.Add('WHERE SER=:S AND NMB=:N AND TYPE=''1''');
    if period <> '' then WorkSQL.SQL.Add(' AND ' + period);
    WorkSQL.SQL.Add('GROUP BY CURR');
    WorkSQL.ParamByName('S').AsString := AVQuerySeria.AsString;
    WorkSQL.ParamByName('N').AsFloat := AVQueryNumber.AsFloat;
    WorkSQL.Open;
    while not WorkSQL.Eof do begin
        if WorkSQL.Fields[1].AsString <> '' then begin
            if WorkSQL.Fields[1].AsString = 'BRB' then
                AVQueryOutSumma.AsFloat := WorkSQL.Fields[0].AsFloat
            else
                s := 'выпл. ' + WorkSQL.Fields[0].AsString + ' ' + WorkSQL.Fields[1].AsString;
        end;
        WorkSQL.Next;
    end;
    WorkSQL.Close;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT SUM(T."SUM"), CURR FROM AVPAYS T');
    WorkSQL.SQL.Add('WHERE SER=:S AND NMB=:N AND TYPE=''0''');
    if period <> '' then WorkSQL.SQL.Add(' AND ' + period);
    WorkSQL.SQL.Add('GROUP BY CURR');
    WorkSQL.ParamByName('S').AsString := AVQuerySeria.AsString;
    WorkSQL.ParamByName('N').AsFloat := AVQueryNumber.AsFloat;
    WorkSQL.Open;
    while not WorkSQL.Eof do begin
        if WorkSQL.Fields[1].AsString <> '' then begin
            if WorkSQL.Fields[1].AsString = 'BRB' then
                AVQueryDeclSumma.AsFloat := WorkSQL.Fields[0].AsFloat
            else
                s := 'заявл. ' + WorkSQL.Fields[0].AsString + ' ' + WorkSQL.Fields[1].AsString;
        end;
        WorkSQL.Next;
    end;
    WorkSQL.Close;

    AVQueryPayInfo.AsString := s;
end;

procedure TAvForm.btnRefreshClick(Sender: TObject);
begin
    AVQuery.Close;
    AVQuery.Open;
end;

end.
