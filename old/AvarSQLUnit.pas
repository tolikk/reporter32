unit AvarSQLUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, RxQuery, StdCtrls, ExtCtrls, Grids, DBGrids, RXDBCtrl;

type
  TAvSQLForm = class(TForm)
    AVQuery: TRxQuery;
    RxDBGrid: TRxDBGrid;
    DataSourceAV: TDataSource;
    Panel1: TPanel;
    Label1: TLabel;
    AvCount: TLabel;
    Panel2: TPanel;
    Button1: TButton;
    Button2: TButton;
    AVQuerySER: TStringField;
    AVQueryNMB: TIntegerField;
    AVQueryN: TSmallintField;
    AVQueryAVDT: TDateTimeField;
    AVQueryREGDT: TDateTimeField;
    AVQueryPLACE_CODE: TStringField;
    AVQueryPLACE: TStringField;
    AVQueryWRKNMB: TStringField;
    AVQueryFIO: TStringField;
    AVQueryISFIZ: TStringField;
    AVQueryISRSD: TStringField;
    AVQueryDOCTYPE: TSmallintField;
    AVQueryDOCSN: TStringField;
    AVQueryVLADEN: TSmallintField;
    AVQueryADRCODE: TStringField;
    AVQueryADDR: TStringField;
    AVQueryK: TFloatField;
    AVQueryMARKA: TStringField;
    AVQueryLTR: TStringField;
    AVQueryCHRCT: TFloatField;
    AVQueryAUTONMB: TStringField;
    AVQueryBODYSHS: TStringField;
    AVQueryBADFIO: TStringField;
    AVQuerySTAJ: TSmallintField;
    AVQueryREGRESS: TIntegerField;
    AVQueryDECISION: TIntegerField;
    AVQueryDAMAGECODE: TStringField;
    AVQueryDTFREE: TDateTimeField;
    AVQueryDTCLOSE: TDateTimeField;
    AVQueryPAY: TFloatField;
    AVQueryCURR: TStringField;
    AVQueryFS1E: TFloatField;
    AVQueryFS1B: TFloatField;
    AVQueryRS1E: TFloatField;
    AVQueryRS1B: TFloatField;
    AVQueryFS2E: TFloatField;
    AVQueryFS2B: TFloatField;
    AVQueryRS2E: TFloatField;
    AVQueryRS2B: TFloatField;
    AVQueryFS3E: TFloatField;
    AVQueryFS3B: TFloatField;
    AVQueryRS3E: TFloatField;
    AVQueryRS3B: TFloatField;
    AVQueryDT1: TDateTimeField;
    AVQueryDT2: TDateTimeField;
    AVQueryDT3: TDateTimeField;
    AVQueryOZ1: TStringField;
    AVQueryOZ2: TStringField;
    AVQueryOZ3: TStringField;
    AVQueryOTH1: TStringField;
    AVQueryOTH2: TStringField;
    AVQueryOTH3: TStringField;
    AVQueryOZS1: TFloatField;
    AVQueryOZS2: TFloatField;
    AVQueryOZS3: TFloatField;
    AVQueryOTHS1: TFloatField;
    AVQueryOTHS2: TFloatField;
    AVQueryOTHS3: TFloatField;
    AVQueryDMG1: TFloatField;
    AVQueryDMG2: TFloatField;
    AVQueryDMG3: TFloatField;
    AVQueryOWNRCODE: TIntegerField;
    AVQueryBASECARCODE: TIntegerField;
    AVQueryCARCODE: TIntegerField;
    AVQueryAVCODE: TIntegerField;
    AVQueryINSCOMP: TSmallintField;
    AVQueryUPDDT: TDateField;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure RxDBGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AvSQLForm: TAvSQLForm;

implementation

uses mandsql, datamod, AvarUnit, Clipbrd;

{$R *.DFM}

procedure TAvSQLForm.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TAvSQLForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     AVQuery.Close;
     Action := caFree
end;

procedure TAvSQLForm.FormDestroy(Sender: TObject);
begin
     AvSQLForm := nil
end;

procedure TAvSQLForm.FormCreate(Sender: TObject);
var
     S : String;
     I : Integer;
begin
     try
         Screen.Cursor := crHourGlass;
         AVQuery.MacroByName('WHERE').AsString := '1=1';
         if MandSQLForm.GetFilter <> '' then
             AVQuery.MacroByName('WHERE').AsString := MandSQLForm.GetFilter;

         if (GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0 then begin
             DataSourceAV.DataSet := nil;
             RxDBGrid.DataSource := nil;
         end;

         AVQuery.Open;

         if (GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0 then begin
             for I := 0 to AVQuery.FieldCount - 1 do
                 S := S + AVQuery.Fields[I].FieldName + #13;

             MessageDlg(S, mtInformation, [mbOk], 0);
             Clipboard.AsText := S;
             S := 'ØÀÃ 1';
             DataSourceAV.DataSet := AVQuery;
             S := 'ØÀÃ 2';
             RxDBGrid.DataSource := DataSourceAV;
             S := 'ØÀÃ 3';
         end;


         AvCount.Caption := IntToStr(AVQuery.RecordCount);
     except
         on E : Exception do
             MessageDlg(e.Message + S, mtInformation, [mbOk], 0);
     end;

     Screen.Cursor := crDefault; 
     Left := MandSQLForm.Left;
     Top := MandSQLForm.Top;
     Width := MandSQLForm.Width;
     Height := MandSQLForm.Height - GetSystemMetrics(SM_CYCAPTION) - 5;
end;

procedure TAvSQLForm.RxDBGridGetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
     if AvQuery.FieldByName('N').AsInteger > 1 then
         Background := clRed;
end;
procedure TAvSQLForm.Button2Click(Sender: TObject);
begin
    DatSetToExcel(AVQuery);
end;

end.
