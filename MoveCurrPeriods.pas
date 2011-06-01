unit MoveCurrPeriods;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, Grids, DBGrids, RXDBCtrl, DBTables, StdCtrls, ExtCtrls;

type
  TMoveCurr = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Button1: TButton;
    MoveCurrTbl: TTable;
    DataSourceMoveCurr: TDataSource;
    RxDBGrid1: TRxDBGrid;
    MoveCurrTblPayDate: TDateField;
    MoveCurrTblPeriodFrom: TDateField;
    MoveCurrTblPeriodTo: TDateField;
    Button2: TButton;
    DelBtn: TButton;
    procedure MoveCurrTblBeforePost(DataSet: TDataSet);
    procedure Button2Click(Sender: TObject);
    procedure DelBtnClick(Sender: TObject);
    procedure MoveCurrTblBeforeDelete(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure RxDBGrid1GetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MoveCurr: TMoveCurr;

implementation

{$R *.DFM}

procedure TMoveCurr.MoveCurrTblBeforePost(DataSet: TDataSet);
begin
     if (not MoveCurrTblPayDate.IsNull) AND (MoveCurrTblPayDate.AsDateTime < MoveCurrTblPeriodTo.AsDateTime) then begin
        MessageDlg('Дата перечисления должна быть после периода', mtInformation, [mbOk], 0);
        Abort;
     end;
     if MoveCurrTblPeriodFrom.AsDateTime >= MoveCurrTblPeriodTo.AsDateTime then begin
        MessageDlg('Даты периода неправильные', mtInformation, [mbOk], 0);
        Abort;
     end;
end;

procedure TMoveCurr.Button2Click(Sender: TObject);
begin
     MoveCurrTbl.Append
end;

procedure TMoveCurr.DelBtnClick(Sender: TObject);
begin
     MoveCurrTbl.Delete
end;

procedure TMoveCurr.MoveCurrTblBeforeDelete(DataSet: TDataSet);
begin
     if not DelBtn.Focused then Abort
end;

procedure TMoveCurr.FormCreate(Sender: TObject);
begin
     MoveCurrTbl.Open
end;

procedure TMoveCurr.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     if MoveCurrTbl.Active then
         if MoveCurrTbl.State in [dsEdit, dsinsert] then
             MoveCurrTbl.Post;

     CanClose := true;
end;

procedure TMoveCurr.RxDBGrid1GetCellProps(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor);
begin
    if MoveCurrTblPayDate.IsNull then Background := clYellow
end;

end.
