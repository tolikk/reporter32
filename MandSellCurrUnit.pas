unit MandSellCurrUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, StdCtrls, ExtCtrls, Grids, DBGrids, RXDBCtrl;

type
  TMandSellCurr = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Button1: TButton;
    Button2: TButton;
    DelButton: TButton;
    SellTbl: TTable;
    DataSourceSellCurr: TDataSource;
    RxDBGrid1: TRxDBGrid;
    SellTblSaleDate: TDateField;
    SellTblUSD: TFloatField;
    SellTblDM: TFloatField;
    SellTblEUR: TFloatField;
    SellTblUAH: TFloatField;
    SellTblRUR: TFloatField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DelButtonClick(Sender: TObject);
    procedure SellTblBeforeDelete(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MandSellCurr: TMandSellCurr;

implementation

{$R *.DFM}

procedure TMandSellCurr.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := cafree;
end;

procedure TMandSellCurr.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
     if SellTbl.Active then
       if SellTbl.State in [dsEdit, dsInsert] then
         SellTbl.Post;

     CanClose := true;
end;

procedure TMandSellCurr.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TMandSellCurr.Button2Click(Sender: TObject);
begin
     if SellTbl.State in [dsEdit, dsInsert] then
         SellTbl.Post;

     SellTbl.Append;
end;

procedure TMandSellCurr.DelButtonClick(Sender: TObject);
begin
     SellTbl.Delete
end;

procedure TMandSellCurr.SellTblBeforeDelete(DataSet: TDataSet);
begin
     if not DelButton.Focused then Abort
end;

procedure TMandSellCurr.FormCreate(Sender: TObject);
begin
     try
        SellTbl.DataBaseName := 'RESTRAX';
        SellTbl.Open;
        exit;
     except
     end;
     SellTbl.DataBaseName := 'BASO';
     SellTbl.Open;
end;

procedure TMandSellCurr.FormDestroy(Sender: TObject);
begin
     MandSellCurr := nil;
end;

end.
