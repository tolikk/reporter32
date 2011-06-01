unit ChClient_Unit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Db, RxLookup, DBTables, RxQuery, Buttons;

type
  TF_ChClient = class(TForm)
    Label1: TLabel;
    OldClient: TEdit;
    Label2: TLabel;
    ChRxQuery: TRxQuery;
    RxDBLookupCombo: TRxDBLookupCombo;
    DataSource: TDataSource;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ChRxQueryCompany_code: TIntegerField;
    ChRxQueryCompany_name: TStringField;
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    SelComp : string;
    { Public declarations }
  end;

var
  F_ChClient: TF_ChClient;

implementation

{$R *.DFM}

procedure TF_ChClient.BitBtn2Click(Sender: TObject);
begin
     if RxDBLookupCombo.Text = '' then
         exit;
     SelComp := ChRxQueryCompany_code.AsString;    
     ModalResult := mrYes;
end;

procedure TF_ChClient.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     Action := caFree;
end;

end.
