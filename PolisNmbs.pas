unit PolisNmbs;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Grids, DBGrids, StdCtrls, datamod, ExtCtrls;

type
  TPolisNmbsComp = class(TForm)
    Table: TTable;
    TableS: TStringField;
    TableSTART: TFloatField;
    TableEND: TFloatField;
    DBGrid: TDBGrid;
    DataSource: TDataSource;
    Agent: TTable;
    AgentAgent_code: TStringField;
    AgentName: TStringField;
    TableAg: TTable;
    TableAgS: TStringField;
    TableAgSTART: TFloatField;
    TableAgEND: TFloatField;
    TableAgAgent: TStringField;
    TableAgAgentCombo: TStringField;
    TableCount: TIntegerField;
    TableAgCount: TIntegerField;
    Panel1: TPanel;
    StatusLine: TLabel;
    Panel2: TPanel;
    Button3: TButton;
    Button2: TButton;
    Button1: TButton;
    procedure TableBeforePost(DataSet: TDataSet);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure DBGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TableCalcFields(DataSet: TDataSet);
    procedure TableAgCalcFields(DataSet: TDataSet);
    procedure TableAgBeforePost(DataSet: TDataSet);
    procedure TableAgAfterScroll(DataSet: TDataSet);
    procedure TableAfterScroll(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PolisNmbsComp: TPolisNmbsComp;

implementation

{$R *.DFM}

procedure TPolisNmbsComp.TableBeforePost(DataSet: TDataSet);
begin
    if TableSTART.AsInteger > TableEND.AsInteger then
    begin
        StatusLine.Caption := 'Начальный больше конечного';
        Abort;
    end
    else
    if (TableEND.AsInteger - TableSTART.AsInteger) > 5000 then
    begin
        StatusLine.Caption := 'Больше 5000 полисов выдано';
        Abort;
    end
    else
    StatusLine.Caption := '';
end;

procedure TPolisNmbsComp.Button1Click(Sender: TObject);
begin
    Close
end;

procedure TPolisNmbsComp.Button3Click(Sender: TObject);
begin
    if TableAg.Active then
    begin
        TableAg.Delete
    end;
    if Table.Active then
    begin
        Table.Delete
    end
end;

procedure TPolisNmbsComp.Button2Click(Sender: TObject);
begin
    if TableAg.Active then
    begin
        TableAg.Append
    end;
    if Table.Active then
    begin
        Table.Append
    end
end;

procedure TPolisNmbsComp.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
    if TableAg.Active then
    begin
        CanClose := false;
        if TableAg.State in [dsEdit, dsInsert] then
            TableAg.Post;
        CanClose := true;
    end;
    if Table.Active then
    begin
        CanClose := false;
        if Table.State in [dsEdit, dsInsert] then
            Table.Post;
        CanClose := true;
    end
end;

procedure TPolisNmbsComp.FormCreate(Sender: TObject);
begin
    Agent.Open;
    try
        DataSource.DataSet := TableAg;
        TableAg.Open;
    except
        DataSource.DataSet := Table;
        Table.Open;
        DBGrid.Columns[4].Visible := false;
    end;
    WindowState := wsMaximized;
     Caption := Caption + ' ' + COMPAMYNAME;
end;

procedure TPolisNmbsComp.DBGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if TableAg.Active then
    if (Key = VK_DELETE) then
    if (DBGrid.Focused) AND (DBGrid.SelectedField = TableAgAgentCombo) then
    begin
        if TableAg.State <> dsEdit then
        begin
            TableAg.Edit
        end;
        TableAgAgentCombo.Clear;
        TableAgAgent.Clear;
        TableAg.Post
    end
end;

procedure TPolisNmbsComp.TableCalcFields(DataSet: TDataSet);
begin
    try
        TableCount.AsInteger := TableEND.AsInteger - TableSTART.AsInteger + 1
    except
    end
end;

procedure TPolisNmbsComp.TableAgCalcFields(DataSet: TDataSet);
begin
    try
        TableAgCount.AsInteger := TableAgEND.AsInteger - TableAgSTART.AsInteger + 1
    except
    end

end;

procedure TPolisNmbsComp.TableAgBeforePost(DataSet: TDataSet);
begin
    if TableAGSTART.AsInteger > TableAGEND.AsInteger then
    begin
        StatusLine.Caption := 'Начальный больше конечного';
        Abort;
    end
    else
    if (TableAGEND.AsInteger - TableAGSTART.AsInteger) > 5000 then
    begin
        StatusLine.Caption := 'Больше 5000 полисов выдано';
        Abort;
    end
    else
    StatusLine.Caption := '';
end;

procedure TPolisNmbsComp.TableAgAfterScroll(DataSet: TDataSet);
begin
    StatusLine.Caption := '';
end;

procedure TPolisNmbsComp.TableAfterScroll(DataSet: TDataSet);
begin
    StatusLine.Caption := '';
end;

end.
