unit UnitImportPx;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, RXSpin, RxLookup, DB, DBTables, RxQuery;

type
  TformImportPx = class(TForm)
    Button1: TButton;
    Label1: TLabel;
    Label2: TLabel;
    AgPercent: TRxSpinEdit;
    DataSource: TDataSource;
    Agent: TRxDBLookupCombo;
    txtFilter: TEdit;
    Agents: TRxQuery;
    procedure Button1Click(Sender: TObject);
    procedure AgentChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure txtFilterChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  formImportPx: TformImportPx;

implementation

{$R *.dfm}

procedure TformImportPx.Button1Click(Sender: TObject);
begin
    if Agent.Text = '' then exit;
    if AgPercent.Value <= 0 then exit;
    ModalResult := mrOk
end;

procedure TformImportPx.AgentChange(Sender: TObject);
begin
    AgPercent.Value := Agents.FieldByName('MANDPCNT').AsFloat
end;

procedure TformImportPx.FormCreate(Sender: TObject);
begin
    Agents.Open 
end;

procedure TformImportPx.txtFilterChange(Sender: TObject);
begin
    Agents.Close;
    Agents.MacroByName('FILTER').AsString := '''%' + txtFilter.Text + '%''';
    Agents.Open
end;

end.
