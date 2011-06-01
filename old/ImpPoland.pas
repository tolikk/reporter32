unit ImpPoland;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Db, DBTables, DBCtrls, RXLookup, RXSpin;

type
  TImportPoland = class(TForm)
    Label1: TLabel;
    FileName: TEdit;
    SelFile: TButton;
    Label2: TLabel;
    Agents: TTable;
    AgentsSource: TDataSource;
    Button1: TButton;
    Button2: TButton;
    AgentsAgent_code: TStringField;
    AgentsName: TStringField;
    CBAgents: TRxDBLookupCombo;
    OpenDialog: TOpenDialog;
    AgPercent: TRxSpinEdit;
    Label3: TLabel;
    IsUridich: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SelFileClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ImportPoland: TImportPoland;

implementation

{$R *.DFM}

procedure TImportPoland.Button1Click(Sender: TObject);
begin
     if FileName.Text = '' then
         exit;

     if CBAgents.Text = '' then
         exit;

     ModalResult := mrOk;
end;

procedure TImportPoland.FormCreate(Sender: TObject);
begin
     Agents.Open
end;

procedure TImportPoland.SelFileClick(Sender: TObject);
begin
     if OpenDialog.Execute then
         FileName.Text := OpenDialog.FileName;
end;

end.
