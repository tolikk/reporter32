unit AddFld;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  RXSpin, StdCtrls, ActnList, dbtables, Mask;

type
  TAddFldForm = class(TForm)
    Label1: TLabel;
    TableName: TEdit;
    lbNameFld: TLabel;
    NameFld: TEdit;
    Label2: TLabel;
    FldType: TComboBox;
    lbSizeFld: TLabel;
    SizeFld: TRxSpinEdit;
    Ok: TButton;
    ActionList: TActionList;
    FldSize: TAction;
    procedure FormCreate(Sender: TObject);
    procedure FldSizeUpdate(Sender: TObject);
    procedure OkClick(Sender: TObject);
  private
    { Private declarations }
  public
    Table : TTable;
    DataTypeName : string;
    { Public declarations }
  end;

var
  AddFldForm: TAddFldForm;

implementation

{$R *.DFM}

procedure TAddFldForm.FormCreate(Sender: TObject);
begin
    FldType.ItemIndex := 0;
end;

procedure TAddFldForm.FldSizeUpdate(Sender: TObject);
begin
    SizeFld.Enabled := FldType.ItemIndex = 0;
    lbSizeFld.Enabled := FldType.ItemIndex = 0;
    Ok.Enabled := (Table.TableLevel > 4) OR not (FldType.ItemIndex in [1,2]);   
end;

procedure TAddFldForm.OkClick(Sender: TObject);
var
    i : integer;
begin
    NameFld.Text := Trim(NameFld.Text);
    if NameFld.Text = '' then exit;
    if Pos(' ', NameFld.Text) <> 0 then exit;
    for i := 0 to Table.Fields.Count - 1 do
        if UpperCase(Table.Fields[i].FieldName) = UpperCase(NameFld.Text) then begin
            MessageDlg('Поле с таким именем есть в таблице', mtInformation, [mbOk], 0);
            exit;
        end;
    case FldType.ItemIndex of
    0 : DataTypeName := 'CHAR(' + SizeFld.Text + ')';
    1 : DataTypeName := 'SMALLINT';
    2 : DataTypeName := 'INTEGER';
    3 : DataTypeName := 'FLOAT';
    4 : DataTypeName := 'DATE';
    end;

    ModalResult := mrOK;
end;

end.
