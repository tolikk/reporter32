unit Export;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, Buttons, RXSpin;

type
  TExportForm = class(TForm)
    FileName: TEdit;
    Label1: TLabel;
    Button1: TButton;
    ExpBtn: TButton;
    SaveDialog: TSaveDialog;
    BitBtn1: TBitBtn;
    Label2: TLabel;
    Monthes: TComboBox;
    Year: TRxSpinEdit;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ExportForm: TExportForm;

implementation

{$R *.DFM}

procedure TExportForm.Button1Click(Sender: TObject);
begin
     if SaveDialog.Execute then begin
        FileName.Text := SaveDialog.FileName;
        ExpBtn.Enabled := true;
     end;
end;

procedure TExportForm.FormCreate(Sender: TObject);
begin
     Monthes.ItemIndex := 0;
end;

end.
