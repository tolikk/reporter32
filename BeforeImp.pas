unit BeforeImp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls;

type
  TBeforeImport = class(TForm)
    Label1: TLabel;
    FileName: TEdit;
    Bevel1: TBevel;
    Label2: TLabel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button1: TButton;
    OpenDialog: TOpenDialog;
    Bevel2: TBevel;
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  BeforeImport: TBeforeImport;

implementation

{$R *.DFM}

procedure TBeforeImport.BitBtn1Click(Sender: TObject);
begin
     if FileName.Text = '' then Exit;
     ModalResult := mrOk;
end;

procedure TBeforeImport.Button1Click(Sender: TObject);
begin
     if OpenDialog.Execute then
        FileName.Text := OpenDialog.FileName;
end;

end.
