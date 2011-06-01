unit TmplNameUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TGetTemplName = class(TForm)
    Label1: TLabel;
    Name: TEdit;
    Button1: TButton;
    OkBtn: TButton;
    procedure OkBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetTemplName: TGetTemplName;

implementation

{$R *.DFM}

procedure TGetTemplName.OkBtnClick(Sender: TObject);
begin
     if Name.Text <> '' then ModalResult := mrOk;
end;

end.
