unit GetFormulaUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TGetFormula = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    Formula: TEdit;
    Button1: TButton;
    Button2: TButton;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetFormula: TGetFormula;

implementation

{$R *.DFM}

procedure TGetFormula.FormShow(Sender: TObject);
begin
     if Formula.Text = '' then
        Formula.Text := Format('* 6000 * %g', [555.55]);
end;

end.
