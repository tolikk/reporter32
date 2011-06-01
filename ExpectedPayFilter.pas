unit ExpectedPayFilter;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TExpectedPaysForm = class(TForm)
    Label2: TLabel;
    DtFrom: TDateTimePicker;
    DtTo: TDateTimePicker;
    OK: TButton;
    Label1: TLabel;
    Name: TEdit;
    Label3: TLabel;
    Agent: TEdit;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ExpectedPaysForm: TExpectedPaysForm;

implementation

{$R *.dfm}

procedure TExpectedPaysForm.FormCreate(Sender: TObject);
var
    d, m, y : WORD;
begin
    DecodeDate(Date, y, m, d);
    DtFrom.Date := EncodeDate(y, m, 1);

    if m = 12 then
    begin
        m := 1;
        y := y + 1;
    end
    else
    begin
        m := m + 1;
    end;
    DtTo.Date := EncodeDate(y, m, 1) - 1;
end;

end.
