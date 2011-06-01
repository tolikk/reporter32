unit GetPeriodOnly;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TGetPeriodForm = class(TForm)
    Label2: TLabel;
    DtFrom: TDateTimePicker;
    DtTo: TDateTimePicker;
    btnOk: TButton;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetPeriodForm: TGetPeriodForm;

implementation

{$R *.dfm}

procedure TGetPeriodForm.FormCreate(Sender: TObject);
var
    d, m, y : WORD;
begin
    DecodeDate(Date, y, m, d);
    DtTo.Date := EncodeDate(y, m, 1) - 1;
    if m > 1 then
    begin
        m := m - 1
    end
    else
    begin
        m := 12;
        y := y - 1;
    end;
    DtFrom.Date := EncodeDate(y, m, 1);
end;

end.
