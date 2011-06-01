unit GetTwoDatesUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TGetTwoDates = class(TForm)
    Label1: TLabel;
    StartDate: TDateTimePicker;
    Label2: TLabel;
    EndDate: TDateTimePicker;
    Button1: TButton;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetTwoDates: TGetTwoDates;

implementation

{$R *.dfm}

procedure TGetTwoDates.FormCreate(Sender: TObject);
var
    Y, M, D : WORD;
begin
    DecodeDate(Date - 30, Y, M, D);
    StartDate.Date := EncodeDate(Y, M, 1);
    M := M + 1;
    if M = 13 then begin
        M := 1;
        Y := Y + 1;
    end;
    EndDate.Date := EncodeDate(Y, M, 1) - 1;
end;

end.
