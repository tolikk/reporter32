unit GetRepDateUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Mask, RXSpin;

type
  TGetRepDate = class(TForm)
    Label1: TLabel;
    RepDate: TDateTimePicker;
    Button1: TButton;
    Label2: TLabel;
    Label3: TLabel;
    FizTax: TRxSpinEdit;
    UrTax: TRxSpinEdit;
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetRepDate: TGetRepDate;

implementation

{$R *.dfm}

procedure TGetRepDate.FormCreate(Sender: TObject);
var
    D, M, Y : WORD;
begin
    DecodeDate(Date, Y, M, D);
    RepDate.Date := EncodeDate(Y, M, 1);
end;

end.
