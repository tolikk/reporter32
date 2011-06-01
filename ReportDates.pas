unit ReportDates;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls;

type
  TGetDates = class(TForm)
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    FromDate: TDateTimePicker;
    ToDate: TDateTimePicker;
    StaticText3: TStaticText;
    Quartal: TComboBox;
    Button1: TButton;
    Button2: TButton;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure SetQuartal(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetDates: TGetDates;

implementation

{$R *.DFM}

procedure TGetDates.FormCreate(Sender: TObject);
var
     y, m, d : WORD;
begin
     DecodeDate(Date, y, m, d);
     Quartal.ItemIndex := TRUNC(m / 4);
     if Quartal.ItemIndex = 0 then
        Quartal.ItemIndex := 3
     else
        Quartal.ItemIndex := Quartal.ItemIndex - 1;

     SetQuartal(nil);
end;

procedure TGetDates.SetQuartal(Sender: TObject);
var
     Year : TDateTime;
     y, m, d : WORD;
begin
     DecodeDate(Date, y, m, d);
     if Quartal.Itemindex < 4 then begin
         FromDate.Date := EncodeDate(y, Quartal.Itemindex * 3 + 1, 1);
         ToDate.Date := IncMonth(EncodeDate(y, Quartal.Itemindex * 3 + 1, 1), 3) - 1;
     end
     else begin
         FromDate.Date := EncodeDate(y, (Quartal.Itemindex - 4) + 1, 1);
         ToDate.Date := IncMonth(EncodeDate(y, (Quartal.Itemindex - 4) + 1, 1), 1) - 1;
     end;
end;

end.
