unit GurnalUbitokUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, Placemnt;

type
  TFilterForUbitok = class(TForm)
    Button2: TButton;
    Label1: TLabel;
    DateFrom: TDateTimePicker;
    DateTo: TDateTimePicker;
    Label2: TLabel;
    Label3: TLabel;
    InsSumma: TEdit;
    ValCode: TEdit;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    CheckedDate1, CheckedDate2 : boolean;
    { Public declarations }
  end;

var
  FilterForUbitok: TFilterForUbitok;

implementation

{$R *.DFM}

procedure TFilterForUbitok.FormShow(Sender: TObject);
begin
     DateFrom.Checked := CheckedDate1;
     DateTo.Checked := CheckedDate2;
end;

procedure TFilterForUbitok.FormCreate(Sender: TObject);
begin
     CheckedDate1 := true;
     CheckedDate2 := true;
end;

end.
