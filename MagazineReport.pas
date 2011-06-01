unit MagazineReport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls;

type
  TMagazineForm = class(TForm)
    Label1: TLabel;
    RepDate: TDateTimePicker;
    Label2: TLabel;
    OtvSumma: TEdit;
    OtvCurr: TEdit;
    btnOk: TButton;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MagazineForm: TMagazineForm;

implementation

{$R *.DFM}

end.
