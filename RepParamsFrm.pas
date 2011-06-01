unit RepParamsFrm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TRepParams = class(TForm)
    Label2: TLabel;
    DtFrom: TDateTimePicker;
    DtTo: TDateTimePicker;
    Button1: TButton;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  RepParams: TRepParams;

implementation

{$R *.dfm}

end.
