unit PcntFormUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, RXSpin, ComCtrls, Placemnt;

type
  TPcntForm = class(TForm)
    btnOk: TButton;
    btnCancel: TButton;
    Label1: TLabel;
    Date: TDateTimePicker;
    Label2: TLabel;
    Label3: TLabel;
    PcntBefore: TRxSpinEdit;
    PcntAfter: TRxSpinEdit;
    FormStorage: TFormStorage;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  PcntForm: TPcntForm;

implementation

{$R *.dfm}

end.
