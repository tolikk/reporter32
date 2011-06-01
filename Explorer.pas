unit Explorer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Buttons, RXSpin, StdCtrls, Placemnt, Mask;

type
  TExplore = class(TForm)
    Seria: TEdit;
    Number: TRxSpinEdit;
    SpeedButton1: TSpeedButton;
    Label1: TLabel;
    Label2: TLabel;
    Draft1Number: TRxSpinEdit;
    SpeedButton2: TSpeedButton;
    FormStorageExplore: TFormStorage;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Explore: TExplore;

implementation

uses MF;

{$R *.DFM}

procedure TExplore.SpeedButton1Click(Sender: TObject);
begin
     try
        MainForm.StateText.Caption := 'Поиск полиса...';
        MAINForm.FindSN(Seria.Text, ROUND(Number.Value));
     except
     end;
     MainForm.StateText.Caption := '';
end;

procedure TExplore.SpeedButton2Click(Sender: TObject);
begin
     try
        MainForm.StateText.Caption := 'Поиск платёжки...';
        MAINForm.FindDraft1Number(ROUND(Draft1Number.Value));
     except
     end;
     MainForm.StateText.Caption := '';
end;

end.
