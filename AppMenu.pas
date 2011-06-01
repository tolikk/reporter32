unit AppMenu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Buttons;

type
  TMainMenu = class(TForm)
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MainMenu: TMainMenu;
  StartParam : string;

implementation

{$R *.DFM}

procedure TMainMenu.SpeedButton2Click(Sender: TObject);
begin
    StartParam := 'POLAND';
    Close;
end;

procedure TMainMenu.SpeedButton1Click(Sender: TObject);
begin
    StartParam := '';
    Close;
end;

end.
