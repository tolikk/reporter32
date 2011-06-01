unit MagazLeben;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  InpRezParamsFor, Placemnt, Grids, StdCtrls, RXSpin, ComCtrls, ExtCtrls, math,
  Mask;

type
  TMagazineLeben = class(TInitRezervLeben)
    procedure CalcBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MagazineLeben: TMagazineLeben;

implementation

{$R *.DFM}

procedure TMagazineLeben.CalcBtnClick(Sender: TObject);
begin
    ModalResult := mrOK;
end;

end.
