unit ProgressDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls;

type
  TProgress = class(TForm)
    ProgressBar: TProgressBar;
    ProcessLabel: TLabel;
    Bevel1: TBevel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Progress: TProgress;

implementation

{$R *.DFM}

procedure TProgress.FormShow(Sender: TObject);
begin
     ProgressBar.Position := 0;
end;

end.
