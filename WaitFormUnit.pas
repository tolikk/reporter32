unit WaitFormUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ComCtrls;

type
  TWaitForm = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Bevel1: TBevel;
    ProgressBar: TProgressBar;
    WorkName: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  WaitForm: TWaitForm;

implementation

{$R *.DFM}

procedure TWaitForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     Action := caFree
end;

procedure TWaitForm.FormDestroy(Sender: TObject);
begin
     WaitForm := nil;
end;

end.
