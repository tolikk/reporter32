unit ShowInfoUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TShowInfo = class(TForm)
    Button1: TButton;
    InfoPanel: TMemo;
    SaveMsg: TButton;
    SaveDialog: TSaveDialog;
    procedure Button1Click(Sender: TObject);
    procedure SaveMsgClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ShowInfo: TShowInfo;

implementation

{$R *.DFM}

procedure TShowInfo.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TShowInfo.SaveMsgClick(Sender: TObject);
var
    F : TextFile;
begin
    if not SaveDialog.Execute then exit;
    AssignFile(F, SaveDialog.FileName);
    Rewrite(F);
    Writeln(F, InfoPanel.Text);
    CloseFile(F);
end;

end.
