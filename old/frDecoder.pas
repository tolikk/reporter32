unit frDecoder;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TDecoderFrm = class(TForm)
    Kode: TEdit;
    Result: TEdit;
    Button1: TButton;
    Kode2: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure KodeChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DecoderFrm: TDecoderFrm;

implementation

{$R *.DFM}

uses mandsql;

procedure TDecoderFrm.Button1Click(Sender: TObject);
begin
    Close
end;

procedure TDecoderFrm.KodeChange(Sender: TObject);
begin
    Result.Text := DecodeMsgId(Kode.Text);
    Kode2.Text := CodeMsgId(DecodeMsgId(Kode.Text));
end;

end.
