unit TechCalc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TTechCalculation = class(TForm)
    Label1: TLabel;
    TextData: TMemo;
    Button1: TButton;
    GroupBox1: TGroupBox;
    Label3: TLabel;
    Nalog: TEdit;
    Label2: TLabel;
    Premium1: TEdit;
    Label4: TLabel;
    Premium2: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure Premium1Change(Sender: TObject);
  private
    { Private declarations }
  public
    Premium : double;
    Good, Bad : integer;
    Period : string;
    { Public declarations }
  end;

var
  TechCalculation: TTechCalculation;

implementation

uses mf, datamod;

{$R *.DFM}

procedure TTechCalculation.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TTechCalculation.Premium1Change(Sender: TObject);
begin
     TextData.Lines.Clear;

     try
       if Period <> '' then begin
           TextData.Lines.Add('�� ������ ' + Period);
           TextData.Lines.Add('');
       end;
       TextData.Lines.Add('��������� ������ ������ 100%            ' + FloatToStr(Premium));
       TextData.Lines.Add('��������� ������ ��� ' + COMPAMYNAME + ' (' + Premium1.Text + '%)' + Copy('                  ', 1, 13 - Length(COMPAMYNAME)) + FloatToStr(Premium / 100 * StrToFloat(Premium1.Text)));
       TextData.Lines.Add('��������� ������ ��� ����������� (' + Premium2.Text + '%)  ' + FloatToStr(Premium / 100 * StrToFloat(Premium2.Text)));
       TextData.Lines.Add('����� �� ����� ����������� (' + Nalog.Text + '%)        ' + FloatToStr(Premium / 100 * StrToFloat(Premium2.Text) * (100 - StrToFloat(Nalog.Text)) / 100));
       TextData.Lines.Add('');
       TextData.Lines.Add('����������� �������');
       TextData.Lines.Add('    �����     ' + IntToStr(Good + Bad));
       TextData.Lines.Add('    ��������  ' + IntToStr(Good));
       TextData.Lines.Add('    ��������� ' + IntToStr(Bad));
     except
     end
end;

end.
