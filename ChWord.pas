unit ChWord;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DataMod;

type
  TChWordsDlg = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Word2: TEdit;
    Button1: TButton;
    Button2: TButton;
    Word1: TComboBox;
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Word1Change(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    _2Part : TStringList;
  public
    { Public declarations }
  end;

var
  ChWordsDlg: TChWordsDlg;

implementation

uses inifiles;

{$R *.DFM}

procedure TChWordsDlg.Button2Click(Sender: TObject);
begin
     if Word1.Text = Word2.Text then
        exit;
     if Word1.Text = '' then
        exit;
     ModalResult := mrOk;
end;

procedure TChWordsDlg.FormCreate(Sender: TObject);
var
     INI : TIniFile;
     i : integer;
     s : string;
begin
     _2Part := TStringList.Create;
     INI := TIniFile.Create(BLANK_INI);
     for i := 0 to 100 do begin
         s := INI.ReadString('MANDATORY', 'Replace' + IntToStr(i), '');
         if s = '' then continue;
         if Pos(' | ', s) = 0 then continue;
         Word1.Items.Add(Copy(s, 1,  Pos(' | ', s) - 1));
         _2Part.Add(Copy(s, Pos(' | ', s) + 3, 1024));
     end;
     INI.Free;
end;

procedure TChWordsDlg.Word1Change(Sender: TObject);
begin
     if Word1.ItemIndex <> -1 then
         Word2.Text := _2Part[Word1.ItemIndex];
end;

procedure TChWordsDlg.FormDestroy(Sender: TObject);
begin
     _2Part.Free;
end;

end.
