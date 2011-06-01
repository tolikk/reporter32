unit Move_word;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TMoveWord = class(TForm)
    Label1: TLabel;
    Word: TEdit;
    Label2: TLabel;
    Direction: TComboBox;
    Button1: TButton;
    OKBTN: TButton;
    Label3: TLabel;
    Filter: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure WordChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  MoveWord: TMoveWord;

implementation

{$R *.DFM}

procedure TMoveWord.FormCreate(Sender: TObject);
begin
     Direction.ItemIndex := 0;
end;

procedure TMoveWord.WordChange(Sender: TObject);
begin
     OKBTN.Enabled := Length(Trim(Word.Text)) > 0;
end;

end.
