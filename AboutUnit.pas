unit AboutUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Animate, GIFCtrl;

type
  TAbout = class(TForm)
    Title: TLabel;
    Company: TLabel;
    Button1: TButton;
    lblINI: TLabel;
    procedure FormShow(Sender: TObject);
    procedure lblINIClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  About: TAbout;

implementation

uses mf, datamod, RXShell;

{$R *.DFM}

procedure TAbout.FormShow(Sender: TObject);
var
     hFile : integer;
     Dt : TDateTime;
begin
     Company.Caption := 'Разработка'#13'Костюка Анатолия для ' + COMPAMYNAME + #13'1999-2011'#13'(последеноминационная)'#13#13'мобильный +375 (29) 5567895'#13'e-mail kaacpp@mail.ru';

     Dt := FileDateToDateTime(FileAge(Application.ExeName));
     Company.Caption := Company.Caption + #13#13'Файл от ' + DateToStr(DT);
     lblINI.Caption := BLANK_INI;
end;

procedure TAbout.lblINIClick(Sender: TObject);
var
    path : array[0..255] of char;
begin
    GetWindowsDirectory(path, 255);
    FileExecute('NOTEPAD.EXE', lblINI.Caption, path, esNormal);
    Close
end;

end.
