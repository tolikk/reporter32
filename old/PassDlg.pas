unit PassDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, inifiles, ExtCtrls;

type
  TPasswordDlg = class(TForm)
    Label1: TLabel;
    DBName: TEdit;
    UserName: TEdit;
    Label2: TLabel;
    Password: TEdit;
    Label3: TLabel;
    Button1: TButton;
    Bevel1: TBevel;
    IsSaveData: TCheckBox;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    AutoConnectStr : string;
    function MyShowModal: Integer;
    { Public declarations }
  end;

var
  PasswordDlg: TPasswordDlg;


implementation

uses mandsql;

{$R *.DFM}

function TPasswordDlg.MyShowModal: Integer;
var
    INI : TINIFile;
    s : string;
    divPos : integer;
begin
    MandSQLForm.SQLDB.Close;

    INI := TINIFile.Create('blank.ini');
    s := INI.ReadString('MANDATORY', 'RPT_DB', 'MANDATOR/SYSADM/');
    INI.Free;

    divPos := Pos('/', s);
    if divPos <> 0 then begin
       DBName.Text := Copy(s, 1, divPos - 1);
       if AutoConnectStr <> '' then
           s := AutoConnectStr
       else
           s := Copy(s, divPos + 1, 1024);
       divPos := Pos('/', s);
       if divPos <> 0 then begin
           UserName.Text := Copy(s, 1, divPos - 1);
           Password.Text := Copy(s, divPos + 1, 1024);
       end;

       if AutoConnectStr <> '' then
           Button1Click(nil);

        if ModalResult <> mrOK then
            MyShowModal := ShowModal;

        if ModalResult = mrOK then
            MyShowModal := ModalResult;
    end;
end;

procedure TPasswordDlg.Button1Click(Sender: TObject);
var
     INI : TIniFile;
begin
     MandSQLForm.SQLDB.AliasName := DBName.Text;
     MandSQLForm.SQLDB.Params.Clear;
     MandSQLForm.SQLDB.Params.Add('USER NAME=' + UserName.Text);
     MandSQLForm.SQLDB.Params.Add('PASSWORD=' + Password.Text);
     try
         MandSQLForm.SQLDB.Open;
         if IsSaveData.Checked then begin
             INI := TINIFile.Create('blank.ini');
             INI.WriteString('MANDATORY', 'RPT_DB', DBName.Text + '/' + UserName.Text + '/');
             INI.Free;
         end;
         ModalResult := mrOK;
     except
         on E : Exception do begin
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
         end;
     end
end;

procedure TPasswordDlg.FormCreate(Sender: TObject);
begin
    if UserName.Text <> '' then
        Password.SetFocus
end;

end.
