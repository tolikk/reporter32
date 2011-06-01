unit MyPreview;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, qrprntr, quickrpt;

type
  TQRMyPreview = class(TForm)
    QRPreview: TQRPreview;
    Panel: TPanel;
    Button1: TButton;
    ToWidth: TButton;
    Button4: TButton;
    First: TButton;
    Next: TButton;
    Prev: TButton;
    Last: TButton;
    Button9: TButton;
    StartPage: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    LastPage: TEdit;
    ZoomCombo: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ToWidthClick(Sender: TObject);
    procedure FirstClick(Sender: TObject);
    procedure PrevClick(Sender: TObject);
    procedure NextClick(Sender: TObject);
    procedure LastClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure QRPreviewPageAvailable(Sender: TObject; PageNum: Integer);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ZoomComboChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Report : TQuickRep;
    procedure SetState;
  end;

var
  QRMyPreview: TQRMyPreview;

implementation

{$R *.DFM}

procedure TQRMyPreview.FormCreate(Sender: TObject);
begin
    DeleteMenu(GetSystemMenu(Handle, false), SC_MOVE, MF_BYCOMMAND);
    DeleteMenu(GetSystemMenu(Handle, false), SC_RESTORE, MF_BYCOMMAND);
    DeleteMenu(GetSystemMenu(Handle, false), SC_MAXIMIZE, MF_BYCOMMAND);
    DeleteMenu(GetSystemMenu(Handle, false), SC_SIZE, MF_BYCOMMAND);
    DeleteMenu(GetSystemMenu(Handle, false), SC_MINIMIZE, MF_BYCOMMAND);
    DeleteMenu(GetSystemMenu(Handle, false), 0, MF_BYPOSITION);
//    SendMessage(ZoomCombo.Handle, CB_SETITEMHEIGHT, 0, ToWidth.Height);
//    SendMessage(ZoomCombo.Handle, CB_SETITEMHEIGHT, -1, ToWidth.Height);
end;

procedure TQRMyPreview.Button1Click(Sender: TObject);
begin
  QRPreview.ZoomToFit;
  ZoomCombo.ItemIndex := -1;
end;

procedure TQRMyPreview.ToWidthClick(Sender: TObject);
begin
  QRPreview.ZoomToWidth;
  ZoomCombo.ItemIndex := -1;
end;

procedure TQRMyPreview.FirstClick(Sender: TObject);
begin
  if QRPreview.PageNumber <> 1 then begin
    QRPreview.PageNumber := 1;
  end;
  SetState;
end;

procedure TQRMyPreview.PrevClick(Sender: TObject);
begin
  if QRPreview.PageNumber > 1 then
    QRPreview.PageNumber := QRPreview.PageNumber - 1;
  SetState;
end;

procedure TQRMyPreview.NextClick(Sender: TObject);
begin
  if QRPreview.PageNumber < QRPreview.QRPrinter.PageCount then
    QRPreview.PageNumber := QRPreview.PageNumber + 1;
  SetState;
end;

procedure TQRMyPreview.LastClick(Sender: TObject);
begin
  if QRPreview.PageNumber < QRPreview.QRPrinter.PageCount then
    QRPreview.PageNumber := QRPreview.QRPrinter.PageCount;
  SetState;
end;

procedure TQRMyPreview.Button4Click(Sender: TObject);
begin
    if StrToInt(StartPage.Text) > StrToInt(LastPage.Text) then begin
       MessageDlg('Неправильные номера страниц.', mtInformation, [mbOK], 0);
       exit;
    end;
    Report.PrinterSettings.FirstPage := StrToInt(StartPage.Text);
    Report.PrinterSettings.LastPage := StrToInt(LastPage.Text);
    Report.Print;
end;

procedure TQRMyPreview.Button9Click(Sender: TObject);
begin
     Close
end;

procedure TQRMyPreview.SetState;
var
      Pages : integer;
begin
      Pages := QRPreview.QRPrinter.PageCount;
      Next.Enabled := QRPreview.PageNumber < Pages;
      Last.Enabled := QRPreview.PageNumber <> Pages;
      First.Enabled := QRPreview.PageNumber <> 1;
      Prev.Enabled := QRPreview.PageNumber > 1;
      if StartPage.Text = '' then
         StartPage.Text := '1';
      if LastPage.Text = '' then
         LastPage.Text := IntToStr(Pages);
end;

procedure TQRMyPreview.QRPreviewPageAvailable(Sender: TObject;
  PageNum: Integer);
begin
    SetState;
    QRPreview.ZoomToWidth;
end;

procedure TQRMyPreview.FormShow(Sender: TObject);
begin
    Left := Application.MainForm.Left;
    Top := Application.MainForm.Top;
    Width := Application.MainForm.Width;
    if Width < 554 then Width := 554;
    Height := Application.MainForm.Height;
end;

procedure TQRMyPreview.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
end;

procedure TQRMyPreview.ZoomComboChange(Sender: TObject);
begin
  QRPreview.Zoom := StrToInt(Copy(ZoomCombo.Text, 1, Length(ZoomCombo.Text) - 1));
end;

end.
