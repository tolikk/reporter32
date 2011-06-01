unit Rep2C;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, inifiles, math, dbtables, ComCtrls, Db, datamod, Clipbrd,
  Menus;

type
  TForm2C = class(TForm)
    SpeedButton1: TSpeedButton;
    Panel: TPanel;
    ProgressBar: TProgressBar;
    DatabaseRpt: TDatabase;
    WorkSQL: TQuery;
    MANDRPT: TTable;
    Timer: TTimer;
    SpeedButton2: TSpeedButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure PanelResize(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure ExecReport(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
  private
    { Private declarations }
    Buttons : TList;
    MSEXCEL : VARIANT;
    procedure AlignButtons;
    function  fExecReport(Sender: TObject) : boolean;
  public
    IsServer : boolean;
    { Public declarations }
  end;

var
  Form2C: TForm2C;

implementation

uses CurrRatesUnit;

{$R *.DFM}

procedure TForm2C.ExecReport(Sender: TObject);
begin
    fExecReport(Sender);
end;

function TForm2C.fExecReport(Sender: TObject) : boolean;
var
    s : string;
begin
    fExecReport := true;
    try
        Screen.Cursor := crHourglass;
        s := Caption;
        Caption := s + ' Думаю...' ;
        OpenExcel2C(MSEXCEL);
        WorkSQL.Close;
        WorkSQL.DatabaseName := DatabaseRpt.DatabaseName;
        RunReport(MSEXCEL, TSpeedButton(Sender).Tag, WorkSQL);
        TSpeedButton(Sender).Font.Color := clRed;
    except
      on E : Exception do begin
        Application.BringToFront;
        Clipboard.AsText := E.Message;
        MessageDlg(E.Message, mtError, [mbOk], 0);
        fExecReport := false;
      end;
    end;
    Caption := s;
    Screen.Cursor := crDefault;
    Application.BringToFront;
end;

procedure TForm2C.FormCreate(Sender: TObject);
var
    Name : string;
    i, n : integer;
    Item : TSpeedButton;
    INI : TTolikIniFile;
    IsFlat : boolean;
begin
    Buttons := TList.Create;
    INI := TTolikIniFile.Create('reports.ini');
    DatabaseRpt.AliasName := INI.ReadString('DB', 'Alias', 'BASO');
    IsServer := INI.ReadString('DB', 'DBTYPE', '') = 'SERVER';
    if IsServer then MANDRPT.TableName := 'SYSADM.' + MANDRPT.TableName;

    n := 0;
    IsFlat := true;//((GetCurrentTime() DIV 1000) MOD 2) = 0;
    for i := 1 to 100 do begin
        Name := INI.ReadString('REPORT' + IntToStr(i), 'Name', '');
        if Name <> '' then begin
            Item := TSpeedButton.Create(Panel);
            Item.Tag := i;                          
            Item.OnClick := ExecReport;
            Item.Caption := Name;
            Buttons.Add(Item);
            Item.Flat := IsFlat;
            Item.Transparent := true;
            Item.Visible := true;
            Item.Hint := 'Сформировать отчёт ' + Name;
            Item.Parent := Panel;
            TSpeedButton(Sender).Font.Color := clBlue;
            TSpeedButton(Sender).Font.Style := TSpeedButton(Sender).Font.Style + [fsUnderline];
            n := n + 1;
        end;
    end;
    AlignButtons;
    INI.Free;
end;

procedure TForm2C.AlignButtons;
var
    H, W, i, j : integer;
begin
    W := Panel.ClientWidth DIV 70;
    H := Panel.ClientHeight DIV 20;
    for i := 0 to W-1 do
       for j := 0 to H-1 do begin
           if i * W + j >= Buttons.Count then break;
           TSpeedButton(Buttons[i * W + j]).Left := 70 * j;
           TSpeedButton(Buttons[i * W + j]).Top := 20 * i;
           TSpeedButton(Buttons[i * W + j]).Width := 70;
           TSpeedButton(Buttons[i * W + j]).Height := 20;
       end;
end;

procedure TForm2C.FormDestroy(Sender: TObject);
begin
    Buttons.Free
end;

procedure TForm2C.PanelResize(Sender: TObject);
begin
    AlignButtons;
end;

procedure TForm2C.SpeedButton1Click(Sender: TObject);
begin
    Close;
end;

procedure TForm2C.TimerTimer(Sender: TObject);
begin
    Timer.Enabled := false;
    if IsServer then
        if MessageDlg('Вы хотите заполнить таблицу видов машин по ОСТС?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
    if not PrepareHiLo(MandRpt, WorkSQL, IsServer, ProgressBar) then Close;
    ProgressBar.Visible := false;
end;

procedure TForm2C.SpeedButton2Click(Sender: TObject);
var
    i : integer;
begin
    for i := 0 to Buttons.Count - 1 do
        if not fExecReport(Buttons[i]) then break;
end;

procedure TForm2C.N1Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

end.
