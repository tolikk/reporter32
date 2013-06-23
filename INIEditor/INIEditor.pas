unit INIEditor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, ComCtrls, Grids;

const
   _Import_str = 'Импортная';
   _NoImport_str = 'Отечественная';

type
  TMainForm = class(TForm)
    DownPanel: TPanel;
    Button1: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    AutoMarkGrid: TStringGrid;
    SplitterAutoMark: TSplitter;
    AutoSubMarkGrid: TStringGrid;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TabSheet1Resize(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AutoMarkGridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure AutoMarkGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure AutoMarkGridSetEditText(Sender: TObject; ACol, ARow: Integer;
      const Value: String);
    procedure AutoMarkGridGetEditText(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure AutoSubMarkGridSetEditText(Sender: TObject; ACol,
      ARow: Integer; const Value: String);
    procedure AutoSubMarkGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure AutoSubMarkGridGetEditText(Sender: TObject; ACol,
      ARow: Integer; var Value: String);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    StartEditText : string;
    BlankWnd : HWND;
    AutoMarkFile : string;
    function SaveAMToIni(ARow, ARow2 : integer; s : string) : string;
    procedure SavePToIni(ARow : integer; s : string);
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

uses inifiles;

{$R *.DFM}

procedure TMainForm.Button1Click(Sender: TObject);
begin
    Close
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TMainForm.TabSheet1Resize(Sender: TObject);
begin
    AutoMarkGrid.ColWidths[0] := AutoMarkGrid.Width DIV 2 - 10;
    AutoMarkGrid.ColWidths[1] := AutoMarkGrid.Width DIV 2 - 10;
    AutoSubMarkGrid.ColWidths[0] := AutoSubMarkGrid.Width;
end;

procedure TMainForm.FormCreate(Sender: TObject);
var
    INI : TINIFile;
    i : integer;
    CanSelect : boolean;
    R : TRect;
    s : string;
begin
    AutoMarkFile := '';

    BlankWnd := 0;
    INI := TINIFile.Create('win.ini');
    s := INI.ReadString('Paradox Engine', 'NetNamePath', '');
    INI.Free;
    if s <> '' then s := s + '\blank.ini';
    if not FileExists(s) then s := 'blank.ini';
    INI := TINIFile.Create(s);
    AutoMarkFile := INI.ReadString('MANDATORY', 'AutoMarkFile', s);

    INI.Free;
    INI := TINIFile.Create(AutoMarkFile);

    for i := 0 to 1000 do begin
        AutoMarkGrid.Cells[0, i] := INI.ReadString('MANDATORY', 'AutoType' + IntToStr(i), '');
        if AutoMarkGrid.Cells[0, i] <> '' then
            AutoMarkGrid.RowCount := AutoMarkGrid.RowCount + 1
        else
            break;
        case INI.ReadInteger('MANDATORY', 'AutoType' + IntToStr(i) + 'IMP', -1) of
        1 : AutoMarkGrid.Cells[1, i] := _Import_str;
        0 : AutoMarkGrid.Cells[1, i] := _NoImport_str;
        else AutoMarkGrid.Cells[1, i] := 'Не указано';
        end;
    end;
    INI.Free;
    if ParamStr(1) = '16' then begin
        BlankWnd := FindWindow('OWLWindow31', nil);
        if BlankWnd <> 0 then begin
            EnableWindow(BlankWnd, false);
            GetWindowRect(BlankWnd, R);
            Left := R.Left;
            Top := R.TOp;
            Width := R.Right - R.Left;
            Height := R.Bottom - R.Left;
        end;
    end;
    AutoMarkGridSelectCell(nil, 0, 0, CanSelect);
end;

procedure TMainForm.AutoMarkGridSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
var
    INI : TINIFile;
    i : integer;
begin
    CanSelect := true;//not (goEditing in AutoMarkGrid.Options);
//    if not CanSelect then exit;
//    if AutoMarkGrid.Cells[ACol, ARow] <> '' then begin
        AutoSubMarkGrid.RowCount := 1;
        INI := TINIFile.Create(AutoMarkFile);
        for i := 0 to 1000 do begin
            AutoSubMarkGrid.Cells[0, i] := INI.ReadString('MANDATORY', 'AutoType' + IntToStr(ARow) + '_' +  IntToStr(i), '');
            if AutoSubMarkGrid.Cells[0, i] <> '' then
                AutoSubMarkGrid.RowCount := AutoSubMarkGrid.RowCount + 1
            else
                break;
        end;
        INI.Free;
//    end;
end;

procedure TMainForm.AutoMarkGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    i, j : integer;
    s : string;
    INI : TINIFile;
    CanSelect : boolean;
begin
    if Key = VK_DELETE then begin
        if AutoMarkGrid.Row = AutoMarkGrid.RowCount - 1 then exit;
        if MessageDlg('Вы хотите удалить марку ' + AutoMarkGrid.Cells[AutoMarkGrid.Col, AutoMarkGrid.Row], mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;

        Screen.Cursor := crHourglass;
        for i := AutoMarkGrid.Row to AutoMarkGrid.RowCount - 2 do begin
            AutoMarkGrid.Cells[0, i] := AutoMarkGrid.Cells[0, i + 1];
            AutoMarkGrid.Cells[0, i] := SaveAMToIni(i, -1, AutoMarkGrid.Cells[0, i]);
            AutoMarkGrid.Cells[1, i] := AutoMarkGrid.Cells[1, i + 1];
            SavePToIni(i, AutoMarkGrid.Cells[1, i]);
            //Перезаписать подмарки
            INI := TINIFile.Create(AutoMarkFile);
            for j := 0 to 1000 do begin
                s := INI.ReadString('MANDATORY', 'AutoType' + IntToStr(i + 1) + '_' + IntToStr(j), '');
                SaveAMToIni(i, j, s);
            end;
            INI.Free;
            AutoMarkGridSelectCell(nil, 0, AutoMarkGrid.Row, CanSelect);
        end;
        AutoMarkGrid.RowCount := AutoMarkGrid.RowCount - 1;
        Screen.Cursor := crDefault;
    end;
    if (Key = VK_ESCAPE) AND AutoMarkGrid.EditorMode then begin
        AutoMarkGrid.Cells[AutoMarkGrid.Col, AutoMarkGrid.Row] := StartEditText;
    end;
end;

procedure TMainForm.AutoMarkGridSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: String);
begin
     if not AutoMarkGrid.EditorMode then begin
         if ACol = 1 then begin
            if Value = '1' then AutoMarkGrid.Cells[ACol, ARow] := _Import_str
            else AutoMarkGrid.Cells[ACol, ARow] := _NoImport_str;
            SavePToIni(ARow, AutoMarkGrid.Cells[ACol, ARow]);
            exit;
         end;
         if Value = '' then begin
            if (ARow + 1) = AutoMarkGrid.RowCount then exit;
            AutoMarkGrid.Cells[ACol, ARow] := 'ПУСТО';
         end;
         AutoMarkGrid.Cells[ACol, ARow] := SaveAMToIni(ARow, -1, AutoMarkGrid.Cells[ACol, ARow]);
         if (ARow + 1) = AutoMarkGrid.RowCount then begin
             AutoMarkGrid.RowCount := AutoMarkGrid.RowCount + 1;
             AutoMarkGrid.Cells[0, AutoMarkGrid.RowCount - 1] := '';
             AutoMarkGrid.Cells[1, AutoMarkGrid.RowCount - 1] := '';
         end;
     end;
end;

procedure TMainForm.SavePToIni(ARow : integer; s : string);
var
    INI : TINIFile;
begin
     INI := TINIFile.Create(AutoMarkFile);
     if s = _Import_str then
         INI.WriteInteger('MANDATORY',  'AutoType' + IntToStr(ARow) + 'IMP', 1)
     else
         INI.WriteInteger('MANDATORY',  'AutoType' + IntToStr(ARow) + 'IMP', 0);
    INI.Free;
end;

function TMainForm.SaveAMToIni(ARow, ARow2 : integer; s : string) : string;
var
    INI : TINIFile;
    bad_s : string;
    i : integer;
begin
    bad_s := '.:-"%$*=';
    s := AnsiUpperCase(s);
    for i := 1 to Length(s) do begin
        //if (s[i] >= 'A') AND (s[i] <= 'Z') then
        //    s[i] := 'Я';
        if Pos(s[i], bad_s) <> 0 then
            s[i] := 'Я';
    end;
    s := Trim(s);

    INI := TINIFile.Create(AutoMarkFile);
    if ARow2 = -1 then begin
        if s = '' then
            INI.DeleteKey('MANDATORY',  'AutoType' + IntToStr(ARow))
        else
            INI.WriteString('MANDATORY',  'AutoType' + IntToStr(ARow), s)
    end
    else begin
        if s = '' then
            INI.DeleteKey('MANDATORY',  'AutoType' + IntToStr(ARow) + '_' + IntToStr(ARow2))
        else
            INI.WriteString('MANDATORY',  'AutoType' + IntToStr(ARow) + '_' + IntToStr(ARow2), s);
    end;
    INI.Free;

    SaveAMToIni := s;
end;

procedure TMainForm.AutoMarkGridGetEditText(Sender: TObject; ACol,
  ARow: Integer; var Value: String);
begin
    if ACol = 1 then begin
        if Value = _Import_str then Value := '1'
        else Value := '0'
    end
    else
        StartEditText := Value;
end;

procedure TMainForm.AutoSubMarkGridSetEditText(Sender: TObject; ACol,
  ARow: Integer; const Value: String);
begin
     if not AutoSubMarkGrid.EditorMode then begin
         if Value = '' then begin
            if (ARow + 1) = AutoSubMarkGrid.RowCount then exit;
            AutoSubMarkGrid.Cells[ACol, ARow] := 'ПУСТО';
         end;
         AutoSubMarkGrid.Cells[ACol, ARow] := SaveAMToIni(AutoMarkGrid.Row, ARow, AutoSubMarkGrid.Cells[ACol, ARow]);
         if (ARow + 1) = AutoSubMarkGrid.RowCount then begin
             AutoSubMarkGrid.RowCount := AutoSubMarkGrid.RowCount + 1;
             AutoSubMarkGrid.Cells[0, AutoSubMarkGrid.RowCount - 1] := '';
         end;
     end;
end;

procedure TMainForm.AutoSubMarkGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    i : integer;
begin
    if (Key = VK_ESCAPE) AND AutoSubMarkGrid.EditorMode then begin
        AutoSubMarkGrid.Cells[AutoSubMarkGrid.Col, AutoSubMarkGrid.Row] := StartEditText;
    end;
    if Key = VK_DELETE then begin
        if AutoSubMarkGrid.Row = AutoSubMarkGrid.RowCount - 1 then exit;
        if MessageDlg('Вы хотите удалить ' + AutoSubMarkGrid.Cells[AutoSubMarkGrid.Col, AutoSubMarkGrid.Row], mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;
        for i := AutoSubMarkGrid.Row to AutoSubMarkGrid.RowCount - 2 do begin
            AutoSubMarkGrid.Cells[0, i] := AutoSubMarkGrid.Cells[0, i + 1];
            SaveAMToIni(AutoMarkGrid.Row, i, AutoSubMarkGrid.Cells[0, i]);
        end;
        AutoSubMarkGrid.RowCount := AutoSubMarkGrid.RowCount - 1;
    end;
end;

procedure TMainForm.AutoSubMarkGridGetEditText(Sender: TObject; ACol,
  ARow: Integer; var Value: String);
begin
    StartEditText := Value;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
    if BlankWnd <> 0 then begin
        EnableWindow(BlankWnd, true);
        BringWindowToTop(BlankWnd);
    end;
end;

end.
