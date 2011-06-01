unit ExpErrsUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, Mask, ToolEdit, Grids, DBGrids, RxQuery;

type
  TExpErrsFlag = class(TForm)
    Panel1: TPanel;
    CloseBtn: TButton;
    SaveBtn: TButton;
    SaveDialog: TSaveDialog;
    DataSource3: TDataSource;
    DataSource4: TDataSource;
    Panel2: TPanel;
    Label1: TLabel;
    DirectoryEdit: TDirectoryEdit;
    ExecBtn: TButton;
    Panel3: TPanel;
    DBGrid3: TDBGrid;
    Panel4: TPanel;
    LabelLeft3: TLabel;
    RxQuery3: TRxQuery;
    Splitter1: TSplitter;
    DBGrid4: TDBGrid;
    RxQuery4: TRxQuery;
    WorkSQL: TQuery;
    procedure CloseBtnClick(Sender: TObject);
    procedure ExecBtnClick(Sender: TObject);
    procedure SaveBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ExpErrsFlag: TExpErrsFlag;

implementation

{$R *.DFM}

procedure TExpErrsFlag.CloseBtnClick(Sender: TObject);
begin
    Close
end;

procedure TExpErrsFlag.ExecBtnClick(Sender: TObject);
var
    s, s1, s2, fs, Table3 : string;
    i : integer;
    sr : TSearchRec;
    FileAttrs : Integer;
begin
    s := '';
    s1 := '';
    s2 := '';
    FileAttrs := faAnyFile + faArchive;
    try
        fs := DirectoryEdit.Text;
        if fs = '' then exit;
        if fs[Length(fs)] <> '\' then fs := fs + '\';
        if FindFirst(fs + 'N3*.dbf', FileAttrs, sr) <> 0 then begin
            MessageDlg('Файл N3 не найден', mtInformation, [mbOk], 0);
            exit;
        end;
        FindClose(sr);
        Table3 := sr.Name;
        Screen.Cursor := crHourglass;
        with WorkSQL, WorkSQL.SQL do begin
            Clear;
            i := 0;
            Add('SELECT MSG_ID FROM ''' + fs + Table3 + ''' GROUP BY MSG_ID HAVING COUNT(*) > 1 ORDER BY MSG_ID');
            Open;
            while not EOF do begin
                if (i >= 0) AND (i < 250) then begin
                    if s <> '' then s := s + ',';
                    s := s + '''' + FieldByName('MSG_ID').AsString + '''';
                end;
                if (i >= 250) AND (i < 500) then begin
                    if s1 <> '' then s1 := s1 + ',';
                    s1 := s1 + '''' + FieldByName('MSG_ID').AsString + '''';
                end;
                if (i >= 500) AND (i < 750) then begin
                    if s2 <> '' then s2 := s2 + ',';
                    s2 := s2 + '''' + FieldByName('MSG_ID').AsString + '''';
                end;
                Next;
                Inc(i);
                if i > 750 then begin
                    MessageDlg('Показаны не все ошибки! Исправите эти - попробуете снова!', mtInformation, [mbOk], 0);
                    break;
                end;
            end;
            Close;
        end;

        if s <> '' then begin
            if s1 = '' then
                s1 := '''''';
            if s2 = '' then
                s2 := '''''';
            RxQuery3.MacroByName('LIST').AsString := s;
            RxQuery3.MacroByName('LIST1').AsString := s1;
            RxQuery3.MacroByName('LIST2').AsString := s2;
            RxQuery3.MacroByName('TABLE3').AsString := '''' + fs + Table3 + '''';
            RxQuery3.Open;
        end;
    except
      on E : Exception do begin
        MessageDlg('Ошибка поиска ошибок :)))' + E.Message, mtError, [mbOk], 0);
      end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TExpErrsFlag.SaveBtnClick(Sender: TObject);
var
    F : TextFile;
    i : integer;
    s : string;
begin
    if SaveDialog.Execute then begin
        AssignFile(F, SaveDialog.FileName);
        Rewrite(F);
        RxQuery3.First;
        while not RxQuery3.Eof do begin
            for i := 1 to RxQuery3.Fields.Count do begin
                s := RxQuery3.Fields[i - 1].AsString;
                while Length(s) < 5 do s := ' ' + s;
                Write(F, s);
                if i <> RxQuery3.Fields.Count then
                   Write(F, ', ');
            end;
            WriteLn(F, '');
            RxQuery3.Next;
        end;
        CloseFile(F);
    end;
end;

end.
