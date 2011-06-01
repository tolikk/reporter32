unit TableModifier;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, Db, DBTables;

type
  TModifier = class(TForm)
    PackTablesBtn: TButton;
    Button3: TButton;
    Button4: TButton;
    SpeedButton1: TSpeedButton;
    OpenDialog: TOpenDialog;
    StatusBar: TStatusBar;
    Table: TTable;
    WorkQuery: TQuery;
    Label1: TLabel;
    ListAliases: TComboBox;
    lbPath: TLabel;
    BatchMove: TBatchMove;
    TableDest: TTable;
    lbInfo: TLabel;
    OpenDialogTxt: TOpenDialog;
    Database: TDatabase;
    iniInfo: TLabel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    function CreateBackup(TableName : string) : boolean;
    function RestoreBackup(TableName : string) : boolean;
    procedure FormCreate(Sender: TObject);
    procedure ListAliasesChange(Sender: TObject);
    procedure PackTablesBtnClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Modifier: TModifier;


implementation

uses AddFld, bdeutils, datamod, bde, math, inifiles, rxstrutils;

{$R *.DFM}

procedure TModifier.SpeedButton1Click(Sender: TObject);
begin
    Close
end;

function GetLanguage(Table : TTable) : string;
var
    FCurProp: CURProps;
begin
    if not Table.Active then Table.Open;
    Check(DbiGetCursorProps(Table.Handle, FCurProp));
    GetLanguage := FCurProp.szLangDriver;
end;

procedure SetLanguage(Table : TTable; Path, NewLang, CurrentLang : string);
var
    BF : File of Char;
    i, Counter, Position : integer;
    Ch : char;
    TableVer : integer;
begin
    if NewLang = CurrentLang then exit;
    try
        Table.Open;
        TableVer := Table.TableLevel;
        Table.Close;
        Table.DatabaseName := '';
        if true{TableVer = 4} then begin
            AssignFile(BF, Path + Table.TableName);
            Reset(BF);
            Counter := 1; //Читать чимволы
            Position := 0;
            for i := 0 to 2048 do begin
                if Position <> 0 then break;
                if Counter = 1 then Read(BF, Ch);
                Counter := 1;
                if (Ch = CurrentLang[1]) then begin
                    Counter := 2; //Проверять строку
                    while true do begin
                        Read(BF, Ch);
                        if Ch <> CurrentLang[Counter] then break;
                        Counter := Counter + 1;
                        if Counter > Length(CurrentLang) then begin
                            Position := FilePos(BF) - Length(CurrentLang);
                            break;
                        end;
                    end;
                end;
            end;
            if Position <> 0 then begin
                Seek(BF, Position);
                for i := 1 to Length(NewLang) do begin
                    Ch := NewLang[i];
                    Write(BF, Ch);
                end;
                for i := i to max(Length(NewLang), Length(CurrentLang)) do begin
                    Ch := #0;
                    Write(BF, Ch);
                end;
            end;
            CloseFile(BF);
        end;
    except
        Table.Close;
        Table.DatabaseName := '';
        MessageDlg('Нельзя изменить язык таблицы ' + Table.TableName, mtInformation, [mbOk], 0);
        //raise;
    end;
end;

procedure TModifier.Button3Click(Sender: TObject);
var
    OldLang : string;
begin
    OpenDialog.Options := OpenDialog.Options - [ofAllowMultiSelect];
    OpenDialog.InitialDir := lbPath.Caption;
    if not OpenDialog.Execute then exit;

    try
        Table.Close;
        Table.TableName := OpenDialog.FileName;
        Table.Open;
        OldLang := GetLanguage(Table);
    except
        MessageDlg('Нельзя открыть таблицу ' + OpenDialog.FileName + ' для модификации. Возможно какое-либо приложение испобльзует её.', mtInformation, [mbOk], 0);
        exit;
    end;
    AddFldForm.Table := Table;
    AddFldForm.TableName.Text := OpenDialog.FileName;
    if AddFldForm.ShowModal <> mrOk then begin
        Table.Close;
        exit;
    end;

    try
       Screen.Cursor := crHourglass;
       if not CreateBackup(Table.TableName) then begin
           Table.Close;
           exit;
       end;
       StatusBar.SimpleText := 'Добавление поля в таблицу';
       StatusBar.Update;
       Update;
       WorkQuery.Close;
       WorkQuery.SQL.Clear;
       WorkQuery.SQL.Add('ALTER TABLE ''' + OpenDialog.FileName + ''' ADD ' + AddFldForm.NameFld.Text + ' ' + AddFldForm.DataTypeName);
       Table.Close;
       WorkQuery.ExecSQL;
       SetLanguage(Table, '', OldLang, GetLanguage(Table));
       MessageDlg('Поле добавлено успешно', mtInformation, [mbOk], 0);
    except
      on E : Exception do begin
        MessageDlg('Ошибка добавления поля ' + AddFldForm.NameFld.Text + ' в таблицу ' + E.Message + #13'Все изменявшиеся таблицы находятся в ' + GetAppDir + 'BACKUP\', mtError, [mbOk], 0);
      end;
    end;
    Table.Close;
    StatusBar.SimpleText := '';
    Screen.Cursor := crDefault;
end;

function TModifier.CreateBackup(TableName : string) : boolean;
var
    s : string;
    SrcName, DestName : array[0..255] of char;
begin
    CreateBackup := true;
    s := StatusBar.SimpleText;
    StatusBar.SimpleText := 'Резервное копирование...';
    StatusBar.Update;
    if CopyFile(StrPCopy(SrcName, TableName), StrPCopy(DestName, GetAppDir + 'BACKUP\' + ExtractFileName(TableName)), False) = false then begin
        MessageDlg('Ошибка резервного копирования файла', mtError, [mbOk], 0);
        CreateBackup := false;   
    end;
    CopyFile(StrPCopy(SrcName, ChangeFileExt(TableName, '.MB')), StrPCopy(DestName, ChangeFileExt(GetAppDir + 'BACKUP\' + ExtractFileName(TableName), '.MB')), False);
    StatusBar.SimpleText := s;
    StatusBar.Update;
{    CreateBackup := false;
    TableDest.TableName := GetAppDir + 'BACKUP\' + ExtractFileName(TableName);
    try
        s := StatusBar.SimpleText;
        StatusBar.SimpleText := 'Резервное копирование...';
        StatusBar.Update;
        BatchMove.Execute;
        if BatchMove.MovedCount <> BatchMove.Source.RecordCount then begin
            MessageDlg('Попытка создать копию данных для ' + TableName + ' не удалась', mtInformation, [mbOk], 0);
            exit;
        end;
        CreateBackup := true;
    except
      on E : Exception do begin
        MessageDlg('Ошибка резервного копирования файла'#13 + E.Message, mtError, [mbOk], 0);
      end;
    end;
    StatusBar.SimpleText := s;
    StatusBar.Update;
    TableDest.Close;
}
end;

function TModifier.RestoreBackup(TableName : string) : boolean;
begin
    RestoreBackup := true;
end;

procedure TModifier.FormCreate(Sender: TObject);
var
    MyStringList : TStringList;
    I : integer;
begin
    iniInfo.Caption := BLANK_INI;
    MyStringList := TStringList.Create;
    try
      Session.GetAliasNames(MyStringList);
      { fill a list box with alias names for the user to select from }
      for I := 0 to MyStringList.Count - 1 do begin
        if Session.GetAliasDriverName(MyStringList[I]) = 'STANDARD' then
           ListAliases.Items.Add(MyStringList[I]);
      end;
    finally
      MyStringList.Free;
    end;
    lbInfo.Caption := 'Копии данных будут в ' + GetAppDir + 'BACKUP\'
end;

procedure TModifier.ListAliasesChange(Sender: TObject);
//var
//    MyStringList : TStringList;
begin
    if ListAliases.ItemIndex >= 0 then begin

//        MyStringList := TStringList.Create;
//        Session.GetAliasParams(ListAliases.Text, MyStringList);
        lbPath.Caption := GetAliasPath(ListAliases.Text); //MyStringList.Values['PATH'];
//        MyStringList.Free;
    end;
end;

procedure TModifier.PackTablesBtnClick(Sender: TObject);
var
   f: file of Byte;
   size1, size2 : integer;
   rec1, rec2 : integer;
   str : string;
   i : integer;
begin
    OpenDialog.InitialDir := lbPath.Caption;
    OpenDialog.Options := OpenDialog.Options + [ofAllowMultiSelect];
    if not OpenDialog.Execute then exit;

    try
        Screen.Cursor := crHourglass;

        for i := 0 to OpenDialog.Files.Count - 1 do begin
            AssignFile(f, OpenDialog.Files[i]);
            Reset(f);
            size1 := FileSize(f);
            CloseFile(f);

            Table.TableName := OpenDialog.Files[i];
            Table.Open;

            if not CreateBackup(Table.TableName) then break;

            rec1 := Table.RecordCount;
            StatusBar.SimpleText := 'Упорядочивание данных в ' + OpenDialog.Files[i] + '...';
            StatusBar.Update;
            Update;
            PackTable(Table);
            rec2 := Table.RecordCount;
            Table.Close;
            Reset(f);
            size2 := FileSize(f);
            CloseFile(f);

            str := str + #13 + OpenDialog.Files[i] + ' : ';
            if size2 < size1 then str := str + 'Размер изменился с ' + IntToStr(size1) + ' до ' + InttoStr(size2) + ' байт, уменьшился на ' + IntToStr(size1 - size2)
            else str := str + 'Размер не изменился';
            if rec2 < rec1 then str := str + #13'ВНИМАНИЕ!!! Пропало ' + IntToStr(rec1 - rec2) + ' записей';
        end;
    except
        MessageDlg('Ошибка при операции на таблицей ' + OpenDialog.Files[i]+ '. Возможно какое-либо приложение испобльзует её.'#13'Все изменявшиеся таблицы находятся в ' + GetAppDir + 'BACKUP\', mtInformation, [mbOk], 0);
    end;
    if str <> '' then
        MessageDlg('Результат упаковки данных' + str, mtInformation, [mbOk], 0);
    StatusBar.SimpleText := '';
    Screen.Cursor := crDefault;
    Table.Close;
end;

procedure TModifier.Button4Click(Sender: TObject);
var
    F : TextFile;
    str : string;
    Alias : string;
    INI : TIniFile;
    N : integer;
    RecCount : integer;
    sect : string;
    sect_key : string;
    value : string;
begin
    Alias := '';
    RecCount := 0;
    if not OpenDialogTxt.Execute then exit;
    AssignFile(F, OpenDialogTxt.FileName);
    Reset(F);
    Screen.Cursor := crHourglass;

    if ListAliases.Text <> '' then begin
        Alias := GetAliasPath(ListAliases.Text);
        Session.AddStandardAlias('TASKCMD', Alias, 'PARADOX');
    end;

    while not Eof(F) do begin
        ReadLn(F, str);
        str := Trim(str);
        if str = '' then continue;
        if Copy(str, 1, 2) = '//' then continue;
        if Copy(str, 1, 4) = 'MSG ' then begin
            MessageDlg(Trim(Copy(str, 5, 1024)), mtInformation, [mbOk], 0);
            continue;
        end;
        if Copy(str, 1, 5) = 'RUSS ' then begin
            Table.Close;    
            Table.DatabaseName := 'TASKCMD';
            Table.TableName := UpperCase(Trim(Copy(str, 6, 1024)));
            if Pos('.DB', Table.TableName) = 0 then Table.TableName := Table.TableName + '.DB';
            SetLanguage(Table, GetAliasPath('TASKCMD'), 'cyrr', GetLanguage(Table));
            continue;
        end;
        if Copy(str, 1, 4) = 'INI ' then begin
            INI := TIniFile.Create(BLANK_INI);
            N := Pos(ExtractWord(2, str, [' ']), Copy(str, 5, Length(str))) + 4;
            N := Pos(ExtractWord(3, str, [' ']), Copy(str, N + 1, Length(str))) + N + Length(ExtractWord(3, str, [' ']));
            sect := ExtractWord(2, str, [' ']);
            sect_key := ExtractWord(3, str, [' ']);
            value := Trim(Copy(str, N, Length(str)));
            INI.WriteString(sect, sect_key, value);
            INI.Free;
            continue;
        end;
        if Copy(str, 1, 3) = 'DB ' then begin
            if Alias <> '' then Session.DeleteAlias('TASKCMD');
            Alias := Trim(Copy(str, 4, 1024));
            if Pos('\', Alias) = 0 then
                Alias := GetAliasPath(Alias);
            Session.AddStandardAlias('TASKCMD', Alias, 'PARADOX');
            WorkQuery.Close;
            WorkQuery.DatabaseName := 'TASKCMD';
            continue;
        end;
        try
            WorkQuery.Close;
            WorkQuery.SQL.Clear;
            WorkQuery.SQL.Add(str);
            WorkQuery.ExecSQL;
            RecCount := RecCount + WorkQuery.RowsAffected;
        except
          on E : Exception do begin
            if MessageDlg('Ошибка при выполнении команды ' + str + #13 + E.Message + #13'Продолжить?', mtError, [mbYes, mbNo], 0) = mrNo then
                break;
          end;
        end;
    end;
    Session.DeleteAlias('TASKCMD');
    WorkQuery.Close;
    WorkQuery.DatabaseName := '';
    Screen.Cursor := crDefault;
    CloseFile(F);
    Database.Close;

    MessageDlg('Обработано записей: ' + IntToStr(RecCount), mtInformation, [mbOk], 0);
end;

end.
