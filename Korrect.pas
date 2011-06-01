unit Korrect;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, Grids, DBGrids, Menus, RXSpin, RXDBCtrl,
  Mask;

type
  TEnumFunc = function (s, s1, s2, filter : string) : string;

  TKorrectData = class(TForm)
    GetStrings: TQuery;
    GetStringsSrc: TDataSource;
    GetStringsSTR: TStringField;
    GetStringsSERIA: TStringField;
    GetStringsNUMBER: TFloatField;
    CommandsMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    ChTable: TTable;
    N5: TMenuItem;
    GetStringsTBL: TStringField;
    Panel2: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    ClassFld: TComboBox;
    SortCombo: TComboBox;
    Button2: TButton;
    TableType: TComboBox;
    NmbFrom: TEdit;
    NmbTo: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Ver: TComboBox;
    IsPeriod: TCheckBox;
    Monthes: TComboBox;
    Year: TRxSpinEdit;
    IsNumber: TCheckBox;
    SaveCh: TButton;
    RxDBGrid: TRxDBGrid;
    UpdateSQL: TUpdateSQL;
    WorkSQL: TQuery;
    Divider1: TMenuItem;
    ChAddress: TMenuItem;
    Splitter: TSplitter;
    OutWnd: TMemo;
    FindNumber: TRxSpinEdit;
    BadManMenu: TMenuItem;
    procedure ClassFldChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SortComboChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N5Click(Sender: TObject);
    procedure CheckData(Sender: TObject);
    procedure IsNumberClick(Sender: TObject);
    procedure GetStringsUpdateRecord(DataSet: TDataSet;
      UpdateKind: TUpdateKind; var UpdateAction: TUpdateAction);
    procedure SaveChClick(Sender: TObject);
    procedure RxDBGridGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure GetStringsAfterEdit(DataSet: TDataSet);
    procedure GetStringsBeforeInsert(DataSet: TDataSet);
    procedure GetStringsBeforeDelete(DataSet: TDataSet);
    procedure CommandsMenuPopup(Sender: TObject);
    procedure ChAddressClick(Sender: TObject);
    procedure GetStringsBeforeEdit(DataSet: TDataSet);
    procedure FindNumberChange(Sender: TObject);
    procedure BadManMenuClick(Sender: TObject);
  private
    function forEach(OpName, TableName, Field, Filter : string; func : TEnumFunc; s1, s2, filt_str : string) : boolean;
    procedure BigForEach(s : string; func : TEnumFunc; s1, s2, filter : string);
    function IsTableN(N : integer) : boolean;
    function NmbFilter : string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  KorrectData: TKorrectData;
  RecChanged : integer;
  SN : string;

function ExpStr(s : string) : string;

implementation

uses MF, ChWord, Move_word, Variants;

{$R *.DFM}

function TKorrectData.IsTableN(N : integer) : boolean;
begin
     IsTableN := (TableType.ItemIndex = 0) OR (TableType.ItemIndex = N);
end;

procedure TKorrectData.ClassFldChange(Sender: TObject);
var
     fldName : string;
     GLFilter : string;
     verFilt : string;
begin
     if SaveCh.Enabled then
          SaveChClick(nil);

     verFilt := '';
     if Ver.ItemIndex = 0 then verFilt := ' AND VERSION=2';
     GLFilter := NmbFilter;
     if GLFilter <> '' then GLFilter := ' AND ' + GLFilter;
     with GetStrings, GetStrings.SQL do begin
          Close;
          Clear;
          if ClassFld.ItemIndex = 0 then begin //FIO
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,NAME AS STR, ''Полисы       '' AS TBL FROM MANDATOR WHERE PAYDATE IS NOT NULL  ' + GLFilter + verFilt);
             if IsTableN(2)then begin
                if SQL.Count > 0 then
                    Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,NAME AS STR, ''Др. владельцы'' AS TBL  FROM MANDOWN WHERE (1>0) ' + GLFilter);
             end;
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,FIO AS STR, ''Выплаты      '' AS TBL  FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'ФИО';
             Fields[2].Size := 50;
          end;
          if ClassFld.ItemIndex = 1 then begin //Address
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,ADDRESS AS STR, ''Полисы       '' AS TBL FROM MANDATOR WHERE PAYDATE IS NOT NULL  ' + GLFilter + verFilt);
             if IsTableN(2)then begin
                if SQL.Count > 0 then
                    Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,ADDRESS AS STR, ''Др. владельцы'' AS TBL  FROM MANDOWN WHERE (1=1) ' + GLFilter);
             end;
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,FIOADDRESS AS STR, ''Выплаты      '' AS TBL  FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'Адреса';
             Fields[2].Size := 50;
          end;
          if ClassFld.ItemIndex = 2 then begin //Marka
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,MARKA AS STR, ''Полисы       '' AS TBL FROM MANDATOR WHERE PAYDATE IS NOT NULL  ' + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,MARKA_P AS STR, ''Выплаты      '' AS TBL FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'Марки автомобилей';
             Fields[2].Size := 30;
          end;
          if ClassFld.ItemIndex = 3 then begin //AutoNumber
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,AUTONUMBER AS STR, ''Полисы       '' AS TBL FROM MANDATOR WHERE PAYDATE IS NOT NULL  ' + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,AUTONMB_P AS STR, ''Выплаты      '' AS TBL FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'Номера автомобилей';
             Fields[2].Size := 10;
          end;
          if ClassFld.ItemIndex = 4 then begin //Кузов
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,NUMBERBODY AS STR, ''Полисы       '' AS TBL FROM MANDATOR M WHERE STATE<>1 ' + GLFilter);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,BODY_NO AS STR, ''Выплаты      '' AS TBL FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'Номера кузова';
          end;
          if ClassFld.ItemIndex = 5 then begin //Шасси
             if IsTableN(1) then
                Add('SELECT SERIA,NUMBER,CHASSIS AS STR, ''Полисы       '' AS TBL FROM MANDATOR M WHERE STATE<>1 ' + GLFilter);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SERIA,NUMBER,BODY_NO AS STR, ''Выплаты      '' AS TBL FROM MANDAV2 WHERE (1=1) ' + GLFilter);
             end;
             fldName := 'Номера шасси';
          end;

          if SortCombo.ItemIndex = 1 then
             Add('ORDER BY 3')
          else
          if SortCombo.ItemIndex = 2 then
             Add('ORDER BY 3 DESC')
          else
             Add('ORDER BY NUMBER, SERIA');

          try
             if ClassFld.Itemindex <> -1 then begin
                 Screen.Cursor := crHourglass;
                 //DBGrid.DataSource := nil;
                 Open;
                 //DBGrid.DataSource := GetStringsSrc;
                 Fields[2].DisplayLabel := fldName;
                 //ShowMessage(IntToStr(RecordCount));
                 if RecordCount = 0 then
                     Close
             end;
          except
             on E : Exception do
                MessageDlg(E.Message, mtInformation, [mbOK], 0);
          end;
          Screen.Cursor := crDefault;
     end;
end;

procedure TKorrectData.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
     KorrectData := nil;
end;

procedure TKorrectData.SortComboChange(Sender: TObject);
begin
     ClassFldChange(nil);
end;

procedure TKorrectData.FormCreate(Sender: TObject);
begin
     SortCombo.ItemIndex := 0;
     TableType.ItemIndex := 0;
     Ver.ItemIndex := 0;
     Monthes.ItemIndex := 7;

     Left := MAINFORM.Left;
     Top := MAINFORM.Top;// + GetSystemMetrics(SM_CYCAPTION);
     Width := MAINFORM.Width;
     Height := MAINFORM.Height - GetSystemMetrics(SM_CYCAPTION) - 5;
end;

function ExpStr(s : string) : string;
var
     signS : string;
     I : integer;
begin
     signS := ';:.,"''!@#$%^&*()_=+<>?~`/[]{}\|';
     s := Trim(s);
     s := AnsiUpperCase(s);

     for i := 1 to Length(signS) do
        while Pos(signS[i], s) <> 0 do
          s[Pos(signS[i], s)] := ' ';

     while Pos('Ё', s) <> 0 do
        s[Pos('Ё', s)] := 'Е';

     while Pos('  ', s) <> 0 do
        s := Copy(s, 1, Pos('  ', s)) + Copy(s, Pos('  ', s) + 2, 1024);

     ExpStr := s;
end;

function ExpStrER(s : string) : string;
var
     enStr : string;
     ruStr : string;
     i : integer;
begin
     enStr := 'ABWVGDEJZIYKQLMNOPRSTUFC';
     ruStr := 'АБВВГДЕЖЗИИККЛМНОПРСТЮФЦ';
     s := ExpStr(s);

     for i := 1 to Length(enStr) do
        while Pos(enStr[i], s) <> 0 do
          s[Pos(enStr[i], s)] := ruStr[i];

     ExpStrER := s;
end;

procedure TKorrectData.Button1Click(Sender: TObject);
var
     p : TPoint;
begin
     BadManMenu.Visible := ((GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0) AND
                           ((GetAsyncKeyState(VK_CONTROL) AND $8000) <> 0);
     SaveChClick(nil);
     GetCursorPos(p);
     CommandsMenu.Popup(p.x, p.y);
end;

function TKorrectData.NmbFilter : string;
var
    GlFilter : string;
begin
     if IsNumber.Checked then begin
         if Length(NmbFrom.Text) > 0 then begin
            GlFilter := ' NUMBER >= ' + NmbFrom.Text;
         end;
         if Length(NmbTo.Text) > 0 then begin
            if GlFilter <> '' then GlFilter := GlFilter + ' AND ';
            GlFilter := GlFilter + ' NUMBER <= ' + NmbTo.Text;
         end;
     end;    

     if IsPeriod.Checked then begin
        if GlFilter <> '' then GlFilter := GlFilter + ' AND (';
        GlFilter := GlFilter + ' UPDATEDATE >=''' + DateToStr(EncodeDate(ROUND(Year.Value), Monthes.ItemIndex + 1, 1)) + ''' AND ' +
                               ' UPDATEDATE <''' + DateToStr(IncMonth(EncodeDate(ROUND(Year.Value), Monthes.ItemIndex + 1, 1), 1)) + ''' OR ';
     end;

     NmbFilter := GlFilter;
end;

function TKorrectData.forEach(OpName, TableName, Field, Filter : string; func : TEnumFunc; s1, s2, filt_str : string) : boolean;
var
     newStr : string;
     GlFilter : string;
     amsg : MSG;

     CurrRec, AllRec : integer;
begin
     forEach := true;
     ChTable.TableName := TableName;
     GlFilter := NmbFilter;

     if (Filter <> '') AND (GlFilter <> '') then
        GlFilter := ' AND ' + GlFilter; //добавить AND

     ChTable.Filter := Filter + GlFilter;

     try
         MAINFORM.StatusBar.Panels[0].Text := OpName;
         MAINFORM.StatusBar.Update;
         ChTable.Open;
         CurrRec := 1;
         AllRec := ChTable.RecordCount;
         while not ChTable.EOF do begin
               newStr := func(ChTable.FieldByName(Field).AsString, s1, s2, filt_str);
               if newStr <> ChTable.FieldByName(Field).AsString then begin
                   OutWnd.Lines.Add('''' + ChTable.FieldByName(Field).AsString + ''' заменил на ''' + newStr + '''');
                   Inc(RecChanged);
                   SN := ChTable.FieldByName('Seria').AsString + '/' + ChTable.FieldByName('Number').AsString;
                   ChTable.Edit;
                   ChTable.FieldByName(Field).AsString := newStr;
                   ChTable.Post;
               end;

               ChTable.Next;
               Inc(CurrRec);
               MAINFORM.StatusBar.Panels[1].Text := IntToStr(CurrRec);
               MAINFORM.StatusBar.Update;
               MAINFORM.ProgressBar.Position := ROUND(CurrRec * 100 / AllRec);
               if(GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 then begin
                   MessageBeep(0);
                   while (GetAsyncKeyState(VK_ESCAPE) AND $8000) <> 0 do;
                   while PeekMessage(amsg, 0, WM_KEYDOWN, WM_KEYDOWN, PM_REMOVE) do;
                   while PeekMessage(amsg, 0, WM_KEYUP, WM_KEYUP, PM_REMOVE) do;
                   if MessageDlg('Вы хотите прервать текущую операцию', mtInformation, [mbYes, mbNo], 0) = mrYes then begin
                       forEach := false;
                       break;
                   end;
               end;
         end;
     except
        on E : Exception do begin
           MessageDlg('Возникла ошибка на полисе ' + SN + #13 + E.Message, mtInformation, [mbOK], 0);
           forEach := false;
        end;
     end;

     MAINFORM.ProgressBar.Position := 0;
     MAINFORM.StatusBar.Panels[0].Text := '';
     MAINFORM.StatusBar.Panels[1].Text := '';
     ChTable.Close;
end;

function auxUpStr(s, s1, s2, filter : string) : string;
begin
     auxUpStr := AnsiUpperCase(s);
end;

procedure TKorrectData.BigForEach(s : string; func : TEnumFunc; s1, s2, filter : string);
Var
     GoOn : boolean;
     verFilt : string;
begin
     verFilt := '';
     RecChanged := 0;
     OutWnd.Lines.Text := '';

     if Ver.ItemIndex = 0 then verFilt := ' AND VERSION=2';

     Screen.Cursor := crHourGlass;
     GoOn := true;
     if ClassFld.ItemIndex = 0 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' ФИО в полисах', 'MANDATOR', 'NAME', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(2) AND GoOn then GoOn := forEach(s + ' ФИО во владельцах', 'MANDOWN', 'NAME', 'SERIA IS NOT NULL', func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' ФИО в выплатах', 'MANDAV2', 'FIO', 'SERIA IS NOT NULL', func, s1, s2, filter);
     end;
     if ClassFld.ItemIndex = 1 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Адрес в полисах', 'MANDATOR', 'ADDRESS', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(2) AND GoOn then GoOn := forEach(s + ' Адрес во владельцах', 'MANDOWN', 'ADDRESS', '(NUMBER>0)', func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Адрес в выплатах', 'MANDAV2', 'FIOADDRESS', '(NUMBER>0)', func, s1, s2, filter);
     end;
     if ClassFld.ItemIndex = 2 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Марка в полисах', 'MANDATOR', 'MARKA', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Марка в выплатах', 'MANDAV2', 'MARKA_P', '(NUMBER>0)', func, s1, s2, filter);
     end;
     if ClassFld.ItemIndex = 3 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Номер в полисах', 'MANDATOR', 'AUTONUMBER', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Номер в выплатах', 'MANDAV2', 'AUTONMB_P', '(NUMBER>0)', func, s1, s2, filter);
     end;
     if ClassFld.ItemIndex = 4 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Номер кузова в полисах', 'MANDATOR', 'NUMBERBODY', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Номер кузова в выплатах', 'MANDAV2', 'BODY_NO', '(NUMBER>0)', func, s1, s2, filter);
     end;
     if ClassFld.ItemIndex = 5 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Номер шасси в полисах', 'MANDATOR', 'CHASSIS', 'PAYDATE IS NOT NULL ' + verFilt, func, s1, s2, filter);
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Номер шасси в выплатах', 'MANDAV2', 'BODY_NO', '(NUMBER>0)', func, s1, s2, filter);
     end;

     OutWnd.Lines.Add('Оперция ' + s + ' произвела ' + IntToStr(RecChanged) + ' замен');

     Screen.Cursor := crDefault;

     if (RecChanged > 0) AND (MessageDlg('Вы хотите перечитать данные в списке?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
         ClassFldChange(nil);
end;

procedure TKorrectData.N1Click(Sender: TObject);
begin
     BigForEach('В верхний регистр', auxUpStr, '', '', '');
end;

function auxDelSpacesAndDivs(s, s1, s2, filter : string) : string;
begin
     auxDelSpacesAndDivs := ExpStr(s);
end;

procedure TKorrectData.N2Click(Sender: TObject);
begin
     BigForEach('Лишние символы в', auxDelSpacesAndDivs, '', '', '');
end;

function auxEnglishToRussian(s, s1, s2, filter : string) : string;
begin
     auxEnglishToRussian := ExpStrER(s);
end;

procedure TKorrectData.N3Click(Sender: TObject);
begin
     BigForEach('Англ. в русские в', auxEnglishToRussian, '', '', '');
end;

function auxReplaceWords(s, s1, s2, filter : string) : string;
var
     divider : integer;
begin
     divider := Pos(s1, s);

     if divider <> 0 then
        s := Copy(s, 1, divider - 1) + s2 + Copy(s, divider + Length(s1), 1024);

     auxReplaceWords := s;
end;

procedure TKorrectData.N4Click(Sender: TObject);
begin
     if ChWordsDlg.ShowModal = mrOK then begin
          BigForEach('Замена слова ' + ChWordsDlg.Word1.Text + ' на ' + ChWordsDlg.Word2.Text, auxReplaceWords, ChWordsDlg.Word1.Text, ChWordsDlg.Word2.Text, '');
     end;
end;

function auxMoveWord(s, s1, s2, filter : string) : string;
var
     divider : integer;
     ClipLen : integer;
begin
     if (Length(filter) > 0) AND (Pos(filter, s) = 0) then begin
        auxMoveWord := s;
        exit;
     end;

     divider := Pos(' ' + s1 + ' ', s);
     ClipLen := 0;
     if divider = 0 then begin
        if (Copy(s, 1, Length(s1) + 1) = (s1 + ' ')) AND (s2 = '0') then begin
           divider := 1;
           ClipLen := 1;
        end;
        if (Copy(s, Length(s) - Length(s1), Length(s1) + 1) = (' ' + s1)) AND (s2 = '1') then begin
           divider := Length(s) - Length(s1);
           ClipLen := 0;
        end
     end;

     if divider <> 0 then begin
        if s2 = '0' then begin //Right
            s := Copy(s, 1, divider - ClipLen) + Copy(s, divider + Length(s1) + 2 - ClipLen, 1024);

            //Пропускаем пробелы
            while (divider <= Length(s)) AND (s[divider] = ' ') do
                  Inc(divider);

            //Пропускаем слово
            while (divider <= Length(s)) AND (s[divider] <> ' ') do
                  Inc(divider);

            if divider > Length(s) then
                s := s + ' ' + s1
            else
                s := Copy(s, 1, divider) + s1 + Copy(s, divider, 1024);
        end;
        if s2 = '1' then begin //Left
            s := Copy(s, 1, divider - ClipLen) + Copy(s, divider + Length(s1) + 2 - ClipLen, 1024);

            //Пропускаем пробелы
            while (divider > 0) AND (s[divider] = ' ') do
                  Dec(divider);

            //Пропускаем слово
            while (divider > 0) AND (s[divider] <> ' ') do
                  Dec(divider);

            if divider = 0 then
                s := s1 + ' ' + s
            else
                s := Copy(s, 1, divider) + s1 + Copy(s, divider, 1024);

        end;
     end;

     auxMoveWord := s;
end;

procedure TKorrectData.N5Click(Sender: TObject);
begin
     if MoveWord.ShowModal = mrOK then begin
          BigForEach('Слово ' + MoveWord.Word.Text + ' ' + MoveWord.Direction.Text, auxMoveWord, MoveWord.Word.Text, IntToStr(MoveWord.Direction.ItemIndex), MoveWord.Filter.Text);
     end;
end;

procedure TKorrectData.CheckData(Sender: TObject);
begin
      if Year.Value = 2000 then
         if Monthes.ItemIndex < 6 then
             Monthes.ItemIndex := 6;

     if IsPeriod.Checked then
         ClassFldChange(nil);
end;

procedure TKorrectData.IsNumberClick(Sender: TObject);
begin
     ClassFldChange(nil);
end;

procedure TKorrectData.GetStringsUpdateRecord(DataSet: TDataSet;
  UpdateKind: TUpdateKind; var UpdateAction: TUpdateAction);
var
     Table, Field : string;
     s : string;
begin
     Table := 'MANDATOR';
     case ClassFld.ItemIndex of
          0 : Field := 'NAME';
          1 : Field := 'ADDRESS';
          2 : Field := 'MARKA';
          3 : Field := 'AUTONUMBER';
          4 : Field := 'NUMBERBODY';
          5 : Field := 'CHASSIS';
     end;

     if Pos('Выпл', GetStringsTBL.NewValue) > 0 then begin
          Table := 'MANDAV2';
          case ClassFld.ItemIndex of
                0 : Field := 'FIO';
                1 : Field := 'FIOADDRESS';
                2 : Field := 'MARKA_P';
                3 : Field := 'AUTONMB_P';
                4 : Field := 'BODY_NO';
                5 : Field := 'BODY_NO';
          end;
     end;

     if Pos('Др', GetStringsTBL.NewValue) > 0 then begin
          Table := 'MANDOWN';
          case ClassFld.ItemIndex of
                0 : Field := 'NAME';
                1 : Field := 'FIOADDRESS';
          end;
     end;

     with WorkSQL, WorkSQL.SQL do begin
          Clear;
          s := 'UPDATE ' + Table +
               ' SET ' + Field + '=''' + VarToStr(GetStringsSTR.NewValue) + '''' +
               ' WHERE SERIA=''' + VarToStr(GetStringsSERIA.OldValue) + '''' +
               ' AND NUMBER= ' + VarToStr(GetStringsNUMBER.OldValue) +
               ' AND ' + Field + '=''' + VarToStr(GetStringsSTR.OldValue) + '''';
          Add(s);
          ExecSQL;
     end;

     UpdateAction := uaApplied;
end;

procedure TKorrectData.SaveChClick(Sender: TObject);
begin
    if not SaveCh.Enabled then exit;
    try
      Screen.Cursor := crHourGlass;
      GetStrings.ApplyUpdates; {try to write the updates to the database};
      SaveCh.Enabled := false;
      GetStrings.CommitUpdates; {on success, clear the cache}
    except
      on E : Exception do begin
          MessageDlg('Ошибка обновления данных'#13 + E.Message, mtInformation, [mbOK], 0);
      end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TKorrectData.RxDBGridGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
     if Field.OldValue <> Field.NewValue then
        Background := $00FFCACA
end;

procedure TKorrectData.GetStringsAfterEdit(DataSet: TDataSet);
begin
      SaveCh.Enabled := true;
end;

procedure TKorrectData.GetStringsBeforeInsert(DataSet: TDataSet);
begin
     Abort
end;

procedure TKorrectData.GetStringsBeforeDelete(DataSet: TDataSet);
begin
     Abort
end;

procedure TKorrectData.CommandsMenuPopup(Sender: TObject);
begin
     ChAddress.Visible := false;
     ChAddress.Visible := GetStrings.Active AND (ClassFld.ItemIndex = 1) AND (GetStringsTBL.AsString = 'Др. владельцы');
     Divider1.Visible := ChAddress.Visible;

     if(ChAddress.Visible) then
        ChAddress.Caption := 'Сменить адреса как ''' + GetStringsSTR.AsString + ''' на адреса владельца';
end;

procedure TKorrectData.ChAddressClick(Sender: TObject);
begin
    if MessageDlg('Операция корректировки адреса будет проведена для всей БД без учёта фильтра! Продолжить?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;

    Screen.Cursor := crHourGlass;
    try
      with WorkSQL, WorkSQL.SQL do begin
           Clear;
           Add('UPDATE MANDOWN O SET O.ADDRESS = (SELECT M.ADDRESS FROM MANDATOR M WHERE M.NUMBER=O.NUMBER AND M.SERIA=O.SERIA)');
           Add('WHERE ADDRESS=:ADDR');
           ParamByName('ADDR').AsString := GetStringsSTR.AsString;
           ExecSQL;
      end;
    except
      on E : Exception do
          MessageDlg(E.Message, mtInformation, [mbOK], 0);
    end;

    ClassFldChange(nil);
    Screen.Cursor := crDefault;
end;

procedure TKorrectData.GetStringsBeforeEdit(DataSet: TDataSet);
begin
//     if SaveCh.Enabled = false then
//        MessageDlg('Необходимо выбрать период, за который вы хотите менять данные', mtInformation, [mbOK], 0);
//     if not IsPeriod.Checked then begin
//        MessageDlg('Необходимо выбрать период, за который вы хотите менять данные', mtInformation, [mbOK], 0);
//        Abort;
//     end;
end;

procedure TKorrectData.FindNumberChange(Sender: TObject);
begin
     if not GetStrings.Active then exit;

     GetStrings.Locate('Number', FindNumber.Value, []);
end;

procedure TKorrectData.BadManMenuClick(Sender: TObject);
begin
     RecChanged := 0;
     OutWnd.Lines.Text := '';

     Screen.Cursor := crHourGlass;
     forEach('ФИО виновника в авариях', 'MANDAV2', 'BADMAN', '', auxUpStr, '', '', '');
     forEach('ФИО виновника в авариях', 'MANDAV2', 'BADMAN', '', auxDelSpacesAndDivs, '', '', '');
     forEach('ФИО потерпевшего в авариях', 'MANDAV2', 'FIO', '', auxUpStr, '', '', '');
     forEach('ФИО потерпевшего в авариях', 'MANDAV2', 'FIO', '', auxDelSpacesAndDivs, '', '', '');
     OutWnd.Lines.Add('Оперция произвела ' + IntToStr(RecChanged) + ' замен');

     Screen.Cursor := crDefault;
end;

end.
