unit KorrectSQL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Db, DBTables, Grids, DBGrids, Menus, RXSpin, RXDBCtrl,
  Mask;

type
  TEnumFunc = function (s, s1, s2, filter : string) : string;

  TKorrectDataSQL = class(TForm)
    GetStrings: TQuery;
    GetStringsSrc: TDataSource;
    CommandsMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
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
    IsNotBURO: TCheckBox;
    ChQuery: TQuery;
    N31: TMenuItem;
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
    procedure N31Click(Sender: TObject);
  private
    function forEach(OpName, TableName, Field, Filter : string; func : TEnumFunc; s1, s2, s3, filt_str, add_fld : string) : boolean;
    procedure BigForEach(s : string; func : TEnumFunc; s1, s2, s3, filter : string);
    function IsTableN(N : integer) : boolean;
    function NmbFilter : string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  KorrectDataSQL: TKorrectDataSQL;
  RecChanged : integer;
  SN : string;

implementation

uses ChWord, Move_word, Korrect, MandSQL, Variants;

{$R *.DFM}

function TKorrectDataSQL.IsTableN(N : integer) : boolean;
begin
     IsTableN := (TableType.ItemIndex = 0) OR (TableType.ItemIndex = N);
end;

function GetBUROCodesFilter(IsNeed : boolean; fld : string) : string;
begin
     if IsNeed then
         Result := '(' + fld + '<=0 OR ' + fld + ' IS NULL)'
     else
         Result := '(Nmb=Nmb)';
end;

procedure TKorrectDataSQL.ClassFldChange(Sender: TObject);
var
     fldName : string;
     GLFilter : string;
     verFilt : string;
begin
     if SaveCh.Enabled then
          SaveChClick(nil);

     verFilt := '';
     if Ver.ItemIndex = 0 then verFilt := ' AND VER=2 ';
     GLFilter := NmbFilter;
     if GLFilter <> '' then GLFilter := ' AND ' + GLFilter;
     with GetStrings, GetStrings.SQL do begin
          Close;
          Clear;
          if ClassFld.ItemIndex = 0 then begin //FIO
             if IsTableN(1) then
                Add('SELECT SER,NMB,NAME, ''Полисы       ''  FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNERCODE') + GLFilter + verFilt);
             if IsTableN(2)then begin
                if SQL.Count > 0 then
                    Add('UNION ALL');
                Add('SELECT SER,NMB,NAME , ''Доп. владельцы''   FROM SYSADM.MANDOWN M WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE') + GLFilter);
             end;
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,FIO, ''Выплаты      ''   FROM SYSADM.MANDAV M WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE') + GLFilter);
             end;
             fldName := 'ФИО';
          end;
          if ClassFld.ItemIndex = 1 then begin //Address
             if IsTableN(1) then
                Add('SELECT SER,NMB,ADDR , ''Полисы       '' FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNERCODE') + GLFilter + verFilt);
             if IsTableN(2)then begin
                if SQL.Count > 0 then
                    Add('UNION ALL');
                Add('SELECT SER,NMB,ADDR , ''Доп. владельцы''  FROM SYSADM.MANDOWN M WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE') + GLFilter);
             end;
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,ADDR , ''Выплаты      ''  FROM SYSADM.MANDAV WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE') + GLFilter);
             end;
             fldName := 'Адреса';
          end;
          if ClassFld.ItemIndex = 2 then begin //Marka
             if IsTableN(1) then
                Add('SELECT SER,NMB,MARKA , ''Полисы       '' FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'BASECARCODE') + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,MARKA , ''Выплаты      '' FROM SYSADM.MANDAV WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'BASECARCODE') + GLFilter);
             end;
             fldName := 'Марки автомобилей';
          end;
          if ClassFld.ItemIndex = 3 then begin //AutoNumber
             if IsTableN(1) then
                Add('SELECT SER,NMB,AUTONMB , ''Полисы       ''  FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,AUTONMB , ''Выплаты      ''  FROM SYSADM.MANDAV WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter);
             end;
             fldName := 'Номера автомобилей';
          end;
          if ClassFld.ItemIndex = 4 then begin //Кузов
             if IsTableN(1) then
                Add('SELECT SER,NMB,NMBODY , ''Полисы       ''  FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,BODYSHS , ''Выплаты      ''  FROM SYSADM.MANDAV WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter);
             end;
             fldName := 'Номера кузова';
          end;
          if ClassFld.ItemIndex = 5 then begin //Шасси
             if IsTableN(1) then
                Add('SELECT SER,NMB,CHASSIS , ''Полисы       ''  FROM SYSADM.MANDATOR M WHERE STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter + verFilt);
             if IsTableN(3)then begin
                if SQL.Count > 0 then
                  Add('UNION ALL');
                Add('SELECT SER,NMB,BODYSHS , ''Выплаты      ''  FROM SYSADM.MANDAV WHERE ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + GLFilter);
             end;
             fldName := 'Номера шасси';
          end;

          if SortCombo.ItemIndex = 1 then
             Add('ORDER BY 3')
          else
          if SortCombo.ItemIndex = 2 then
             Add('ORDER BY 3 DESC')
          else
             Add('ORDER BY 1, 2');

          try
             if ClassFld.Itemindex <> -1 then begin
                 Screen.Cursor := crHourglass;
                 //DBGrid.DataSource := nil;
                 Open;
                 //DBGrid.DataSource := GetStringsSrc;
                 Fields[0].DisplayLabel := 'Серия';
                 Fields[1].DisplayLabel := 'Номер';
                 Fields[2].DisplayLabel := fldName;
                 Fields[3].DisplayLabel := 'Таблица';
                 Fields[0].ReadOnly := true;
                 Fields[1].ReadOnly := true;
                 Fields[3].ReadOnly := true;
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

procedure TKorrectDataSQL.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
     KorrectDataSQL := nil;
end;

procedure TKorrectDataSQL.SortComboChange(Sender: TObject);
begin
     ClassFldChange(nil);
end;

procedure TKorrectDataSQL.FormCreate(Sender: TObject);
begin
     SortCombo.ItemIndex := 0;
     TableType.ItemIndex := 0;
     Ver.ItemIndex := 0;
     Monthes.ItemIndex := 7;

     Left := MandSQLForm.Left;
     Top := MandSQLForm.Top;// + GetSystemMetrics(SM_CYCAPTION);
     Width := MandSQLForm.Width;
     Height := MandSQLForm.Height - GetSystemMetrics(SM_CYCAPTION) - 5;
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

     while (Length(s) > 0) AND (s[1] = ' ') do
         s := Copy(s, 2, Length(s));

     while (Length(s) > 0) AND (s[Length(s)] = ' ') do
         s := Copy(s, 1, Length(s) - 1);

     ExpStr := s;
end;

function ExpStrER(s : string) : string;
var
     enStr : string;
     ruStr : string;
     i : integer;
begin
     enStr := 'HABWVGDEJZIYKQLMNOPRSTUFC';
     ruStr := 'ХАБВВГДЕЖЗИИККЛМНОПРСТЮФЦ';
     s := ExpStr(s);

     for i := 1 to Length(enStr) do
        while Pos(enStr[i], s) <> 0 do
          s[Pos(enStr[i], s)] := ruStr[i];

     ExpStrER := s;
end;

function ExpStrERMarka(s : string) : string;
var
     enStr : string;
     ruStr : string;
     i : integer;
begin
     enStr := 'HABWVGDEJZIYKQLMNOPRSTUFC';
     ruStr := 'ХАБВВГДЕЖЗИИККЛМНОПРСТУФК';
     s := ExpStr(s);

     for i := 1 to Length(enStr) do
        while Pos(enStr[i], s) <> 0 do
          s[Pos(enStr[i], s)] := ruStr[i];

     ExpStrERMarka := s;
end;

procedure TKorrectDataSQL.Button1Click(Sender: TObject);
var
     p : TPoint;
begin
     BadManMenu.Visible := ((GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0) AND
                           ((GetAsyncKeyState(VK_CONTROL) AND $8000) <> 0);
     SaveChClick(nil);
     GetCursorPos(p);
     CommandsMenu.Popup(p.x, p.y);
end;

function TKorrectDataSQL.NmbFilter : string;
var
    GlFilter : string;
begin
     if IsNumber.Checked then begin
         if Length(NmbFrom.Text) > 0 then begin
            GlFilter := ' NMB >= ' + NmbFrom.Text;
         end;
         if Length(NmbTo.Text) > 0 then begin
            if GlFilter <> '' then GlFilter := GlFilter + ' AND ';
            GlFilter := GlFilter + ' NMB <= ' + NmbTo.Text;
         end;
     end;    

     if IsPeriod.Checked then begin
        if GlFilter <> '' then GlFilter := GlFilter + ' AND ';
        GlFilter := GlFilter + ' UPDDT >=''' + DateToStr(EncodeDate(ROUND(Year.Value), Monthes.ItemIndex + 1, 1)) + ''' AND ' +
                               ' UPDDT <''' + DateToStr(IncMonth(EncodeDate(ROUND(Year.Value), Monthes.ItemIndex + 1, 1), 1)) + '''';
     end;

     NmbFilter := GlFilter;
end;

function TKorrectDataSQL.forEach(OpName, TableName, Field, Filter : string; func : TEnumFunc; s1, s2, s3, filt_str, add_fld : string) : boolean;
var
     newStr : string;
     GlFilter : string;
     amsg : MSG;
     curr_filt : string;
     CurrRec, AllRec : integer;
begin
     forEach := true;
     ChQuery.SQL.Clear;
     ChQuery.SQL.Add('SELECT SER,NMB,' + Field);
     if add_fld <> '' then
         ChQuery.SQL.Add(',' + add_fld);
     ChQuery.SQL.Add(' FROM ' + TableName);
     GlFilter := NmbFilter;//MERCEDES
     curr_filt := Filter;
     if GlFilter <> '' then begin
        if curr_filt <> '' then curr_filt := curr_filt + ' AND ';
        curr_filt := curr_filt + GlFilter;
     end;
     if s3 <> '' then begin
        if curr_filt <> '' then curr_filt := curr_filt + ' AND ';
        curr_filt := curr_filt + s3;
     end;

     if curr_filt <> '' then
         ChQuery.SQL.Add('WHERE ' + curr_filt);

     try
         WorkSQL.Close;
         WorkSQL.Params.Clear;
         WorkSQL.SQL.Clear;
         WorkSQL.SQL.Add('UPDATE ' + TableName);
         WorkSQL.SQL.Add('SET ' + Field + '=:VAL');
         WorkSQL.SQL.Add('WHERE SER=:S AND NMB=:Nm');
         if add_fld <> '' then
             WorkSQL.SQL.Add('AND ' + add_fld + '=:' + add_fld);
         WorkSQL.Prepare;
         WorkSQL.ParamByName('VAL').ParamType := ptInput;
         WorkSQL.ParamByName('VAL').DataType := ftString;

         MandSQLForm.StatusBar.Panels[0].Text := OpName;
         MandSQLForm.StatusBar.Update;
         ChQuery.Open;
         CurrRec := 1;
         AllRec := ChQuery.RecordCount;
         while not ChQuery.EOF do begin
               newStr := func(ChQuery.FieldByName(Field).AsString, s1, s2, filt_str);
               if newStr <> ChQuery.FieldByName(Field).AsString then begin
                   SN := ChQuery.FieldByName('Ser').AsString + '/' + ChQuery.FieldByName('Nmb').AsString;

                   MandSQLForm.SQLDB.StartTransaction;
                   WorkSQL.ParamByName('S').AsString := ChQuery.FieldByName('Ser').AsString;
                   WorkSQL.ParamByName('Nm').AsInteger := ChQuery.FieldByName('Nmb').AsInteger;
                   if newStr = '' then 
                       WorkSQL.ParamByName('VAL').Clear
                   else
                       WorkSQL.ParamByName('VAL').AsString := newStr;
                   if add_fld <> '' then begin
                       if ChQuery.FieldByName(add_fld).DataType = ftString then
                           WorkSQL.ParamByName(add_fld).AsString := ChQuery.FieldByName(add_fld).AsString
                       else
                           WorkSQL.ParamByName(add_fld).AsInteger := ChQuery.FieldByName(add_fld).AsInteger;
                   end;
                   WorkSQL.ExecSQL;
                   MandSQLForm.SQLDB.Commit;

                   OutWnd.Lines.Add(SN + '  ''' + ChQuery.FieldByName(Field).AsString + ''' заменил на ''' + newStr + '''');
                   Inc(RecChanged);
               end;

               ChQuery.Next;
               Inc(CurrRec);
               MandSQLForm.StatusBar.Panels[1].Text := IntToStr(CurrRec);
               MandSQLForm.StatusBar.Update;
               MandSQLForm.ProgressBar.Position := ROUND(CurrRec * 100 / AllRec);
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
           if MandSQLForm.SQLDB.InTransaction then
               MandSQLForm.SQLDB.Rollback;
           MessageDlg('Возникла ошибка на полисе ' + SN + #13 + E.Message, mtInformation, [mbOK], 0);
           forEach := false;
        end;
     end;

     MandSQLForm.ProgressBar.Position := 0;
     MandSQLForm.StatusBar.Panels[0].Text := '';
     MandSQLForm.StatusBar.Panels[1].Text := '';
     ChQuery.Close;
end;

function auxUpStr(s, s1, s2, filter : string) : string;
begin
     auxUpStr := AnsiUpperCase(s);
end;

procedure TKorrectDataSQL.BigForEach(s : string; func : TEnumFunc; s1, s2, s3, filter : string);
Var
     GoOn : boolean;
     verFilt : string;
begin
     verFilt := '';
     RecChanged := 0;
     OutWnd.Lines.Text := '';

     if Ver.ItemIndex = 0 then verFilt := ' AND VER=2';

     Screen.Cursor := crHourGlass;
     GoOn := true;
     if ClassFld.ItemIndex = 0 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' ФИО в полисах', 'SYSADM.MANDATOR', 'NAME', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNERCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(2) AND GoOn then GoOn := forEach(s + ' ФИО во владельцах', 'SYSADM.MANDOWN', 'NAME', GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE'), func, s1, s2, s3, filter, 'NAME');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' ФИО в выплатах', 'SYSADM.MANDAV', 'FIO', GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE'), func, s1, s2, s3, filter, 'N');
     end;
     if ClassFld.ItemIndex = 1 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Адрес в полисах', 'SYSADM.MANDATOR', 'ADDR', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'OWNERCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(2) AND GoOn then GoOn := forEach(s + ' Адрес во владельцах', 'SYSADM.MANDOWN', 'ADDR', GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE'), func, s1, s2, s3, filter, 'NAME');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Адрес в выплатах', 'SYSADM.MANDAV', 'ADDR', GetBUROCodesFilter(IsNotBURO.Checked, 'OWNRCODE'), func, s1, s2, s3, filter, 'N');
     end;
     if ClassFld.ItemIndex = 2 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Марка в полисах', 'SYSADM.MANDATOR', 'MARKA', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'BASECARCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Марка в выплатах', 'SYSADM.MANDAV', 'MARKA', GetBUROCodesFilter(IsNotBURO.Checked, 'BASECARCODE'), func, s1, s2, s3, filter, 'N');
     end;
     if ClassFld.ItemIndex = 3 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Номер в полисах', 'SYSADM.MANDATOR', 'AUTONMB', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Номер в выплатах', 'SYSADM.MANDAV', 'AUTONMB', GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE'), func, s1, s2, s3, filter, 'N');
     end;
     if ClassFld.ItemIndex = 4 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Кузов в полисах', 'SYSADM.MANDATOR', 'NMBODY', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Кузов в выплатах', 'SYSADM.MANDAV', 'BODYSHS', GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE'), func, s1, s2, s3, filter, 'N');
     end;
     if ClassFld.ItemIndex = 5 then begin
       if IsTableN(1) AND GoOn then GoOn := forEach(s + ' Шасси в полисах', 'SYSADM.MANDATOR', 'CHASSIS', 'STATE<>1 AND ' + GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE') + verFilt, func, s1, s2, s3, filter, '');
       if IsTableN(3) AND GoOn then GoOn := forEach(s + ' Шасси в выплатах', 'SYSADM.MANDAV', 'BODYSHS', GetBUROCodesFilter(IsNotBURO.Checked, 'CARCODE'), func, s1, s2, s3, filter, 'N');
     end;

     OutWnd.Lines.Add('Операция ' + s + ' произвела ' + IntToStr(RecChanged) + ' замен');

     Screen.Cursor := crDefault;

     if (RecChanged > 0) AND (MessageDlg('Вы хотите перечитать данные в списке?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
         ClassFldChange(nil);
end;

procedure TKorrectDataSQL.N1Click(Sender: TObject);
begin
     BigForEach('В верхний регистр', auxUpStr, '', '', '', '');
end;

function auxDelSpacesAndDivs(s, s1, s2, filter : string) : string;
begin
     auxDelSpacesAndDivs := ExpStr(s);
end;

procedure TKorrectDataSQL.N2Click(Sender: TObject);
begin
     BigForEach('Лишние символы в', auxDelSpacesAndDivs, '', '', '', '');
end;

function auxEnglishToRussian(s, s1, s2, filter : string) : string;
begin
     auxEnglishToRussian := ExpStrER(s);
end;

function auxEnglishToRussianMarka(s, s1, s2, filter : string) : string;
begin
     auxEnglishToRussianMarka := ExpStrERMarka(s);
end;

procedure TKorrectDataSQL.N3Click(Sender: TObject);
begin
     BigForEach('Англ. в русские в', auxEnglishToRussian, '', '', '', '');
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

procedure TKorrectDataSQL.N4Click(Sender: TObject);
var
     filter_s : string;
begin
     filter_s := '';
     if ChWordsDlg.ShowModal = mrOK then begin
          if (TableType.ItemIndex = 1) then begin
              case ClassFld.ItemIndex of 
                  0 : filter_s := 'NAME LIKE ''%' + ChWordsDlg.Word1.Text + '%''';
                  2 : filter_s := 'MARKA LIKE ''%' + ChWordsDlg.Word1.Text + '%''';
              end;
          end;
          BigForEach('Замена слова ' + ChWordsDlg.Word1.Text + ' на ' + ChWordsDlg.Word2.Text, auxReplaceWords, ChWordsDlg.Word1.Text, ChWordsDlg.Word2.Text, filter_s, '');
     end;
end;

function auxMoveWord(s, s1, direction, filter : string) : string;
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
        if (Copy(s, 1, Length(s1) + 1) = (s1 + ' ')) AND (direction = '0') then begin
           divider := 1;
           ClipLen := 1;
        end;
        if (Copy(s, Length(s) - Length(s1), Length(s1) + 1) = (' ' + s1)) AND (direction = '1') then begin
           divider := Length(s) - Length(s1);
           ClipLen := 0;
        end
     end;

     if divider <> 0 then begin
        if direction = '0' then begin //Right
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
        if direction = '1' then begin //Left
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

procedure TKorrectDataSQL.N5Click(Sender: TObject);
begin
     if MoveWord.ShowModal = mrOK then begin
          BigForEach('Слово ' + MoveWord.Word.Text + ' ' + MoveWord.Direction.Text, auxMoveWord, MoveWord.Word.Text, IntToStr(MoveWord.Direction.ItemIndex), '', MoveWord.Filter.Text);
     end;
end;

procedure TKorrectDataSQL.CheckData(Sender: TObject);
begin
      if Year.Value = 2000 then
         if Monthes.ItemIndex < 6 then
             Monthes.ItemIndex := 6;

     if IsPeriod.Checked then
         ClassFldChange(nil);
end;

procedure TKorrectDataSQL.IsNumberClick(Sender: TObject);
begin
     ClassFldChange(nil);
end;

procedure TKorrectDataSQL.GetStringsUpdateRecord(DataSet: TDataSet;
  UpdateKind: TUpdateKind; var UpdateAction: TUpdateAction);
var
     Table, Field : string;
     s : string;
begin
     Table := 'SYSADM.MANDATOR';
     case ClassFld.ItemIndex of
          0 : Field := 'NAME';
          1 : Field := 'ADDR';
          2 : Field := 'MARKA';
          3 : Field := 'AUTONMB';
          4 : Field := 'NMBODY';
          5 : Field := 'CHASSIS';
     end;

     if Pos('Выпл', GetStrings.Fields[3].NewValue) > 0 then begin
          Table := 'SYSADM.MANDAV';
          case ClassFld.ItemIndex of
                0 : Field := 'FIO';
                1 : Field := 'ADDR';
                2 : Field := 'MARKA';
                3 : Field := 'AUTONMB';
                4 : Field := 'BODYSHS';
                5 : Field := 'BODYSHS';
          end;
     end;

     if Pos('Доп', GetStrings.Fields[3].NewValue) > 0 then begin
          Table := 'SYSADM.MANDOWN';
          case ClassFld.ItemIndex of
                0 : Field := 'NAME';
                1 : Field := 'ADDR';
          end;
     end;

     with WorkSQL, WorkSQL.SQL do begin
          Clear;
          s := 'UPDATE ' + Table +
               ' SET ' + Field + '=''' + VarToStr(GetStrings.Fields[2].NewValue) + '''' +
               ' WHERE SER=''' + VarToStr(GetStrings.Fields[0].OldValue) + '''' +
               ' AND NMB= ' + VarToStr(GetStrings.Fields[1].OldValue) +
               ' AND ' + Field + '=''' + VarToStr(GetStrings.Fields[2].OldValue) + '''';
          Add(s);
//          MessageDlg(s, mtInformation, [mbOk], 0);
          ExecSQL;
     end;

     UpdateAction := uaApplied;
end;

procedure TKorrectDataSQL.SaveChClick(Sender: TObject);
begin
    if not SaveCh.Enabled then exit;
    try
      MandSQLForm.SQLDB.StartTransaction;
      Screen.Cursor := crHourGlass;
      GetStrings.ApplyUpdates; {try to write the updates to the database};
      SaveCh.Enabled := false;
      GetStrings.CommitUpdates; {on success, clear the cache}
      MandSQLForm.SQLDB.Commit;
    except
      on E : Exception do begin
          MandSQLForm.SQLDB.Rollback;
          MessageDlg('Ошибка обновления данных'#13 + E.Message, mtInformation, [mbOK], 0);
      end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TKorrectDataSQL.RxDBGridGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
     if Field.OldValue <> Field.NewValue then
        Background := $00FFCACA
end;

procedure TKorrectDataSQL.GetStringsAfterEdit(DataSet: TDataSet);
begin
      SaveCh.Enabled := true;
end;

procedure TKorrectDataSQL.GetStringsBeforeInsert(DataSet: TDataSet);
begin
     Abort
end;

procedure TKorrectDataSQL.GetStringsBeforeDelete(DataSet: TDataSet);
begin
     Abort
end;

procedure TKorrectDataSQL.CommandsMenuPopup(Sender: TObject);
begin
{     ChAddress.Visible := false;
     ChAddress.Visible := GetStrings.Active AND (ClassFld.ItemIndex = 1) AND (GetStringsTBL.AsString = 'Др. владельцы');
     Divider1.Visible := ChAddress.Visible;

     if(ChAddress.Visible) then
        ChAddress.Caption := 'Сменить адреса как ''' + GetStringsSTR.AsString + ''' на адреса владельца';
        }
end;

procedure TKorrectDataSQL.ChAddressClick(Sender: TObject);
begin
{    if MessageDlg('Операция корректировки адреса будет проведена для всей БД без учёта фильтра! Продолжить?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then exit;

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
    Screen.Cursor := crDefault;}
end;

procedure TKorrectDataSQL.GetStringsBeforeEdit(DataSet: TDataSet);
begin
//     if SaveCh.Enabled = false then
//        MessageDlg('Необходимо выбрать период, за который вы хотите менять данные', mtInformation, [mbOK], 0);
//     if not IsPeriod.Checked then begin
//        MessageDlg('Необходимо выбрать период, за который вы хотите менять данные', mtInformation, [mbOK], 0);
//        Abort;
//     end;
end;

procedure TKorrectDataSQL.FindNumberChange(Sender: TObject);
begin
     if not GetStrings.Active then exit;

     GetStrings.Locate('Nmb', FindNumber.Value, []);
end;

procedure TKorrectDataSQL.BadManMenuClick(Sender: TObject);
begin
     RecChanged := 0;
     OutWnd.Lines.Text := '';

     Screen.Cursor := crHourGlass;
     forEach('ФИО виновника в авариях', 'SYSADM.MANDAV', 'BADMAN', '', auxUpStr, '', '', '', '', 'N');
     forEach('ФИО виновника в авариях', 'SYSADM.MANDAV', 'BADMAN', '', auxDelSpacesAndDivs, '', '', '', '', 'N');
     forEach('ФИО потерпевшего в авариях', 'SYSADM.MANDAV', 'FIO', '', auxUpStr, '', '', '', '', 'N');
     forEach('ФИО потерпевшего в авариях', 'SYSADM.MANDAV', 'FIO', '', auxDelSpacesAndDivs, '', '', '', '', 'N');
     OutWnd.Lines.Add('Операция произвела ' + IntToStr(RecChanged) + ' замен');

     Screen.Cursor := crDefault;
end;

procedure TKorrectDataSQL.N31Click(Sender: TObject);
begin
     BigForEach('Англ. в русские в', auxEnglishToRussianMarka, '', '', '', '');
end;

end.
