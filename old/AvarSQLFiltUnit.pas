unit AvarSQLFiltUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, RxQuery, Grids, DBGrids, RXDBCtrl, ExtCtrls, StdCtrls,
  ComCtrls, Buttons;

type
  TListAvariasSQL = class(TForm)
    Panel1: TPanel;
    AvariaGrid: TRxDBGrid;
    RxAvariaQuery: TRxQuery;
    DataSource: TDataSource;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label4: TLabel;
    RegDateFrom: TDateTimePicker;
    PayDateFrom: TDateTimePicker;
    Label3: TLabel;
    Label5: TLabel;
    RegDateTo: TDateTimePicker;
    PayDateTo: TDateTimePicker;
    Panel2: TPanel;
    ApplyFilterBtn: TSpeedButton;
    Button1: TButton;
    Timer: TTimer;
    RxAvariaQuerySERIA: TStringField;
    RxAvariaQueryNWORK: TStringField;
    RxAvariaQueryFIO: TStringField;
    RxAvariaQueryBADMAN: TStringField;
    RxAvariaQueryISTC1: TStringField;
    RxAvariaQueryOCENTC: TFloatField;
    RxAvariaQueryISTC2: TStringField;
    RxAvariaQueryOTHTC: TFloatField;
    RxAvariaQueryISHL1: TStringField;
    RxAvariaQueryOCENHL: TFloatField;
    RxAvariaQueryISHL2: TStringField;
    RxAvariaQueryOTHHL: TFloatField;
    RxAvariaQueryISIM1: TStringField;
    RxAvariaQueryOCENIM: TFloatField;
    RxAvariaQueryISIM2: TStringField;
    RxAvariaQueryOTHIM: TFloatField;
    RxAvariaQuerySUMPAY: TFloatField;
    RxAvariaQuerySum1_1: TFloatField;
    RxAvariaQuerySum1_1V: TFloatField;
    RxAvariaQuerySum2_1: TFloatField;
    RxAvariaQuerySum2_1V: TFloatField;
    RxAvariaQuerySum3_1: TFloatField;
    RxAvariaQuerySum3_1V: TFloatField;
    MsgPanel: TPanel;
    Label1: TLabel;
    WorkSQL: TRxQuery;
    ShowSummBtn: TButton;
    RxAvariaQueryPayDates: TStringField;
    RxAvariaQueryNUMBER: TIntegerField;
    RxAvariaQueryN: TSmallintField;
    RxAvariaQueryD1: TDateTimeField;
    RxAvariaQueryD2: TDateTimeField;
    RxAvariaQueryD3: TDateTimeField;
    RxAvariaQueryAVDATE: TDateTimeField;
    RxAvariaQueryWRDATE: TDateTimeField;
    RxAvariaQueryDECISION: TIntegerField;
    RxAvariaQueryDecisionText: TStringField;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure ApplyFilterBtnClick(Sender: TObject);
    procedure AvariaGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure ShowSummBtnClick(Sender: TObject);
    procedure RxAvariaQueryCalcFields(DataSet: TDataSet);
  private
     function GetFilterText : string;
     function GetSummaText(Fld, CondFld, FldDate : string) : string;
     function GetSummaText2(Fld, FldDate : string) : string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ListAvariasSQL: TListAvariasSQL;

implementation

uses ShowInfoUnit, inifiles;

{$R *.DFM}

procedure TListAvariasSQL.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TListAvariasSQL.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     RxAvariaQuery.Close;
     WorkSQL.Close;
     Action := caFree;
end;

procedure TListAvariasSQL.FormDestroy(Sender: TObject);
begin
     ListAvariasSQL := nil;
end;

procedure TListAvariasSQL.FormCreate(Sender: TObject);
begin
    try
         //AvTable.Open;
         //RxAvariaQueryNWORK.Size := AvTable.FieldByName('NWork').Size;
         //AvTable.Close;

{         Screen.Cursor := crHourGlass;
         AVQuery.MacroByName('WHERE').AsString := '1=1';
         if MainForm.GetFilter <> '' then
             AVQuery.MacroByName('WHERE').AsString := MainForm.GetFilter;
         AVQuery.Open;
         AvCount.Caption := IntToStr(AVQuery.RecordCount);}
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;

     Screen.Cursor := crDefault;
     Left := Application.MainForm.Left;
     Top := Application.MainForm.Top;
     Width := Application.MainForm.Width;
     Height := Application.MainForm.Height - GetSystemMetrics(SM_CYCAPTION) - 5;
end;

procedure TListAvariasSQL.TimerTimer(Sender: TObject);
begin
     ApplyFilterBtn.Visible := not ApplyFilterBtn.Visible;
end;

function TListAvariasSQL.GetFilterText : string;
var
    s : string;
begin
    if RegDateFrom.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + 'REGDT >= ''' + DateToStr(RegDateFrom.Date) + '''';
    end;
    if RegDateTo.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + 'REGDT <= ''' + DateToStr(RegDateTo.Date) + '''';
    end;
    if PayDateFrom.Checked AND not PayDateTo.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + '(DT1 >= ''' + DateToStr(PayDateFrom.Date) + ''' OR ' +
                'DT2 >= ''' + DateToStr(PayDateFrom.Date) + ''' OR ' +
                'DT3 >= ''' + DateToStr(PayDateFrom.Date) + ''')';
    end;
    if PayDateTo.Checked AND not PayDateFrom.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + '(DT1 <= ''' + DateToStr(PayDateTo.Date) + ''' OR ' +
                'DT2 <= ''' + DateToStr(PayDateTo.Date) + ''' OR ' +
                'DT3 <= ''' + DateToStr(PayDateTo.Date) + ''')';
    end;
    if PayDateTo.Checked AND PayDateFrom.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + '(DT1 >= ''' + DateToStr(PayDateFrom.Date) + ''' AND DT1 <= ''' + DateToStr(PayDateTo.Date) + ''' OR ' +
                'DT2 >= ''' + DateToStr(PayDateFrom.Date) + ''' AND DT2 <= ''' + DateToStr(PayDateTo.Date) + ''' OR ' +
                'DT3 >= ''' + DateToStr(PayDateFrom.Date) + ''' AND DT3 <= ''' + DateToStr(PayDateTo.Date) + ''')';
    end;
    if s = '' then GetFilterText := s
    else GetFilterText := 'WHERE ' + s;
end;

procedure TListAvariasSQL.ApplyFilterBtnClick(Sender: TObject);
begin
     AvariaGrid.SetFocus;

     MsgPanel.Visible := false;
     ApplyFilterBtn.Visible := true;
     Timer.Enabled := false;
     ShowSummBtn.Visible := true;

     RxAvariaQuery.Close;
     RxAvariaQuery.MacroByName('WHERE').AsString := GetFilterText;
     try
         Screen.Cursor := crHourGlass;
         RxAvariaQuery.Open;
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

procedure TListAvariasSQL.AvariaGridGetCellProps(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor);
begin
     if (Field = RxAvariaQueryOCENTC) or
        (Field = RxAvariaQueryOTHTC) then
         Background := $00FFFF00;
     if (Field = RxAvariaQueryOCENTC) and
        (RxAvariaQueryISTC1.AsString = 'N') then
         AFont.Color := clRed;
     if (Field = RxAvariaQueryOTHTC) and
        (RxAvariaQueryISTC2.AsString = 'N') then
         AFont.Color := clRed;

     if (Field = RxAvariaQueryOCENHL) or
        (Field = RxAvariaQueryOTHHL) then
         Background := $0088FF00;
     if (Field = RxAvariaQueryOCENHL) and
        (RxAvariaQueryISHL1.AsString = 'N') then
         AFont.Color := clRed;
     if (Field = RxAvariaQueryOTHHL) and
        (RxAvariaQueryISHL2.AsString = 'N') then
         AFont.Color := clRed;

     if (Field = RxAvariaQueryOCENIM) or
        (Field = RxAvariaQueryOTHIM) then
         Background := $00FF8800;
     if (Field = RxAvariaQueryOCENIM) and
        (RxAvariaQueryISIM1.AsString = 'N') then
         AFont.Color := clRed;
     if (Field = RxAvariaQueryOTHIM) and
        (RxAvariaQueryISIM2.AsString = 'N') then
         AFont.Color := clRed;
end;

function TListAvariasSQL.GetSummaText(Fld, CondFld, FldDate : string) : string;
var
     s, filt_str : string;
begin
    if PayDateFrom.Checked then
       filt_str := FldDate + ' >= ''' + DateToStr(PayDateFrom.Date) + '''';

    if PayDateTo.Checked then begin
       if filt_str <> '' then filt_str := filt_str + ' AND ';
       filt_str := filt_str + FldDate + ' <= ''' + DateToStr(PayDateTo.Date) + '''';
    end;

    if filt_str <> '' then filt_str := filt_str + ' AND ';
    filt_str := filt_str + 'CURR=''BRB''';

     with WorkSQL, WorkSQL.SQL do begin
         Clear;
         Add('SELECT SUM(' + Fld + ') AS S ');
         Add('FROM SYSADM.MANDAV');
         if GetFilterText <> '' then begin
            Add(GetFilterText);
            Add('AND ' + CondFld + '=''Y''' + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + CondFld + '=''Y''' + 'AND ' + filt_str);

         Open;
         s := FloatToStr(FieldByName('S').AsFloat) + ' Руб.';
         Close;

         Clear;
         Add('SELECT SUM(' + Fld + ') AS S ');
         Add('FROM SYSADM.MANDAV');
         if GetFilterText <> '' then begin
            Add(GetFilterText);
            Add('AND ' + CondFld + '<>''Y''' + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + CondFld + '<>''Y''' + 'AND ' + filt_str);

         Open;
         s := s + ', ' + FloatToStr(FieldByName('S').AsFloat) + ' Руб.';
         Close;
     end;

     GetSummaText := s;
end;

function TListAvariasSQL.GetSummaText2(Fld, FldDate : string) : string;
var
     s : string;
     filt_str : string;
begin
    if PayDateFrom.Checked then
       filt_str := FldDate + ' >= ''' + DateToStr(PayDateFrom.Date) + '''';

    if PayDateTo.Checked then begin
       if filt_str <> '' then filt_str := filt_str + ' AND ';
       filt_str := filt_str + FldDate + ' <= ''' + DateToStr(PayDateTo.Date) + '''';
    end;

    if filt_str <> '' then filt_str := filt_str + ' AND ';
    filt_str := filt_str + 'CURR=''BRB''';

     with WorkSQL, WorkSQL.SQL do begin
         Clear;
         Add('SELECT SUM(' + Fld + 'E) AS S, SUM(' + Fld + 'B) AS S2 ');
         Add('FROM SYSADM.MANDAV');
         if GetFilterText <> '' then begin
            Add(GetFilterText + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + filt_str);

         Open;
         s := FieldByName('S').AsString + ' Евро, ' + FieldByName('S2').AsString + ' Руб.';
         Close;
{
         Clear;
         Add('SELECT SUM(' + Fld + 'V) AS S');
         Add('FROM MANDAV2');
         if GetFilterText <> '' then begin
            Add(GetFilterText + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + filt_str);

         Open;
         s := s + ', ' + FieldByName('S').AsString + ' Руб.';
         Close;}
     end;

     GetSummaText2 := s;
end;

procedure TListAvariasSQL.ShowSummBtnClick(Sender: TObject);
var
     s : string;
begin
    Screen.Cursor := crHourGlass;
    try
     s := s + 'Оценка TC        ' + GetSummaText('OZS1', 'OZ1', 'DT1') + #13;
     s := s + 'Прочее TC        ' + GetSummaText('OTHS1', 'OTH1', 'DT1') + #13;
     s := s + 'Оценка Жизнь     ' + GetSummaText('OZS2', 'OZ2', 'DT2') + #13;
     s := s + 'Прочее Жизнь     ' + GetSummaText('OTHS2', 'OTH2', 'DT2') + #13;
     s := s + 'Оценка Имущество ' + GetSummaText('OZS3', 'OZ3', 'DT3') + #13;
     s := s + 'Прочее Имущество ' + GetSummaText('OTHS3', 'OTH3', 'DT3') + #13;
     s := s + #13;
     s := s + 'Страх. возмещение TC        ' + GetSummaText2('RS1', 'DT1') + #13;
     s := s + 'Страх. возмещение Жизнь     ' + GetSummaText2('RS2', 'DT2') + #13;
     s := s + 'Страх. возмещение Имущество ' + GetSummaText2('RS3', 'DT3') + #13;
    except
     on E : Exception do
         s := E.Message;
    end;
    Screen.Cursor := crDefault;

    ShowInfo.InfoPanel.Lines.Text := s;
    ShowInfo.Caption := 'Результаты';
    ShowInfo.ShowModal
end;

procedure TListAvariasSQL.RxAvariaQueryCalcFields(DataSet: TDataSet);
var
     INI : TIniFile;
     i, P : integer;
     s, s2 : string;
begin
     RxAvariaQueryPayDates.AsString := RxAvariaQueryD1.AsString + ', ' +
                                       RxAvariaQueryD2.AsString + ', ' +
                                       RxAvariaQueryD3.AsString;

     if Length(RxAvariaQueryPayDates.AsString) = 4 then
         RxAvariaQueryPayDates.AsString := '';

     INI := TIniFile.Create('blank.ini');
     RxAvariaQueryDecisionText.AsString := 'НЕ ОПРЕДЕЛЕНО';
     for i := 0 to 5 do begin
         s := INI.ReadString('MANDATORY', 'AVDesision' + IntToStr(i), '');
         if s = '' then break;
         P := Pos(',', s);
         if P = 0 then break;
         s2 := Trim(Copy(s, 1, P - 1));
         if RxAvariaQueryDecision.AsString = s2 then begin
             RxAvariaQueryDecisionText.AsString := Copy(s, P + 1, 100);
             break;
         end;
     end;
     INI.Free;
end;

end.
