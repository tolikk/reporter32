unit ArariasFilterUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, RxQuery, Grids, DBGrids, RXDBCtrl, ExtCtrls, StdCtrls,
  ComCtrls, Buttons, Placemnt;

type
  TListAvarias = class(TForm)
    Panel1: TPanel;
    AvariaGrid: TRxDBGrid;
    RxAvariaQuery: TRxQuery;
    DataSource: TDataSource;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label4: TLabel;
    RegDateFrom: TDateTimePicker;
    PeriodFrom: TDateTimePicker;
    Label3: TLabel;
    Label5: TLabel;
    RegDateTo: TDateTimePicker;
    PeriodTo: TDateTimePicker;
    Panel2: TPanel;
    ApplyFilterBtn: TSpeedButton;
    Button1: TButton;
    RxAvariaQuerySERIA: TStringField;
    RxAvariaQueryNUMBER: TFloatField;
    RxAvariaQueryN: TFloatField;
    RxAvariaQueryAVDATE: TDateField;
    RxAvariaQueryWRDATE: TDateField;
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
    RxAvariaQueryV1_Date: TDateField;
    RxAvariaQueryV2_Date: TDateField;
    RxAvariaQueryV3_Date: TDateField;
    RxAvariaQueryPayDates: TStringField;
    AVTable: TTable;
    Button2: TButton;
    RxAvariaQueryOutSumma: TFloatField;
    RxAvariaQueryDeclSumma: TFloatField;
    RxAvariaQueryPayInfo: TStringField;
    RxAvariaQueryStateName: TStringField;
    RxAvariaQueryWhat: TFloatField;
    StateList: TComboBox;
    IsShowPaysDetail: TCheckBox;
    RxAvariaQueryTYPE: TStringField;
    RxAvariaQueryD: TDateField;
    RxAvariaQueryRISK: TStringField;
    RxAvariaQueryAvPayType: TStringField;
    RxAvariaQuerySUMMA: TFloatField;
    RxAvariaQueryV1: TStringField;
    FormStorage: TFormStorage;
    RxAvariaQueryCURR: TStringField;
    RxAvariaQueryRISKSTR: TStringField;
    RxAvariaQueryZATRSTR: TStringField;
    RxAvariaQueryZATR: TSmallintField;
    IsZayav: TCheckBox;
    RxAvariaQueryDtClLoss: TDateField;
    Button3: TButton;
    RxAvariaQueryUpdateDate: TDateField;
    procedure Button1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ApplyFilterBtnClick(Sender: TObject);
    procedure AvariaGridGetCellProps(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor);
    procedure ShowSummBtnClick(Sender: TObject);
    procedure RxAvariaQueryCalcFields(DataSet: TDataSet);
    procedure Button2Click(Sender: TObject);
    procedure StateListChange(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
     function GetFilterText : string;
     function GetSummaText(Fld, CondFld, FldDate : string) : string;
     function GetSummaText2(Fld, FldDate : string) : string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ListAvarias: TListAvarias;

implementation

uses MF, ShowInfoUnit, datamod, inifiles;

{$R *.DFM}

procedure TListAvarias.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TListAvarias.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
end;

procedure TListAvarias.FormDestroy(Sender: TObject);
begin
     ListAvarias := nil;
end;

procedure TListAvarias.FormCreate(Sender: TObject);
begin
    try
         AvTable.Open;
         RxAvariaQueryNWORK.Size := AvTable.FieldByName('NWork').Size;
         AvTable.Close;

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
     Left := MAINFORM.Left;
     Top := MAINFORM.Top;
     Width := MAINFORM.Width;
     Height := MAINFORM.Height - GetSystemMetrics(SM_CYCAPTION) - 5;

end;

function TListAvarias.GetFilterText : string;
var
    s, s1 : string;
begin
    if RegDateFrom.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + 'WRDATE >= ''' + DateToStr(RegDateFrom.Date) + '''';
    end;
    if RegDateTo.Checked then begin
       if s <> '' then s := s + ' AND ';
       s := s + 'WRDATE <= ''' + DateToStr(RegDateTo.Date) + '''';
    end;

    if not IsShowPaysDetail.Checked then begin
        if PeriodFrom.Checked AND not PeriodTo.Checked then begin
           if s <> '' then s := s + ' AND ';
           s := s + '(V1_Date >= ''' + DateToStr(PeriodFrom.Date) + ''' OR ' +
                    'V2_Date >= ''' + DateToStr(PeriodFrom.Date) + ''' OR ' +
                    'V3_Date >= ''' + DateToStr(PeriodFrom.Date) + ''')';
        end;
        if PeriodTo.Checked AND not PeriodFrom.Checked then begin
           if s <> '' then s := s + ' AND ';
           s := s + '(V1_Date <= ''' + DateToStr(PeriodTo.Date) + ''' OR ' +
                    'V2_Date <= ''' + DateToStr(PeriodTo.Date) + ''' OR ' +
                    'V3_Date <= ''' + DateToStr(PeriodTo.Date) + ''')';
        end;
        if PeriodTo.Checked AND PeriodFrom.Checked then begin
           if s <> '' then s := s + ' AND ';
           s := s + '(V1_Date >= ''' + DateToStr(PeriodFrom.Date) + ''' AND V1_Date <= ''' + DateToStr(PeriodTo.Date) + ''' OR ' +
                    'V2_Date >= ''' + DateToStr(PeriodFrom.Date) + ''' AND V2_Date <= ''' + DateToStr(PeriodTo.Date) + ''' OR ' +
                    'V3_Date >= ''' + DateToStr(PeriodFrom.Date) + ''' AND V3_Date <= ''' + DateToStr(PeriodTo.Date) + ''')';
        end;
    end
    else begin
        if PeriodFrom.Checked then begin
            s1 := ' D >= ''' + DateToStr(PeriodFrom.Date) + '''';
        end;
        if PeriodTo.Checked then begin
            if s1 <> '' then s1 := s1 + ' AND ';
            s1 := s1 + ' D <= ''' + DateToStr(PeriodTo.Date) + '''';
        end;
        if (s1 <> '') and (s <> '') then s := s + ' AND ';
        s := s + s1;
    end;

    if StateList.ItemIndex = 1 then begin
       if s <> '' then s := s + ' AND ';
       s := s + 'NOT (WHAT IN(0,1)) ';
    end;
    if StateList.ItemIndex = 1 then begin
       if s <> '' then s := s + ' AND ';
       s := s + '(WHAT IN(0,1)) ';
    end;

    if IsShowPaysDetail.Checked then begin
        if s <> '' then s := ' AND ' + s;
        s := ' ,AVPAYS AP WHERE A.SERIA=AP.SER AND A.NUMBER=AP.NMB AND A.N=AP.N' + s;
        if not IsZayav.Checked then begin
            s := s + ' AND AP.TYPE=''1''';
        end;
    end;


    if (s = '') or IsShowPaysDetail.Checked then GetFilterText := s
    else GetFilterText := 'WHERE ' + s;
end;

procedure TListAvarias.ApplyFilterBtnClick(Sender: TObject);
begin
     AvariaGrid.SetFocus;

     RxAvariaQueryD.Visible := IsShowPaysDetail.Checked;
     RxAvariaQuerySUMMA.Visible := IsShowPaysDetail.Checked;
     RxAvariaQueryCURR.Visible := IsShowPaysDetail.Checked;
     RxAvariaQueryAvPayType.Visible := IsShowPaysDetail.Checked;
     //RxAvariaQueryPayInfo.Visible := not IsShowPaysDetail.Checked;
     AvariaGrid.Columns[4].Visible := not IsShowPaysDetail.Checked;
     AvariaGrid.Columns[5].Visible := not IsShowPaysDetail.Checked;
     AvariaGrid.Columns[12].Visible := IsShowPaysDetail.Checked;
     AvariaGrid.Columns[13].Visible := IsShowPaysDetail.Checked;
     AvariaGrid.Columns[14].Visible := IsShowPaysDetail.Checked;
     AvariaGrid.Columns[15].Visible := IsShowPaysDetail.Checked;
     AvariaGrid.Columns[16].Visible := IsShowPaysDetail.Checked;
     AvariaGrid.Columns[17].Visible := IsShowPaysDetail.Checked;
     //AvariaGrid.Columns[6].Visible := not IsShowPaysDetail.Checked;

     MsgPanel.Visible := false;
     ApplyFilterBtn.Visible := true;

     RxAvariaQuery.Close;
     RxAvariaQuery.MacroByName('WHERE').AsString := GetFilterText;
     if IsShowPaysDetail.Checked then
        RxAvariaQuery.MacroByName('FIELDS').AsString := 'AP.TYPE, AP.D, AP."SUM" AS SUMMA, AP.CURR, AP.ZATR AS ZATR, AP.RISK AS RISK'
     else
        RxAvariaQuery.MacroByName('FIELDS').AsString := ''''' AS TYPE, ''01.01.2000'' AS D, 0 AS SUMMA, '''' AS CURR, cast (-1 as SMALLINT) AS ZATR, '''' AS RISK';

     try
         Screen.Cursor := crHourGlass;
         RxAvariaQuery.Open;
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;
     Screen.Cursor := crDefault;
end;

procedure TListAvarias.AvariaGridGetCellProps(Sender: TObject;
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

function TListAvarias.GetSummaText(Fld, CondFld, FldDate : string) : string;
var
     s, filt_str : string;
begin
{    if PayDateFrom.Checked then
       filt_str := FldDate + ' >= ''' + DateToStr(PayDateFrom.Date) + '''';

    if PayDateTo.Checked then begin
       if filt_str <> '' then filt_str := filt_str + ' AND ';
       filt_str := filt_str + FldDate + ' <= ''' + DateToStr(PayDateTo.Date) + '''';
    end;
    if filt_str = '' then filt_str := '1=1';

     with WorkSQL, WorkSQL.SQL do begin
         Clear;
         Add('SELECT SUM(' + Fld + ') AS S');
         Add('FROM MANDAV2');
         if GetFilterText <> '' then begin
            Add(GetFilterText);
            Add(' AND ' + CondFld + '=''Y''' + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + CondFld + '=''Y''' + 'AND ' + filt_str);

         Open;
         s := FloatToStr(FieldByName('S').AsFloat) + ' Руб.';
         Close;

         Clear;
         Add('SELECT SUM(' + Fld + ') AS S');
         Add('FROM MANDAV2');
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
}
     GetSummaText := s;
end;

function TListAvarias.GetSummaText2(Fld, FldDate : string) : string;
var
     s : string;
     filt_str : string;
begin
{    if PayDateFrom.Checked then
       filt_str := FldDate + ' >= ''' + DateToStr(PayDateFrom.Date) + '''';

    if PayDateTo.Checked then begin
       if filt_str <> '' then filt_str := filt_str + ' AND ';
       filt_str := filt_str + FldDate + ' <= ''' + DateToStr(PayDateTo.Date) + '''';
    end;
    if filt_str = '' then filt_str := '1=1';

     with WorkSQL, WorkSQL.SQL do begin
         Clear;
         Add('SELECT SUM(' + Fld + ') AS S, SUM(' + Fld + 'V) AS S2');
         Add('FROM MANDAV2');
         if GetFilterText <> '' then begin
            Add(GetFilterText + 'AND ' + filt_str);
         end
         else
            Add('WHERE ' + filt_str);

         Open;
         s := FieldByName('S').AsString + ' Евро, ' + FieldByName('S2').AsString + ' Руб.';
         Close;

     end;
}
     GetSummaText2 := s;
end;

procedure TListAvarias.ShowSummBtnClick(Sender: TObject);
var
     s : string;
begin
    Screen.Cursor := crHourGlass;
    try
     s := s + 'Оценка TC        ' + GetSummaText('OCENTC', 'ISTC1', 'V1_Date') + #13;
     s := s + 'Прочее TC        ' + GetSummaText('OTHTC', 'ISTC2', 'V1_Date') + #13;
     s := s + 'Оценка Жизнь     ' + GetSummaText('OCENHL', 'ISHL1', 'V2_Date') + #13;
     s := s + 'Прочее Жизнь     ' + GetSummaText('OTHHL', 'ISHL2', 'V2_Date') + #13;
     s := s + 'Оценка Имущество ' + GetSummaText('OCENIM', 'ISIM1', 'V3_Date') + #13;
     s := s + 'Прочее Имущество ' + GetSummaText('OTHIM', 'ISIM2', 'V3_Date') + #13;
     s := s + #13;
     s := s + 'Страх. возмещение TC        ' + GetSummaText2('SUM1_1', 'V1_Date') + #13;
     s := s + 'Страх. возмещение Жизнь     ' + GetSummaText2('SUM2_1', 'V2_Date') + #13;
     s := s + 'Страх. возмещение Имущество ' + GetSummaText2('SUM3_1', 'V3_Date') + #13;
    except
     on E : Exception do
         s := E.Message;
    end;
    Screen.Cursor := crDefault;

    ShowInfo.InfoPanel.Lines.Text := s;
    ShowInfo.Caption := 'Результаты';
    ShowInfo.ShowModal
end;

procedure TListAvarias.RxAvariaQueryCalcFields(DataSet: TDataSet);
var
    s : string;
    period : string;
    z : integer;
    INI : TIniFile;
begin
    if PeriodFrom.Checked then period := 'D >= ''' + DateToStr(PeriodFrom.Date) + '''';
    if PeriodTo.Checked then begin
        if period <> '' then period := period + ' AND ';
        period := period + 'D <= ''' + DateToStr(PeriodTo.Date) + '''';
    end;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT SUM(T."SUM"), CURR FROM AVPAYS T');
    WorkSQL.SQL.Add('WHERE SER=:S AND NMB=:N AND TYPE=''1''');
    if period <> '' then WorkSQL.SQL.Add(' AND ' + period);
    WorkSQL.SQL.Add('GROUP BY CURR');
    WorkSQL.ParamByName('S').AsString := RxAvariaQuerySERIA.AsString;
    WorkSQL.ParamByName('N').AsFloat := RxAvariaQueryNumber.AsFloat;
    WorkSQL.Open;
    while not WorkSQL.Eof do begin
        if WorkSQL.Fields[1].AsString <> '' then begin
            if WorkSQL.Fields[1].AsString = 'BRB' then
                RxAvariaQueryOutSumma.AsFloat := WorkSQL.Fields[0].AsFloat
            else
                s := 'выпл. ' + WorkSQL.Fields[0].AsString + ' ' + WorkSQL.Fields[1].AsString;
        end;
        WorkSQL.Next;
    end;
    WorkSQL.Close;

    WorkSQL.Close;
    WorkSQL.SQL.Clear;
    WorkSQL.SQL.Add('SELECT SUM(T."SUM"), CURR FROM AVPAYS T');
    WorkSQL.SQL.Add('WHERE SER=:S AND NMB=:N AND TYPE=''0''');
    if period <> '' then WorkSQL.SQL.Add(' AND ' + period);
    WorkSQL.SQL.Add('GROUP BY CURR');
    WorkSQL.ParamByName('S').AsString := RxAvariaQuerySeria.AsString;
    WorkSQL.ParamByName('N').AsFloat := RxAvariaQueryNumber.AsFloat;
    WorkSQL.Open;
    while not WorkSQL.Eof do begin
        if WorkSQL.Fields[1].AsString <> '' then begin
            if WorkSQL.Fields[1].AsString = 'BRB' then
                RxAvariaQueryDeclSumma.AsFloat := WorkSQL.Fields[0].AsFloat
            else
                s := 'заявл. ' + WorkSQL.Fields[0].AsString + ' ' + WorkSQL.Fields[1].AsString;
        end;
        WorkSQL.Next;
    end;
    WorkSQL.Close;

    if IsShowPaysDetail.Checked then begin
        RxAvariaQueryOutSumma.Clear;
        RxAvariaQueryDeclSumma.Clear;
    end;

    if not RxAvariaQueryV1_Date.IsNull then begin
        if s <> '' then s := '; ';
        s := s + RxAvariaQueryV1_Date.AsString + ' ' + RxAvariaQuerySum1_1V.AsString + ' ' + RxAvariaQueryV1.AsString;
    end;
    if not RxAvariaQueryV2_Date.IsNull then begin
        if s <> '' then s := '; ';
        s := s + RxAvariaQueryV2_Date.AsString + ' ' + RxAvariaQuerySum2_1V.AsString + ' ' + RxAvariaQueryV1.AsString;
    end;
    if not RxAvariaQueryV3_Date.IsNull then begin
        if s <> '' then s := '; ';
        s := s + RxAvariaQueryV3_Date.AsString + ' ' + RxAvariaQuerySum3_1V.AsString + ' ' + RxAvariaQueryV1.AsString;
    end;

    RxAvariaQueryPayInfo.AsString := s;

    INI := TIniFile.Create(BLANK_INI);
    RxAvariaQueryStateName.AsString := INI.ReadString('MANDATORY', 'AVDesision' + RxAvariaQueryWhat.AsString, '<НЕ ОПРЕДЕЛЕНО>');
    RxAvariaQueryRISKSTR.AsString := '';
    if RxAvariaQueryRISK.AsString = '0' then RxAvariaQueryRISKSTR.AsString := 'ТС';
    if RxAvariaQueryRISK.AsString = '1' then RxAvariaQueryRISKSTR.AsString := 'ЖИЗНЬ,ЗДОРОВЬЕ';
    if RxAvariaQueryRISK.AsString = '2' then RxAvariaQueryRISKSTR.AsString := 'ИМУЩЕСТВО';
    for z := 0 to 30 do begin
        RxAvariaQueryZATRSTR.AsString := INI.ReadString('MANDATORY', 'Zatr' + IntToStr(z), '');
        if((RxAvariaQueryZATRSTR.AsString = '') or (Pos(RxAvariaQueryZATR.AsString+',', RxAvariaQueryZATRSTR.AsString) <> 0)) then break;
    end;
    INI.Free;

    RxAvariaQueryAvPayType.AsString := 'заявлено';
    if RxAvariaQueryTYPE.AsString = '1' then RxAvariaQueryAvPayType.AsString := 'оплачено';
end;

procedure TListAvarias.Button2Click(Sender: TObject);
begin
    DatSetToExcel(RxAvariaQuery);
end;

procedure TListAvarias.StateListChange(Sender: TObject);
begin
    ApplyFilterBtnClick(Sender);
end;


procedure TListAvarias.Button3Click(Sender: TObject);
var
    i : integer;
    Dt : TDateTime;
begin
    if AvariaGrid.SelCount = 0 then
    begin
        MessageDlg('Выбери дела для закрытия', mtInformation, [mbOk], 0);
        exit;
    end;

    if MessageDlg('Дата закрытия будет установлена равной дате последней выплаты или дате обновления данных. Продолжить?', mtInformation, [mbYes, mbNo], 0) = mrNo then
    begin
        exit;
    end;

    try
        Screen.Cursor := crHourGlass;
        for i := 1 to AvariaGrid.SelCount do
        begin
            AvariaGrid.GotoSelection(i - 1);
            if RxAvariaQuery.FieldByName('DtClLoss').IsNull then
            begin
                Dt := RxAvariaQuery.FieldByName('UpdateDate').AsDateTime;
                if (not RxAvariaQuery.FieldByName('V1_Date').IsNull) and (RxAvariaQuery.FieldByName('V1_Date').AsDateTime > Dt) then
                begin
                    Dt := RxAvariaQuery.FieldByName('V1_Date').AsDateTime;
                end;
                if (not RxAvariaQuery.FieldByName('V2_Date').IsNull) and (RxAvariaQuery.FieldByName('V1_Date').AsDateTime > Dt) then
                begin
                    Dt := RxAvariaQuery.FieldByName('V2_Date').AsDateTime;
                end;
                if (not RxAvariaQuery.FieldByName('V3_Date').IsNull) and (RxAvariaQuery.FieldByName('V1_Date').AsDateTime > Dt) then
                begin
                    Dt := RxAvariaQuery.FieldByName('V4_Date').AsDateTime;
                end;

                MainForm.HandBookSQL.Close;
                MainForm.HandBookSQL.SQL.Clear;
                MainForm.HandBookSQL.SQL.Add('SELECT MAX(D) AS D FROM AVPAYS WHERE TYPE=''1'' AND SER = ''' + RxAvariaQuery.FieldByName('SERIA').AsString + ''' AND NMB = ' + RxAvariaQuery.FieldByName('NUMBER').AsString + ' AND N=' + RxAvariaQuery.FieldByName('N').AsString);
                MainForm.HandBookSQL.Open;
                if (not MainForm.HandBookSQL.FieldByName('D').IsNull) and (MainForm.HandBookSQL.FieldByName('D').AsDateTime > Dt) then
                begin
                    Dt := RxAvariaQuery.FieldByName('D').AsDateTime;
                end;
                MainForm.HandBookSQL.Close;

                MainForm.HandBookSQL.Close;
                MainForm.HandBookSQL.SQL.Clear;
                MainForm.HandBookSQL.SQL.Add('UPDATE MANDAV2 SET DTCLLOSS=:Dt WHERE SERIA = ''' + RxAvariaQuery.FieldByName('SERIA').AsString + '''AND NUMBER = ' + RxAvariaQuery.FieldByName('NUMBER').AsString + ' AND N=' + RxAvariaQuery.FieldByName('N').AsString);
                MainForm.HandBookSQL.ParamByName('Dt').AsDateTime := Dt;
                MainForm.HandBookSQL.ExecSQL;
                MainForm.HandBookSQL.Close;

            end;
        end;
        RxAvariaQuery.Close;
        RxAvariaQuery.Open;
    except
        on E : Exception do
        begin
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
        end
    end;

    Screen.Cursor := crDefault;
end;

end.
