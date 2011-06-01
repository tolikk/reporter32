unit FilterMandatory;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Placemnt, ComCtrls, CheckLst, Db, DBTables, Buttons,
  RXSpin, inifiles, Menus, Mask, DataMod;

type
  TMandatoryFilterDlg = class(TForm)
    DownPanel: TPanel;
    Panel2: TPanel;
    DataPanel: TPanel;
    IsPayDate: TCheckBox;
    Label1: TLabel;
    Pay1Start_: TDateTimePicker;
    Label2: TLabel;
    Pay1End_: TDateTimePicker;
    BitBtn1: TBitBtn;
    IsInsurer: TCheckBox;
    InsNameMode: TComboBox;
    PartName: TEdit;
    IsInsStart: TCheckBox;
    Label4: TLabel;
    InsStart1: TDateTimePicker;
    Label5: TLabel;
    InsStart2: TDateTimePicker;
    IsInsEnd: TCheckBox;
    Label6: TLabel;
    InsEnd1: TDateTimePicker;
    Label7: TLabel;
    InsEnd2: TDateTimePicker;
    IsPeriod: TCheckBox;
    Period: TComboBox;
    Bevel2: TBevel;
    Is2Fee: TCheckBox;
    SecondFee: TComboBox;
    IsPay2: TCheckBox;
    SecondPay: TComboBox;
    IsPutDate: TCheckBox;
    Label10: TLabel;
    PutDate1: TDateTimePicker;
    Label11: TLabel;
    PutDate2: TDateTimePicker;
    Label12: TLabel;
    IsPayNumber: TCheckBox;
    Label13: TLabel;
    PayNumber1: TRxSpinEdit;
    Label14: TLabel;
    PayNumber2: TRxSpinEdit;
    FormStorageFilter: TFormStorage;
    IsWasActive: TCheckBox;
    Label8: TLabel;
    WasActiveFrom: TDateTimePicker;
    Label9: TLabel;
    WasActiveTo: TDateTimePicker;
    Panel1: TPanel;
    AgentList: TCheckListBox;
    Divider: TBevel;
    Label3: TLabel;
    Panel3: TPanel;
    CntChecked: TLabel;
    FormStorage: TFormStorage;
    IsFee2Date: TCheckBox;
    Label15: TLabel;
    Fee2DateFrom: TDateTimePicker;
    Label16: TLabel;
    Fee2DateTo: TDateTimePicker;
    ListTemplates: TComboBox;
    SpeedButton: TSpeedButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    AgentsTbl: TQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    States: TListBox;
    IsState: TCheckBox;
    Label17: TLabel;
    IsRepDate: TCheckBox;
    Label18: TLabel;
    RepDate1: TDateTimePicker;
    Label19: TLabel;
    RepDate2: TDateTimePicker;
    Label20: TLabel;
    IsRepDate2: TCheckBox;
    Rep2Date1: TDateTimePicker;
    Label21: TLabel;
    Rep2Date2: TDateTimePicker;
    InsurType: TComboBox;
    IsUpDT: TCheckBox;
    Label22: TLabel;
    UpDtFrom: TDateTimePicker;
    Label23: TLabel;
    UpDtTo: TDateTimePicker;
    Label24: TLabel;
    ListNumbers: TEdit;
    Is1SU: TCheckBox;
    IsStopped: TCheckBox;
    IsReport: TComboBox;
    Label25: TLabel;
    Label26: TLabel;
    IsOther: TComboBox;
    Bevel1: TBevel;
    IsAdditionalPay: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AgentListClick(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    AgCodes : TStringList;
  public
    { Public declarations }
    function  GetFilter(s : string) : string;
    function  GetFilter1 : string;
    function  GetFilter2 : string;
    procedure ConnectFilter(var filter : string; Operation, Condition : string);
    function  GetCommonFilter(AgentField : string) : string;
  end;

var
  MandatoryFilterDlg: TMandatoryFilterDlg;

implementation

uses MF, TmplNameUnit, registry, rxstrutils;

{$R *.DFM}

function ExpandPeriod(s : string) : string;
var
    s1, s2, res : string;
    i : integer;
begin
    if Pos('-', s) = 0 then res := s
    else begin
        s1 := ExtractWord(1, s, ['-']);
        s2 := ExtractWord(2, s, ['-']);
        for i := StrToInt(s1) to StrToInt(s2) do begin
            if res <> '' then res := res + ',';
            res := res + IntToStr(i);
        end;
    end;
    ExpandPeriod := res;
end;

function ExpandNmb(s : string) : string;
var
    res, s1 : string;
    i : integer;
begin
    for i := 1 to 100 do begin
        s1 := ExtractWord(i, s, [',']);
        if s1 = '' then break;
        if res <> '' then res := res + ',';
        res := res + ExpandPeriod(s1);
    end;

    ExpandNmb := res;
end;

function TMandatoryFilterDlg.GetCommonFilter(AgentField : string) : string;
var
     filter : string;
     str, state_filt : string;
     i, sel_i : integer;
     s : string;
begin
     filter := '';

     if ListNumbers.Text <> '' then begin
         ConnectFilter(filter, 'AND', 'NUMBER IN (' + ExpandNmb(ListNumbers.Text) + ')');
     end;

     if IsWasActive.Checked then begin
         //Начали действовать в период или испорчены в период
         ConnectFilter(filter, 'AND', '(' +
                                          '(STATE <> 3 AND FROMDATE BETWEEN ''' + DateToStr(WasActiveFrom.Date) + ''' AND ''' + DateToStr(WasActiveTo.Date) + ''')' +
                                          ' OR ' +
                                          '(PAYDATE IS NULL AND REGDATE BETWEEN ''' + DateToStr(WasActiveFrom.Date) + ''' AND ''' + DateToStr(WasActiveTo.Date) + ''')' +
                                      ')');
     end;

     if Is1SU.Checked then
          ConnectFilter(filter, 'AND', 'ALLFEETEXT IN (''Y'', ''0'')');

     //Агенты
     str := '';
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if Length(str) > 0 then begin
          if(AgentField = '') then
              ConnectFilter(filter, 'AND', '(AgentCode IN (' + Copy(str, 1, Length(str) - 1) + ') OR Fee2Text IN (' + Copy(str, 1, Length(str) - 1) + '))');
          if(AgentField <> '') then
              ConnectFilter(filter, 'AND', '(' + AgentField + ' IN (' + Copy(str, 1, Length(str) - 1) + '))');
     end;

     //Страхователь
     if (IsInsurer.Checked) AND (Trim(PartName.Text) <> '') then begin
        if InsNameMode.Itemindex = 0 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''' + Trim(PartName.Text) + '%''');
        if InsNameMode.Itemindex = 1 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''%' + Trim(PartName.Text) + '''');
        if InsNameMode.Itemindex = 2 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''%' + Trim(PartName.Text) + '%''');
     end;

     if IsInsurer.Checked and (InsurType.ItemIndex > 0) then begin
        if InsurType.ItemIndex = 1 then
          ConnectFilter(filter, 'AND', 'insurertype = 1');
        if InsurType.ItemIndex = 2 then
          ConnectFilter(filter, 'AND', 'insurertype = 0');
        if InsurType.ItemIndex = 3 then
          ConnectFilter(filter, 'AND', 'insurertype = 2');
     end;

     //Начало страхования
     if IsInsStart.Checked then begin
        if InsStart1.Checked then
           ConnectFilter(filter, 'AND', 'FromDate >= ''' + DateToStr(InsStart1.Date) + '''');
        if InsStart2.Checked then
           ConnectFilter(filter, 'AND', 'FromDate <= ''' + DateToStr(InsStart2.Date) + '''');
     end;

     //Отчётная дата 1
     if IsRepDate.Checked then begin
        if RepDate1.Checked then
           ConnectFilter(filter, 'AND', 'RET1 >= ''' + DateToStr(RepDate1.Date) + '''');
        if RepDate2.Checked then
           ConnectFilter(filter, 'AND', 'RET1 <= ''' + DateToStr(RepDate2.Date) + '''');
     end;

     //Отчётная дата 1
     if IsUpDT.Checked then begin
        if UpDtFrom.Checked then
           ConnectFilter(filter, 'AND', 'updatedate >= ''' + DateToStr(UpDtFrom.Date) + '''');
        if UpDtTo.Checked then
           ConnectFilter(filter, 'AND', 'updatedate <= ''' + DateToStr(UpDtTo.Date) + '''');
     end;

     //Отчётная дата 2
     if IsRepDate2.Checked then begin
        if Rep2Date1.Checked then
           ConnectFilter(filter, 'AND', 'RET2 >= ''' + DateToStr(Rep2Date1.Date) + '''');
        if Rep2Date2.Checked then
           ConnectFilter(filter, 'AND', 'RET2 <= ''' + DateToStr(Rep2Date2.Date) + '''');
     end;

     //2Й взнос
     if IsFee2Date.Checked then begin
        if Fee2DateFrom.Checked then
           ConnectFilter(filter, 'AND', 'Fee2date >= ''' + DateToStr(Fee2DateFrom.Date) + '''');
        if Fee2DateTO.Checked then
           ConnectFilter(filter, 'AND', 'Fee2date <= ''' + DateToStr(Fee2DateTo.Date) + '''');
     end;

     //Конец страхования
     if IsInsEnd.Checked then begin
        if InsEnd1.Checked then
           ConnectFilter(filter, 'AND', 'ToDate >= ''' + DateToStr(InsEnd1.Date) + '''');
        if InsEnd2.Checked then
           ConnectFilter(filter, 'AND', 'ToDate <= ''' + DateToStr(InsEnd2.Date) + '''');
     end;

     //Период страхования
     if IsPeriod.Checked AND (Period.ItemIndex <> -1) then begin
        ConnectFilter(filter, 'AND', 'Period = ''' + Period.Text + '''');
     end;

     //2й взнос
     if Is2Fee.Checked then begin
        ConnectFilter(filter, 'AND', 'IsFee2=' + IntToStr(SecondFee.ItemIndex));
     end;

     //2й платёж
     if IsPay2.Checked then begin
        ConnectFilter(filter, 'AND', 'IsPay2=' + IntToStr(SecondPay.ItemIndex));
     end;

     //Скидка
     if IsReport.ItemIndex > 0 then begin
        ConnectFilter(filter, 'AND', 'BaseCarCode=' + IntToStr(99 + IsReport.ItemIndex));
     end;

     if IsOther.ItemIndex = 1 then begin
        ConnectFilter(filter, 'AND', 'AgentCode!=''''');
     end;

     if IsOther.ItemIndex = 2 then begin
        ConnectFilter(filter, 'AND', 'AgentCode=''''');
     end;

     //Дата выдачи
     if IsPutDate.Checked then begin
        if PutDate1.Checked then
           ConnectFilter(filter, 'AND', 'RegDate >= ''' + DateToStr(PutDate1.Date) + '''');
        if PutDate2.Checked then
           ConnectFilter(filter, 'AND', 'RegDate <= ''' + DateToStr(PutDate2.Date) + '''');
     end;

     //Номер платёжки
     if IsPayNumber.Checked then begin
        ConnectFilter(filter, 'AND', 'Draftnumber>=' + PayNumber1.Text);
        ConnectFilter(filter, 'AND', 'Draftnumber<=' + PayNumber2.Text);
     end;

     if IsState.Checked AND (States.SelCount > 0) then begin
        for sel_i := 1 to States.Items.Count - 2 do //не учитывать последнее
            if States.Selected[sel_i] then begin
                if state_filt <> '' then
                    state_filt := state_filt + ',';
                state_filt := state_filt + IntToStr(sel_i);
            end;

        if ((not States.Selected[6]) AND (not States.Selected[0])) then
            s := 'STATE IN (' + state_filt + ')'
        else
        if ((States.Selected[6]) AND (States.Selected[0])) then begin
            if state_filt <> '' then
                state_filt := state_filt + ',';
            state_filt := state_filt + '0';
            s := 'STATE IN (' + state_filt + ')'
        end
        else
        if (States.Selected[6]) AND (Not States.Selected[0]) then begin //Дубликат!!!
            if state_filt <> '' then
                s := 'STATE IN (' + state_filt + ') OR STATE=0 AND PREVNUMBER>0'
            else
                s := 'STATE=0 AND PREVNUMBER>0'
        end
        else begin //Номальные, не дубликаты
            if state_filt <> '' then
                s := 'STATE IN (' + state_filt + ') OR STATE=0 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)'
            else
                s := 'STATE=0 AND (PREVNUMBER=0 OR PREVNUMBER IS NULL)'
        end;

        ConnectFilter(filter, 'AND', '(' + s + ')');
     end;


     //Отчётная дата 1
     //Цепляем испорченные
     if IsRepDate.Checked and (RepDate1.Checked or RepDate2.Checked) and (not (IsState.Checked AND (States.SelCount > 0))) and (not IsInsurer.Checked) then begin
        filter := '(' + filter + ') OR ( STATE=1 ';

        if RepDate1.Checked then
           ConnectFilter(filter, 'AND ', 'REGDATE >= ''' + DateToStr(RepDate1.Date) + '''');
        if RepDate2.Checked then
           ConnectFilter(filter, 'AND', 'REGDATE <= ''' + DateToStr(RepDate2.Date) + '''');
           
        filter := filter + ')';
     end;


     Result := filter;
end;

//то же что и GetFilter1 + GetFilter2
function TMandatoryFilterDlg.GetFilter(s : string) : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
{       str := '((PAYDATE IS NULL AND REGDATE BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''') OR ' +
              '(PAYDATE IS NOT NULL AND PAYDATE BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''') OR ' +
              '(PAY2DATE IS NOT NULL AND PAY2DATE BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + '''))';}
       str := '((RET1 IS NULL AND REGDATE BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + ''') OR ' +
              '(RET1 BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + ''') OR ' +
              '(RET2 BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + '''))';
       if IsStopped.Checked then
            str := '(' + str + ' OR STOPDATE BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + ''')';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter(s)) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter(s) + ')');

     if IsAdditionalPay.Checked then begin
          ConnectFilter(filter, 'AND', '(TX.N IS NOT NULL)');
     end;

     Result := filter;
end;

//Испорченные в период и оплаченные 1-й этап(или всё) в этот же период
function TMandatoryFilterDlg.GetFilter1 : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
       str := '((PAYDATE IS NULL AND REGDATE BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + ''') OR ' +
              '(RET1 BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + '''))';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter('')) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter('') + ')');

     Result := filter;
end;

// Второй платёж был в определённый период
function TMandatoryFilterDlg.GetFilter2 : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
       str := '(RET2 BETWEEN ''' +
              DateToStr(Pay1Start_.Date) + ''' AND ''' +
              DateToStr(Pay1End_.Date) + ''')';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter('')) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter('') + ')');

     Result := filter;
end;

procedure TMandatoryFilterDlg.ConnectFilter(var filter : string; Operation, Condition : string);
begin
     if Length(filter) <> 0 then
        filter := filter + ' ' + Operation + ' ';
     filter := filter + Condition;
end;

procedure TMandatoryFilterDlg.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
     str : string;
begin
    IsReport.ItemIndex := 0;
    INI := TRegIniFile.Create('Обязательное страхование');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
     end;
     INI.Free;

     ListTemplates.ItemIndex := 0;
     AgCodes := TStringList.Create;
     AgentsTbl.Open;
     while not AgentsTbl.EOF do begin
           AgentList.Items.Add(AgentsTblName.AsString);
           AgCodes.Add(AgentsTblAgent_code.AsString);
           AgentsTbl.Next;
     end;
     AgentsTbl.Close;
     InsNameMode.ItemIndex := 0;
     SecondFee.ItemIndex := 0;
     SecondPay.ItemIndex := 0;
end;

procedure TMandatoryFilterDlg.FormResize(Sender: TObject);
begin
    Divider.Height := Height - DataPanel.Height - DownPanel.Height - GetSystemMetrics(sm_CyCaption) - 5;
    AgentList.Height := Height - DataPanel.Height - DownPanel.Height - AgentList.Top - GetSystemMetrics(sm_CyCaption) - 10;
end;

procedure TMandatoryFilterDlg.FormShow(Sender: TObject);
var
    INI : TIniFile;
    rstr : string;
    i : integer;
begin
     if Period.Items.Count = 0 then begin
         INI := TIniFile.Create(BLANK_INI);
         for i := 0 to 999 do begin
             rstr := INI.ReadString('MANDATORY', 'Period' + IntToStr(i), '');
             if rstr = '' then break;
             Period.Items.Add(rstr);
         end;
         Period.ItemIndex := 0;
         INI.Free;
     end;

     AgentListClick(nil);
{     if Resident.Items.Count = 0 then begin
         INI := TIniFile.Create(BLANK_INI);
         for i := 0 to 99999 do begin
             rstr := INI.ReadString('MANDATORY', 'TarifTable' + IntToStr(i), '');
             if rstr = '' then break;
             Resident.Items.Add( Copy(rstr, 1, Pos('|', rstr) - 1) );
         end;
         Resident.ItemIndex := 0;
         INI.Free;
     end;}
end;

procedure TMandatoryFilterDlg.AgentListClick(Sender: TObject);
var
     Count, i : integer;
begin
     Count := 0;
     for i := 0 to AgentList.Items.Count - 1 do
         if AgentList.Checked[i] then
             Inc(Count);
     CntChecked.Caption := 'Помечено ' + IntToStr(Count) + ' из ' + IntToStr(AgentList.Items.Count);

     if Sender <> nil then
         ListTemplates.ItemIndex := -1;
end;

procedure TMandatoryFilterDlg.ListTemplatesChange(Sender: TObject);
var
     i, fnd_idx, fnd_token : integer;
     INI : TRegIniFile;
     str : string;
     Code : string;
begin
     //if ListTemplates.ItemIndex = 0 then begin
       for i := 0 to AgentList.Items.Count - 1 do
            AgentList.Checked[i] := false;
       //exit;
     //end;

     Code := '';
     INI := TRegIniFile.Create('Обязательное страхование');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str = ListTemplates.Text then begin
             str := INI.ReadString('AgTemplates', IntToStr(i), '');

             for fnd_token := 1 to Length(str) do begin
                 if str[fnd_token] = ',' then begin
                     for fnd_idx := 0 to AgentList.Items.Count - 1 do
                         if AgCodes[fnd_idx] = Code then
                             AgentList.Checked[fnd_idx] := true;

                     Code := '';
                 end
                 else
                     Code := Code + str[fnd_token];
             end;
         end;
     end;
     INI.Free;

     AgentListClick(nil);
end;

procedure TMandatoryFilterDlg.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TMandatoryFilterDlg.N1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Обязательное страхование');
     if GetTemplName.ShowModal = mrOk then begin
         for i := 0 to 100 do
           if INI.ReadString('AgTemplates', IntToStr(i), '') = '' then break;
         index := i;
         str := '';
         for i := 0 to AgentList.Items.Count - 1 do
             if AgentList.Checked[i] then
                 str := str + AgCodes[i] + ',';
         INI.WriteString('AgTemplates', IntToStr(index), str);
         INI.WriteString('AgTemplates', 'Name' + IntToStr(index), GetTemplName.Name.Text);
         ListTemplates.ItemIndex := ListTemplates.Items.Add(GetTemplName.Name.Text);
     end;
     INI.Free;
end;

procedure TMandatoryFilterDlg.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Обязательное страхование');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

procedure TMandatoryFilterDlg.FormDestroy(Sender: TObject);
begin
    AgCodes.Free
end;

end.
