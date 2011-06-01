unit FilterMandSQL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Placemnt, ComCtrls, CheckLst, Db, DBTables, Buttons,
  RXSpin, inifiles, Menus, RxQuery;

type
  TMandSQLFilterDlg = class(TForm)
    DownPanel: TPanel;
    Panel2: TPanel;
    DataPanel: TPanel;
    IsPayDate: TCheckBox;
    Label1: TLabel;
    Pay1Start: TDateTimePicker;
    Label2: TLabel;
    Pay1End: TDateTimePicker;
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
    States: TListBox;
    IsState: TCheckBox;
    Label17: TLabel;
    AgentsTbl: TRxQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    IsPayDoc: TCheckBox;
    Label13: TLabel;
    Label14: TLabel;
    Pay1_1: TEdit;
    Pay1_2: TEdit;
    Label18: TLabel;
    Pay2_1: TEdit;
    Label19: TLabel;
    Pay2_2: TEdit;
    IsDup: TCheckBox;
    IsSpecial: TCheckBox;
    SpecCombo: TComboBox;
    Params: TEdit;
    Label20: TLabel;
    IsAvar: TCheckBox;
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
    SpecList : TStringList;
  public
    { Public declarations }
    function  GetFilter : string;
    function  GetFilter1 : string;
    function  GetFilter2 : string;
    procedure ConnectFilter(var filter : string; Operation, Condition : string);
    function  GetCommonFilter : string;
  end;

var
  MandSQLFilterDlg: TMandSQLFilterDlg;

implementation

uses MF, TmplNameUnit, registry, strutils, RxStrUtils;

{$R *.DFM}

function TMandSQLFilterDlg.GetCommonFilter : string;
var
     filter : string;
     str, str2, state_filt, s : string;
     i, sel_i : integer;
begin
     filter := '';

     if IsAvar.Checked then begin
         ConnectFilter(filter, 'AND', 'EXISTS (SELECT 1 FROM SYSADM.MANDAV A WHERE M.SER=A.SER AND M.NMB=A.NMB)');
     end;

     if IsWasActive.Checked then begin
         //Начали действовать в период или испорчены в период
         ConnectFilter(filter, 'AND', '(' +
                                          '(STATE <> 3 AND FRDT BETWEEN ''' + DateToStr(WasActiveFrom.Date) + ''' AND ''' + DateToStr(WasActiveTo.Date) + ''')' +
                                          ' OR ' +
                                          '(PAY1DT IS NULL AND REGDATE BETWEEN ''' + DateToStr(WasActiveFrom.Date) + ''' AND ''' + DateToStr(WasActiveTo.Date) + ''')' +
                                      ')');
     end;

     //Агенты
     str := '';
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if Length(str) > 0 then
          ConnectFilter(filter, 'AND', '(AGNT IN (' + Copy(str, 1, Length(str) - 1) + '))');

     if IsDup.Checked then
          ConnectFilter(filter, 'AND', ' (PNMB > 0) ');
     //Страхователь
     if (IsInsurer.Checked) AND (Trim(PartName.Text) <> '') then begin
        if InsNameMode.Itemindex = 0 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''' + Trim(PartName.Text) + '%''');
        if InsNameMode.Itemindex = 1 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''%' + Trim(PartName.Text) + '''');
        if InsNameMode.Itemindex = 2 then
          ConnectFilter(filter, 'AND', 'Name LIKE ''%' + Trim(PartName.Text) + '%''');
     end;

     //Начало страхования
     if IsInsStart.Checked then begin
        if InsStart1.Checked then
           ConnectFilter(filter, 'AND', 'FrDt >= ''' + DateToStr(InsStart1.Date) + '''');
        if InsStart2.Checked then
           ConnectFilter(filter, 'AND', 'FrDt <= ''' + DateToStr(InsStart2.Date) + '''');
     end;

     //2Й взнос
     if IsFee2Date.Checked then begin
        if Fee2DateFrom.Checked then
           ConnectFilter(filter, 'AND', 'Fee2dt >= ''' + DateToStr(Fee2DateFrom.Date) + '''');
        if Fee2DateTO.Checked then
           ConnectFilter(filter, 'AND', 'Fee2dt <= ''' + DateToStr(Fee2DateTo.Date) + '''');
     end;

     //Конец страхования
     if IsInsEnd.Checked then begin
        if InsEnd1.Checked then
           ConnectFilter(filter, 'AND', 'ToDt >= ''' + DateToStr(InsEnd1.Date) + '''');
        if InsEnd2.Checked then
           ConnectFilter(filter, 'AND', 'ToDt <= ''' + DateToStr(InsEnd2.Date) + '''');
     end;

     //2й взнос
     if Is2Fee.Checked then begin
        if SecondFee.ItemIndex = 0 then
            ConnectFilter(filter, 'AND', 'IsFee2=''N''')
        else
            ConnectFilter(filter, 'AND', 'IsFee2=''Y''');
     end;

     //2й платёж
     if IsPay2.Checked then begin
        if SecondPay.ItemIndex = 0 then
            ConnectFilter(filter, 'AND', 'IsPay2=''N''')
        else
            ConnectFilter(filter, 'AND', 'IsPay2=''Y''');
     end;

     //Дата выдачи
     if IsPutDate.Checked then begin
        if PutDate1.Checked then
           ConnectFilter(filter, 'AND', 'RegDate >= ''' + DateToStr(PutDate1.Date) + '''');
        if PutDate2.Checked then
           ConnectFilter(filter, 'AND', 'RegDate <= ''' + DateToStr(PutDate2.Date) + '''');
     end;

     if IsState.Checked AND (States.SelCount > 0) then begin
        for sel_i := 0 to States.Items.Count - 1 do
            if States.Selected[sel_i] then begin
                if state_filt <> '' then
                    state_filt := state_filt + ',';
                state_filt := state_filt + IntToStr(sel_i);
            end;
        ConnectFilter(filter, 'AND', 'STATE IN (' + state_filt + ')');
     end;

     if IsPayDoc.Checked then begin
         str := '';
         if Trim(Pay1_1.Text) <> '' then
             ConnectFilter(str, 'AND', 'PAY1DOC>=' + Trim(Pay1_1.Text));
         if Trim(Pay1_2.Text) <> '' then
             ConnectFilter(str, 'AND', 'PAY1DOC<=' + Trim(Pay1_2.Text));

         str2 := '';
         if Trim(Pay2_1.Text) <> '' then
             ConnectFilter(str2, 'AND', 'PAY2DOC>=' + Trim(Pay2_1.Text));
         if Trim(Pay2_2.Text) <> '' then
             ConnectFilter(str2, 'AND', 'PAY2DOC<=' + Trim(Pay2_2.Text));

          if (str <> '') AND (str2 <> '') then begin
             str := '(' + str + ')';
             str2 := '(' + str2 + ')';
             ConnectFilter(str, 'OR', str2);
             str := '(' + str + ')';
          end else
          if str = '' then
             str := str2;

          ConnectFilter(filter, 'AND', str);
     end;

     if (IsSpecial.Checked) AND (SpecCombo.ItemIndex >= 0) then begin
          s := SpecList[SpecCombo.ItemIndex];
          i := 1;
          while true do begin
              if Pos('::' + IntToStr(i), s) <> 0 then begin
                  s := Copy(s, 1, Pos('::' + IntToStr(i), s) - 1) + ExtractWord(i, Params.Text, [',']) + Copy(s, Pos('::' + IntToStr(i), s) + 3, Length(s));
              end
              else
                  Inc(i);
              if i > 5 then break;
          end;
          ConnectFilter(filter, 'AND', s);
     end;

     Result := filter;
end;

//то же что и GetFilter1 + GetFilter2
function TMandSQLFilterDlg.GetFilter : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
       str := '((PAY1DT IS NULL AND REGDATE BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''') OR ' +
              '(PAY1DT IS NOT NULL AND PAY1DT BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''') OR ' +
              '(PAY2DT IS NOT NULL AND PAY2DT BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + '''))';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter + ')');

     Result := filter;
end;

//Испорченные в период и оплаченные 1-й этап(или всё) в этот же период
function TMandSQLFilterDlg.GetFilter1 : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
       str := '((PAY1DT IS NULL AND REGDATE BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''') OR ' +
              '(PAY1DT IS NOT NULL AND PAY1DT BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + '''))';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter + ')');

     Result := filter;
end;

// Второй платёж был в определённый период
function TMandSQLFilterDlg.GetFilter2 : string;
var
     filter : string;
     str : string;
begin
     filter := '';

     if(IsPayDate.Checked) then begin
       str := '(PAY2DT IS NOT NULL AND PAY2DT BETWEEN ''' +
              DateToStr(Pay1Start.Date) + ''' AND ''' +
              DateToStr(Pay1End.Date) + ''')';
       ConnectFilter(filter, 'AND', str);
     end;

     if Length(GetCommonFilter) > 0 then
          ConnectFilter(filter, 'AND', '(' + GetCommonFilter + ')');

     Result := filter;
end;

procedure TMandSQLFilterDlg.ConnectFilter(var filter : string; Operation, Condition : string);
begin
     if Length(filter) <> 0 then
        filter := filter + ' ' + Operation + ' ';
     filter := filter + Condition;
end;

procedure TMandSQLFilterDlg.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     INIF : TIniFile;
     i : integer;
     str, s : string;
begin
     SpecList := TStringList.Create;

     INI := TRegIniFile.Create('Обязательное страхование');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
     end;
     INI.Free;

     AgentsTbl.DatabaseName := 'SQLDB';
     AgentsTbl.MacroByName('TABLE').AsString := 'SYSADM.AGENT';
     INIF := TIniFile.Create(BLANK_INI);
     if INIF.ReadString('MANDATORY', 'LOCAL_DB', '') <> '' then begin
         AgentsTbl.DatabaseName := 'LOCALDB';
         AgentsTbl.MacroByName('TABLE').AsString := 'AGENT';
     end;

     for i := 0 to 100 do begin
         s := INIF.ReadString('MANDATORY', 'SpecFilter' + IntToStr(i), '');
         if (s = '') OR (Pos(';', s) = 0) then break;
         SpecCombo.Items.Add(Copy(s, 1, Pos(';', s) - 1));
         SpecList.Add(Copy(s, Pos(';', s) + 1, Length(s)));
     end;
     INIF.Free;

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

procedure TMandSQLFilterDlg.FormResize(Sender: TObject);
begin
    Divider.Height := Height - DataPanel.Height - DownPanel.Height - GetSystemMetrics(sm_CyCaption) - 5;
    AgentList.Height := Height - DataPanel.Height - DownPanel.Height - AgentList.Top - GetSystemMetrics(sm_CyCaption) - 10;
end;

procedure TMandSQLFilterDlg.FormShow(Sender: TObject);
var
    INI : TIniFile;
    rstr : string;
    i : integer;
begin
{     if Period.Items.Count = 0 then begin
         INI := TIniFile.Create(BLANK_INI);
         for i := 0 to 999 do begin
             rstr := INI.ReadString('MANDATORY', 'Period' + IntToStr(i), '');
             if rstr = '' then break;
             Period.Items.Add(rstr);
         end;
         Period.ItemIndex := 0;
         INI.Free;
     end;
}
     AgentListClick(nil);
end;

procedure TMandSQLFilterDlg.AgentListClick(Sender: TObject);
var
     Count, i : integer;
begin
     Count := 0;
     for i := 0 to AgentList.Items.Count - 1 do
         if AgentList.Checked[i] then
             Inc(Count);
     CntChecked.Caption := 'Помечено ' + IntToStr(Count);

     if Sender <> nil then
         ListTemplates.ItemIndex := -1;
end;

procedure TMandSQLFilterDlg.ListTemplatesChange(Sender: TObject);
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

procedure TMandSQLFilterDlg.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TMandSQLFilterDlg.N1Click(Sender: TObject);
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

procedure TMandSQLFilterDlg.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
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

procedure TMandSQLFilterDlg.FormDestroy(Sender: TObject);
begin
    AgCodes.Free;
    SpecList.Free
end;

end.
