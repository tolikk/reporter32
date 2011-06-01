unit LebenFilterUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, CheckLst, ComCtrls, Buttons, ExtCtrls, Db, DBTables, Placemnt,
  Menus;

type
  TLebenFilter = class(TForm)
    Label2: TLabel;
    DateFrom: TDateTimePicker;
    DateTo: TDateTimePicker;
    Label3: TLabel;
    AgentList: TCheckListBox;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    IsAgent: TCheckBox;
    CntChecked: TLabel;
    AgentsTbl: TQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    StartFrom: TDateTimePicker;
    Label4: TLabel;
    StartTo: TDateTimePicker;
    Label5: TLabel;
    EndFrom: TDateTimePicker;
    Label6: TLabel;
    EndTo: TDateTimePicker;
    Label1: TLabel;
    Label7: TLabel;
    DurationsFrom: TEdit;
    DurationsTo: TEdit;
    IsShowBad: TCheckBox;
    FormStorage: TFormStorage;
    ListTemplates: TComboBox;
    SpeedButton: TSpeedButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    Label8: TLabel;
    ListAssist: TCheckListBox;
    IsState: TCheckBox;
    State: TComboBox;
    IsHaveUbitok: TCheckBox;
    listInsType: TComboBox;
    IsInsurerType: TCheckBox;
    Label9: TLabel;
    StopDateFrom: TDateTimePicker;
    Label10: TLabel;
    StopDateTo: TDateTimePicker;
    procedure FormCreate(Sender: TObject);
    procedure AgentListClickCheck(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
    procedure AgentListClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    AgCodes : TStringList;
    { Private declarations }
    procedure ConnectFilter(var filter : string; Operation, Condition : string);
  public
    function getFilterStr : string;
    { Public declarations }
  end;

var
  LebenFilter: TLebenFilter;

implementation

uses TmplNameUnit, registry, LebenUnit;

{$R *.DFM}

procedure TLebenFilter.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
     str : string;
begin
     State.ItemIndex := 0;
     AgCodes := TStringList.Create;
     AgentsTbl.Open;
     while not AgentsTbl.EOF do begin
           AgentList.Items.Add(AgentsTblName.AsString);
           AgCodes.Add(AgentsTblAgent_code.AsString);
           AgentsTbl.Next;
     end;
     AgentsTbl.Close;

     INI := TRegIniFile.Create('Страхование жизни');
     for i := 0 to 100 do begin
         str := INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '');
         if str <> '' then
             ListTemplates.Items.Add(str);
         //else
         //    break;
     end;
     INI.Free;
end;

procedure TLebenFilter.AgentListClickCheck(Sender: TObject);
var
     Count, i : integer;
begin
     Count := 0;
     for i := 0 to AgentList.Items.Count - 1 do
         if AgentList.Checked[i] then
             Inc(Count);
     CntChecked.Caption := 'Помечено ' + IntToStr(Count);
end;

function TLebenFilter.getFilterStr : string;
var
     str, filter, bad_s : string;
     i : integer;
begin
     str := '';
     bad_s := '';
     filter := '';

     if DateFrom.Checked then
        ConnectFilter(filter, 'AND', 'PaymentDate >= ''' + DateToStr(DateFrom.Date) + '''');
     if DateTo.Checked then
        ConnectFilter(filter, 'AND', 'PaymentDate <= ''' + DateToStr(DateTo.Date) + '''');

     if StopDateFrom.Checked then
        ConnectFilter(filter, 'AND', 'StopDate >= ''' + DateToStr(StopDateFrom.Date) + '''');
     if StopDateTo.Checked then
        ConnectFilter(filter, 'AND', 'StopDate <= ''' + DateToStr(StopDateTo.Date) + '''');

     if StartFrom.Checked then
        ConnectFilter(filter, 'AND', 'P."Begin" >= ''' + DateToStr(StartFrom.Date) + '''');
     if StartTo.Checked then
        ConnectFilter(filter, 'AND', 'P."Begin" <= ''' + DateToStr(StartTo.Date) + '''');

     if EndFrom.Checked then
        ConnectFilter(filter, 'AND', 'P."End" >= ''' + DateToStr(EndFrom.Date) + '''');
     if EndTo.Checked then
        ConnectFilter(filter, 'AND', 'P."End" <= ''' + DateToStr(EndTo.Date) + '''');

     if DurationsFrom.Text <> '' then
        ConnectFilter(filter, 'AND', 'Duration >= ''' + DurationsFrom.Text + '''');
     if DurationsTo.Text <> '' then
        ConnectFilter(filter, 'AND', 'Duration <= ''' + DurationsTo.Text + '''');

     if IsState.Checked and (State.ItemIndex >= 0) then begin
         ConnectFilter(filter, 'AND', 'St=''' + IntToStr(State.ItemIndex) + '''')
     end;

     if IsInsurerType.Checked then begin
         if listInsType.ItemIndex = 0 then
             ConnectFilter(filter, 'AND', 'Insurtype=''F''')
         else
             ConnectFilter(filter, 'AND', 'Insurtype=''U''')
     end;

     if filter <> '' then begin
         if IsShowBad.Checked then begin
             bad_s := ' WriteDate is NULL ';
             if DateFrom.Checked then
                ConnectFilter(bad_s, 'AND', 'PaymentDate >= ''' + DateToStr(DateFrom.Date) + '''');
             if DateTo.Checked then
                ConnectFilter(bad_s, 'AND', 'PaymentDate <= ''' + DateToStr(DateTo.Date) + '''');

             filter := '((' + filter + ') OR (' + bad_s + '))'
         end
     end;

     //АГЕНТ
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if (Length(str) > 0) AND IsAgent.Checked then
          ConnectFilter(filter, 'AND', '(Agent_Code IN (' + Copy(str, 1, Length(str) - 1) + ')) ');

     //АССИСТАНС
     str := '';
     for i := 0 to ListAssist.Items.Count - 1 do
          if ListAssist.Checked[i] then begin
             if str <> '' then str := str + ' OR ';
             if i = 0 then
                 str := str + 'Restrax is null or Restrax=0'
             else
                 str := str + 'Restrax=' + IntToStr(i);
          end;

     if Length(str) > 0 then
          ConnectFilter(filter, 'AND', '(' + str + ')');

     if IsHaveUbitok.Checked then
          ConnectFilter(filter, 'AND', 'EXISTS (SELECT 1 FROM PRIV2UB WHERE SER=P.SERIA AND NMB=P.NUMBER)');

     if filter <> '' then begin
         {if IsShowBad.Checked then begin
             bad_s := ' WriteDate is NULL ';
             if DateFrom.Checked then
                ConnectFilter(bad_s, 'AND', 'PaymentDate >= ''' + DateToStr(DateFrom.Date) + '''');
             if DateTo.Checked then
                ConnectFilter(bad_s, 'AND', 'PaymentDate <= ''' + DateToStr(DateTo.Date) + '''');

             getFilterStr := '(' + filter + ') OR (' + bad_s + ')'
         end
         else}
             getFilterStr := filter
     end
     else
         getFilterStr := '1=1'
end;

procedure TLebenFilter.ConnectFilter(var filter : string; Operation, Condition : string);
begin
     if Length(filter) <> 0 then
        filter := filter + ' ' + Operation + ' ';
     filter := filter + Condition;
end;

procedure TLebenFilter.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top + SpeedButton.Height;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TLebenFilter.ListTemplatesChange(Sender: TObject);
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
     INI := TRegIniFile.Create('Страхование жизни');
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

procedure TLebenFilter.AgentListClick(Sender: TObject);
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

procedure TLebenFilter.N1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Страхование жизни');
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

procedure TLebenFilter.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Страхование жизни');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

procedure TLebenFilter.FormDestroy(Sender: TObject);
begin
    AgCodes.Free
end;

procedure TLebenFilter.FormShow(Sender: TObject);
var
     i : integer;
begin
     if ListAssist.Items.Count = 0 then begin
        for i := 0 to 10 do
            if (InsLeben2.AssistNames[i] <> '') then
                ListAssist.Items.Add(InsLeben2.AssistNames[i])
            else
                break;
     end;
end;

end.
