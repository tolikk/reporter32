unit PolandFilterUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls, DBCtrls, Db, DBTables, RxDBComb, RXLookup,
  CheckLst, Menus, Buttons, registry, Placemnt;

type
  TPolandFilter = class(TForm)
    IsNumber: TCheckBox;
    IsRegDate: TCheckBox;
    Label1: TLabel;
    NumberFrom: TEdit;
    Label2: TLabel;
    NumberTo: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    RegDateFrom: TDateTimePicker;
    RegDateTo: TDateTimePicker;
    IsStartDate: TCheckBox;
    Label5: TLabel;
    StartDateFrom: TDateTimePicker;
    Label6: TLabel;
    StartDateTo: TDateTimePicker;
    Button1: TButton;
    Bevel1: TBevel;
    Button2: TButton;
    IsOwner: TCheckBox;
    OwnerLike: TEdit;
    IsAutoNumber: TCheckBox;
    AutoNumberLike: TEdit;
    IsWasActive: TCheckBox;
    Label8: TLabel;
    WasActiveFrom: TDateTimePicker;
    Label9: TLabel;
    WasActiveTo: TDateTimePicker;
    IsOformDate: TCheckBox;
    Label7: TLabel;
    OformDateFrom: TDateTimePicker;
    Label10: TLabel;
    OformDateTo: TDateTimePicker;
    AgentList: TCheckListBox;
    Label11: TLabel;
    CntChecked: TLabel;
    ListTemplates: TComboBox;
    SpeedButton: TSpeedButton;
    PopupMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    AgentsTbl: TQuery;
    AgentsTblAGENT_CODE: TStringField;
    AgentsTblNAME: TStringField;
    FormStorage: TFormStorage;
    IsState: TCheckBox;
    State_: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure SpeedButtonClick(Sender: TObject);
    procedure AgentListClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure ListTemplatesChange(Sender: TObject);
  private
    AgCodes : TStringList;
    procedure Connect(var s: string; cond : string);
    { Private declarations }
  public
    function FilterText : string;
    function GetPeriodText : string;
    { Public declarations }
  end;

var
  PolandFilter: TPolandFilter;

implementation

uses TmplNameUnit;

{$R *.DFM}

procedure TPolandFilter.Connect(var s: string; cond : string);
begin
     if s <> '' then
         s := s + ' AND ';
     s := s + cond;
end;

function TPolandFilter.FilterText : string;
var
     s, str : string;
     i : integer;
begin
     if IsNumber.Checked then begin
        if Trim(NumberFrom.Text) <> '' then
            Connect(s, 'NUMBER>=' + NumberFrom.Text);
        if Trim(NumberTo.Text) <> '' then
            Connect(s, 'NUMBER<=' + NumberTo.Text);
     end;

     if IsRegDate.Checked then begin
        if RegDateFrom.Checked then
            Connect(s, 'REPDATE>=''' + DateToStr(RegDateFrom.Date) + '''');
        if RegDateTo.Checked then
            Connect(s, 'REPDATE<=''' + DateToStr(RegDateTo.Date) + '''');
     end;

     if IsOformDate.Checked then begin
        if OformDateFrom.Checked then
            Connect(s, 'PAYDATE>=''' + DateToStr(OformDateFrom.Date) + '''');
        if OformDateTo.Checked then
            Connect(s, 'PAYDATE<=''' + DateToStr(OformDateTo.Date) + '''');
     end;

     if IsStartDate.Checked then begin
        if StartDateFrom.Checked then
            Connect(s, 'STARTDATE>=''' + DateToStr(StartDateFrom.Date) + '''');
        if StartDateTo.Checked then
            Connect(s, 'STARTDATE<=''' + DateToStr(StartDateTo.Date) + '''');
     end;

     //Агенты
     str := '';
     for i := 0 to AgentList.Items.Count - 1 do
          if AgentList.Checked[i] then
             str := str + '''' + AgCodes[i] + ''',';
     if Length(str) > 0 then
          Connect(s, '(AgentCode IN (' + Copy(str, 1, Length(str) - 1) + '))');
//     if IsAgent.Checked then begin
//         if SelAgent.Text <> '' then
//             Connect(s, 'AGENTCODE=''' + Agent.FieldByName('AGENT_CODE').AsString + '''');
//     end;

     if IsOwner.Checked AND (OwnerLike.Text <> '') then begin
         Connect(s, '(NAME LIKE ''%' + OwnerLike.Text + '%'' OR NAME LIKE ''%' + AnsiUpperCase(OwnerLike.Text) + '%'' OR NAME LIKE ''%' + AnsiLowerCase(OwnerLike.Text) + '%'')');
     end;

     if IsAutoNumber.Checked AND (AutoNumberLike.Text <> '') then begin
         Connect(s, '(AutoNumber LIKE ''%' + AutoNumberLike.Text + '%'' OR AutoNumber LIKE ''%' + AnsiUpperCase(AutoNumberLike.Text) + '%'' OR AutoNumber LIKE ''%' + AnsiLowerCase(AutoNumberLike.Text) + '%'')');
     end;

     if IsState.Checked then begin
         //if State.ItemIndex <> 0 then
         //    Connect(s, '(Name = ''ИСПОРЧЕН'')')
         //else
         //    Connect(s, '(Name <> ''ИСПОРЧЕН'')')
         Connect(s, '(State = ''' + IntToStr(State_.ItemIndex) + ''')')
     end;

     //Начали действовать или испорчены
     if IsWasActive.Checked then begin
         Connect(s, '((STARTDATE>=''' + DateToStr(WasActiveFrom.Date) + ''' AND ' +
                    'STARTDATE<=''' + DateToStr(WasActiveTo.Date) + ''') OR ' +
                    '(PremiumPay is null AND REPDATE>=''' + DateToStr(WasActiveFrom.Date) + ''' AND ' +
                    'REPDATE<=''' + DateToStr(WasActiveTo.Date) + '''))');
     end;

     if s <> '' then s := ' WHERE ' + s;
     FilterText := s;
end;

procedure TPolandFilter.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
     str : string;
begin
     State_.ItemIndex := 0;
     INI := TRegIniFile.Create('Польша');
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
end;

function TPolandFilter.GetPeriodText : string;
var
     s : string;
begin
     s := '';
     if IsRegDate.Checked then begin
        if RegDateFrom.Checked then
            s := 'C ' + DateToStr(RegDateFrom.Date);
        if RegDateTo.Checked then begin
            s := s + ' ПО ' + DateToStr(RegDateTo.Date);
        end;
     end;

     GetPeriodText := s;
end;

procedure TPolandFilter.SpeedButtonClick(Sender: TObject);
var
     p : TPOINT;
begin
     p.x := SpeedButton.Left;
     p.y := SpeedButton.Top;
     p := ClientToScreen(p);
     PopupMenu.Popup(p.x, p.y);
end;

procedure TPolandFilter.AgentListClick(Sender: TObject);
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

procedure TPolandFilter.FormShow(Sender: TObject);
begin
     AgentListClick(nil);
end;

procedure TPolandFilter.N2Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     if MessageDlg('Удалить набор ' + ListTemplates.Text + '?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then exit;

     INI := TRegIniFile.Create('Польша');
     for i := 0 to 100 do
         if INI.ReadString('AgTemplates', 'Name' + IntToStr(i), '') = ListTemplates.Text then begin
             INI.DeleteKey('AgTemplates', IntToStr(i));
             INI.DeleteKey('AgTemplates', 'Name' + IntToStr(i));
             ListTemplates.Items.Delete(ListTemplates.ItemIndex);
             break;
         end;
     INI.Free;
end;

procedure TPolandFilter.N1Click(Sender: TObject);
var
     INI : TRegIniFile;
     i, index : integer;
     str : string;
begin
     INI := TRegIniFile.Create('Польша');
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

procedure TPolandFilter.ListTemplatesChange(Sender: TObject);
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
     INI := TRegIniFile.Create('Польша');
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

end.
