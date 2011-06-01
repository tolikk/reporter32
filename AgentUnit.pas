unit AgentUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Db, DBTables, Grids, DBGrids, RXDBCtrl,
  RxQuery, ActnList, DataMod;

type
  TF_Agent = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    RxDBGrid: TRxDBGrid;
    Database: TDatabase;
    DataSource: TDataSource;
    Bevel1: TBevel;
    LeftDivider: TBevel;
    RightDivider: TBevel;
    Timer: TTimer;
    LabelLeft: TLabel;
    Panel3: TPanel;
    CloseBtn: TBitBtn;
    SaveButton: TBitBtn;
    LabelRight: TLabel;
    AddBtn: TBitBtn;
    DelBtn: TBitBtn;
    CancelCh: TBitBtn;
    MainQuery: TRxQuery;
    MainQueryAGENT_CODE: TStringField;
    MainQueryNAME: TStringField;
    MainQueryADDRESS: TStringField;
    MainQueryPHONE: TStringField;
    MainQueryCONTACT: TStringField;
    MainQueryDIVISIONCODE: TFloatField;
    MainQueryTYPE: TStringField;
    MainQueryMANDPCNT: TFloatField;
    MainQueryFOREIGNPCNT: TFloatField;
    MainQueryLEBEN2PCNT: TFloatField;
    MainQueryBULSTRADPCNT: TFloatField;
    UpdateSQL: TUpdateSQL;
    ActionList: TActionList;
    ActionSave: TAction;
    ActionCancel: TAction;
    MainQueryBELGREEN: TFloatField;
    Label1: TLabel;
    MainQueryRUSS: TFloatField;
    MainQueryPOLANDPCNT: TFloatField;
    MainQuerySOVAG: TFloatField;
    MainQueryUID: TStringField;
    FreeAgents: TQuery;
    Panel4: TPanel;
    LostAgents: TQuery;
    Panel5: TPanel;
    btnLost: TButton;
    btnFreeAgents: TButton;
    CheckAgent: TQuery;
    AgFilter: TEdit;
    Label2: TLabel;
    IdxTable: TTable;
    procedure CloseBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure SaveButtonClick(Sender: TObject);
    procedure AddBtnClick(Sender: TObject);
    procedure DelBtnClick(Sender: TObject);
    procedure MainQueryVTYPEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure MainQueryBeforeInsert(DataSet: TDataSet);
    procedure MainQueryTYPESetText(Sender: TField; const Text: String);
    procedure MainQueryTYPEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure MainQueryBeforeDelete(DataSet: TDataSet);
    procedure MainQueryAfterDelete(DataSet: TDataSet);
    procedure MainQueryAfterInsert(DataSet: TDataSet);
    procedure MainQueryAfterPost(DataSet: TDataSet);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure CancelChClick(Sender: TObject);
    procedure NameFilterChange(Sender: TObject);
    procedure ActionSaveUpdate(Sender: TObject);
    procedure MainQueryCalcFields(DataSet: TDataSet);
    procedure btnFreeAgentsClick(Sender: TObject);
    procedure btnLostClick(Sender: TObject);
    procedure AgFilterChange(Sender: TObject);
    procedure AgFilterEnter(Sender: TObject);
    procedure AgFilterExit(Sender: TObject);
  private
    { Private declarations }
    procedure CreateIdx;
  public
    { Public declarations }
  end;

var
  F_Agent: TF_Agent;

implementation

uses inifiles, appmenu, QClipbrd;

{$R *.DFM}

procedure TF_Agent.CloseBtnClick(Sender: TObject);
begin
     Close;
end;
procedure TF_Agent.CreateIdx;
var
    FoundIdx : Boolean;
    I : integer;
begin
    try
        IdxTable.TableName := 'MANDATOR';
//        MANDATOR.Exclusive := True;
//        MANDATOR.Active := False;
//        IdxTable.Active := True;
        IdxTable.IndexDefs.Update;
        FoundIdx := False;
        for I := 0 to IdxTable.IndexDefs.Count - 1 do
          if IdxTable.IndexDefs.Items[I].Fields = 'Agentcode' then
            FoundIdx := True;
        if not FoundIdx then
            IdxTable.AddIndex('', 'Agentcode', []);

        FoundIdx := False;
        for I := 0 to IdxTable.IndexDefs.Count - 1 do
          if IdxTable.IndexDefs.Items[I].Fields = 'Fee2text' then
            FoundIdx := True;
        if not FoundIdx then
            IdxTable.AddIndex('', 'Fee2text', []);
    finally
        IdxTable.close;
    end;
end;

procedure TF_Agent.FormCreate(Sender: TObject);
var
     INI : TIniFile;
     IsMain : Boolean;
     i : integer;
begin
     Database.Close;
     INI := TIniFile.Create(BLANK_INI);
     Database.LoginPrompt := StartParam = 'AGENTSSQL';
     //INI.ReadInteger('AGENTS', 'ISLOCAL', 1) <> 1;
     //s := INI.ReadString('AGENTS', 'QUOTE', '');
     if not Database.LoginPrompt then begin
         MainQuery.MacroByName('TABLE').AsString := 'AGENT';
         Caption := Caption + ' [Paradox]';
         Database.AliasName := INI.ReadString('AGENTS', 'ALIAS', 'BASO');
     end
     else begin
         Caption := Caption + ' [SQLBase]';
         Database.AliasName := INI.ReadString('AGENTS', 'ALIASSQL', 'BASO');
     end;
     INI.Free;

     Caption := Caption + ' ' + COMPAMYNAME;

     try
         Database.Open;
     except
       on E : Exception do begin
         MessageDlg('Ошибка подключения к БД'#13 + E.Message, mtError, [mbOk], 0);
         Halt
       end;
     end;

     IsMain := false;
     try
         MainQuery.Open;
         IsMain := true;
     except
         on E : Exception do begin
            MessageDlg(e.Message + #13 + 'Основной запрос', mtInformation, [mbOk], 0);
         end;
     end;

     Timer.Enabled := true;
end;

procedure TF_Agent.TimerTimer(Sender: TObject);
var
     i : integer;
     X : integer;
begin
     X := 13;
     for i := RxDBGrid.LeftCol - 1 to 7 do
     begin
         X := X + RxDBGrid.Columns[i].Width + 1;
         if i = 0 then
         begin
            AgFilter.Left := X;
            AgFilter.Width := RxDBGrid.Columns[i + 1].Width;
         end
     end;
     LeftDivider.Left := X;
     LabelLeft.Width := X;       
     LabelRight.Left := X + 2;

     for i := 8 to RxDBGrid.Columns.Count - 1 do
         X := X + RxDBGrid.Columns[i].Width + 1;

     LabelRight.Width := X - LabelRight.Left - 4;
     RightDivider.Left := X;
end;

procedure TF_Agent.SaveButtonClick(Sender: TObject);
var
     s : string;
begin
     try
        s := MainQueryAGENT_CODE.Value;
        Screen.Cursor := crHourGlass;
        Database.StartTransaction;
        MainQuery.ApplyUpdates;
        Database.Commit;
        MainQuery.CommitUpdates;
        SaveButton.Enabled := false;
        CancelCh.Enabled := false;
        MainQuery.Close;
        MainQuery.Open;
        MainQuery.Locate('AGENT_CODE', s, []);
     except
         on E : Exception do begin
            if Database.InTransaction then
                Database.Rollback;
            MessageDlg(e.Message, mtInformation, [mbOk], 0);
         end;
     end;
     Screen.Cursor := crDefault; 
end;

procedure TF_Agent.AddBtnClick(Sender: TObject);
var
     Code : integer;
     s : string;
begin
     MainQuery.Last;
     if MainQuery.RecordCount = 0 then
         Code := -1
     else begin
         while not MainQuery.Bof do begin
             try
             Code := -1; 
             Code := MainQueryAGENT_CODE.AsInteger;
             break;
             except
             end;
             MainQuery.Prior;
         end;
     end;
     MainQuery.Append;
     MainQueryAGENT_CODE.ReadOnly := false;
     s := IntToStr(Code + 1);
     while Length(s) < 4 do s := '0' + s;
     MainQueryAGENT_CODE.AsString := s;
     MainQueryAGENT_CODE.ReadOnly := true;
     MainQueryTYPE.AsString := 'F';
end;

procedure TF_Agent.DelBtnClick(Sender: TObject);
var
    Capt : string;
    s : string;
begin
    Capt := Caption;
    Caption := Capt + ' - ИНДЕКСИРОВАНИЕ';
    CreateIdx;
    With CheckAgent do
    begin
        try
            Screen.Cursor := crHourGlass;
            ParamByName('AGENT').AsString := MainQueryAGENT_CODE.AsString;
            Caption := Capt + ' - ПРОВЕРКА НАЛИЧИЯ ПОЛИСОВ';
            Open;
            Caption := Capt + ' - ЗАГРУЗКА ДАННЫХ';
            while not Eof do
            begin
                s := s + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' ' + FieldByName('INS').AsString + ' Агент ' + FieldByName('AGENT').AsString + #13;
                Next;
                if Length(s) > 10000 then
                begin
                    s := s + #13'...';
                    break;
                end;
            end;
        except
        end;
        Caption := Capt;
        Screen.Cursor := crDefault;
        if s <> '' then
        begin
            Clipboard.AsText := 'На агента ' + MainQueryAGENT_CODE.AsString + ' ' + MainQueryNAME.AsString + ' зарегистрированы полисы ' + #13 + s;
            MessageDlg(Copy(s, 1, 1000) + #13'Указанные ПОЛИСЫ записаны на агента ' + MainQueryAGENT_CODE.AsString + ' ' + MainQueryNAME.AsString + #13'Полный список полисов скопирован в буфер обмена.', mtInformation, [mbOk], 0);
            Screen.Cursor := crDefault;
            exit;
        end;
        Close;
    end;

    if MessageDlg('Удалить ' + MainQueryNAME.AsString + '?', mtInformation, [mbYes, mbNo], 0) = mrYes then
        MainQuery.Delete;
end;

procedure TF_Agent.MainQueryVTYPEGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
begin
     MessageBeep(0);
end;

procedure TF_Agent.MainQueryBeforeInsert(DataSet: TDataSet);
begin
     if not AddBtn.Focused then
         Abort
end;

procedure TF_Agent.MainQueryTYPESetText(Sender: TField;
  const Text: String);
begin
     if Text = 'Ю' then
         Sender.AsString := 'U'
     else
     if Text = 'Ф' then
         Sender.AsString := 'F'
     else
         Sender.Clear;
end;                            

procedure TF_Agent.MainQueryTYPEGetText(Sender: TField; var Text: String;
  DisplayText: Boolean);
begin
     if Sender.AsString = 'U' then
         Text := 'Ю'
     else
     if Sender.AsString = 'F' then
         Text := 'Ф'
     else
         Text := '';
end;

procedure TF_Agent.MainQueryBeforeDelete(DataSet: TDataSet);
begin
     if not DelBtn.Focused then
         Abort
end;

procedure TF_Agent.MainQueryAfterDelete(DataSet: TDataSet);
begin
     SaveButton.Enabled := true;
     CancelCh.Enabled := true;
end;

procedure TF_Agent.MainQueryAfterInsert(DataSet: TDataSet);
begin
     SaveButton.Enabled := true;
     CancelCh.Enabled := true;
end;

procedure TF_Agent.MainQueryAfterPost(DataSet: TDataSet);
begin
     SaveButton.Enabled := true;
     CancelCh.Enabled := true;
end;

procedure TF_Agent.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
     if SaveButton.Enabled = true then
       SaveButtonClick(nil);
end;

procedure TF_Agent.CancelChClick(Sender: TObject);
begin
     if MessageDlg('Вы хотите отменить сделанные изменения?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         exit;

     MainQuery.CancelUpdates;
     SaveButton.Enabled := false;
     CancelCh.Enabled := false;
end;

procedure TF_Agent.NameFilterChange(Sender: TObject);
begin
{     if CancelCh.Enabled then begin
         NameFilter.Text := '';
         exit;
     end;

     AddBtn.Enabled := NameFilter.Text = '';

     MainQuery.Close;
     if NameFilter.Text = '' then begin
         MainQuery.MacroByName('WHERE').AsString := '';
         MainQuery.RequestLive := true;
         MainQuery.CachedUpdates := true;
     end
     else begin
         MainQuery.CachedUpdates := false;
         MainQuery.RequestLive := false;
         MainQuery.MacroByName('WHERE').AsString := 'WHERE UPPER(NAME) LIKE ''%' + AnsiUpperCase(NameFilter.Text) + '%'' ';
     end;
     MainQuery.Open;   }
end;

procedure TF_Agent.ActionSaveUpdate(Sender: TObject);
begin
     ActionSave.Enabled := MainQuery.UpdatesPending OR (MainQuery.State in [dsEdit, dsInsert]);
     ActionCancel.Enabled := MainQuery.UpdatesPending OR (MainQuery.State in [dsEdit, dsInsert]);
end;

procedure TF_Agent.MainQueryCalcFields(DataSet: TDataSet);
var
    code : array[0..16] of char;
    val : double;
begin
    val := MainQuery.FieldByName('POLANDPCNT').AsFloat;
    Move(val, code[0], 8);
    val := MainQuery.FieldByName('SOVAG').AsFloat;
    Move(val, code[8], 8);
    code[16] := #0;
    MainQueryUID.AsString := code;
end;

procedure TF_Agent.btnFreeAgentsClick(Sender: TObject);
var
    s : string;
begin
    CreateIdx;
    With FreeAgents do
    begin
        try
            Screen.Cursor := crHourGlass;
            Open;
            while not Eof do
            begin
                s := s + FieldByName('Agent_code').AsString + ' ' + FieldByName('Name').AsString + #13;
                Next;
            end;
        except
        end;
        Close;
    end;
    Screen.Cursor := crDefault;

    if s <> '' then
    begin
        Clipboard.AsText := 'Указанные Агенты нигде не зарегистрированы' + #13 + s;
        MessageDlg(Copy(s, 1, 1000) + 'Указанные агенты нигде не зарегистрированыю Вы можете их удалиить'#13'Полный список агентов скопирован в буфер обмена.', mtInformation, [mbOk], 0);
    end
    else
    begin
        MessageDlg('База данных без проблем', mtInformation, [mbOk], 0);
    end
end;

procedure TF_Agent.btnLostClick(Sender: TObject);
var
    s : string;
begin
    CreateIdx;
    With LostAgents do
    begin
        try
            Screen.Cursor := crHourGlass;
            Open;
            while not Eof do
            begin
                s := s + FieldByName('SERIA').AsString + '/' + FieldByName('NUMBER').AsString + ' ' + FieldByName('INS').AsString + ' Агент ' + FieldByName('AGENT').AsString + #13;
                Next;
            end;
        except
        end;
        Close;
    end;
    Screen.Cursor := crDefault;

    if s <> '' then
    begin
        Clipboard.AsText := 'Указанные ПОЛИСЫ имеют неопределенных агентов' + #13 + s;
        MessageDlg(Copy(s, 1, 1000) + #13'Указанные ПОЛИСЫ имеют неопределенных агентов'#13'Полный список полисов скопирован в буфер обмена.', mtInformation, [mbOk], 0);
    end
    else
    begin
        MessageDlg('База данных без проблем', mtInformation, [mbOk], 0);
    end
end;

procedure TF_Agent.AgFilterChange(Sender: TObject);
begin
    if (MainQuery.State <> dsEdit) AND (MainQuery.State <> dsInsert) then
    begin
        MainQuery.Close;
        MainQuery.MacroByName('WHERE').AsString := 'WHERE UPPER(NAME) LIKE ''%' + AnsiUpperCase(AgFilter.Text) + '%'' ';
        MainQuery.Open;
        AddBtn.Enabled := false;
    end
end;

procedure TF_Agent.AgFilterEnter(Sender: TObject);
begin
    AgFilter.Text := '';
end;

procedure TF_Agent.AgFilterExit(Sender: TObject);
begin
    AddBtn.Enabled := AgFilter.Text = '';
end;

end.
