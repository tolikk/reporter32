unit ImpFormParam;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Db, DBTables, DataMod;

type
  TAskImport = class(TForm)
    Button1: TButton;
    btnOk: TButton;
    GroupBox1: TGroupBox;
    IsSeria: TCheckBox;
    IsAgent: TCheckBox;
    NameAgent: TEdit;
    NewSeria: TComboBox;
    WorkQuery: TQuery;
    BatchMove: TBatchMove;
    SrcTable: TTable;
    DestTable: TTable;
    Label1: TLabel;
    Mode: TComboBox;
    Label2: TLabel;
    AgTable: TTable;
    AgTableAgent_code: TStringField;
    AgTableName: TStringField;
    Lbl: TLabel;
    BatchMoveDel: TBatchMove;
    Memo: TMemo;
    NewAgent: TComboBox;
    procedure IsSeriaClick(Sender: TObject);
    procedure IsAgentClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure NewAgentChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  AskImport: TAskImport;

implementation

uses impUnit, inifiles;

{$R *.DFM}

procedure TAskImport.IsSeriaClick(Sender: TObject);
begin
     NewSeria.Enabled := IsSeria.Checked;
end;

procedure TAskImport.IsAgentClick(Sender: TObject);
begin
     NewAgent.Enabled := IsAgent.Checked;
     NameAgent.Enabled := IsAgent.Checked;
end;

procedure TAskImport.btnOkClick(Sender: TObject);
var
     i, j, l : integer;
     AgentFld : string;
     INI : TIniFile;
     IsFound : boolean;
     IdxList : string;
begin
     if IsSeria.Checked AND (NewSeria.Text = '') then begin
         exit;
     end;

     if NewSeria.Text <> '' then begin
        IsFound := false;
        for i := 0 to NewSeria.Items.Count - 1 do
            if NewSeria.Items[i] = NewSeria.Text then
                IsFound := true;

        if not IsFound then begin
            NewSeria.Items.Add(NewSeria.Text);
            INI := TIniFile.Create(BLANK_INI);
            for i := 0 to NewSeria.Items.Count - 1 do
                INI.WriteString('IMPORT', 'ChSer' + IntToStr(i), NewSeria.Items[i]);
            INI.Free;
        end;
     end;

     if IsAgent.Checked AND ((Length(NewAgent.Text) <> 4) OR (NameAgent.Text = '')) then
         exit;

     if NewAgent.Text <> '' then begin
        IsFound := false;
        for i := 0 to NewAgent.Items.Count - 1 do
            if NewAgent.Items[i] = NewAgent.Text then
                IsFound := true;

        if not IsFound then begin
            NewAgent.Items.Add(NewAgent.Text);
            INI := TIniFile.Create(BLANK_INI);
            for i := 0 to NewAgent.Items.Count - 1 do
                INI.WriteString('IMPORT', 'ChAgent' + IntToStr(i), NewAgent.Items[i]);
            INI.Free;
        end;
     end;

     try
         AgentFld := '';
         Screen.Cursor := crHourGlass;
         DestTable.TableName := ImportData.DestFile.Text;
         DestTable.Open;

       for l := 0 to Memo.Lines.Count - 1 do begin
         SrcTable.Close;
         SrcTable.TableName := Memo.Lines[l];
         SrcTable.Open;
         if SrcTable.IndexFieldCount = 0 then begin
             Lbl.Caption := 'Создание индекса...';
             Lbl.Update;
             IdxList := '';
             for j := 0 to DestTable.IndexFieldCount - 1 do begin
                 if IdxList <> '' then IdxList := IdxList + ';';
                 IdxList := IdxList + SrcTable.Fields[j].FieldName;
             end;
             SrcTable.Close;
             SrcTable.AddIndex('PK', IdxList, [ixPrimary]);
         end;
         SrcTable.Close;

         if IsSeria.Checked then begin
            with WorkQuery, WorkQuery.SQL do begin
                 Clear;
                 Add('UPDATE ''' + Memo.Lines[l] + '''');
                 Add('SET SERIA = ''' + NewSeria.Text + '''');
                 Lbl.Caption := 'Замена серий...';
                 Lbl.Update;
                 ExecSQL;
            end;
         end;
         if IsAgent.Checked then begin
            if AgentFld = '' then begin
                SrcTable.Open;
                for i := 0 to SrcTable.Fields.Count - 1 do begin
                    if (SrcTable.Fields[i].DataType = ftString) AND (SrcTable.Fields[i].DataSize <= 7) AND (SrcTable.Fields[i].DataSize >= 4) AND (Pos('AG', UpperCase(SrcTable.Fields[i].FieldName)) = 1) then
                        AgentFld := SrcTable.Fields[i].FieldName;
                end;
                SrcTable.Close;
                if AgentFld = '' then begin
                   Screen.Cursor := crDefault;
                   MessageDlg('Поля агента не найдено!!!', mtInformation, [mbOk], 0);
                   exit;
                end else
                if MessageDlg('Поле c кодом агента называеться ''' + AgentFld  + '''?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then begin
                   Screen.Cursor := crDefault;
                   exit;
                end;
            end;

            with WorkQuery, WorkQuery.SQL do begin
                 Clear;
                 Add('UPDATE ''' + Memo.Lines[l] + '''');
                 Add('SET ' + AgentFld + ' = ''' + NewAgent.Text + '''');
                 Lbl.Caption := 'Замена агентов...';
                 Lbl.Update;
                 ExecSQL;
            end;
         end;

         SrcTable.Close;
         if Mode.ItemIndex = 0 then begin
             BatchMove.Mode := batAppend;
             Lbl.Caption := 'Чистка...';
             Lbl.Update;
             BatchMoveDel.Execute;
         end;
         if Mode.ItemIndex = 1 then
             BatchMove.Mode := batAppendUpdate;
         if Mode.ItemIndex = 2 then
             BatchMove.Mode := batUpdate;

         Lbl.Caption := 'Добавление...';
         Lbl.Update;
         BatchMove.Execute;
       end;
       Screen.Cursor := crDefault;
     except
         on E : Exception do begin
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
             Lbl.Caption := '';
             Screen.Cursor := crDefault;
             DestTable.Close;
             SrcTable.Close;
             exit;
         end;
     end;

     SrcTable.Close;
     DestTable.Close;
     Lbl.Caption := '';
     ModalResult := mrOk;
end;

procedure TAskImport.FormCreate(Sender: TObject);
var
     INI : TIniFile;
     s : string;
     i : integer;
begin
     Mode.ItemIndex := 0;
     INI := TIniFile.Create(BLANK_INI);
     for i := 0 to 100 do begin
         s := INI.ReadString('IMPORT', 'ChSer' + IntToStr(i), '');
         if s <> '' then
             NewSeria.Items.Add(s);
     end;
     for i := 0 to 100 do begin
         s := INI.ReadString('IMPORT', 'ChAgent' + IntToStr(i), '');
         if s <> '' then
             NewAgent.Items.Add(s);
     end;
     INI.Free;

     AgTable.Open;
end;

procedure TAskImport.NewAgentChange(Sender: TObject);
begin
     NameAgent.Text := '';
     if AgTable.Locate('Agent_code', NewAgent.Text, []) then
         NameAgent.Text := AgTableName.AsString;
end;

end.
