unit PolPolisUnit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RxDBComb, RxLookup, Mask, DBCtrls, ExtCtrls, DB, Menus,
  Buttons;

type
  TPolPolisData = class(TForm)
    PanelData: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label6: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    dbAgentType: TDBCheckBox;
    dbAgentPcnt: TDBEdit;
    dbAgent: TRxDBLookupCombo;
    dbOwn: TDBEdit;
    dbOwnCity: TDBEdit;
    dbMarka: TDBEdit;
    dnAutoNmb: TDBEdit;
    dbBodyNmb: TDBEdit;
    dbIns: TDBEdit;
    dbInsCity: TDBEdit;
    dbReg: TDBEdit;
    dbPay: TDBEdit;
    dbPay1: TDBEdit;
    dbCurr1: TRxDBComboBox;
    dbSeria: TDBEdit;
    dbNumber: TDBEdit;
    dbPay2: TDBEdit;
    dbCurr2: TRxDBComboBox;
    Label3: TLabel;
    dbFrom: TDBEdit;
    Label15: TLabel;
    dbTo: TDBEdit;
    dbFromTime: TDBEdit;
    dbPeriod: TDBEdit;
    Label16: TLabel;
    dbCharact: TDBEdit;
    dbCountry: TDBComboBox;
    Label17: TLabel;
    Label18: TLabel;
    dbTarif: TDBEdit;
    DataSource: TDataSource;
    Save: TButton;
    New: TButton;
    dbLetter: TDBComboBox;
    btnContinue: TButton;
    Label19: TLabel;
    btnDublicat: TButton;
    PopupMenuContinue: TPopupMenu;
    N151: TMenuItem;
    N301: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N111: TMenuItem;
    N112: TMenuItem;
    dbPNumber: TDBEdit;
    Label21: TLabel;
    btnCancel: TButton;
    DBText1: TDBText;
    dbIsDup: TDBCheckBox;
    Label20: TLabel;
    dbPremium: TDBEdit;
    dbOwnType: TDBCheckBox;
    dbHouse: TDBEdit;
    dbFlat: TDBEdit;
    Label22: TLabel;
    Label24: TLabel;
    Label23: TLabel;
    dbEngine: TDBEdit;
    Label25: TLabel;
    Label26: TLabel;
    dbAutoType: TDBComboBox;
    dbInsType: TDBCheckBox;
    Label27: TLabel;
    dbHouse2: TDBEdit;
    dbFlat2: TDBEdit;
    Label28: TLabel;
    dbCountry2: TDBComboBox;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Label29: TLabel;
    dbRetDate: TDBEdit;
    Label30: TLabel;
    dbRetSum: TDBEdit;
    dbRetCurr: TRxDBComboBox;
    btnCopyInsData: TSpeedButton;
    dbStreet: TDBEdit;
    dbInsStreet: TDBEdit;
    Timer: TTimer;
    dbAutoVid: TDBLookupComboBox;
    PopupMenuAuto: TPopupMenu;
    procedure FormCreate(Sender: TObject);
    procedure SaveClick(Sender: TObject);
    procedure NewClick(Sender: TObject);
    procedure btnContinueClick(Sender: TObject);
    procedure COntinue(Sender: TObject);
    procedure btnDublicatClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dbFromTimeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnCopyInsDataClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure PopupMenuAutoPopup(Sender: TObject);
    procedure SetMarkaClick(Sender: TObject);
  private
    { Private declarations }
  public
    //IsEditNumber : boolean;
    procedure ShowStore;
    { Public declarations }
  end;

var
  PolPolisData: TPolPolisData;

implementation

uses PolandUnit, MsgOkForm, inifiles, math;

{$R *.dfm}

procedure TPolPolisData.FormCreate(Sender: TObject);
begin
    //Parent := PolandPolises
end;

procedure TPolPolisData.SaveClick(Sender: TObject);
begin
    if DataSource.DataSet.State in [dsEdit, dsInsert] then
        DataSource.DataSet.Post;
end;

procedure TPolPolisData.NewClick(Sender: TObject);
var
    FieldsBuf : VARIANT;
    i : integer;
begin
//    if not IsEditNumber then begin
        if PolandPolises.PolandQuery.State in [dsEdit, dsInsert] then
            PolandPolises.PolandQuery.Post;

        FieldsBuf := VarArrayCreate([0, 0, 0, DataSource.DataSet.Fields.Count], varVariant);
        for i := 0 to DataSource.DataSet.Fields.Count - 1 do
            FieldsBuf[0, i] := DataSource.DataSet.Fields[i].Value;

        DataSource.DataSet.Next;
        DataSource.DataSet.Insert;

        for i := 0 to DataSource.DataSet.Fields.Count - 1 do
            DataSource.DataSet.Fields[i].Value := FieldsBuf[0, i];

        DataSource.DataSet.FieldByName('Number').AsInteger := DataSource.DataSet.FieldByName('Number').AsInteger + 1;
        dbNumber.SetFocus;    
        VarClear(FieldsBuf);

//        IsEditNumber := true;
//    end;
end;

procedure TPolPolisData.ShowStore;
begin
    Timer.Enabled := true;
    {MsgOk.Parent := self;
    MsgOk.Left := 0;//(Width - MsgOk.Width) DIV 2;
    MsgOk.Top := (Height - MsgOk.Height) DIV 2;}
    MsgOk.Show;
    MsgOk.Update;

end;

procedure TPolPolisData.btnContinueClick(Sender: TObject);
var
    p : TPoint;
begin
    if DataSource.DataSet.FieldByName('State').AsString <> NORMAL_POLIS then
        exit;

    p.x := btnContinue.Left;
    p.y := btnContinue.Top + btnContinue.Height;
    windows.ClientToScreen(Handle, p);
    PopupMenuContinue.Popup(p.x, p.y);
end;

procedure TPolPolisData.Continue(Sender: TObject);
var
    FieldsBuf : VARIANT;
    i : integer;
    Number : integer;
    EndDate : TDateTime;
begin
    if DataSource.DataSet.State in [dsEdit, dsInsert] then
        DataSource.DataSet.Post;

    FieldsBuf := VarArrayCreate([0, 0, 0, DataSource.DataSet.Fields.Count], varVariant);
    for i := 0 to DataSource.DataSet.Fields.Count - 1 do
        FieldsBuf[0, i] := DataSource.DataSet.Fields[i].Value;

    Number := DataSource.DataSet.FieldByName('NUMBER').AsInteger;
    EndDate := DataSource.DataSet.FieldByName('ENDDATE').AsDateTime;

    DataSource.DataSet.Next;
    DataSource.DataSet.Insert;

    for i := 0 to DataSource.DataSet.Fields.Count - 1 do
        DataSource.DataSet.Fields[i].Value := FieldsBuf[0, i];

    VarClear(FieldsBuf);

    DataSource.DataSet.FieldByName('PNUMBER').AsInteger := Number;
    DataSource.DataSet.FieldByName('NUMBER').Clear;
    DataSource.DataSet.FieldByName('REPDATE').AsDateTime := Date;
    DataSource.DataSet.FieldByName('STARTDATE').AsDateTime := EndDate + 1;

    DataSource.DataSet.FieldByName('PREMIUMPAY').Clear;
    DataSource.DataSet.FieldByName('PREMIUMCURR').Clear;
    DataSource.DataSet.FieldByName('PREMIUMPAY2').Clear;
    DataSource.DataSet.FieldByName('PREMIUMCURR2').Clear;

    dbPeriod.SetFocus;
    if TControl(Sender).Tag in [1,2,3,4,5,6,7,8,9,10,11] then
        dbPeriod.Text := '*' + IntToStr(TControl(Sender).Tag);
    if TControl(Sender).Tag in [15,30] then
        dbPeriod.Text := '+' + IntToStr(TControl(Sender).Tag);
    dbNumber.SetFocus;
end;

procedure TPolPolisData.btnDublicatClick(Sender: TObject);
var
    FieldsBuf : VARIANT;
    i : integer;
    Number : integer;
begin
    if DataSource.DataSet.FieldByName('State').AsString <> NORMAL_POLIS then
        exit;

    if DataSource.DataSet.State in [dsEdit, dsInsert] then
        DataSource.DataSet.Post;

    FieldsBuf := VarArrayCreate([0, 0, 0, DataSource.DataSet.Fields.Count], varVariant);
    for i := 0 to DataSource.DataSet.Fields.Count - 1 do
        FieldsBuf[0, i] := DataSource.DataSet.Fields[i].Value;

    Number := DataSource.DataSet.FieldByName('NUMBER').AsInteger;

    DataSource.DataSet.Next;
    DataSource.DataSet.Insert;

    for i := 0 to DataSource.DataSet.Fields.Count - 1 do
        DataSource.DataSet.Fields[i].Value := FieldsBuf[0, i];

    VarClear(FieldsBuf);

    DataSource.DataSet.FieldByName('PNUMBER').AsInteger := Number;
    DataSource.DataSet.FieldByName('NUMBER').Clear;
    DataSource.DataSet.FieldByName('REPDATE').AsDateTime := Date;

    DataSource.DataSet.FieldByName('PREMIUMPAY').Clear;
    DataSource.DataSet.FieldByName('PREMIUMCURR').Clear;
    DataSource.DataSet.FieldByName('PREMIUMPAY2').Clear;
    DataSource.DataSet.FieldByName('PREMIUMCURR2').Clear;

    DataSource.DataSet.FieldByName('ISDUP').AsString := 'Y';

    dbNumber.SetFocus;
end;

procedure TPolPolisData.btnCancelClick(Sender: TObject);
begin
    if DataSource.DataSet.State in [dsEdit, dsInsert] then
        DataSource.DataSet.Cancel;
end;

procedure TPolPolisData.dbFromTimeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if Key = Ord(' ') then
        dbFromTime.Text := FormatDateTime('hh:nn', Time);
end;

procedure TPolPolisData.btnCopyInsDataClick(Sender: TObject);
begin
    if not (dbInsType.DataSource.DataSet.State in [dsEdit, dsInsert]) then
        dbInsType.DataSource.DataSet.Edit;

    dbIns.SetFocus;     
    dbInsType.Field.Value := dbOwnType.Field.Value;
    dbIns.Field.Value := dbOwn.Field.Value;
    dbCountry2.Field.Value := dbCountry.Field.Value;
    dbInsCity.Field.Value := dbOwnCity.Field.Value;
    dbInsStreet.Field.Value := dbStreet.Field.Value;
    dbHouse2.Field.Value := dbHouse.Field.Value;
    dbFlat2.Field.Value := dbFlat.Field.Value;
end;

procedure TPolPolisData.TimerTimer(Sender: TObject);
begin
    MsgOk.Hide;
end;

procedure TPolPolisData.PopupMenuAutoPopup(Sender: TObject);
var
    INI : TIniFile;
    ININame, key, s : string;
    i, j, fi : integer;
    KeyMenu, Menu : TMenuItem;
begin
    if PopupMenuAuto.Items.Count = 0 then begin
        INI := TIniFile.Create(BLANK_INI);
        ININame := INI.ReadString('POLAND', 'AutoMarkFile', '');
        INI.Free;
        if ININame = '' then exit;
        INI := TIniFile.Create(ININame);
        for i := 0 to 100 do begin
            key := INI.ReadString('POLAND', 'AutoType' + InttoStr(i), '');
            if key = '' then break;
            KeyMenu := TMenuItem.Create(PopupMenuAuto);
            KeyMenu.Caption := key;

            for fi := 0 to PopupMenuAuto.Items.Count - 1 do
                if PopupMenuAuto.Items[fi].Caption > KeyMenu.Caption then break;

            fi := min(fi, PopupMenuAuto.Items.Count);
            PopupMenuAuto.Items.Insert(fi, KeyMenu);
            KeyMenu.Tag := (i SHL 8);
            KeyMenu.OnClick := SetMarkaClick;
            for j := 0 to 100 do begin
                s := INI.ReadString('POLAND', 'AutoType' + InttoStr(i) + '_' + InttoStr(j), '');
                if s = '' then break;
                Menu := TMenuItem.Create(KeyMenu);
                Menu.Caption := s;
                Menu.OnClick := SetMarkaClick;
                Menu.Tag := (i SHL 8) + j + 1;
                KeyMenu.Add(Menu);
                KeyMenu.OnClick := nil;
            end;
        end;
        INI.Free;

        for i := 0 to PopupMenuAuto.Items.Count - 1 do begin
            if (i MOD 30) = 0 then
                PopupMenuAuto.Items[i].Break := mbBreak;
        end
    end;
end;

procedure TPolPolisData.SetMarkaClick(Sender: TObject);
var
    i, j : integer;
    INI : TIniFile;
    ININame, s, _type : string;
    _vid : integer;
begin
    j := TControl(Sender).Tag AND $FF;
    i := (TControl(Sender).Tag AND $FF00) SHR 8;
    INI := TIniFile.Create(BLANK_INI);
    ININame := INI.ReadString('POLAND', 'AutoMarkFile', '');
    INI.Free;
    if ININame = '' then exit;
    INI := TIniFile.Create(ININame);
    s := INI.ReadString('POLAND', 'AutoType' + InttoStr(i), '');
    _type := INI.ReadString('POLAND', 'AutoType' + InttoStr(i) + 'ÒÈÏ', '');
    _vid := INI.ReadInteger('POLAND', 'AutoType' + InttoStr(i) + 'ÂÈÄ', 0);
    if j <> 0 then begin
        s := s + ' ' + INI.ReadString('POLAND', 'AutoType' + InttoStr(i) + '_' + InttoStr(j - 1), '');
        _type := INI.ReadString('POLAND', 'AutoType' + InttoStr(i) + '_' + InttoStr(j - 1) + 'ÒÈÏ', _type);
        _vid := INI.ReadInteger('POLAND', 'AutoType' + InttoStr(i) + '_' + InttoStr(j - 1) + 'ÂÈÄ', _vid);
    end;

    if Length(s) > 2 then begin
        if not (PolandPolises.PolandQuery.State in [dsEdit, dsInsert]) then
            PolandPolises.PolandQuery.Edit;
        dbMarka.Text := s;
        dbAutoType.Text := _type;
        PolandPolises.PolandQuery.FieldByName(dbAutoVid.DataField).AsInteger := _vid;
    end;

    INI.Free;
end;

end.
