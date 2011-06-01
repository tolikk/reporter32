unit RepParamsAgntFrm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, DB, DBTables, CheckLst;

type
  TRepParamsAgnt = class(TForm)
    Label2: TLabel;
    DtFrom: TDateTimePicker;
    DtTo: TDateTimePicker;
    btnOk: TButton;
    Label1: TLabel;
    Query: TQuery;
    listAgnt: TCheckListBox;
    CountInfo: TLabel;
    IsFullRep: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure listAgntClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    { Private declarations }
  public
    agCodeList : TStringList;
    agResultCodeList : TStringList;
    agResultNameList : TStringList;
    agSQL : string;
    filtParam : string;
    { Public declarations }
    function GetAgentName(Code : string) : string;
  end;

var
  RepParamsAgnt: TRepParamsAgnt;

implementation

{$R *.dfm}

procedure TRepParamsAgnt.FormCreate(Sender: TObject);
var
    d, m, y : WORD;
begin
    agCodeList := TStringList.Create;
    agResultCodeList := TStringList.Create;
    agResultNameList := TStringList.Create;
    DecodeDate(Date, y, m, d);
    DtTo.Date := EncodeDate(y, m, 1) - 1;
    if m > 1 then
    begin
        m := m - 1
    end
    else
    begin
        m := 12;
        y := y - 1;
    end;
    DtFrom.Date := EncodeDate(y, m, 1);
end;

procedure TRepParamsAgnt.FormDestroy(Sender: TObject);
begin
    agCodeList.Free;
    agResultCodeList.Free;
    agResultNameList.Free;
end;

procedure TRepParamsAgnt.FormShow(Sender: TObject);
begin
    agResultCodeList.Clear;
    agResultNameList.Clear;
    CountInfo.Caption := '';
    agCodeList.Clear;
    listAgnt.Clear;
    Query.SQL.Clear;
    if filtParam = '' then filtParam := '1=1';
    Query.SQL.Add('SELECT * FROM AGENT WHERE ' + filtParam + ' ORDER BY Name');
    Query.Open;
    while not Query.Eof do begin
        agCodeList.Add(Query.FieldByName('Agent_code').AsString);
        listAgnt.Items.Add(Query.FieldByName('Name').AsString);
        Query.Next
    end;
end;

procedure TRepParamsAgnt.listAgntClick(Sender: TObject);
var
    i, cnt : integer;
begin
    cnt := 0;
    for i := 0 to listAgnt.Items.Count - 1 do
    begin
        if listAgnt.Checked[i] then Inc(cnt);
    end;
    CountInfo.Caption := 'Выбрано ' + IntToStr(cnt);
end;

procedure TRepParamsAgnt.btnOkClick(Sender: TObject);
var
    i : integer;
begin
    agSQL := '(';
    for i := 0 to listAgnt.Items.Count - 1 do
    begin
        if listAgnt.Checked[i] then
        begin
            if Length(agSQL) > 3 then agSQL := agSQL + ',';
            agSQL := agSQL + '''' + agCodeList[i] + '''';
            agResultCodeList.Add(agCodeList[i]);
            agResultNameList.Add(listAgnt.Items[i]);
        end
    end;
    agSQL := agSQL + ')';
end;

function TRepParamsAgnt.GetAgentName(Code : string) : string;
var
    i : integer;
begin
    GetAgentName := Code;
    for i := 0 to agCodeList.Count - 1 do
    begin
        if agCodeList[i] = Code then
        begin
            GetAgentName := listAgnt.Items[i];
            exit;
        end;
    end;
end;

end.
