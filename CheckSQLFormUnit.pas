unit CheckSQLFormUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, CheckLst, Db, DBTables;

type
  TCheckSQLForm = class(TForm)
    SQLListItems: TCheckListBox;
    Panel1: TPanel;
    Button1: TButton;
    OKBtn: TButton;
    SQLQuery: TQuery;
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure SQLListItemsClick(Sender: TObject);
  private
    { Private declarations }
  public
    SQLText, InitValues : string;
    IsSelOne : boolean;
    { Public declarations }
  end;

var
  CheckSQLForm: TCheckSQLForm;

implementation

{$R *.DFM}

function ExistsInListStr(Str : string; Val : integer) : boolean;
var
     DivPos : integer;
     strVal : string;
begin
     ExistsInListStr := false;

     while true do begin
         DivPos := Pos(',', Str);
         strVal := Trim(Str);

         if strVal = '' then break;

         if DivPos <> 0 then
             strVal := Trim(Copy(Str, 1, DivPos - 1));
         if StrToInt(strVal) = Val then begin
             ExistsInListStr := true;
             exit;
         end;
         Str := Trim(Copy(Str, DivPos + 1, Length(Str)));
         if DivPos = 0 then break;
     end;
end;

procedure TCheckSQLForm.FormShow(Sender: TObject);
begin
     OKBtn.Enabled := true;
     SQLListItems.Clear;
     with SQLQuery, SQLQuery.SQl do begin
          Clear;
          Add(SQLText);
          try
              Open;
              while not EOF do begin
                    SQLListItems.Items.Add(FieldByName(Fields[1].FieldName).AsString);
                    SQLListItems.Items.Objects[SQLListItems.Items.Count - 1] := Tobject(FieldByName(Fields[0].FieldName).AsInteger);
                    if ExistsInListStr(InitValues, FieldByName(Fields[0].FieldName).AsInteger) then
                        SQLListItems.Checked[SQLListItems.Items.Count - 1] := true;
                    Next;
              end;
          except
              on E : Exception do begin
                  MessageDlg(e.Message, mtInformation, [mbOk], 0);
                  OKBtn.Enabled := false;
              end;
          end;
          Close;
     end;
end;

procedure TCheckSQLForm.OKBtnClick(Sender: TObject);
var
     i : integer;
begin
     if SQLListItems.SelCount = 0 then exit;
     InitValues := '';
     for i := 0 to SQLListItems.Items.Count - 1 do
       if SQLListItems.Checked[i] then
         InitValues := InitValues + IntToStr(Integer(SQLListItems.Items.Objects[i])) + ', ';

     InitValues := Copy(InitValues, 1, Length(InitValues) - 2);

     ModalResult := mrOk;
end;

procedure TCheckSQLForm.SQLListItemsClick(Sender: TObject);
var
     i : integer;
begin
     if IsSelOne then begin
        if SQLListItems.Checked[SQLListItems.ItemIndex] then
            for i := 0 to SQLListItems.Items.Count - 1 do
                if i <> SQLListItems.ItemIndex then
                    SQLListItems.Checked[i] := false;
     end;
end;

end.
