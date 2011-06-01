unit Sort;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, CheckLst, db, Buttons, ActnList, ExtCtrls;

type
  TSortForm = class(TForm)
    List: TCheckListBox;
    Label1: TLabel;
    Up: TSpeedButton;
    Down: TSpeedButton;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    ActionList: TActionList;
    PosUP: TAction;
    PosDOWN: TAction;
    Bevel1: TBevel;
    procedure UpClick(Sender: TObject);
    procedure DownClick(Sender: TObject);
    procedure PosUPExecute(Sender: TObject);
    procedure PosDOWNExecute(Sender: TObject);
    procedure PosUPUpdate(Sender: TObject);
    procedure PosDOWNUpdate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    m_Query : TDataSet;
    { Public declarations }
    function  GetOrder(PrevString : string) : string;
  end;

var
  SortForm: TSortForm;

implementation

uses MF;

{$R *.DFM}

function TSortForm.GetOrder(PrevString : string) : string;
var
     filtstr : string;
     i, j : integer;
begin
     filtstr := '';

     for i := 0 to List.Items.Count - 1 do begin
       if List.Checked[i] then begin
         for j := 0 to m_Query.Fields.Count - 1 do begin
             if List.Items[i] = m_Query.Fields[j].DisplayLabel then
                filtstr := filtstr + m_Query.Fields[j].FieldName + ',';
         end;
       end
     end;

     if PrevString <> '' then PrevString := PrevString + ' ';
     Result := PrevString + Copy(filtstr, 1, Length(filtstr) - 1);
end;

procedure TSortForm.UpClick(Sender: TObject);
var
     IsCheck : boolean;
     s : string;
     Index : integer;
begin
     //Up
     Index := List.ItemIndex;
     if Index > 0 then begin
        IsCheck := List.Checked[Index];
        s := List.Items[Index];
        List.Items.Delete(Index);

        List.Items.Insert(Index - 1, s);
        List.Checked[Index - 1] := IsCheck;

        List.ItemIndex := Index - 1;
     end;
end;

procedure TSortForm.DownClick(Sender: TObject);
var
     IsCheck : boolean;
     s : string;
     Index : integer;
begin
     //Down
     Index := List.ItemIndex;
     if Index < List.Items.Count - 1 then begin
        IsCheck := List.Checked[Index];
        s := List.Items[Index];
        List.Items.Delete(Index);

        List.Items.Insert(Index + 1, s);
        List.Checked[Index + 1] := IsCheck;

        List.ItemIndex := Index + 1;
     end;
end;

procedure TSortForm.PosUPExecute(Sender: TObject);
begin
     UpClick(nil);
end;

procedure TSortForm.PosDOWNExecute(Sender: TObject);
begin
     DownClick(nil);
end;

procedure TSortForm.PosUPUpdate(Sender: TObject);
begin
     PosUP.Enabled := List.ItemIndex > 0;
end;

procedure TSortForm.PosDOWNUpdate(Sender: TObject);
begin
     PosDOWN.Enabled := List.ItemIndex < (List.Items.Count - 1);
end;

procedure TSortForm.FormShow(Sender: TObject);
var
     i : integer;
begin
     if List.Items.Count > 0 then exit;
     for i := 0 to m_Query.Fields.Count - 1 do begin
         if m_Query.Fields[i].FieldKind = fkData then
           if m_Query.Fields[i].Visible then
              List.Items.Add(m_Query.Fields[i].DisplayLabel);
     end;
end;

procedure TSortForm.FormCreate(Sender: TObject);
begin
     m_Query := nil;
end;

end.
