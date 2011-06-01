unit BGBuroUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, RXSpin, Placemnt, Mask;

type
  TBG_Buro = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Path: TEdit;
    OK: TButton;
    FormStorage: TFormStorage;
    NExport: TComboBox;
    GroupBox1: TGroupBox;
    listExports: TListBox;
    Button1: TButton;
    Button2: TButton;
    IsTest: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  BG_Buro: TBG_Buro;

implementation

uses BelGreenRpt;

{$R *.DFM}

procedure TBG_Buro.FormCreate(Sender: TObject);
//var
//    _Year, Month, Day : Word;
begin
    NExport.ItemIndex := 0;
//    DecodeDate(IncMonth(Date, -1), _Year, Month, Day);
//    Year.Value := _Year;
//    MonthLst.ItemIndex := Month;
end;

procedure TBG_Buro.Button2Click(Sender: TObject);
var
    Index : integer;
begin
    listExports.Clear;
    try
        with BelGreenForm.WorkSQL, BelGreenForm.WorkSQL.SQL do begin
            Close;
            Clear;
            Add('SELECT NEXPORT, UPD, COUNT(*) AS CNT FROM BELGREENEXP GROUP BY NEXPORT, UPD ORDER BY 1 DESC');
            Screen.Cursor := crHourGlass;
            Open;
            while not Eof do begin
                if FieldByName('NEXPORT').AsString = '' then break;
                Index := listExports.Items.Add(FieldByName('NEXPORT').AsString + ', ' + FieldByName('UPD').AsString + ', ' + FieldByName('CNT').AsString + ' штук');
                listExports.Items.Objects[Index] := TObject(FieldByName('NEXPORT').AsInteger);
                Next;
            end;
            Close;
        end;
    except
      on E : Exception do begin
        MessageDlg(''#13 + E.Message, mtInformation, [mbOk], 0);
      end;
    end;
    Screen.Cursor := crDefault;
end;

procedure TBG_Buro.Button1Click(Sender: TObject);
var
    s : string;
begin
    if listExports.ItemIndex <> 0 then exit;
    s := listExports.Items[listExports.ItemIndex];
    s := Copy(s, 1, Pos(',', s) - 1);
    if (GetAsyncKeyState(VK_SHIFT) AND $8000) = 0 then exit;
    if MessageDlg('Вы хотите удалить экспорт N' + s + '?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
        if (GetAsyncKeyState(VK_SHIFT) AND $8000) = 0 then exit;
        try
            with BelGreenForm.WorkSQL, BelGreenForm.WorkSQL.SQL do begin
                Close;
                Clear;
                Add('DELETE FROM BELGREENEXP WHERE NEXPORT=' + s);
                ExecSQL;
                Close;
                Button2Click(nil);
            end;
        except
          on E : Exception do begin
            MessageDlg(''#13 + E.Message, mtInformation, [mbOk], 0);
          end;
        end;
    end;
end;

end.
