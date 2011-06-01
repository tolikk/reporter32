unit GetPeriodUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, DBTables;

type
  TGetPeriod = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    StartDate: TDateTimePicker;
    EndDate: TDateTimePicker;
    Ok: TBitBtn;
    Label3: TLabel;
    ExpMode: TComboBox;
    ListExports: TListBox;
    Label4: TLabel;
    Query: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure StartDateExit(Sender: TObject);
    procedure ExpModeChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  GetPeriod: TGetPeriod;

implementation

{$R *.dfm}

procedure TGetPeriod.FormCreate(Sender: TObject);
begin
    ExpMode.ItemIndex := 0;
    EndDate.Date := Date;
    StartDate.Date := EncodeDate(2004, 1, 1);
end;

procedure TGetPeriod.StartDateExit(Sender: TObject);
begin
    if (GetAsyncKeyState(VK_SHIFT) AND $8000) <> 0 then
        exit;
    if StartDate.Date <  EncodeDate(2004, 1, 1) then
        StartDate.Date := EncodeDate(2004, 1, 1);
end;

procedure TGetPeriod.ExpModeChange(Sender: TObject);
begin
    StartDate.Enabled := ExpMode.ItemIndex = 0;
    EndDate.Enabled := ExpMode.ItemIndex = 0;
    Label1.Enabled := ExpMode.ItemIndex = 0;
    Label2.Enabled := ExpMode.ItemIndex = 0;

    ListExports.Enabled := ExpMode.ItemIndex = 3;
    Label4.Enabled := ExpMode.ItemIndex = 3;

    if ExpMode.ItemIndex <> 3 then
        ListExports.Clear
    else begin
        Query.Open;
        while not Query.Eof do begin
            ListExports.Items.Add(Query.Fields[0].AsString);
            Query.Next;
        end;
        Query.Close;
    end    
end;

end.
