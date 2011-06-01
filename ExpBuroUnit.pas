unit ExpBuroUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Mask, ToolEdit, ComCtrls, inifiles, ExtCtrls, Placemnt;

type
  TExpBUROForm = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    PeriodFrom: TDateTimePicker;
    PeriodTo: TDateTimePicker;
    Label3: TLabel;
    OutDirectory: TDirectoryEdit;
    IsTestExport: TCheckBox;
    Button1: TButton;
    Bevel1: TBevel;
    Label4: TLabel;
    PeriodCombo: TComboBox;
    Label5: TLabel;
    FormStorage: TFormStorage;
    IsBackgroundMode: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ExpBUROForm: TExpBUROForm;

implementation

{$R *.DFM}

procedure TExpBUROForm.FormCreate(Sender: TObject);
var
     Year, Month, Day : Word;
     INI : TINIFile;
     i : integer;
begin
{     DecodeDate(Date, Year, Month, Day);
     if Month > 5 then
         Month := Month - 5
     else begin
         Year := Year - 1;
         Month := 12 - (5 - Month);
     end;
     for i := 1 to 12 do begin
         PeriodCombo.Items.Add(IntToStr(Month) + '.' + IntToStr(Year));
         Month := Month + 1;
         if Month = 13 then begin
             Month := 1;
             Year := Year + 1;
         end;
     end;
     PeriodCombo.ItemIndex :=
}
{     DecodeDate(Date, Year, Month, Day);
     PeriodFrom.Date := EncodeDate(Year, Month, 1);
     if Month = 12 then begin
        Month := 1;
        Year := Year + 1;
     end
     else
        Month := Month + 1;

     PeriodTo.Date := EncodeDate(Year, Month, 1) - 1;
}
     INI := TINIFile.Create(BLANK_INI);
     OutDirectory.Text := INI.ReadString('MANDATORY', 'BuroOut', 'A:\');
     INI.Free;
end;

procedure TExpBUROForm.Button1Click(Sender: TObject);
begin
    if IsTestExport.Checked then
        IsBackgroundMode.Checked := false;
end;

end.
