unit InpRezParams;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, RXSpin, ExtCtrls, Placemnt, Mask, Grids, registry, datamod;

type
  TInitRezerv = class(TForm)
    Label1: TLabel;
    RepData: TDateTimePicker;
    Button1: TButton;
    Bevel1: TBevel;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Label2: TLabel;
    Label3: TLabel;
    FizichTax1: TRxSpinEdit;
    UridichTax1: TRxSpinEdit;
    Label4: TLabel;
    CompanyPercent1: TRxSpinEdit;
    FormStorage: TFormStorage;
    IsFast: TCheckBox;
    Label5: TLabel;
    OutFonds: TStringGrid;
    FizTax2: TRxSpinEdit;
    UrTax2: TRxSpinEdit;
    FizTaxDate: TDateTimePicker;
    UrTaxDate: TDateTimePicker;
    RestraxPart: TStringGrid;
    IsRestPart: TCheckBox;
    CompanyPercentDate: TDateTimePicker;
    CompanyPercent2: TRxSpinEdit;
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    List : TASKOneTaxList;
    ListRest : TASKOneTaxList;
    procedure RestoreName(name : string);
    { Public declarations }
  end;

var
  InitRezerv: TInitRezerv;

implementation

uses CurrRatesUnit, DaysForSellCurrUnit;

{$R *.DFM}

procedure TInitRezerv.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
begin
     RepData.Date := Date;
     OutFonds.Cells[0, 0] := 'Дата начала';
     OutFonds.Cells[1, 0] := 'Процент';
     INI := TRegIniFile.Create('Страхование');
     OutFonds.ColWidths[0] := OutFonds.Width div 2 - 5;
     OutFonds.ColWidths[1] := OutFonds.Width div 2 - 5;
     for i := 0 to 10 do begin
        OutFonds.Cells[0, i + 1] := INI.ReadString('Отчисления в фонды', 'Дата' + IntToStr(i), '');
        OutFonds.Cells[1, i + 1] := INI.ReadString('Отчисления в фонды', 'Процент' + IntToStr(i), '');
        RestraxPart.Cells[0, i + 1] := INI.ReadString('Доля перестраховщика', 'Дата' + IntToStr(i), '');
        RestraxPart.Cells[1, i + 1] := INI.ReadString('Доля перестраховщика', 'Процент' + IntToStr(i), '');
        if OutFonds.Cells[0, i + 1] = '' then break;
        if RestraxPart.Cells[0, i + 1] = '' then break;
     end;
     RestraxPart.Cells[0, 0] := 'Дата начала';
     RestraxPart.Cells[1, 0] := 'Процент';
     RestraxPart.ColWidths[0] := RestraxPart.Width div 2 - 5;
     RestraxPart.ColWidths[1] := RestraxPart.Width div 2 - 5;
     INI.Free;
end;

procedure TInitRezerv.Button3Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

procedure TInitRezerv.Button4Click(Sender: TObject);
begin
     DaysForSellCurr.ShowModal
end;

procedure TInitRezerv.Button2Click(Sender: TObject);
begin
    FillChar(ListRest, sizeof(ListRest), 0);
    if not FillOutTaxList(List, OutFonds) then exit;
    if IsRestPart.Checked then
        if not FillOutTaxList(ListRest, RestraxPart) then exit;
    modalresult := mrok;
end;

procedure TInitRezerv.FormDestroy(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
begin
     INI := TRegIniFile.Create('Страхование');
     for i := 0 to 10 do begin
        if OutFonds.Cells[0, i + 1] <> '' then begin
            INI.WriteString('Отчисления в фонды', 'Дата' + IntToStr(i), OutFonds.Cells[0, i + 1]);
            INI.WriteString('Отчисления в фонды', 'Процент' + IntToStr(i), OutFonds.Cells[1, i + 1]);
        end;
        if RestraxPart.Cells[0, i + 1] <> '' then begin
            INI.WriteString('Доля перестраховщика', 'Дата' + IntToStr(i), RestraxPart.Cells[0, i + 1]);
            INI.WriteString('Доля перестраховщика', 'Процент' + IntToStr(i), RestraxPart.Cells[1, i + 1]);
        end;
     end;
     INI.Free;
end;

procedure TInitRezerv.RestoreName(name : string);
begin
    FormStorage.IniSection := name;
    FormStorage.RestoreFormPlacement;
    FormStorage.Active := true;
end;

end.
