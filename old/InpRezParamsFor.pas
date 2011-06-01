unit InpRezParamsFor;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, RXSpin, ExtCtrls, Placemnt, Grids, Math, Mask, datamod;

type
  RestPercentsData = record
      PutPercent, ComisPercent : double;
  end;

  RestPercentsDataArray = array [1..10] of RestPercentsData;

  TInitRezervLeben = class(TForm)
    Label1: TLabel;
    RepData: TDateTimePicker;
    Button1: TButton;
    Bevel1: TBevel;
    CalcBtn: TButton;
    Button3: TButton;
    Button4: TButton;
    Label2: TLabel;
    Label3: TLabel;
    FizichTax: TRxSpinEdit;
    FormStorage: TFormStorage;
    UridichTax: TRxSpinEdit;
    IsFast: TCheckBox;
    StringGrid: TStringGrid;
    Button5: TButton;
    Label5: TLabel;
    OutFonds: TStringGrid;
    Label4: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure CalcBtnClick(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure FormHide(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
     S : RestPercentsDataArray;
  end;
var
  InitRezervLeben: TInitRezervLeben;

implementation

uses CurrRatesUnit, DaysForSellCurrUnit, 
  MandSellCurrUnit, MoveCurrPeriods, registry;

{$R *.DFM}

procedure TInitRezervLeben.FormCreate(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
begin
     RepData.Date := Date;
     StringGrid.Cells[0, 0] := 'Ассистанс';
     StringGrid.ColWidths[0] := ROUND(StringGrid.Width * 0.4);
     StringGrid.Cells[1, 0] := 'Передано ответственности';
     StringGrid.Cells[2, 0] := 'Перестр. комиссия';
     OutFonds.Cells[0, 0] := 'Дата начала';
     OutFonds.Cells[1, 0] := 'Процент';
     INI := TRegIniFile.Create('Страхование');
     OutFonds.ColWidths[0] := OutFonds.Width div 2 - 7;
     OutFonds.ColWidths[1] := OutFonds.Width div 2 - 7;
     for i := 0 to 10 do begin
        OutFonds.Cells[0, i + 1] := INI.ReadString('Отчисления в фонды', 'Дата' + IntToStr(i), '');
        OutFonds.Cells[1, i + 1] := INI.ReadString('Отчисления в фонды', 'Процент' + IntToStr(i), '');
        if OutFonds.Cells[0, i + 1] = '' then break;
     end;
     INI.Free;
end;

procedure TInitRezervLeben.CalcBtnClick(Sender: TObject);
var
     i : integer;
     List : TASKOneTaxList;
begin
     if not FillOutTaxList(List, InitRezervLeben.OutFonds) then exit;
     for i := 1 to min(10, StringGrid.RowCount - 1) do begin
         S[i].PutPercent := 0;
         S[i].ComisPercent := 0;
         if StringGrid.Cells[1, i] <> '' then
             S[i].PutPercent := StrToFloat(StringGrid.Cells[1, i]);
         if StringGrid.Cells[2, i] <> '' then
             S[i].ComisPercent := StrToFloat(StringGrid.Cells[2, i]);
     end;
end;

procedure TInitRezervLeben.Button3Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

procedure TInitRezervLeben.Button4Click(Sender: TObject);
begin
     DaysForSellCurr.ShowModal
end;

procedure TInitRezervLeben.FormShow(Sender: TObject);
var
     i : integer;
     INI : TRegIniFile;
begin
     StringGrid.RowCount := 1;
     for i := 0 to 10 do
       if InsLeben2.AssistNames[i] <> '' then begin
         StringGrid.RowCount := StringGrid.RowCount + 1;
         StringGrid.Cells[0, StringGrid.RowCount - 1] := InsLeben2.AssistNames[i];
       end
       else
         break;

     INI := TRegIniFile.Create('Страхование жизни');
     for i := 1 to StringGrid.RowCount - 1 do begin
         StringGrid.Cells[1, i] := INI.ReadString('Percents', StringGrid.Cells[0, i] + '1', '');
         StringGrid.Cells[2, i] := INI.ReadString('Percents', StringGrid.Cells[0, i] + '2', '');
     end;
     INI.Free;

     StringGrid.FixedRows := 1;
end;

procedure TInitRezervLeben.Button5Click(Sender: TObject);
begin
    MoveCurr.ShowModal
end;

procedure TInitRezervLeben.FormHide(Sender: TObject);
var
     i : integer;
     INI : TRegIniFile;
begin
     INI := TRegIniFile.Create('Страхование жизни');
     for i := 1 to StringGrid.RowCount - 1 do begin
         INI.WriteString('Percents', StringGrid.Cells[0, i] + '1', StringGrid.Cells[1, i]);
         INI.WriteString('Percents', StringGrid.Cells[0, i] + '2', StringGrid.Cells[2, i]);
     end;
     INI.Free;
end;

end.
