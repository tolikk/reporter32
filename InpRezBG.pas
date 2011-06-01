unit InpRezBG;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, RXSpin, ExtCtrls, Placemnt, Mask, Grids,
  registry, datamod;

type
  TInitRezervBG_nouse = class(TForm)
    Label1: TLabel;
    RepData: TDateTimePicker;
    Button1: TButton;
    Bevel1: TBevel;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Label2: TLabel;
    Label3: TLabel;
    FizichTax: TRxSpinEdit;
    UridichTax: TRxSpinEdit;
    Label4: TLabel;
    CompanyPercent: TRxSpinEdit;
    FormStorage: TFormStorage;
    IsFast: TCheckBox;
    Label5: TLabel;
    OutFonds: TStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  InitRezervBG_nouse: TInitRezervBG_nouse;

implementation

uses CurrRatesUnit, DaysForSellCurrUnit, BelGreenRpt;

{$R *.DFM}

procedure TInitRezervBG_nouse.FormCreate(Sender: TObject);
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
        if OutFonds.Cells[0, i + 1] = '' then break;
     end;
     INI.Free;
end;

procedure TInitRezervBG_nouse.Button3Click(Sender: TObject);
begin
     CurrencyRates.ShowModal
end;

procedure TInitRezervBG_nouse.Button4Click(Sender: TObject);
begin
     DaysForSellCurr.ShowModal
end;

procedure TInitRezervBG_nouse.FormDestroy(Sender: TObject);
var
     INI : TRegIniFile;
     i : integer;
begin
     INI := TRegIniFile.Create('Страхование');
     for i := 0 to 10 do begin
        if OutFonds.Cells[0, i + 1] = '' then break;
        INI.WriteString('Отчисления в фонды', 'Дата' + IntToStr(i), OutFonds.Cells[0, i + 1]);
        INI.WriteString('Отчисления в фонды', 'Процент' + IntToStr(i), OutFonds.Cells[1, i + 1]);
     end;
     INI.Free;
end;

procedure TInitRezervBG_nouse.Button2Click(Sender: TObject);
//var
    //List : TASKOneTaxList;
begin
    //if not FillOutTaxList(list, PTStringGrid(@OutFonds)) then exit; 
    modalresult := mrOk;
end;

end.
