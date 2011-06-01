unit CurrRatesUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Placemnt, Db, DBTables, Buttons, registry;

type
  TCurrencyRates = class(TForm)
    GroupBox1: TGroupBox;
    SpeedButton1: TSpeedButton;
    L1: TLabel;
    L2: TLabel;
    L4: TLabel;
    L5: TLabel;
    RateTableName: TEdit;
    FieldDate: TComboBox;
    FieldUSD: TComboBox;
    FieldEUR: TComboBox;
    FieldRUR: TComboBox;
    OpenDialog: TOpenDialog;
    RateTable: TTable;
    Button1: TButton;
    Field1: TComboBox;
    Label1: TLabel;
    Field2: TComboBox;
    Label2: TLabel;
    Field3: TComboBox;
    Label3: TLabel;
    Field4: TComboBox;
    Label4: TLabel;
    Label5: TLabel;
    Field5: TComboBox;
    OtherCurrQuery: TQuery;
    FieldDM_: TComboBox;
    L3: TLabel;
    procedure SpeedButton1Click(Sender: TObject);
    procedure RateTableNameChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    procedure FillCombo(Tbl : TTable; Combo : TComboBox);
    procedure SetIndex(Combo : TComboBox; s : string);
    procedure ReadConfig;
    procedure SaveConfig;
  public
    UseRatesCnt : integer;
    InitTableStr : string;
    UseCurr : array [1..50] of string[5];
    Labels : array [1..5] of TLabel;
    Combos : array [1..5] of TComboBox;

    { Public declarations }
  end;

var
  CurrencyRates: TCurrencyRates;
  UseOtherCurrSQL : string;
  UseDB : string;

implementation

{$R *.DFM}

procedure TCurrencyRates.SpeedButton1Click(Sender: TObject);
begin
     OpenDialog.FileName := RateTableName.Text;
     if OpenDialog.Execute then
         RateTableName.Text := OpenDialog.FileName
end;

procedure TCurrencyRates.RateTableNameChange(Sender: TObject);
var
     i : integer;
begin
     FieldDate.Clear;
     FieldUSD.Clear;
//     FieldDM.Clear;
     FieldEUR.Clear;
     FieldRUR.Clear;

     RateTable.TableName := RateTableName.Text;
     try
         RateTable.Open;

         FillCombo(RateTable, FieldDate);
         FillCombo(RateTable, FieldUSD);
//         FillCombo(RateTable, FieldDM);
         FillCombo(RateTable, FieldEUR);
         FillCombo(RateTable, FieldRUR);

         for i := 1 to UseRatesCnt do begin
             if Combos[i].Visible then begin
                 FillCombo(RateTable, Combos[i]);
             end;
         end;

         RateTable.Close;
     except
     end;
end;

procedure TCurrencyRates.FillCombo(Tbl : TTable; Combo : TComboBox);
var
    i : integer;
begin
    for i := 0 to Tbl.FieldCount - 1 do
        Combo.Items.Add(Tbl.Fields[i].FieldName);
end;

procedure TCurrencyRates.SetIndex(Combo : TComboBox; s : string);
var
    i : integer;
begin
    for i := 0 to Combo.Items.Count - 1 do
        if Combo.Items[i] = s then
            Combo.ItemIndex := i;
end;

procedure TCurrencyRates.SaveConfig;
var
     INI : TRegIniFile;
     i : integer;
begin
     INI := TRegIniFile.Create('Страхование');
     INI.WriteString('Курсы валют', 'Table', RateTableName.Text);
     INI.WriteString('Курсы валют', 'Date', FieldDate.Text);
     INI.WriteString('Курсы валют', 'USD', FieldUSD.Text);
//     INI.WriteString('Курсы валют', 'DM', FieldDM.Text);
     INI.WriteString('Курсы валют', 'EUR', FieldEUR.Text);
     INI.WriteString('Курсы валют', 'RUR', FieldRUR.Text);
     for i := 1 to UseRatesCnt do begin
         if Combos[i].Visible then
             INI.WriteString('Курсы валют', UseCurr[i], Combos[i].Text);
     end;
     INI.Free;
end;

procedure TCurrencyRates.ReadConfig;
var
    INI : TRegIniFile;
    i : integer;
begin
     INI := TRegIniFile.Create('Страхование');
     RateTableName.Text := INI.ReadString('Курсы валют', 'Table', '');
     SetIndex(FieldDate, INI.ReadString('Курсы валют', 'Date', ''));
     SetIndex(FieldUSD, INI.ReadString('Курсы валют', 'USD', ''));
//     SetIndex(FieldDM, INI.ReadString('Курсы валют', 'DM', ''));
     SetIndex(FieldEUR, INI.ReadString('Курсы валют', 'EUR', ''));
     SetIndex(FieldRUR, INI.ReadString('Курсы валют', 'RUR', ''));

     for i := 1 to UseRatesCnt do begin
         if Combos[i].Visible then
             SetIndex(Combos[i], INI.ReadString('Курсы валют', UseCurr[i], ''));
     end;

     INI.Free;
end;

procedure TCurrencyRates.FormCreate(Sender: TObject);
var
     i : integer;
begin
     Labels[1] := Label1;
     Labels[2] := Label2;
     Labels[3] := Label3;
     Labels[4] := Label4;
     Labels[5] := Label5;

     Combos[1] := Field1;
     Combos[2] := Field2;
     Combos[3] := Field3;
     Combos[4] := Field4;
     Combos[5] := Field5;

     UseRatesCnt := 0;
     if UseOtherCurrSQL <> '' then begin
         OtherCurrQuery.SQL.Add(UseOtherCurrSQL);
         OtherCurrQuery.DatabaseName := UseDB;
         OtherCurrQuery.Open;
         i := 1;

         while not OtherCurrQuery.EOF do begin
             UseCurr[i] := OtherCurrQuery.FieldByName('CURR').AsString;
             Combos[i].Visible := true;
             Labels[i].Visible := true;
             Labels[i].Caption := 'Курс ' + UseCurr[i];
             if i = 5 then break;
             Inc(i);
             OtherCurrQuery.Next;
             UseRatesCnt := UseRatesCnt + 1;
         end;

         OtherCurrQuery.Close;

         //Resize
         Height := Height + UseRatesCnt * (Field2.Top - Field1.Top);
     end;
         
     ReadConfig;

     InitTableStr := RateTableName.Text;
end;

procedure TCurrencyRates.FormDestroy(Sender: TObject);
begin
     if InitTableStr <> RateTableName.Text then
         SaveConfig
end;

end.
