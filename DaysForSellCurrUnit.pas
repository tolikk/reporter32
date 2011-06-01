unit DaysForSellCurrUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, RXSpin, Db, DBTables, Mask;

type
  TDaysForSellCurr = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    USD: TRxSpinEdit;
    DM: TRxSpinEdit;
    EUR: TRxSpinEdit;
    RUR: TRxSpinEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Bevel1: TBevel;
    Button1: TButton;
    LabelNew1: TLabel;
    Days1: TRxSpinEdit;
    DaysLbl1: TLabel;
    LabelNew2: TLabel;
    Days2: TRxSpinEdit;
    DaysLbl2: TLabel;
    LabelNew3: TLabel;
    LabelNew4: TLabel;
    Days3: TRxSpinEdit;
    Days4: TRxSpinEdit;
    DaysLbl4: TLabel;
    DaysLbl3: TLabel;
    LabelNew5: TLabel;
    Days5: TRxSpinEdit;
    DaysLbl5: TLabel;
    OtherCurrQuery: TQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormHide(Sender: TObject);
  private
    { Private declarations }

    procedure SaveConfig;
    procedure ReadConfig;

  public
    UseRatesCnt : integer;
    UseCurr : array [1..50] of string[5];
    Labels : array [1..5] of TLabel;
    Labels2 : array [1..5] of TLabel;
    Combos : array [1..5] of TRXSpinEdit;

    { Public declarations }
  end;

var
  UseOtherCurrSQL : string;
  UseDB : string;
  DaysForSellCurr: TDaysForSellCurr;

implementation

  uses registry;

{$R *.DFM}

procedure TDaysForSellCurr.SaveConfig;
var
     INI : TRegIniFile;
     i : integer;
begin
     INI := TRegIniFile.Create('�����������');
     INI.WriteInteger('����� ������� �����', 'USD', ROUND(USD.Value));
     INI.WriteInteger('����� ������� �����', 'DM', ROUND(DM.Value));
     INI.WriteInteger('����� ������� �����', 'EUR', ROUND(EUR.Value));
     INI.WriteInteger('����� ������� �����', 'RUR', ROUND(RUR.Value));
     for i := 1 to UseRatesCnt do begin
         if Combos[i].Visible then
             INI.WriteString('���� �������', UseCurr[i], IntToStr(ROUND(Combos[i].Value)));
     end;
     INI.Free;
end;

procedure TDaysForSellCurr.ReadConfig;
var
     INI : TRegIniFile;
     i : integer;
begin
     INI := TRegIniFile.Create('�����������');
     USD.Value := INI.ReadInteger('����� ������� �����', 'USD', 0);
     DM.Value := INI.ReadInteger('����� ������� �����', 'DM', 0);
     EUR.Value := INI.ReadInteger('����� ������� �����', 'EUR', 0);
     RUR.Value := INI.ReadInteger('����� ������� �����', 'RUR', 0);
     for i := 1 to UseRatesCnt do begin
         if Combos[i].Visible then
             Combos[i].Text := INI.ReadString('���� �������', UseCurr[i], '1');
     end;
     INI.Free;
end;

procedure TDaysForSellCurr.FormCreate(Sender: TObject);
var
     i : integer;
begin
     Labels[1] := LabelNew1;
     Labels[2] := LabelNew2;
     Labels[3] := LabelNew3;
     Labels[4] := LabelNew4;
     Labels[5] := LabelNew5;

     Labels2[1] := DaysLbl1;
     Labels2[2] := DaysLbl2;
     Labels2[3] := DaysLbl3;
     Labels2[4] := DaysLbl4;
     Labels2[5] := DaysLbl5;

     Combos[1] := Days1;
     Combos[2] := Days2;
     Combos[3] := Days3;
     Combos[4] := Days4;
     Combos[5] := Days5;

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
             Labels2[i].Visible := true;
             Labels[i].Caption := UseCurr[i];
             if i = 5 then break;
             Inc(i);
             OtherCurrQuery.Next;
             UseRatesCnt := UseRatesCnt + 1;
         end;

         OtherCurrQuery.Close;

         //Resize
         Height := Height + UseRatesCnt * (Days2.Top - Days1.Top);
     end;

     ReadConfig;
end;

procedure TDaysForSellCurr.FormHide(Sender: TObject);
begin
     SaveConfig
end;

end.
