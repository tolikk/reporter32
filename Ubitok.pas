unit Ubitok;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, StdCtrls, RXSpin, bde, dbtables, db, Mask;

type
  TUbitokFrm = class(TForm)
    Label1: TLabel;
    RegNmb: TEdit;
    Label2: TLabel;
    RegDate: TDateTimePicker;
    Label6: TLabel;
    Summa: TRxSpinEdit;
    Currency: TComboBox;
    GroupBox1: TGroupBox;
    Label7: TLabel;
    PaySum: TRxSpinEdit;
    PayDate: TDateTimePicker;
    Label8: TLabel;
    PayCurr: TComboBox;
    GroupBox2: TGroupBox;
    Label9: TLabel;
    Label10: TLabel;
    FreeSum: TRxSpinEdit;
    FreeDate: TDateTimePicker;
    FreeCurr: TComboBox;
    Label11: TLabel;
    CloseDate: TDateTimePicker;
    GroupBox3: TGroupBox;
    Label5: TLabel;
    ComplainName: TEdit;
    Label3: TLabel;
    ComplainDate: TDateTimePicker;
    Label4: TLabel;
    ComplainNmb: TEdit;
    SaveBtn: TButton;
    Button2: TButton;
    Label12: TLabel;
    Comments: TMemo;
    Label13: TLabel;
    Country: TEdit;
    Label14: TLabel;
    Bordero: TRxSpinEdit;
    BorderoCurr: TComboBox;
    Label99: TLabel;
    StateCb: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure SaveBtnClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FreeDateChange(Sender: TObject);
    procedure PayDateChange(Sender: TObject);
    procedure StateCbChange(Sender: TObject);
  private
    { Private declarations }
  public
    DataQuery : TDataSet;
    SaveQuery : TDataSet;
    MinDate, MaxDate : TDateTime;
    { Public declarations }
  end;

var
  UbitokFrm: TUbitokFrm;

implementation

{$R *.DFM}

procedure TUbitokFrm.FormShow(Sender: TObject);
begin
  RegNmb.Text := '';
  RegDate.Date := Date;
  ComplainDate.Date := Date;
  ComplainName.Text := 'ЗАЯВЛЕНИЕ';
  ComplainNmb.Text := '';
  Summa.Value := 0.01;
  Bordero.Value := 0.01;
  Currency.ItemIndex := -1;
  BorderoCurr.ItemIndex := -1;
  PayDate.Checked := false;
  PaySum.Value := 0;
  PayCurr.ItemIndex := -1;
  FreeDate.Checked := false;
  FreeSum.Value := 0;
  FreeCurr.ItemIndex := -1;
  Country.Text := '';
  Comments.Text := '';
  CloseDate.Date := Date;
  StateCb.ItemIndex := 0;
  MinDate := 0;
  MaxDate := 0;
  if DataQuery <> nil then
      with DataQuery do begin
        RegNmb.Text := FieldByName('regnmb').AsString;
        RegDate.Date := FieldByName('regdate').AsDateTime;
        ComplainDate.Date := FieldByName('ComplDate').AsDateTime;
        ComplainName.Text := FieldByName('ComplName').AsString;
        ComplainNmb.Text := FieldByName('ComplNmb').AsString;
        Summa.Value := FieldByName('UbSum').AsFloat;
        Currency.ItemIndex := Currency.Items.IndexOf(FieldByName('UbCurr').AsString);
        Bordero.Value := FieldByName('BrSum').AsFloat;
        BorderoCurr.ItemIndex := BorderoCurr.Items.IndexOf(FieldByName('BrCurr').AsString);
        StateCB.ItemIndex := FieldByName('State').AsInteger;
        if not FieldByName('paydate').IsNull then
            PayDate.Date := FieldByName('paydate').AsDateTime;
        PaySum.Value := FieldByName('paysum').AsFloat;
        PayCurr.ItemIndex := PayCurr.Items.IndexOf(FieldByName('paycurr').AsString);

        if not FieldByName('freedate').IsNull then
            FreeDate.Date := FieldByName('freedate').AsDateTime;
        FreeSum.Value := FieldByName('freesum').AsFloat;
        FreeCurr.ItemIndex := FreeCurr.Items.IndexOf(FieldByName('freecurr').AsString);
        Comments.Text := FieldByName('text').AsString;
        Country.Text := FieldByName('Country').AsString;
        CloseDate.Date := FieldByName('cldate').AsDateTime;
      end;
end;

procedure TUbitokFrm.SaveBtnClick(Sender: TObject);
begin
  if (RegDate.Date < MinDate) OR ((MaxDate > 0) AND (RegDate.Date > (MaxDate + 60))) then begin
      MessageDlg('Дата регистрации ' + DateToStr(RegDate.Date) + ' не в диапазоне', mtInformation, [mbOk], 0);
      exit;
  end;
  if SaveQuery <> nil then
      with SaveQuery do begin
        FieldByName('regnmb').AsString := RegNmb.Text;
        FieldByName('regdate').AsDateTime := RegDate.Date;
        FieldByName('ComplDate').AsDateTime := ComplainDate.Date;
        FieldByName('ComplName').AsString := ComplainName.Text;
        FieldByName('ComplNmb').AsString := ComplainNmb.Text;
        FieldByName('UbSum').AsFloat := Summa.Value;
        FieldByName('UbCurr').AsString := Currency.Text;
        FieldByName('BrSum').AsFloat := Bordero.Value;
        FieldByName('BrCurr').AsString := BorderoCurr.Text;
        FieldByName('State').AsInteger := StateCb.ItemIndex;
        if PayDate.Checked then begin
            FieldByName('paydate').AsDateTime := PayDate.Date;
            FieldByName('paysum').AsFloat := PaySum.Value;
            FieldByName('paycurr').AsString := PayCurr.Text;
        end
        else begin
            FieldByName('paydate').Clear;
            FieldByName('paysum').Clear;
            FieldByName('paycurr').Clear;
        end;
        if FreeDate.Checked then begin
            FieldByName('freedate').AsDateTime := FreeDate.Date;
            FieldByName('freesum').AsFloat := FreeSum.Value;
            FieldByName('freecurr').AsString := FreeCurr.Text;
        end
        else begin
            FieldByName('freedate').Clear;
            FieldByName('freesum').Clear;
            FieldByName('freecurr').Clear;
        end;
        FieldByName('text').AsString := Comments.Text;
        FieldByName('Country').AsString := Country.Text;
        FieldByName('cldate').AsDateTime := CloseDate.Date;
      end;

  ModalResult := mrOk;    
end;

procedure TUbitokFrm.Button2Click(Sender: TObject);
begin
    Close
end;

procedure TUbitokFrm.FreeDateChange(Sender: TObject);
begin
    if (FreeDate.Checked) AND (FreeCurr.Text = '') then
        FreeCurr.ItemIndex := Currency.ItemIndex + 1;
    if (not FreeDate.Checked) then
        FreeCurr.ItemIndex := 0;
end;

procedure TUbitokFrm.PayDateChange(Sender: TObject);
begin
    if (PayDate.Checked) AND (PayCurr.Text = '') then
        PayCurr.ItemIndex := Currency.ItemIndex + 1;
    if (not PayDate.Checked) then
        PayCurr.ItemIndex := 0;
end;

procedure TUbitokFrm.StateCbChange(Sender: TObject);
begin
    PayDate.Checked := StateCb.ItemIndex = 2;
    PayDateChange(nil);
end;

end.
