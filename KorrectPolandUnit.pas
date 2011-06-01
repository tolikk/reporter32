unit KorrectPolandUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, RXSpin, RxLookup, Db, DBTables, DBCtrls, RxQuery, ComCtrls,
  Mask;

type
  TKorrectPoland = class(TForm)
    Button1: TButton;
    Button2: TButton;
    GroupBox1: TGroupBox;
    IsSetAgPercent: TCheckBox;
    IsSetUrFizich: TCheckBox;
    AgPercent: TRxSpinEdit;
    IsUrFizich: TComboBox;
    GroupBox2: TGroupBox;
    IsFiltSeria: TCheckBox;
    IsFiltNumber: TCheckBox;
    IsFiltAgent: TCheckBox;
    IsFiltNull: TCheckBox;
    NumberMin: TEdit;
    NumberMax: TEdit;
    Label1: TLabel;
    DataSourceAg: TDataSource;
    DataSourceSeries: TDataSource;
    SeriesCombo: TRxDBLookupCombo;
    AgentCombo: TRxDBLookupCombo;
    WorkSQL: TQuery;
    AgTable: TQuery;
    AgTableAgent_code: TStringField;
    AgTableName: TStringField;
    Series: TRxQuery;
    IsRegDate: TCheckBox;
    RegDateFrom: TDateTimePicker;
    RegDateTo: TDateTimePicker;
    Label2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    DbName : string;
    TableName : string;
    AgentCodeFld : string;
    AgentPercentFld : string;
    UridichFizichFld : string;
    NULLFld : string;
    IsUridichFldStr : boolean;
    RegDateFld : string;
    UR_FIZ_VLA : string;
  end;

var
  KorrectPoland: TKorrectPoland;

implementation

{$R *.DFM}

procedure TKorrectPoland.FormCreate(Sender: TObject);
begin
     IsUrFizich.ItemIndex := 0;
     DbName := '';
     TableName := '';
     AgentCodeFld := '';
     AgentPercentFld := '';
     UridichFizichFld := '';
     NULLFld := '';
     IsUridichFldStr := false;
     RegDateFld := '';
     UR_FIZ_VLA := 'YN';
end;

procedure TKorrectPoland.FormShow(Sender: TObject);
begin
     if AgTable.Active = false then begin
         AgTable.DataBaseName := DbName;
         AgTable.Open;
     end;
     if Series.Active = false then begin
         Series.DataBaseName := DbName;
         Series.MacrobyName('TABLE').AsString := TableName;
         Series.Open;
     end;
     if RegDateFld = '' then begin
        IsRegDate.Enabled := false;
        RegDateFrom.Enabled := false;
        RegDateTo.Enabled := false;
     end;
     AgTable.Refresh;
     Series.Refresh;

     if UridichFizichFld = '' then begin
         IsSetUrFizich.Enabled := false;
         IsSetUrFizich.Checked := false;
         IsUrFizich.Enabled := false;
     end;
end;

procedure TKorrectPoland.Button2Click(Sender: TObject);
var
     sqlText : string;
     whereText : string;
begin
     if not IsSetUrFizich.Checked AND not IsSetAgPercent.Checked then begin
        MessageDlg('Ќе выбрано ни одно поле дл€ изменени€', mtError, [mbOk], 0);
        exit;
     end;

     if not IsFiltSeria.Checked AND not IsFiltNumber.Checked AND not IsFiltAgent.Checked AND not IsRegDate.Checked then begin
        if MessageDlg('¬ы хотите изменить данные без каких либо ограничений?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
            exit;
     end;

     WorkSQL.DataBaseName := DbName;
     whereText := ' ' + NULLFld + ' IS NOT NULL ';
     sqlText := 'UPDATE ' + TableName + ' SET ';
     if IsSetUrFizich.Checked then sqlText := sqlText + ' ' + UridichFizichFld + '=:URIDICH,';
     if IsSetAgPercent.Checked then sqlText := sqlText + ' ' + AgentPercentFld + '=:AGPERCENT,';
     sqlText := Copy(sqlText, 1, Length(sqlText) - 1);

     if IsFiltSeria.Checked then begin
         if SeriesCombo.Text = '' then begin
            SeriesCombo.DropDown;
            exit;
         end;
         if whereText <> '' then whereText := whereText + ' AND ';
         whereText := whereText + 'SERIA=''' + Series.FieldByName('SERIA').AsString + '''';
     end;
     if IsFiltNumber.Checked then begin
         if NumberMin.Text <> '' then begin
            if whereText <> '' then whereText := whereText + ' AND ';
            whereText := whereText + 'NUMBER>=' + NumberMin.Text;
         end;
         if NumberMax.Text <> '' then begin
            if whereText <> '' then whereText := whereText + ' AND ';
            whereText := whereText + 'NUMBER<=' + NumberMax.Text;
         end;
     end;
     if IsFiltAgent.Checked then begin
         if AgentCombo.Text = '' then begin
            AgentCombo.DropDown;
            exit;
         end;
         if whereText <> '' then whereText := whereText + ' AND ';
         whereText := whereText + AgentCodeFld + '=''' + AgTableAgent_code.AsString + '''';
     end;
     if IsRegDate.Checked then begin
        if RegDateFrom.Checked then begin
            if whereText <> '' then whereText := whereText + ' AND ';
            whereText := whereText + RegDateFld + '>=''' + DateToStr(RegDateFrom.Date) + '''';
        end;
        if RegDateTo.Checked then begin
            if whereText <> '' then whereText := whereText + ' AND ';
            whereText := whereText + RegDateFld + '<=''' + DateToStr(RegDateTo.Date) + '''';
        end;
     end;

     if whereText <> '' then sqlText := sqlText + ' WHERE ' + whereText;

     try
         Screen.Cursor := crHourGlass;
         WorkSQL.SQL.Clear;
         WorkSQL.SQL.Add(sqlText);
         if IsSetUrFizich.Checked then begin
             if IsUridichFldStr then begin
                 if IsUrFizich.ItemIndex = 1 then
                     WorkSQL.ParamByName('URIDICH').AsString := UR_FIZ_VLA[1]
                 else
                     WorkSQL.ParamByName('URIDICH').AsString := UR_FIZ_VLA[2];
             end
             else
                 WorkSQL.ParamByName('URIDICH').AsBoolean := IsUrFizich.ItemIndex = 1;
         end;
         if IsSetAgPercent.Checked then WorkSQL.ParamByName('AGPERCENT').AsFloat := AgPercent.Value;
         WorkSQL.ExecSQL;
     except
         on E : Exception do
             MessageDlg(e.Message, mtInformation, [mbOk], 0);
     end;

     Screen.Cursor := crDefault;
end;

end.
