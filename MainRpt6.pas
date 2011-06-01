unit MainRpt6;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, RXSpin, Mask, ToolEdit, filectrl, inifiles, rxshell, DataMod;

type
  TMainReport6 = class(TForm)
    Label1: TLabel;
    Monthes: TComboBox;
    Year: TRxSpinEdit;
    Label2: TLabel;
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    Label3: TLabel;
    OutDirectory: TDirectoryEdit;
    UseTab: TCheckBox;
    IsUsePrepare: TCheckBox;
    Label4: TLabel;
    IsTestExport: TCheckBox;
    FlagPeriod: TComboBox;
    IsIndex: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure MonthesChange(Sender: TObject);
    procedure YearChange(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure IsTestExportClick(Sender: TObject);
  private
    { Private declarations }
    procedure CheckData;
  public
    { Public declarations }
  end;

var
  MainReport6: TMainReport6;

implementation

uses WaitFormUnit;

{$R *.DFM}

procedure TMainReport6.FormCreate(Sender: TObject);
var
     INI : TINIFile;
begin
     Monthes.ItemIndex := 5;
     Year.Value := 2000;

     INI := TINIFile.Create(BLANK_INI);
     OutDirectory.Text := INI.ReadString('MANDATORY', 'BuroOut', 'A:\');
     INI.Free;

     FlagPeriod.ItemIndex := 0;
end;

procedure TMainReport6.MonthesChange(Sender: TObject);
begin
     CheckData;
end;

procedure TMainReport6.YearChange(Sender: TObject);
begin
     CheckData;
end;

procedure TMainReport6.CheckData;
begin
      if Year.Value = 2000 then
         if Monthes.ItemIndex < 6 then
             Monthes.ItemIndex := 6;
end;

procedure TMainReport6.Button1Click(Sender: TObject);
var
     INI : TIniFile;
begin
     INI := TINIFile.Create(BLANK_INI);
     INI.WriteString('MANDATORY', 'BuroOut', OutDirectory.Text);
     INI.WriteString('MANDATORY', 'Create Index', 'None');

     if not DirectoryExists(OutDirectory.Text) then begin
        MessageDlg('Путь ' + OutDirectory.Text + ' не существует', mtInformation, [mbOK], 0);
     end
     else begin
        if IsIndex.Checked then
           if FileExecuteWait('ExpIdx.exe', 'CREATE', '', esHidden) = -1 then
               MessageDlg('Программа создания индексов ExpIdx.exe не найдена', mtInformation, [mbOk], 0)
           else begin
               Application.CreateForm(TWaitForm, WaitForm);
               WaitForm.Label1.Caption := 'Создание индексов...';
               WaitForm.Update;
               while INI.ReadString('MANDATORY', 'Create Index', '') = 'None' do
                  Sleep(1000);
               WaitForm.Close;
           end;
        ModalResult := mrOk;
     end;

     INI.Free;
end;

procedure TMainReport6.IsTestExportClick(Sender: TObject);
begin
     IsIndex.Enabled := not IsTestExport.Checked;
     if not IsIndex.Enabled then
         IsIndex.Checked := false;
end;

end.
