unit BasoMonthRptUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, ExtCtrls, StdCtrls, FR_Class, FR_DSet, FR_DBSet, FR_View;

type
  TBasoReportsForm = class(TForm)
    DataSourceQuery: TDataSource;
    frDBDataSetMonth: TfrDBDataSet;
    frReportMonth: TfrReport;
    frPreview: TfrPreview;
    Panel1: TPanel;
    ZoomCombo: TComboBox;
    Button2: TButton;
    Button3: TButton;
    frReportQuarter: TfrReport;
    frReportPolandInside: TfrReport;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure ZoomComboChange(Sender: TObject);
    procedure PrintBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  BasoReportsForm: TBasoReportsForm;

implementation

uses MF;

{$R *.DFM}

procedure TBasoReportsForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
     BasoReportsForm := nil;
end;

procedure TBasoReportsForm.FormCreate(Sender: TObject);
begin
     Left := MainForm.Left;
     Top := MainForm.Top;
     Width := MainForm.Width;
     Height := MainForm.Height;

     ZoomCombo.ItemIndex := 0;
     frPreview.PageWidth;

     DeleteMenu(GetSystemMenu(Handle, FALSE), SC_MOVE, MF_DISABLED OR MF_BYCOMMAND);
end;

procedure TBasoReportsForm.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TBasoReportsForm.ZoomComboChange(Sender: TObject);
begin
     if ZoomCombo.ItemIndex = 0 then
         frPreview.PageWidth
     else
         frPreview.Zoom := ZoomCombo.ItemIndex * 50;
end;

procedure TBasoReportsForm.PrintBtnClick(Sender: TObject);
begin
    frPreview.Print;
end;

end.
