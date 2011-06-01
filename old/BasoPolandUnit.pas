unit BasoPolandUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FR_Class, FR_DSet, FR_DBSet, Db, ExtCtrls, FR_View, FR_Dock, StdCtrls;

type
  TBasoPolandReportsForm = class(TForm)
    DataSourceQuery: TDataSource;
    frDBDataSetMonth: TfrDBDataSet;
    frReportEUROPA: TfrReport;
    frPreview: TfrPreview;
    Panel1: TPanel;
    ZoomCombo: TComboBox;
    Button2: TButton;
    Button3: TButton;
    frReportSVOD: TfrReport;
    frReportQuarter: TfrReport;
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
  BasoPolandReportsForm: TBasoPolandReportsForm;

implementation

uses MF, PolandUnit;

{$R *.DFM}

procedure TBasoPolandReportsForm.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
     BasoPolandReportsForm := nil;
end;

procedure TBasoPolandReportsForm.FormCreate(Sender: TObject);
begin
     Left := PolandPolises.Left;
     Top := PolandPolises.Top;
     Width := PolandPolises.Width;
     Height := PolandPolises.Height;

     ZoomCombo.ItemIndex := 0;
     frPreview.PageWidth;

     DeleteMenu(GetSystemMenu(Handle, FALSE), SC_MOVE, MF_DISABLED OR MF_BYCOMMAND);
end;

procedure TBasoPolandReportsForm.Button1Click(Sender: TObject);
begin
     Close
end;

procedure TBasoPolandReportsForm.ZoomComboChange(Sender: TObject);
begin
     if ZoomCombo.ItemIndex = 0 then
         frPreview.PageWidth
     else
         frPreview.Zoom := ZoomCombo.ItemIndex * 50;
end;

procedure TBasoPolandReportsForm.PrintBtnClick(Sender: TObject);
begin
    frPreview.Print;
end;

end.
