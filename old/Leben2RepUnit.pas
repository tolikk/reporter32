unit Leben2RepUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FR_DSet, FR_DBSet, FR_Class, ExtCtrls, FR_View, StdCtrls, Db;

type
  TLebenReport = class(TForm)
    frPreview: TfrPreview;
    frReport: TfrReport;
    frDBDataSet: TfrDBDataSet;
    Panel1: TPanel;
    ZoomCombo: TComboBox;
    Button2: TButton;
    Button3: TButton;
    DataSource: TDataSource;
    frReportCoris: TfrReport;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button3Click(Sender: TObject);
    procedure ZoomComboChange(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  LebenReport: TLebenReport;

implementation

uses LebenUnit;

{$R *.DFM}

procedure TLebenReport.FormCreate(Sender: TObject);
begin
     Left := InsLeben2.Left;
     Top := InsLeben2.Top;
     Width := InsLeben2.Width;
     Height := InsLeben2.Height;

     ZoomCombo.ItemIndex := 0;
     frPreview.PageWidth;

     DeleteMenu(GetSystemMenu(Handle, FALSE), SC_MOVE, MF_DISABLED OR MF_BYCOMMAND);
end;

procedure TLebenReport.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
     Action := caFree;
     LebenReport := nil;
end;

procedure TLebenReport.Button3Click(Sender: TObject);
begin
     Close;
end;

procedure TLebenReport.ZoomComboChange(Sender: TObject);
begin
     if ZoomCombo.ItemIndex = 0 then
         frPreview.PageWidth
     else
         frPreview.Zoom := ZoomCombo.ItemIndex * 50;
end;

procedure TLebenReport.Button2Click(Sender: TObject);
begin
     frPreview.Print
end;

end.
