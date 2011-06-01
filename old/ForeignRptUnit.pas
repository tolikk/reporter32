unit ForeignRptUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, FR_View, FR_DSet, FR_DBSet, FR_Class, Db;

type
  TForeignRpt = class(TForm)
    Panel1: TPanel;
    ZoomCombo: TComboBox;
    Button2: TButton;
    Button3: TButton;
    frPreview: TfrPreview;
    DataSource: TDataSource;
    frReport: TfrReport;
    frDBDataSet: TfrDBDataSet;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure ZoomComboChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  ForeignRpt: TForeignRpt;

implementation

uses ForeignUnit;

{$R *.DFM}

procedure TForeignRpt.FormClose(Sender: TObject; var Action: TCloseAction);
begin
     Action := caFree;
end;

procedure TForeignRpt.FormDestroy(Sender: TObject);
begin
     ForeignRpt := nil;
end;

procedure TForeignRpt.ZoomComboChange(Sender: TObject);
begin
     if ZoomCombo.ItemIndex = 0 then
         frPreview.PageWidth
     else
         frPreview.Zoom := ZoomCombo.ItemIndex * 50;
end;

procedure TForeignRpt.FormCreate(Sender: TObject);
begin
     Left := InsForeign.Left;
     Top := InsForeign.Top;
     Width := InsForeign.Width;
     Height := InsForeign.Height;

     ZoomCombo.ItemIndex := 0;
     frPreview.PageWidth;

     DeleteMenu(GetSystemMenu(Handle, FALSE), SC_MOVE, MF_DISABLED OR MF_BYCOMMAND);
end;

procedure TForeignRpt.Button3Click(Sender: TObject);
begin
     Close
end;

procedure TForeignRpt.Button2Click(Sender: TObject);
begin
     frPreview.Print
end;

end.
