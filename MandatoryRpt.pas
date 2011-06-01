unit MandatoryRpt;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, Quickrpt, QRCtrls, QrPrntr;
type
  TMandatoryReport = class(TQuickRep)
    TitleBand1: TQRBand;
    SummaryBand1: TQRBand;
    ColumnHeaderBand1: TQRBand;
    DetailBand: TQRBand;
    PageFooterBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRSysData1: TQRSysData;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRDBText8: TQRDBText;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRDBText11: TQRDBText;
    QRSysData2: TQRSysData;
    QRLabel17: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRShape15: TQRShape;
    QRShape16: TQRShape;
    QRShape17: TQRShape;
    QRShape18: TQRShape;
    QRShape19: TQRShape;
    QRShape20: TQRShape;
    QRShape21: TQRShape;
    QRShape22: TQRShape;
    QRShape23: TQRShape;
    QRLabel18: TQRLabel;
    QRShape24: TQRShape;
    QRShape25: TQRShape;
    QRShape26: TQRShape;
    QRShape27: TQRShape;
    QRShape28: TQRShape;
    QRShape29: TQRShape;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    PolicesCount: TQRLabel;
    QRLabel22: TQRLabel;
    BadCount: TQRLabel;
    QRLabel23: TQRLabel;
    RubSumma: TQRLabel;
    QRLabel25: TQRLabel;
    EuroSumma: TQRLabel;
    QRLabel24: TQRLabel;
    Valutas: TQRLabel;
    QRShape30: TQRShape;
    QRShape31: TQRShape;
    QRShape32: TQRShape;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    LeftText: TQRDBText;
    RightText: TQRDBText;
    QRShape33: TQRShape;
    QRDBText12: TQRDBText;
    QRShape34: TQRShape;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    Polices2Count: TQRLabel;
    procedure MandatoryReportPreview(Sender: TObject);
    procedure DetailBandBeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private

  public

  end;

var
  MandatoryReport: TMandatoryReport;
  Line : integer;

implementation

uses MyPreview;

{$R *.DFM}

procedure TMandatoryReport.MandatoryReportPreview(Sender: TObject);
begin
  QRMyPreview.QRPreview.QRPrinter := Sender as TQRPrinter;
  QRMyPreview.Report := MandatoryReport;
  QRMyPreview.Show;
end;

procedure TMandatoryReport.DetailBandBeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
     if (Line MOD 2) = 1 then
        DetailBand.Color := clWhite
     else
        DetailBand.Color := RGB(223, 223, 223);

     Inc(Line);
end;

end.
