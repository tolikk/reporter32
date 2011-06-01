unit DM;

interface

uses
  SysUtils, Classes, DB, DBTables;

type
  TDMM = class(TDataModule)
    HandBookSQL: TQuery;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  DMM: TDMM;

implementation

{$R *.dfm}

end.
