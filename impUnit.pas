unit impUnit;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ActnList, Mask, ToolEdit, Db, DBTables, DataMod;

type
  TImportData = class(TForm)
    Label1: TLabel;
    DestFile: TEdit;
    Label2: TLabel;
    ListInsType: TComboBox;
    Label3: TLabel;
    SetDestFile: TButton;
    ListFiles: TListBox;
    Button2: TButton;
    ActionList: TActionList;
    actionImport: TAction;
    actionSelSourceFiles: TAction;
    OpenDialog: TOpenDialog;
    Directory: TDirectoryEdit;
    DestTable: TTable;
    SrcTable: TTable;
    procedure actionImportUpdate(Sender: TObject);
    procedure actionImportExecute(Sender: TObject);
    procedure actionSelSourceFilesUpdate(Sender: TObject);
    procedure SetDestFileClick(Sender: TObject);
    procedure DirectoryChange(Sender: TObject);
    procedure DestFileChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ListInsTypeChange(Sender: TObject);
  private
    { Private declarations }
    function CompareTables(tblNam : string) : boolean;
  public
    { Public declarations }
  end;

var
  ImportData: TImportData;

implementation

uses ImpFormParam, inifiles;

{$R *.DFM}

procedure TImportData.actionImportUpdate(Sender: TObject);
begin
     actionImport.Enabled := ListFiles.SelCount > 0;
end;

procedure TImportData.actionImportExecute(Sender: TObject);
var
     i : integer;
begin
     AskImport.Memo.Clear;

     for i := 0 to (ListFiles.Items.Count - 1) do
         if ListFiles.Selected[i] then
             AskImport.Memo.Lines.Add(ListFiles.Items[i]);

     if AskImport.ShowModal = mrOk then begin
         for i := ListFiles.Items.Count - 1 downto 0 do
             if ListFiles.Selected[i] then
                 ListFiles.Items.Delete(i);

         ListFiles.ItemIndex := 0;

         if (ListFiles.Items.Count = 0) AND (ListInsType.ItemIndex <> -1) then
             ListInsType.Items.Delete(ListInsType.ItemIndex);
     end;
end;

procedure TImportData.actionSelSourceFilesUpdate(Sender: TObject);
begin
     actionSelSourceFiles.Enabled := (DestFile.Text <> '') AND FileExists(DestFile.Text);
end;

procedure TImportData.SetDestFileClick(Sender: TObject);
begin
     if OpenDialog.Execute then begin
        DestFile.Text := OpenDialog.FileName;
     end;
end;

function TImportData.CompareTables(tblNam : string) : boolean;
var
     i : integer;
begin
     DestTable.TableName := DestFile.Text;
     SrcTable.TableName := tblNam;
     CompareTables := false;

     try
         DestTable.Open;
         SrcTable.Open;
         //if SrcTable.IndexFieldCount <> 0 then
           if DestTable.Fields.Count = SrcTable.Fields.Count then begin
               CompareTables := true;
               for i := 0 to DestTable.Fields.Count - 1 do
                   if SrcTable.FindField(DestTable.Fields[i].FieldName) = nil then
                       CompareTables := false;
           end;
     except
     end;

     DestTable.Close;
     SrcTable.Close;
end;

procedure TImportData.DirectoryChange(Sender: TObject);
var

  SearchRec: TSearchRec;
begin
     ListFiles.Items.Clear;
     if not ((DestFile.Text <> '') AND FileExists(DestFile.Text)) then exit;
     Screen.Cursor := crHourGlass;
     if 0 = FindFirst(Directory.Text + '\*.db', faAnyFile, SearchRec) then
         repeat
             if UpperCase((Directory.Text + '\' + SearchRec.Name)) <> UpperCase(DestFile.Text) then
               if CompareTables(Directory.Text + '\' + SearchRec.Name) then begin
                   ListFiles.Items.Add(Directory.Text + '\' + SearchRec.Name);
                   ListFiles.Update
               end
         until FindNext(SearchRec) <> 0;

     ListFiles.ItemIndex := 0;
     FindClose(SearchRec);
     Screen.Cursor := crDefault;
end;

procedure TImportData.DestFileChange(Sender: TObject);
begin
     Directory.Text := '';
end;

procedure TImportData.FormCreate(Sender: TObject);
var
     INI : TIniFile;
     s : string;
     i : integer;
begin
     INI := TIniFile.Create(BLANK_INI);
     for i := 0 to 100 do begin
         s := INI.ReadString('IMPORT', 'Vid' + IntToStr(i), '');
         if s <> '' then begin
             ListInsType.Items.Objects[ListInsType.Items.Add(s)] := TObject(i);
         end;
     end;
     INI.Free;
end;

procedure TImportData.ListInsTypeChange(Sender: TObject);
var
     INI : TIniFile;
     s : string;
     index : integer;
begin
     if ListInsType.ItemIndex = -1 then exit;
     INI := TIniFile.Create(BLANK_INI);
     index := Integer(ListInsType.Items.Objects[ListInsType.ItemIndex]);
//     MessageDlg(IntToStr(index), mtInformation, [mbOk], 0);
     DestFile.Text := INI.ReadString('IMPORT', 'Dest' + IntToStr(index), '');
     if DestFile.Text <> '' then
         Directory.Text := INI.ReadString('IMPORT', 'Src' + IntToStr(index), '');
     INI.Free;
end;

end.
