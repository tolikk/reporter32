program INIEditorPrj;

uses
  Forms,
  INIEditor in 'INIEditor.pas' {MainForm};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'Редактор blank.ini';
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
