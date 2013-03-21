program MykolayProject;

uses
  Forms,
  ColumsViewModeUnit in 'ColumsViewModeUnit.pas' {ColumsViewModeForm},
  EditUnit in 'EditUnit.pas' {EditForm},
  MykolayUnit in 'MykolayUnit.pas' {MykolayMainForm},
  StatisticUnit in 'StatisticUnit.pas' {StatisticForm},
  UnitAbout in 'UnitAbout.pas' {FormAbout};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMykolayMainForm, MykolayMainForm);
  Application.CreateForm(TColumsViewModeForm, ColumsViewModeForm);
  Application.CreateForm(TEditForm, EditForm);
  Application.CreateForm(TStatisticForm, StatisticForm);
  Application.CreateForm(TFormAbout, FormAbout);
  Application.Run;
end.
