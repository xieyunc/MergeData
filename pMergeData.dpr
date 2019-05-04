program pMergeData;

uses
  Forms,
  uMain in 'source\uMain.pas' {Main},
  Unit_DBGridEhToExcel in 'source\Unit_DBGridEhToExcel.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMain, Main);
  Application.Run;
end.
