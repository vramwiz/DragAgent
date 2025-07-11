program DragAgentSample;

uses
  Vcl.Forms,
  MainForm in 'MainForm.pas' {FormMain},
  DragAgent in 'DragAgent.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
