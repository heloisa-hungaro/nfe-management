program Projeto;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, UMenu, urel, datetimectrls, uremove, uvisualiza;

{$R *.res}

begin
  Application.Title:='Controle NF-e';
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TFMenu, FMenu);
  Application.CreateForm(TFRel, FRel);
  Application.CreateForm(TFRemove, FRemove);
  Application.CreateForm(TFVisualiza, FVisualiza);
  Application.Run;
end.

