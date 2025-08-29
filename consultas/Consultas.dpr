program Consultas;





uses
  Forms,
  UFrmConsultaWinthor in 'UFrmConsultaWinthor.pas' {FrmConsultaWinthor},
  UConsultasWinthor in 'UConsultasWinthor.pas',
  UMensagens in '..\LIBS\UMensagens.pas',
  ULibrary in '..\LIBSODAC\ULibrary.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFrmConsultaWinthor, FrmConsultaWinthor);
  Application.Run;
end.

