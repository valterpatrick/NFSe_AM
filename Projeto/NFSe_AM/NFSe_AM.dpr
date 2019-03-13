program NFSe_AM;





{$R *.dres}

uses
  Vcl.Forms,
  UFrmPrincipal in 'UFrmPrincipal.pas' {FrmPrincipal},
  UNFSe_ArqMagSigISS in 'UNFSe_ArqMagSigISS.pas',
  funcoes in '..\Shared\funcoes.pas',
  UFrmAdvertencias in 'UFrmAdvertencias.pas' {FrmAdvertencias},
  UFrmSobre in 'UFrmSobre.pas' {FrmSobre},
  UNFSe_ArqMagEL in 'UNFSe_ArqMagEL.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmPrincipal, FrmPrincipal);
  Application.CreateForm(TFrmAdvertencias, FrmAdvertencias);
  Application.CreateForm(TFrmSobre, FrmSobre);
  Application.Run;
end.
