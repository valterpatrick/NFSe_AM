unit UNFSe_ArqMagSigISS;

interface

uses
  System.SysUtils, System.Variants, System.Classes, System.Generics.Collections,
  System.StrUtils, System.Math;

type
  TUNFSeSituacaoNotaSigISS = (snTributada, snRetida, snCancelada, snIsenta, snNaoTributada, snImune, snErro);

  TNotaExcelSigISS = class
  public
    A_CPFCNPJ: String;
    B_ValorServico: String;
    C_Email: String;
    D_CodigoServico: String;
    E_AliqSimplesNacional : String;
    F_SituacaoNotaFiscal: String;
    G_ValorRetencaoISS: String;
    H_ValorRetencaoINSS: String;
    I_ValorRetencaoCOFINS: String;
    J_ValorRetencaoPIS: String;
    K_ValorRetencaoIR: String;
    L_ValorRetencaoCSLL: String;
    M_ValorAproximadoTributos: String;
    N_IdentificadorSistema_Legado: String;
    O_InscricaoMunicipalTomador: String;
    P_InscricaoEstadualTomador: String;
    Q_NomeTomador: String;
    R_EnderecoTomador: String;
    S_NumeroEndereco: String;
    T_ComplementoEndereco: String;
    U_BairroTomador: String;
    V_CodCidadeTomador: String;
    W_UnidadeFederalTomador: String;
    X_CEPTomador: String;
    Y_CodCidPrestServ: String;
    Z_DesricaoServico: String;
  end;

  TNotaExcelListaSigISS = class(TList<TNotaExcelSigISS>)
  public
    function New: TNotaExcelSigISS;
    procedure Clear; reintroduce;
    procedure Remove(Valor: TNotaExcelSigISS); reintroduce;
    procedure Delete(Index: Integer); reintroduce;
    destructor Destroy; override;
  end;

  TNotaFiscalSigISS = class
  private
    FDetTipoRegistro: String;
    FDetIdentSistema: String;
    FDetTipoCodificacao: Integer;
    FDetCodServico: Integer;
    FDetSitNota: TUNFSeSituacaoNotaSigISS;
    FDetValorServico: Double;
    FDetValorBaseServico: Double;
    FDetAliqSimplesNacional: Double;
    FDetValorRetencaoISS: Double;
    FDetValorRetencaoINSS: Double;
    FDetValorRetencaoCOFINS: Double;
    FDetValorRetencaoPIS: Double;
    FDetValorRetencaoIR: Double;
    FDetValorRetencaoCSLL: Double;
    FDetValorAproximadoTributos: Double;
    FDetTomadorCPF_CNPJ: String;
    FDetTomadorInscMunicipal: String;
    FDetTomadorInscEstadual: String;
    FDetTomadorNome: String;
    FDetTomadorEndLogradouro: String;
    FDetTomadorEndNumero: String;
    FDetTomadorEndComplemento: String;
    FDetTomadorEndBairro: String;
    FDetTomadorEndCodCidade: Integer;
    FDetTomadorEndUF: String;
    FDetTomadorEndCEP: String;
    FDetTomadorEmail: String;
    FDetCodCidadePrestServ: Integer;
    FDetDescriminacaoServ: String;
  public
    property DetTipoRegistro: String read FDetTipoRegistro write FDetTipoRegistro;
    property DetIdentSistema: String read FDetIdentSistema write FDetIdentSistema;
    property DetTipoCodificacao: Integer read FDetTipoCodificacao write FDetTipoCodificacao;
    property DetCodServico: Integer read FDetCodServico write FDetCodServico;
    property DetSitNota: TUNFSeSituacaoNotaSigISS read FDetSitNota write FDetSitNota;
    property DetValorServico: Double read FDetValorServico write FDetValorServico;
    property DetValorBaseServico: Double read FDetValorBaseServico write FDetValorBaseServico;
    property DetAliqSimplesNacional: Double read FDetAliqSimplesNacional write FDetAliqSimplesNacional;
    property DetValorRetencaoISS: Double read FDetValorRetencaoISS write FDetValorRetencaoISS;
    property DetValorRetencaoINSS: Double read FDetValorRetencaoINSS write FDetValorRetencaoINSS;
    property DetValorRetencaoCOFINS: Double read FDetValorRetencaoCOFINS write FDetValorRetencaoCOFINS;
    property DetValorRetencaoPIS: Double read FDetValorRetencaoPIS write FDetValorRetencaoPIS;
    property DetValorRetencaoIR: Double read FDetValorRetencaoIR write FDetValorRetencaoIR;
    property DetValorRetencaoCSLL: Double read FDetValorRetencaoCSLL write FDetValorRetencaoCSLL;
    property DetValorAproximadoTributos: Double read FDetValorAproximadoTributos write FDetValorAproximadoTributos;
    property DetTomadorCPF_CNPJ: String read FDetTomadorCPF_CNPJ write FDetTomadorCPF_CNPJ;
    property DetTomadorInscMunicipal: String read FDetTomadorInscMunicipal write FDetTomadorInscMunicipal;
    property DetTomadorInscEstadual: String read FDetTomadorInscEstadual write FDetTomadorInscEstadual;
    property DetTomadorNome: String read FDetTomadorNome write FDetTomadorNome;
    property DetTomadorEndLogradouro: String read FDetTomadorEndLogradouro write FDetTomadorEndLogradouro;
    property DetTomadorEndNumero: String read FDetTomadorEndNumero write FDetTomadorEndNumero;
    property DetTomadorEndComplemento: String read FDetTomadorEndComplemento write FDetTomadorEndComplemento;
    property DetTomadorEndBairro: String read FDetTomadorEndBairro write FDetTomadorEndBairro;
    property DetTomadorEndCodCidade: Integer read FDetTomadorEndCodCidade write FDetTomadorEndCodCidade;
    property DetTomadorEndUF: String read FDetTomadorEndUF write FDetTomadorEndUF;
    property DetTomadorEndCEP: String read FDetTomadorEndCEP write FDetTomadorEndCEP;
    property DetTomadorEmail: String read FDetTomadorEmail write FDetTomadorEmail;
    property DetCodCidadePrestServ: Integer read FDetCodCidadePrestServ write FDetCodCidadePrestServ;
    property DetDescriminacaoServ: String read FDetDescriminacaoServ write FDetDescriminacaoServ;
  end;

  TNotaFiscalListaSigISS = class(TList<TNotaFiscalSigISS>)
  public
    function New: TNotaFiscalSigISS;
    class function SituacaoNotaToStr(Valor: TUNFSeSituacaoNotaSigISS): String;
    class function StrToSituacaoNota(Valor: String): TUNFSeSituacaoNotaSigISS;
    procedure Clear; reintroduce;
    procedure Remove(Valor: TNotaFiscalSigISS); reintroduce;
    procedure Delete(Index: Integer); reintroduce;
    destructor Destroy; override;
  end;

  TNFSe_ArqMagSigISS = class
  private
    FCabTipoRegistro: String;
    FCabTipoArquivo: String;
    FCabInscPrestador: String;
    FCabVersaoArquivo: String;
    FCabDataArquivo: TDate;
    FListaNotas: TNotaFiscalListaSigISS;
    FRodTipoRegistro: String;
    FRodNumLinhas: Integer;
    FRodValorTotalServ: Double;
    FRodValorTotalBaseServ: Double;
  public
    property CabTipoRegistro: String read FCabTipoRegistro write FCabTipoRegistro;
    property CabTipoArquivo: String read FCabTipoArquivo write FCabTipoArquivo;
    property CabInscPrestador: String read FCabInscPrestador write FCabInscPrestador;
    property CabVersaoArquivo: String read FCabVersaoArquivo write FCabVersaoArquivo;
    property CabDataArquivo: TDate read FCabDataArquivo write FCabDataArquivo;
    property ListaNotas: TNotaFiscalListaSigISS read FListaNotas write FListaNotas;
    property RodTipoRegistro: String read FRodTipoRegistro write FRodTipoRegistro;
    property RodNumLinhas: Integer read FRodNumLinhas;
    property RodValorTotalServ: Double read FRodValorTotalServ;
    property RodValorTotalBaseServ: Double read FRodValorTotalBaseServ;
    procedure CalcTotalizadores;
    procedure SaveToFile(FileName: String);
  end;

implementation

{ TNotaFiscalLista }

uses funcoes;

procedure TNotaFiscalListaSigISS.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited Clear;
end;

procedure TNotaFiscalListaSigISS.Delete(Index: Integer);
begin
  Items[Index].Free;
  inherited Delete(Index);
end;

destructor TNotaFiscalListaSigISS.Destroy;
var
  I: TNotaFiscalSigISS;
begin
  for I in Self do
    I.Free;
  inherited Destroy;
end;

function TNotaFiscalListaSigISS.New: TNotaFiscalSigISS;
begin
  Result := TNotaFiscalSigISS.Create;
  Add(Result);
end;

procedure TNotaFiscalListaSigISS.Remove(Valor: TNotaFiscalSigISS);
begin
  inherited Remove(Valor);
  Valor.Free;
end;

class function TNotaFiscalListaSigISS.SituacaoNotaToStr(Valor: TUNFSeSituacaoNotaSigISS): String;
begin
  case Valor of
    snTributada:
      Result := 'T';
    snRetida:
      Result := 'R';
    snCancelada:
      Result := 'C';
    snIsenta:
      Result := 'I';
    snNaoTributada:
      Result := 'N';
    snImune:
      Result := 'U';
    snErro:
      Result := '-';
  end;
end;

class function TNotaFiscalListaSigISS.StrToSituacaoNota(Valor: String): TUNFSeSituacaoNotaSigISS;
begin
  if Valor = 'T' then
    Result := snTributada
  else if Valor = 'R' then
    Result := snRetida
  else if Valor = 'C' then
    Result := snCancelada
  else if Valor = 'I' then
    Result := snIsenta
  else if Valor = 'N' then
    Result := snNaoTributada
  else if Valor = 'U' then
    Result := snImune
  else
    Result := snErro;
end;

{ TNFSe_ArqMag }

procedure TNFSe_ArqMagSigISS.CalcTotalizadores;
var
  vValorTotalServ, vValorTotalBaseServ: Double;
  I: Integer;
begin
  vValorTotalServ := 0;
  vValorTotalBaseServ := 0;
  for I := 0 to ListaNotas.Count - 1 do
  begin
    vValorTotalServ := vValorTotalServ + ListaNotas[I].FDetValorServico;
    vValorTotalBaseServ := vValorTotalBaseServ + ListaNotas[I].FDetValorBaseServico;
  end;

  Self.FRodNumLinhas := ListaNotas.Count;
  Self.FRodValorTotalServ := vValorTotalServ;
  Self.FRodValorTotalBaseServ := vValorTotalBaseServ;
end;

procedure TNFSe_ArqMagSigISS.SaveToFile(FileName: String);
var
  StrList: TStringList;
  N: TNotaFiscalSigISS;
begin
  // Layout Versão 3.0 - 01 de fevereiro de 2018
  StrList := TStringList.Create;
  try
    // Calcula os totalizadores do Rodapé
    CalcTotalizadores;
    // Cabeçalho
    StrList.Add( //
      Copy(FillChar(CabTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
      Copy(FillChar(CabTipoArquivo, 12, ' ', True), 1, 12) + // 002 a 013 - Tipo do Arquivo
      Copy(FillChar(CabInscPrestador, 15, '0', False), 1, 15) + // 014 a 028 - Inscrição do Prestador
      Copy(FillChar(CabVersaoArquivo, 3, '0', False), 1, 3) + // 029 a 031 - Versão do Arquivo
      Copy(FormatDateTime('YYYYMMDD', CabDataArquivo), 1, 8) // 032 a 039 - Data do Arquivo
      );

    // Detalhe
    for N in ListaNotas do
    begin
      StrList.Add( //
        Copy(FillChar(N.DetTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
        Copy(FillChar(N.DetIdentSistema, 12, ' ', True), 1, 12) + // 002 a 013 - Identificador Sistema Legado
        Copy(FillChar(IntToStr(N.DetTipoCodificacao), 1, '0', False), 1, 1) + // 014 a 014 - Tipo de Codificação
        Copy(FillChar(IntToStr(N.DetCodServico), 7, '0', False), 1, 7) + // 015 a 021 - Código do Serviço
        Copy(FillChar(ListaNotas.SituacaoNotaToStr(N.DetSitNota), 1, ' ', False), 1, 1) + // 022 a 022 - Situação da Nota Fiscal
        Copy(FillChar(FormatFloat('#', N.DetValorServico * 100), 15, '0', False), 1, 15) + // 023 a 037 - Valor dos Serviços
        Copy(FillChar(FormatFloat('#', N.DetValorBaseServico * 100), 15, '0', False), 1, 15) + // 038 a 052 - Valor da Base de Calculo
        Copy(FillChar(FormatFloat('###', N.DetAliqSimplesNacional * 100), 3, '0', False), 1, 3) + // 053 a 055- - Aliquota Simples Nacional
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoISS * 100), 15, '0', False), 1, 15) + // 056 a 070 - Valor Retenção ISS
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoINSS * 100), 15, '0', False), 1, 15) + // 071 a 085 - Valor Retenção INSS
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoCOFINS * 100), 15, '0', False), 1, 15) + // 086 a 100 - Valor Retenção COFINS
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoPIS * 100), 15, '0', False), 1, 15) + // 101 a 115 - Valor Retenção PIS
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoIR * 100), 15, '0', False), 1, 15) + // 116 a 130 - Valor Retenção IR
        Copy(FillChar(FormatFloat('#', N.DetValorRetencaoCSLL * 100), 15, '0', False), 1, 15) + // 131 a 145 - Valor Retenção CSLL
        Copy(FillChar(FormatFloat('#', N.DetValorAproximadoTributos * 100), 15, '0', False), 1, 15) + // 146 a 160 - Valor Aproximado Tributos
        Copy(IfThen(UpperCase(N.DetTomadorCPF_CNPJ.Trim) = 'PFNI', FillChar('PFNI', 15, ' ', True), FillChar(Ifthen(N.DetTomadorCPF_CNPJ.Length <= 11, FillChar(N.DetTomadorCPF_CNPJ, 11, '0', False), FillChar(N.DetTomadorCPF_CNPJ, 14, '0', False)), 15, ' ', True)), 1, 15) + // 161 a 175 - CPF/CNPJ do tomador
        Copy(FillChar(N.DetTomadorInscMunicipal, 15, '0', False), 1, 15) + // 176 a 190 - Incrição Municipal do Tomador
        Copy(FillChar(N.DetTomadorInscEstadual, 15, '0', False), 1, 15) + // 191 a 205 - Incrição Estadual do Tomador
        Copy(FillChar(N.DetTomadorNome, 100, ' ', True), 1, 100) + // 206 a 305 - Nome/Razão Social do Tomador
        Copy(FillChar(N.DetTomadorEndLogradouro, 50, ' ', True), 1, 50) + // 306 a 355 - Endereço do Tomador
        Copy(FillChar(N.DetTomadorEndNumero, 10, ' ', True), 1, 10) + // 356 a 365 - Número do Tomador
        Copy(FillChar(N.DetTomadorEndComplemento, 30, ' ', True), 1, 30) + // 366 a 395 - Complemento do Tomador
        Copy(FillChar(N.DetTomadorEndBairro, 30, ' ', True), 1, 30) + // 396 a 425 - Bairro do Tomador
        Copy(FillChar(IntToStr(N.DetTomadorEndCodCidade), 7, '0', False), 1, 7) + // 426 a 432 - Código da Cidade do Tomador
        Copy(FillChar(N.DetTomadorEndUF, 2, ' ', True), 1, 2) + // 433 a 434 - UF do Tomador
        Copy(FillChar(N.DetTomadorEndCEP, 8, '0', False), 1, 8) + // 435 a 442 - CEP do Tomador
        Copy(FillChar(N.DetTomadorEmail, 100, ' ', True), 1, 100) + // 443 a 542 - Email do Tomador
        Copy(FillChar(IntToStr(N.DetCodCidadePrestServ), 7, '0', False), 1, 7) + // 543 a 549 - Código da Cidade do local de prestação de serviço
        Copy(N.DetDescriminacaoServ.Trim, 1, IfThen(N.DetDescriminacaoServ.Trim.Length < 1000, N.DetDescriminacaoServ.Trim.Length, 1000))
        // 550 a 550 (N - 1) - Discriminação do Serviço [Não foi utilizado o FillChar, pois será gravado menos dados no arquivo]
        );
    end;

    // Rodapé
    StrList.Add( //
      Copy(FillChar(RodTipoRegistro, 1, '0', True), 1, 1) + // 001 a 001 - Tipo de Registro
      Copy(FillChar(IntToStr(RodNumLinhas), 10, '0', False), 1, 10) + // 002 a 011 - Número de linhas detalhes
      Copy(FillChar(FormatFloat('#', RodValorTotalServ * 100), 15, '0', False), 1, 15) + // 012 a 026 - Valor Total dos Serviços
      Copy(FillChar(FormatFloat('#', RodValorTotalBaseServ * 100), 15, '0', False), 1, 15) // 027 a 041 - Valor Total da Base
      );
    StrList.SaveToFile(FileName);
  finally
    StrList.Free;
  end;
end;

{ TNotaExcelLista }

procedure TNotaExcelListaSigISS.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited Clear;
end;

procedure TNotaExcelListaSigISS.Delete(Index: Integer);
begin
  Items[Index].Free;
  inherited Delete(Index);
end;

destructor TNotaExcelListaSigISS.Destroy;
var
  I: TNotaExcelSigISS;
begin
  for I in Self do
    I.Free;
  inherited Destroy;
end;

function TNotaExcelListaSigISS.New: TNotaExcelSigISS;
begin
  Result := TNotaExcelSigISS.Create;
  Add(Result);
end;

procedure TNotaExcelListaSigISS.Remove(Valor: TNotaExcelSigISS);
begin
  inherited Remove(Valor);
  Valor.Free;
end;

end.
