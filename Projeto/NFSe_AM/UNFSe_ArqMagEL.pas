unit UNFSe_ArqMagEL;

interface

uses
  System.SysUtils, System.Variants, System.Classes, System.Generics.Collections,
  System.StrUtils, System.Math, Funcoes;

type
  TUNFSeSituacaoNotaEL = (snOperacaoNormal, snIsenta_NaoTributadaMuni, snIsenta_NaoTributadaOutMuni, snCancelada, snExtraviada, snISS_SuspensoDecisaoJudicial, snErro);
  TUNFSeTipoRecolhimentoEL = (trNormal, trRetidoFonte, trSemRecolhimento, trErro);
  TUNFSeLocalPrestacaoEL = (lpForaMunicipio, lpMunicipio, lpErro);
  TUNFSeIndicadorTomadorEL = (itCPF, itCNPJ, itNaoIdentificado, itEstrangeiro, itErro);
  TUNFSeIndicadorPrestadorEL = (ipCPF, ipCNPJ, ipErro);
  TUNFSeOptanteSimples = (osNaoOptanteSimplesFederal_Municipal, osOptanteSimplesFederal_UmPorCento, osOptanteSimplesFederal_MeioPorCento, osOptanteSimplesMunicipal, osOptanteSimplesNacional, osErro);
  TUNFSeCasasDecimais = (cdDois, cdTres, cdQuatro, cdErro);

  TNotaExcelEL = class
  public
  end;

  TNotaExcelListaEL = class(TList<TNotaExcelEL>)
  public
    function New: TNotaExcelEL;
    procedure Clear; reintroduce;
    procedure Remove(Valor: TNotaExcelEL); reintroduce;
    procedure Delete(Index: Integer); reintroduce;
    destructor Destroy; override;
  end;

  TNotaFiscalEL = class
  private
    FCabNotTipoRegistro: String;
    FCabNotSequencialNota: String;
    FCabNotDataHora: TDateTime;
    FCabNotTipoRecolhimento: TUNFSeTipoRecolhimentoEL;
    FCabNotSitNota: TUNFSeSituacaoNotaEL;
    FCabNotDataCancelamento: TDate;
    FCabNotCodMuniPrestServ: Integer;
    FCabNotValorTotServ: Double;
    FCabNotValorTotDeducoes: Double;
    FCabNotValorRetencaoPIS: Double;
    FCabNotValorRetencaoCOFINS: Double;
    FCabNotValorRetencaoINSS: Double;
    FCabNotValorRetencaoIR: Double;
    FCabNotValorRetencaoCSLL: Double;
    FCabNotValorISSQN: Double;
    FCabNotLocalPrestacao: TUNFSeLocalPrestacaoEL;
    FCabNotSeqNotaSerSubstituida: String;
    FCabNotOutDescontos: Double;
    FCabNotSequencialReg: Integer;

    FTomTipoRegistro: String;
    FTomSequencialNota: String;
    FTomIndCPF_CNPJ: TUNFSeIndicadorTomadorEL;
    FTomCPF_CNPJ: String;
    FTomNome: String;
    FTomNomeFantasia: String;
    FTomEndTipo: String;
    FTomEndLogradouro: String;
    FTomEndNumero: String;
    FTomEndComplemento: String;
    FTomEndBairro: String;
    FTomEndCidade: String;
    FTomEndUF: String;
    FTomEndCEP: String;
    FTomEmail: String;
    FTomInscEstadual: String;
    FTomSequencialReg: Integer;

    FSrvTipoRegistro: String;
    FSrvSequencialNota: String;
    FSrvCodServPrestado: String;
    FSrvCodTribMuni: String;
    FSrvValorServ: Double;
    FSrvValorDeducao: Double;
    FSrvAliq: Double;
    FSrvUnidade: String;
    FSrvQuantidade: Double;
    FSrvDescricaoServ: String;
    FSrvAlvara: String;
    FSrvSequencialReg: Integer;
  public
    property CabNotTipoRegistro: String read FCabNotTipoRegistro write FCabNotTipoRegistro;
    property CabNotSequencialNota: String read FCabNotSequencialNota write FCabNotSequencialNota;
    property CabNotDataHora: TDateTime read FCabNotDataHora write FCabNotDataHora;
    property CabNotTipoRecolhimento: TUNFSeTipoRecolhimentoEL read FCabNotTipoRecolhimento write FCabNotTipoRecolhimento;
    property CabNotSitNota: TUNFSeSituacaoNotaEL read FCabNotSitNota write FCabNotSitNota;
    property CabNotDataCancelamento: TDate read FCabNotDataCancelamento write FCabNotDataCancelamento;
    property CabNotCodMuniPrestServ: Integer read FCabNotCodMuniPrestServ write FCabNotCodMuniPrestServ;
    property CabNotValorTotServ: Double read FCabNotValorTotServ write FCabNotValorTotServ;
    property CabNotValorTotDeducoes: Double read FCabNotValorTotDeducoes write FCabNotValorTotDeducoes;
    property CabNotValorRetencaoPIS: Double read FCabNotValorRetencaoPIS write FCabNotValorRetencaoPIS;
    property CabNotValorRetencaoCOFINS: Double read FCabNotValorRetencaoCOFINS write FCabNotValorRetencaoCOFINS;
    property CabNotValorRetencaoINSS: Double read FCabNotValorRetencaoINSS write FCabNotValorRetencaoINSS;
    property CabNotValorRetencaoCSLL: Double read FCabNotValorRetencaoCSLL write FCabNotValorRetencaoCSLL;
    property CabNotValorRetencaoIR: Double read FCabNotValorRetencaoIR write FCabNotValorRetencaoIR;
    property CabNotValorISSQN: Double read FCabNotValorISSQN write FCabNotValorISSQN;
    property CabNotLocalPrestacao: TUNFSeLocalPrestacaoEL read FCabNotLocalPrestacao write FCabNotLocalPrestacao;
    property CabNotSeqNotaSerSubstituida: String read FCabNotSeqNotaSerSubstituida write FCabNotSeqNotaSerSubstituida;
    property CabNotOutDescontos: Double read FCabNotOutDescontos write FCabNotOutDescontos;
    property CabNotSequencialReg: Integer read FCabNotSequencialReg write FCabNotSequencialReg;

    property TomTipoRegistro: String read FTomTipoRegistro write FTomTipoRegistro;
    property TomSequencialNota: String read FTomSequencialNota write FTomSequencialNota;
    property TomIndCPF_CNPJ: TUNFSeIndicadorTomadorEL read FTomIndCPF_CNPJ write FTomIndCPF_CNPJ;
    property TomCPF_CNPJ: String read FTomCPF_CNPJ write FTomCPF_CNPJ;
    property TomNome: String read FTomNome write FTomNome;
    property TomNomeFantasia: String read FTomNomeFantasia write FTomNomeFantasia;
    property TomEndTipo: String read FTomEndTipo write FTomEndTipo;
    property TomEndLogradouro: String read FTomEndLogradouro write FTomEndLogradouro;
    property TomEndNumero: String read FTomEndNumero write FTomEndNumero;
    property TomEndComplemento: String read FTomEndComplemento write FTomEndComplemento;
    property TomEndBairro: String read FTomEndBairro write FTomEndBairro;
    property TomEndCidade: String read FTomEndCidade write FTomEndCidade;
    property TomEndUF: String read FTomEndUF write FTomEndUF;
    property TomEndCEP: String read FTomEndCEP write FTomEndCEP;
    property TomEmail: String read FTomEmail write FTomEmail;
    property TomInscEstadual: String read FTomInscEstadual write FTomInscEstadual;
    property TomSequencialReg: Integer read FTomSequencialReg write FTomSequencialReg;

    property SrvTipoRegistro: String read FSrvTipoRegistro write FSrvTipoRegistro;
    property SrvSequencialNota: String read FSrvSequencialNota write FSrvSequencialNota;
    property SrvCodServPrestado: String read FSrvCodServPrestado write FSrvCodServPrestado;
    property SrvCodTribMuni: String read FSrvCodTribMuni write FSrvCodTribMuni;
    property SrvValorServ: Double read FSrvValorServ write FSrvValorServ;
    property SrvValorDeducao: Double read FSrvValorDeducao write FSrvValorDeducao;
    property SrvAliq: Double read FSrvAliq write FSrvAliq;
    property SrvUnidade: String read FSrvUnidade write FSrvUnidade;
    property SrvQuantidade: Double read FSrvQuantidade write FSrvQuantidade;
    property SrvDescricaoServ: String read FSrvDescricaoServ write FSrvDescricaoServ;
    property SrvAlvara: String read FSrvAlvara write FSrvAlvara;
    property SrvSequencialReg: Integer read FSrvSequencialReg write FSrvSequencialReg;
  end;

  TNotaFiscalListaEL = class(TList<TNotaFiscalEL>)
  public
    function New: TNotaFiscalEL;
    class function SituacaoNotaToStr(Valor: TUNFSeSituacaoNotaEL): String;
    class function StrToSituacaoNota(Valor: String): TUNFSeSituacaoNotaEL;
    class function TipoRecolhimentoToStr(Valor: TUNFSeTipoRecolhimentoEL): String;
    class function StrToTipoRecolhimento(Valor: String): TUNFSeTipoRecolhimentoEL;
    class function LocalPrestacaoToStr(Valor: TUNFSeLocalPrestacaoEL): String;
    class function StrToLocalPrestacao(Valor: String): TUNFSeLocalPrestacaoEL;
    class function IndicadorTomadorToStr(Valor: TUNFSeIndicadorTomadorEL): String;
    class function StrToIndicadorTomador(Valor: String): TUNFSeIndicadorTomadorEL;
    procedure Clear; reintroduce;
    procedure Remove(Valor: TNotaFiscalEL); reintroduce;
    procedure Delete(Index: Integer); reintroduce;
    destructor Destroy; override;
  end;

  TNFSe_ArqMagEL = class
  private
    FCabArqTipoRegistro: String;
    FCabArqVersaoLayout: Integer;
    FCabArqInscMunicipalPrestador: String;
    FCabArqPrestadorIndCPF_CNPJ: TUNFSeIndicadorPrestadorEL;
    FCabArqPrestadorCPF_CNPJ: String;
    FCabArqOptanteSimples: TUNFSeOptanteSimples;
    FCabArqDataIniPeriodo: TDate;
    FCabArqDataFimPeriodo: TDate;
    FCabArqQuantNFSeInformadas: Integer;
    FCabArqCasasDec_ValorServ: TUNFSeCasasDecimais;
    FCabArqCasasDec_QuantServ: TUNFSeCasasDecimais;
    FCabArqSequencialReg: Integer;
    FListaNotas: TNotaFiscalListaEL;
    FRodArqTipoRegistro: String;
    FRodArqSequencialReg: Integer;
    function FormataDoubleCD(Valor: Double; CasasDecimais: TUNFSeCasasDecimais): String;
  public
    property CabArqTipoRegistro: String read FCabArqTipoRegistro write FCabArqTipoRegistro;
    property CabArqVersaoLayout: Integer read FCabArqVersaoLayout write FCabArqVersaoLayout;
    property CabArqInscMunicipalPrestador: String read FCabArqInscMunicipalPrestador write FCabArqInscMunicipalPrestador;
    property CabArqPrestadorIndCPF_CNPJ: TUNFSeIndicadorPrestadorEL read FCabArqPrestadorIndCPF_CNPJ write FCabArqPrestadorIndCPF_CNPJ;
    property CabArqPrestadorCPF_CNPJ: String read FCabArqPrestadorCPF_CNPJ write FCabArqPrestadorCPF_CNPJ;
    property CabArqOptanteSimples: TUNFSeOptanteSimples read FCabArqOptanteSimples write FCabArqOptanteSimples;
    property CabArqDataIniPeriodo: TDate read FCabArqDataIniPeriodo write FCabArqDataIniPeriodo;
    property CabArqDataFimPeriodo: TDate read FCabArqDataFimPeriodo write FCabArqDataFimPeriodo;
    property CabArqQuantNFSeInformadas: Integer read FCabArqQuantNFSeInformadas write FCabArqQuantNFSeInformadas;
    property CabArqCasasDec_ValorServ: TUNFSeCasasDecimais read FCabArqCasasDec_ValorServ write FCabArqCasasDec_ValorServ;
    property CabArqCasasDec_QuantServ: TUNFSeCasasDecimais read FCabArqCasasDec_QuantServ write FCabArqCasasDec_QuantServ;
    property CabArqSequencialReg: Integer read FCabArqSequencialReg write FCabArqSequencialReg;
    property ListaNotas: TNotaFiscalListaEL read FListaNotas write FListaNotas;
    property RodArqTipoRegistro: String read FRodArqTipoRegistro write FRodArqTipoRegistro;
    property RodArqSequencialReg: Integer read FRodArqSequencialReg write FRodArqSequencialReg;
    class function IndicadorPrestadorToStr(Valor: TUNFSeIndicadorPrestadorEL): String;
    class function StrToIndicadorPrestador(Valor: String): TUNFSeIndicadorPrestadorEL;
    class function OptanteSimplesToStr(Valor: TUNFSeOptanteSimples): String;
    class function StrToOptanteSimples(Valor: String): TUNFSeOptanteSimples;
    class function CasasDecimaisToStr(Valor: TUNFSeCasasDecimais): String;
    class function StrToCasasDecimais(Valor: String): TUNFSeCasasDecimais;
    procedure SaveToFile(FileName: String);
  end;

implementation

{ TNotaExcelLista }

procedure TNotaExcelListaEL.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited Clear;
end;

procedure TNotaExcelListaEL.Delete(Index: Integer);
begin
  Items[Index].Free;
  inherited Delete(Index);
end;

destructor TNotaExcelListaEL.Destroy;
var
  I: TNotaExcelEL;
begin
  for I in Self do
    I.Free;
  inherited Destroy;
end;

function TNotaExcelListaEL.New: TNotaExcelEL;
begin
  Result := TNotaExcelEL.Create;
  Add(Result);
end;

procedure TNotaExcelListaEL.Remove(Valor: TNotaExcelEL);
begin
  inherited Remove(Valor);
  Valor.Free;
end;

{ TNotaFiscalLista }

procedure TNotaFiscalListaEL.Clear;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Free;
  inherited Clear;
end;

procedure TNotaFiscalListaEL.Delete(Index: Integer);
begin
  Items[Index].Free;
  inherited Delete(Index);
end;

destructor TNotaFiscalListaEL.Destroy;
var
  I: TNotaFiscalEL;
begin
  for I in Self do
    I.Free;
  inherited Destroy;
end;

class function TNotaFiscalListaEL.IndicadorTomadorToStr(Valor: TUNFSeIndicadorTomadorEL): String;
begin
  case Valor of
    itCPF:
      Result := '1';
    itCNPJ:
      Result := '2';
    itNaoIdentificado:
      Result := '3';
    itEstrangeiro:
      Result := '9';
    itErro:
      Result := '-';
  end;
end;

class function TNotaFiscalListaEL.StrToIndicadorTomador(Valor: String): TUNFSeIndicadorTomadorEL;
begin
  if Valor = '1' then
    Result := itCPF
  else if Valor = '2' then
    Result := itCNPJ
  else if Valor = '3' then
    Result := itNaoIdentificado
  else if Valor = '9' then
    Result := itEstrangeiro
  else
    Result := itErro;
end;

function TNotaFiscalListaEL.New: TNotaFiscalEL;
begin
  Result := TNotaFiscalEL.Create;
  Add(Result);
end;

procedure TNotaFiscalListaEL.Remove(Valor: TNotaFiscalEL);
begin
  inherited Remove(Valor);
  Valor.Free;
end;

class function TNotaFiscalListaEL.SituacaoNotaToStr(Valor: TUNFSeSituacaoNotaEL): String;
begin
  case Valor of
    snOperacaoNormal:
      Result := 'T';
    snIsenta_NaoTributadaMuni:
      Result := 'I';
    snIsenta_NaoTributadaOutMuni:
      Result := 'F';
    snCancelada:
      Result := 'C';
    snExtraviada:
      Result := 'E';
    snISS_SuspensoDecisaoJudicial:
      Result := 'J';
    snErro:
      Result := '-';
  end;
end;

class function TNotaFiscalListaEL.StrToLocalPrestacao(Valor: String): TUNFSeLocalPrestacaoEL;
begin
  if Valor = 'F' then
    Result := lpForaMunicipio
  else if Valor = 'M' then
    Result := lpMunicipio
  else
    Result := lpErro;
end;

class function TNotaFiscalListaEL.LocalPrestacaoToStr(Valor: TUNFSeLocalPrestacaoEL): String;
begin
  case Valor of
    lpForaMunicipio:
      Result := 'F';
    lpMunicipio:
      Result := 'M';
    lpErro:
      Result := '-';
  end;
end;

class function TNotaFiscalListaEL.StrToSituacaoNota(Valor: String): TUNFSeSituacaoNotaEL;
begin
  if Valor = 'T' then
    Result := snOperacaoNormal
  else if Valor = 'I' then
    Result := snIsenta_NaoTributadaMuni
  else if Valor = 'F' then
    Result := snIsenta_NaoTributadaOutMuni
  else if Valor = 'C' then
    Result := snCancelada
  else if Valor = 'E' then
    Result := snExtraviada
  else if Valor = 'J' then
    Result := snISS_SuspensoDecisaoJudicial
  else
    Result := snErro;
end;

class function TNotaFiscalListaEL.StrToTipoRecolhimento(Valor: String): TUNFSeTipoRecolhimentoEL;
begin
  if Valor = 'N' then
    Result := trNormal
  else if Valor = 'R' then
    Result := trRetidoFonte
  else if Valor = 'S' then
    Result := trSemRecolhimento
  else
    Result := trErro;
end;

class function TNotaFiscalListaEL.TipoRecolhimentoToStr(Valor: TUNFSeTipoRecolhimentoEL): String;
begin
  case Valor of
    trNormal:
      Result := 'N';
    trRetidoFonte:
      Result := 'R';
    trSemRecolhimento:
      Result := 'S';
    trErro:
      Result := '-';
  end;
end;

{ TNFSe_ArqMagEL }

class function TNFSe_ArqMagEL.CasasDecimaisToStr(Valor: TUNFSeCasasDecimais): String;
begin
  case Valor of
    cdDois:
      Result := '2';
    cdTres:
      Result := '3';
    cdQuatro:
      Result := '4';
    cdErro:
      Result := '-';
  end;
end;

class function TNFSe_ArqMagEL.StrToCasasDecimais(Valor: String): TUNFSeCasasDecimais;
begin
  if Valor = '2' then
    Result := cdDois
  else if Valor = '3' then
    Result := cdTres
  else if Valor = '4' then
    Result := cdQuatro
  else
    Result := cdErro;
end;

function TNFSe_ArqMagEL.FormataDoubleCD(Valor: Double; CasasDecimais: TUNFSeCasasDecimais): String;
begin
  case CasasDecimais of
    cdDois:
      Result := FormatFloat('#', Valor * 100);
    cdTres:
      Result := FormatFloat('#', Valor * 1000);
    cdQuatro:
      Result := FormatFloat('#', Valor * 10000);
    cdErro:
      Result := FormatFloat('#', Valor * 100); // O padrão é duas casas decimais
  end;
end;

class function TNFSe_ArqMagEL.IndicadorPrestadorToStr(Valor: TUNFSeIndicadorPrestadorEL): String;
begin
  case Valor of
    ipCPF:
      Result := '1';
    ipCNPJ:
      Result := '2';
    ipErro:
      Result := '-';
  end;
end;

class function TNFSe_ArqMagEL.StrToIndicadorPrestador(Valor: String): TUNFSeIndicadorPrestadorEL;
begin
  if Valor = '1' then
    Result := ipCPF
  else if Valor = '2' then
    Result := ipCNPJ
  else
    Result := ipErro;
end;

class function TNFSe_ArqMagEL.OptanteSimplesToStr(Valor: TUNFSeOptanteSimples): String;
begin
  case Valor of
    osNaoOptanteSimplesFederal_Municipal:
      Result := '0';
    osOptanteSimplesFederal_UmPorCento:
      Result := '1';
    osOptanteSimplesFederal_MeioPorCento:
      Result := '2';
    osOptanteSimplesMunicipal:
      Result := '3';
    osOptanteSimplesNacional:
      Result := '4';
    osErro:
      Result := '-';
  end;
end;

class function TNFSe_ArqMagEL.StrToOptanteSimples(Valor: String): TUNFSeOptanteSimples;
begin
  if Valor = '0' then
    Result := osNaoOptanteSimplesFederal_Municipal
  else if Valor = '1' then
    Result := osOptanteSimplesFederal_UmPorCento
  else if Valor = '2' then
    Result := osOptanteSimplesFederal_MeioPorCento
  else if Valor = '3' then
    Result := osOptanteSimplesMunicipal
  else if Valor = '4' then
    Result := osOptanteSimplesNacional
  else
    Result := osErro;
end;

procedure TNFSe_ArqMagEL.SaveToFile(FileName: String);
var
  StrList: TStringList;
  N: TNotaFiscalEL;
begin
  // Layout Versão 108
  StrList := TStringList.Create;
  try
    // Cabeçalho
    StrList.Add( //
      Copy(FillChar(CabArqTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
      Copy(FillChar(IntToStr(CabArqVersaoLayout), 3, '0', False), 1, 3) + // 002 a 004 - Versão do Layout
      Copy(FillChar(CabArqInscMunicipalPrestador, 26, ' ', True), 1, 26) + // 005 a 030 - Inscrição Municipal do Prestador
      Copy(FillChar(IndicadorPrestadorToStr(CabArqPrestadorIndCPF_CNPJ), 1, '0', False), 1, 1) + // 031 a 031 - Indicador de CPF/CNPJ do Prestador
      Copy(FillChar(CabArqPrestadorCPF_CNPJ, 14, ' ', True), 1, 14) + // 032 a 045 - CNPJ/CPF do Prestador
      Copy(FillChar(OptanteSimplesToStr(CabArqOptanteSimples), 1, ' ', False), 1, 1) + // 046 a 046 - Optante pelo Simples
      Copy(FormatDateTime('YYYYMMDD', CabArqDataIniPeriodo), 1, 8) + // 047 a 054 - Data de Início do Período
      Copy(FormatDateTime('YYYYMMDD', CabArqDataFimPeriodo), 1, 8) + // 055 a 062 - Data de Fim do Período
      Copy(FillChar(IntToStr(ListaNotas.Count), 5, '0', False), 1, 5) + // 063 a 067 - Quantidade de NFS-e informadas
      Copy(FillChar(CasasDecimaisToStr(CabArqCasasDec_ValorServ), 1, '0', False), 1, 1) + // 068 a 068 - Quantidade de Casas Decimais para o Valor de Serviço
      Copy(FillChar(CasasDecimaisToStr(CabArqCasasDec_QuantServ), 1, '0', False), 1, 1) + // 069 a 069 - Quantidade de Casas Decimais para a Quantidade de um Serviço
      FillChar('', 322, ' ', True) + // 070 a 391 - Preencher com Brancos
      Copy(FillChar(IntToStr(CabArqSequencialReg), 8, '0', False), 1, 8) // 392 a 400 - Sequencial do registro
      );

    // Notas Fiscais
    for N in ListaNotas do
    begin
      // Cabeçalho da NFS-e
      StrList.Add( //
        Copy(FillChar(N.CabNotTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
        Copy(FillChar(N.CabNotSequencialNota, 20, '0', False), 1, 20) + // 002 a 021 - Sequencial da NFS-e
        Copy(FormatDateTime('YYYYMMDDHHNNSS', N.CabNotDataHora), 1, 14) + // 022 a 035 - Data e Hora da NFS-e
        Copy(FillChar(ListaNotas.TipoRecolhimentoToStr(N.CabNotTipoRecolhimento), 1, ' ', True), 1, 1) + // 036 a 036 - Tipo de Recolhimento
        Copy(FillChar(ListaNotas.SituacaoNotaToStr(N.CabNotSitNota), 1, ' ', True), 1, 1) + // 037 a 037 - Situação da Nota Fiscal
        Copy(IfThen(N.CabNotSitNota = snCancelada, FormatDateTime('YYYYMMDD', N.CabNotDataCancelamento), FillChar('', 8, ' ', True)), 1, 8) + // 038 a 045 - Data de Cancelamento
        Copy(FillChar(IntToStr(N.CabNotCodMuniPrestServ), 7, '0', False), 1, 7) + // 046 a 052 - Município de prestação do serviço
        Copy(FillChar(IfThen(N.CabNotSitNota in [snCancelada, snExtraviada], '0', FormataDoubleCD(N.CabNotValorTotServ, CabArqCasasDec_ValorServ)), 15, '0', False), 1, 15) + // 053 a 067 - Valor Total dos Serviços
        Copy(FillChar(FormataDoubleCD(N.CabNotValorTotDeducoes, cdDois), 15, '0', False), 1, 15) + // 068 a 082 - Valor Total das Deduções
        Copy(FillChar(FormataDoubleCD(N.CabNotValorRetencaoPIS, cdDois), 15, '0', False), 1, 15) + // 083 a 097 - Valor da retenção do PIS
        Copy(FillChar(FormataDoubleCD(N.CabNotValorRetencaoCOFINS, cdDois), 15, '0', False), 1, 15) + // 098 a 112 - Valor da retenção do COFINS
        Copy(FillChar(FormataDoubleCD(N.CabNotValorRetencaoINSS, cdDois), 15, '0', False), 1, 15) + // 113 a 127 - Valor da retenção do INSS
        Copy(FillChar(FormataDoubleCD(N.CabNotValorRetencaoIR, cdDois), 15, '0', False), 1, 15) + // 128 a 142 - Valor da retenção do IR
        Copy(FillChar(FormataDoubleCD(N.CabNotValorRetencaoCSLL, cdDois), 15, '0', False), 1, 15) + // 143 a 157 - Valor da retenção do CSLL
        Copy(FillChar(FormataDoubleCD(N.CabNotValorISSQN, cdDois), 15, '0', False), 1, 15) + // 158 a 172 - Valor do ISSQN
        Copy(FillChar(ListaNotas.LocalPrestacaoToStr(N.FCabNotLocalPrestacao), 1, ' ', False), 1, 1) + // 173 a 173 - Local da Prestação
        Copy(FillChar(N.CabNotSeqNotaSerSubstituida, 20, '0', False), 1, 20) + // 174 a 193 - Sequencial da NFS-e à ser Substituída
        Copy(FillChar(FormataDoubleCD(N.CabNotOutDescontos, cdDois), 15, '0', False), 1, 15) + // 194 a 208 - Outros Descontos
        FillChar('', 183, ' ', True) + // 209 a 391 - Preencher com Brancos
        Copy(FillChar(IntToStr(N.CabNotSequencialReg), 8, '0', False), 1, 8) // 392 a 400 - Sequencial do registro
        );

      // Identificação do Tomador da NFS-e
      StrList.Add( //
        Copy(FillChar(N.TomTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
        Copy(FillChar(N.TomSequencialNota, 20, '0', False), 1, 20) + // 002 a 021 - Sequencial da NFS-e
        Copy(FillChar(ListaNotas.IndicadorTomadorToStr(N.TomIndCPF_CNPJ), 1, ' ', False), 1, 1) + // 022 a 022 - Indicador de CPF/CNPJ do Tomador
        Copy(FillChar(N.TomCPF_CNPJ, 14, ' ', True), 1, 14) + // 023 a 036 - CNPJ, CPF do Tomador/(Documento)
        Copy(FillChar(N.TomNome, 50, ' ', True), 1, 50) + // 037 a 086 - Nome do Tomador
        Copy(FillChar(IfThen(N.TomIndCPF_CNPJ = itCNPJ, N.TomNome, ''), 50, ' ', True), 1, 50) + // 087 a 136 - Nome Fantasia
        Copy(FillChar(N.TomEndTipo, 3, ' ', True), 1, 3) + // 137 a 139 - Tipo de Endereço do Tomador
        Copy(FillChar(N.TomEndLogradouro, 50, ' ', True), 1, 50) + // 140 a 189 - Endereço do Tomador
        Copy(FillChar(N.TomEndNumero, 10, ' ', True), 1, 10) + // 190 a 199 - Número do Endereço do Tomador
        Copy(FillChar(N.TomEndComplemento, 20, ' ', True), 1, 20) + // 200 a 219 - Complemento do Endereço do Tomador
        Copy(FillChar(N.TomEndBairro, 30, ' ', True), 1, 30) + // 220 a 249 - Bairro do Tomador
        Copy(FillChar(N.TomEndCidade, 50, ' ', True), 1, 50) + // 250 a 299 - Cidade do Tomador
        Copy(FillChar(N.TomEndUF, 2, ' ', True), 1, 2) + // 300 a 301 - UF do Tomador
        Copy(FillChar(N.TomEndCEP, 8, ' ', True), 1, 8) + // 302 a 309 - CEP do Tomador
        Copy(FillChar(N.TomEmail, 60, ' ', True), 1, 60) + // 310 a 369 - Email do Tomador
        Copy(FillChar(N.TomInscEstadual, 20, ' ', True), 1, 20) + // 370 a 389 - Inscrição Estadual Tomador
        FillChar('', 2, ' ', True) + // 390 a 391 - Preencher com Brancos
        Copy(FillChar(IntToStr(N.TomSequencialReg), 8, '0', False), 1, 8) // 392 a 400 - Sequencial do registro
        );

      // Descrição do Serviço Realizado
      StrList.Add( //
        Copy(FillChar(N.SrvTipoRegistro, 1, '0', False), 1, 1) + // 001 a 001 - Tipo de Registro
        Copy(FillChar(N.SrvSequencialNota, 20, '0', False), 1, 20) + // 002 a 021 - Sequencial da NFS-e
        Copy(FillChar(N.SrvCodServPrestado, 4, ' ', True), 1, 4) + // 022 a 025 - Código do serviço prestado
        Copy(FillChar(N.SrvCodTribMuni, 20, ' ', True), 1, 20) + // 026 a 045 - Código Tributação Município
        Copy(FillChar(IfThen(N.CabNotSitNota in [snCancelada, snExtraviada], '0', FormataDoubleCD(N.SrvValorServ, CabArqCasasDec_ValorServ)), 15, '0', False), 1, 15) + // 046 a 060 - Valor do Serviço
        Copy(FillChar(FormataDoubleCD(N.SrvValorDeducao, cdDois), 15, '0', False), 1, 15) + // 061 a 075 - Valor Dedução
        Copy(FillChar(FormatFloat('####', N.SrvAliq * 100), 4, '0', False), 1, 4) + // 076 a 079 - Alíquota
        Copy(FillChar(N.SrvUnidade, 20, ' ', True), 1, 20) + // 080 a 099 - Unidade do Serviço
        Copy(FillChar(FormataDoubleCD(N.SrvQuantidade, CabArqCasasDec_QuantServ), 8, '0', False), 1, 8) + // 100 a 107 - Quantidade
        Copy(FillChar(N.SrvDescricaoServ, 255, ' ', True), 1, 255) + // 108 a 362 - Descrição do Serviço
        Copy(FillChar(N.SrvAlvara, 20, ' ', True), 1, 20) + // 363 a 382 - Alvará
        FillChar('', 9, ' ', True) + // 383 a 391 - Preencher com Brancos
        Copy(FillChar(IntToStr(N.SrvSequencialReg), 8, '0', False), 1, 8) // 392 a 400 - Sequencial do registro
        );
    end;
  finally
    StrList.Free;
  end;
end;

end.
