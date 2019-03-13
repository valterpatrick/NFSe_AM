unit UFrmPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ComCtrls, ComObj, Math, StrUtils,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ExtDlgs, System.Generics.Collections, IniFiles, UNFSe_ArqMagSigISS;

type
  TTipoSistema = (tsSigISS, tsEL, tsNaoSelecionado);

  TFrmPrincipal = class(TForm)
    Pn_Top: TPanel;
    Label1: TLabel;
    Edt_Arquivo: TEdit;
    Btn_Arquivo: TBitBtn;
    Cmb_Sistema: TComboBox;
    Label2: TLabel;
    Pc_Sistemas: TPageControl;
    Ts_SigISS: TTabSheet;
    Pn_Bot: TPanel;
    Btn_SalvarArqMag: TBitBtn;
    Btn_Fechar: TBitBtn;
    Label3: TLabel;
    Edt_SigISS_InscPrestador: TEdit;
    Bevel1: TBevel;
    Label7: TLabel;
    Edt_SigISS_CodigoServico: TEdit;
    Label8: TLabel;
    Btn_Sobre: TBitBtn;
    Cmb_SigISS_SituacaoNota: TComboBox;
    Edt_SigISS_AliqSimplesNacional: TEdit;
    Label9: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Bevel2: TBevel;
    Label13: TLabel;
    OpenFile: TOpenTextFileDialog;
    Mem_SigISS_DescricaoAtividade: TMemo;
    SaveText: TSaveTextFileDialog;
    TabSheet1: TTabSheet;
    Btn_Configuracoes: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure Btn_ArquivoClick(Sender: TObject);
    procedure Btn_FecharClick(Sender: TObject);
    procedure Btn_SalvarArqMagClick(Sender: TObject);
    procedure Btn_SobreClick(Sender: TObject);
    procedure Cmb_SistemaChange(Sender: TObject);
    procedure Btn_ConfiguracoesClick(Sender: TObject);
  private
    ListaNotasExcelSigISS: TNotaExcelListaSigISS;
    ArqMag: TNFSe_ArqMagSigISS;
    procedure ConfigPageControl;
    procedure GerarArquivoMagnetico;
    procedure GerarArquivoMagnetico_SigISS;
    procedure ImportarDadosExcel;
    procedure ImportarDadosExcel_SigISS;
    procedure ValidarDados;
    procedure ValidarDados_SigISS;
    function SigISS_IndiceSitNotaToStr(Valor: Integer): String;
    function VerifTamArqGerado(Arquivo: String): Double;
    function IndiceToTipoSistema(Indice: Integer): TTipoSistema;
    procedure SaveConfig(Sistema: TTipoSistema);
    procedure ReadConfig(Sistema: TTipoSistema);
  public
  end;

var
  FrmPrincipal: TFrmPrincipal;

implementation

{$R *.dfm}

uses funcoes, UFrmAdvertencias, UFrmSobre;

procedure TFrmPrincipal.Btn_ArquivoClick(Sender: TObject);
begin
  if OpenFile.Execute then
    Edt_Arquivo.Text := OpenFile.FileName;
end;

procedure TFrmPrincipal.Btn_ConfiguracoesClick(Sender: TObject);
begin
  // Configurações
  Application.MessageBox('Em Desenvolvimento. Aguarde', 'Informação', MB_ICONINFORMATION + MB_OK);
end;

procedure TFrmPrincipal.Btn_FecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmPrincipal.Btn_SalvarArqMagClick(Sender: TObject);
begin
  if IndiceToTipoSistema(Cmb_Sistema.ItemIndex) = tsNaoSelecionado then
  begin
    if Cmb_Sistema.CanFocus then
      Cmb_Sistema.SetFocus;
    raise Exception.Create('Sistema não selecionado.');
  end;

  if not(Length(Trim(Edt_Arquivo.Text)) > 0) or not(FileExists(Trim(Edt_Arquivo.Text))) then
  begin
    if Edt_Arquivo.CanFocus then
      Edt_Arquivo.SetFocus;
    raise Exception.Create('Arquivo do Excel não selecionado.');
  end;

  if Application.MessageBox('Confirma a geração do arquivo magnético?', 'Confirmação', MB_ICONQUESTION + MB_YESNO) = IDNO then
    Exit;

  ImportarDadosExcel;
  ValidarDados;
  GerarArquivoMagnetico;
end;

procedure TFrmPrincipal.Btn_SobreClick(Sender: TObject);
begin
  FrmSobre := TFrmSobre.Create(Self);
  FrmSobre.showmodal;
  FrmSobre.Free;
end;

procedure TFrmPrincipal.SaveConfig(Sistema: TTipoSistema);

  procedure SaveSigISS;
  begin
    with TIniFile.Create(ExtractFileDir(Application.ExeName) + '\config.ini') do
    begin
      WriteString('SigISS', 'Inscricao_Prestador', Trim(Edt_SigISS_InscPrestador.Text));
      WriteInteger('SigISS', 'Codigo_Servico', StrToIntDef(Trim(Edt_SigISS_CodigoServico.Text), 0));
      WriteFloat('SigISS', 'Aliq_Simples_Nacional', StrToFloatDef(Trim(Edt_SigISS_AliqSimplesNacional.Text), 0));
      WriteInteger('SigISS', 'Indice_Situacao_Nota', Cmb_SigISS_SituacaoNota.ItemIndex);
      WriteString('SigISS', 'Descricao_Atividade', StringReplace(Trim(Mem_SigISS_DescricaoAtividade.Text), #13#10, ' ', [rfReplaceAll]));
    end;
  end;

  procedure SaveEL;
  begin
    //
  end;

begin
  case Sistema of
    tsSigISS:
      SaveSigISS;
    tsEL:
      SaveEL;
  end;
end;

procedure TFrmPrincipal.ReadConfig(Sistema: TTipoSistema);

  procedure ReadSigISS;
  begin
    with TIniFile.Create(ExtractFileDir(Application.ExeName) + '\config.ini') do
    begin
      Edt_SigISS_InscPrestador.Text := Trim(ReadString('SigISS', 'Inscricao_Prestador', ''));
      Edt_SigISS_CodigoServico.Text := IntToStr(ReadInteger('SigISS', 'Codigo_Servico', 0));
      Edt_SigISS_AliqSimplesNacional.Text := FloatToStr(ReadFloat('SigISS', 'Aliq_Simples_Nacional', 0));
      Cmb_SigISS_SituacaoNota.ItemIndex := ReadInteger('SigISS', 'Indice_Situacao_Nota', -1);
      Mem_SigISS_DescricaoAtividade.Lines.Clear;
      if Trim(ReadString('SigISS', 'Descricao_Atividade', '')) <> '' then
        Mem_SigISS_DescricaoAtividade.Lines.Add(ReadString('SigISS', 'Descricao_Atividade', ''));
    end;
  end;

  procedure ReadEL;
  begin
    //
  end;

begin
  with TIniFile.Create(ExtractFileDir(Application.ExeName) + '\config.ini') do
  begin
    case Sistema of
      tsSigISS:
        ReadSigISS;
      tsEL:
        ReadEL;
    end;
  end;
end;

procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
{$IFDEF DEBUG}
  Cmb_Sistema.ItemIndex := 0;
  Pc_Sistemas.ActivePageIndex := 0;
  Edt_Arquivo.Text := 'D:\valterpatrick\Desenvolvimento\Projetos\Desktop\Delphi\NFSe - AM\Win32\ArqTeste.xlsx';
  Edt_SigISS_InscPrestador.Text := '657726';
  Edt_SigISS_CodigoServico.Text := '802';
  Edt_SigISS_AliqSimplesNacional.Text := '3,96';
  Mem_SigISS_DescricaoAtividade.Lines.Add('Curso Pré Vestibular Extensivo Noturno');
{$ENDIF}
  ConfigPageControl;
  Cmb_SistemaChange(nil);
  Self.Enabled := True;
end;

procedure TFrmPrincipal.Cmb_SistemaChange(Sender: TObject);
begin
  Pc_Sistemas.ActivePageIndex := IfThen(Cmb_Sistema.ItemIndex = -1, 0, Cmb_Sistema.ItemIndex);
  ReadConfig(IndiceToTipoSistema(Pc_Sistemas.ActivePageIndex));
end;

procedure TFrmPrincipal.ConfigPageControl;
begin
  Pc_Sistemas.MultiLine := False;
  Pc_Sistemas.Align := alNone;
  Pc_Sistemas.Left := -4;
  Pc_Sistemas.Top := Pc_Sistemas.Top - 26;
  Pc_Sistemas.Width := ClientWidth + 8;
  Pc_Sistemas.Height := ClientHeight - Pn_Top.Height - Pn_Bot.Height + 27;
  Cmb_Sistema.ItemIndex := 0;
  Pc_Sistemas.ActivePageIndex := IfThen(Cmb_Sistema.ItemIndex = -1, 0, Cmb_Sistema.ItemIndex);
end;

procedure TFrmPrincipal.ValidarDados_SigISS;
var
  vAliq: Double;
  vCodServ: Integer;
  I: Integer;
  ListaAdvertenciasNotas: TStringList;

  procedure AddAdvertencia(Linha: Integer; Advertencia: String);
  begin
    ListaAdvertenciasNotas.Add('LINHA: ' + IntToStr(Linha) + '  -  ADVERTÊNCIA: ' + Advertencia.Trim);
  end;

begin
  ListaAdvertenciasNotas := TStringList.Create;
  vCodServ := StrToIntDef(Trim(Edt_SigISS_CodigoServico.Text), 0);
  vAliq := StrToFloatDef(Trim(Edt_SigISS_AliqSimplesNacional.Text), 0);

  if Trim(Edt_Arquivo.Text) = '' then
  begin
    if Edt_Arquivo.CanFocus then
      Edt_Arquivo.SetFocus;
    raise Exception.Create('Arquivo do Excel não informado.');
  end;

  if Trim(Edt_SigISS_InscPrestador.Text) = '' then
  begin
    if Edt_SigISS_InscPrestador.CanFocus then
      Edt_SigISS_InscPrestador.SetFocus;
    raise Exception.Create('Inscrição do prestador do serviço não informada.');
  end;

  if Length(Trim(Edt_SigISS_InscPrestador.Text)) > 15 then
  begin
    if Edt_SigISS_InscPrestador.CanFocus then
      Edt_SigISS_InscPrestador.SetFocus;
    raise Exception.Create('Inscrição do prestador maior que o permitido [15 caracteres].');
  end;

  if (Trim(Edt_SigISS_CodigoServico.Text) = '') or (Length(Trim(Edt_SigISS_CodigoServico.Text)) > 7) or (vCodServ <= 0) then
  begin
    if Edt_SigISS_CodigoServico.CanFocus then
      Edt_SigISS_CodigoServico.SetFocus;
    raise Exception.Create('Código do serviço padrão não informado, inválido ou maior que o permitido [7 digitos].');
  end;

  if (Trim(Edt_SigISS_AliqSimplesNacional.Text) = '') or (vAliq < 0) then
  begin
    if Edt_SigISS_AliqSimplesNacional.CanFocus then
      Edt_SigISS_AliqSimplesNacional.SetFocus;
    raise Exception.Create('Alíquota Simples Nacional padrão não informada ou inválida.');
  end;

  if vAliq = 0 then
    Application.MessageBox('A alíquota Simples Nacional padrão está zerada.', 'Aviso', MB_ICONWARNING + MB_OK);

  if Cmb_SigISS_SituacaoNota.ItemIndex = -1 then
  begin
    if Cmb_SigISS_SituacaoNota.CanFocus then
      Cmb_SigISS_SituacaoNota.SetFocus;
    raise Exception.Create('Situação da nota padrão não selecionada');
  end;

  if Mem_SigISS_DescricaoAtividade.GetTextLen = 0 then
  begin
    if Mem_SigISS_DescricaoAtividade.CanFocus then
      Mem_SigISS_DescricaoAtividade.SetFocus;
    raise Exception.Create('Descrição da Atividade não informada.');
  end;

  for I := 0 to ListaNotasExcelSigISS.Count - 1 do
  begin
    ListaNotasExcelSigISS.Items[I].A_CPFCNPJ := SomenteNumeros(ListaNotasExcelSigISS.Items[I].A_CPFCNPJ);
    if not Length(ListaNotasExcelSigISS.Items[I].A_CPFCNPJ) in [11, 14] then
      AddAdvertencia(I + 2, '"A - CPF/CNPJ" inválido.');

    if (StrToFloatDef(ListaNotasExcelSigISS.Items[I].B_ValorServico, 0) <= 0) then
      AddAdvertencia(I + 2, '"B - Valor do Serviço" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].C_Email) > 0) and (Pos('@', ListaNotasExcelSigISS.Items[I].C_Email) <= 0) then
      AddAdvertencia(I + 2, '"C - Email" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].D_CodigoServico) > 0) then
    begin
      if (StrToIntDef(ListaNotasExcelSigISS.Items[I].D_CodigoServico, 0) <= 0) then
        AddAdvertencia(I + 2, '"D - Código do Serviço" inválido.');
      if (Length(ListaNotasExcelSigISS.Items[I].D_CodigoServico) > 7) then
        AddAdvertencia(I + 2, '"D - Código do Serviço" maior que o permitido [7 caracteres].');
    end;

    if (Length(ListaNotasExcelSigISS.Items[I].E_AliqSimplesNacional) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].E_AliqSimplesNacional, 0) <= 0) then
      AddAdvertencia(I + 2, '"E - Alíquota Simples Nacional" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].F_SituacaoNotaFiscal) > 0) then
    begin
      if (Length(ListaNotasExcelSigISS.Items[I].F_SituacaoNotaFiscal) > 1) then
        AddAdvertencia(I + 2, '"F - Situação da Nota Fiscal" maior que o permitido [1 caracter].');
      if (StrInSet(ListaNotasExcelSigISS.Items[I].F_SituacaoNotaFiscal, ['T', 'R', 'C', 'I', 'N', 'U'])) then
        AddAdvertencia(I + 2, '"F - Situação da Nota Fiscal" inválido, o valor tem de estar entre os valores [T, R, C, I, N, U].');
    end;

    if (Length(ListaNotasExcelSigISS.Items[I].G_ValorRetencaoISS) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].G_ValorRetencaoISS, 0) < 0) then
      AddAdvertencia(I + 2, '"G - Valor Retenção ISS" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].H_ValorRetencaoINSS) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].H_ValorRetencaoINSS, 0) < 0) then
      AddAdvertencia(I + 2, '"H - Valor Retenção INSS" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].I_ValorRetencaoCOFINS) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].I_ValorRetencaoCOFINS, 0) < 0) then
      AddAdvertencia(I + 2, '"I - Valor Retenção COFINS" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].J_ValorRetencaoPIS) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].J_ValorRetencaoPIS, 0) < 0) then
      AddAdvertencia(I + 2, '"J - Valor Retenção PIS" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].K_ValorRetencaoIR) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].K_ValorRetencaoIR, 0) < 0) then
      AddAdvertencia(I + 2, '"K - Valor Retenção IR" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].L_ValorRetencaoCSLL) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].L_ValorRetencaoCSLL, 0) < 0) then
      AddAdvertencia(I + 2, '"L - Valor Retenção CSLL" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].M_ValorAproximadoTributos) > 0) and (StrToFloatDef(ListaNotasExcelSigISS.Items[I].M_ValorAproximadoTributos, 0) < 0) then
      AddAdvertencia(I + 2, '"M - Valor Aproximado dos Tributos" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].N_IdentificadorSistema_Legado) > 12) then
      AddAdvertencia(I + 2, '"N - Identificador Sistema Legado" maior que o permitido [12 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].O_InscricaoMunicipalTomador) > 15) then
      AddAdvertencia(I + 2, '"O - Inscrição Municipal Tomador" maior que o permitido [15 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].P_InscricaoEstadualTomador) > 15) then
      AddAdvertencia(I + 2, '"P - Inscrição Estadual Tomador" maior que o permitido [15 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].Q_NomeTomador) > 100) then
      AddAdvertencia(I + 2, '"Q - Nome/Razão Social do Tomador" maior que o permitido [100 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].R_EnderecoTomador) > 50) then
      AddAdvertencia(I + 2, '"R - Endereço do Tomador" maior que o permitido [50 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].S_NumeroEndereco) > 10) then
      AddAdvertencia(I + 2, '"S - Número do Endereço do Tomador" maior que o permitido [10 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].T_ComplementoEndereco) > 30) then
      AddAdvertencia(I + 2, '"T - Complemento do Endereço do Tomador" maior que o permitido [30 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].U_BairroTomador) > 30) then
      AddAdvertencia(I + 2, '"U - Bairro do Tomador" maior que o permitido [30 caracteres].');

    if (Length(ListaNotasExcelSigISS.Items[I].V_CodCidadeTomador) > 0) then
    begin
      if (StrToIntDef(ListaNotasExcelSigISS.Items[I].V_CodCidadeTomador, 0) < 0) then
        AddAdvertencia(I + 2, '"V - Código da Cidade do Tomador" inválido.');
      if (Length(ListaNotasExcelSigISS.Items[I].V_CodCidadeTomador) > 7) then
        AddAdvertencia(I + 2, '"V - Código da Cidade do Tomador" maior que o permitido [7 caracteres].');
    end;

    if (Length(ListaNotasExcelSigISS.Items[I].W_UnidadeFederalTomador) > 0) then
    begin
      if (Length(ListaNotasExcelSigISS.Items[I].W_UnidadeFederalTomador) > 2) then
        AddAdvertencia(I + 2, '"W - Unidade Federal do Tomador" maior que o permitido [2 caracteres].');
      if IsNumber(ListaNotasExcelSigISS.Items[I].W_UnidadeFederalTomador) or TemCaracteresEspeciais(ListaNotasExcelSigISS.Items[I].W_UnidadeFederalTomador) then
        AddAdvertencia(I + 2, '"W - Unidade Federal do Tomador" inválido.');
    end;

    ListaNotasExcelSigISS.Items[I].X_CEPTomador := SomenteNumeros(ListaNotasExcelSigISS.Items[I].X_CEPTomador);
    if (Length(ListaNotasExcelSigISS.Items[I].X_CEPTomador) > 0) and (Length(ListaNotasExcelSigISS.Items[I].X_CEPTomador) <> 8) then
      AddAdvertencia(I + 2, '"X - CEP do Tomador" inválido.');

    if (Length(ListaNotasExcelSigISS.Items[I].Y_CodCidPrestServ) > 0) then
    begin
      if (StrToIntDef(ListaNotasExcelSigISS.Items[I].Y_CodCidPrestServ, 0) <= 0) then
        AddAdvertencia(I + 2, '"Y - Código da Cidade do Local de Prestação de Serviço" inválido.');
      if (Length(ListaNotasExcelSigISS.Items[I].Y_CodCidPrestServ) > 7) then
        AddAdvertencia(I + 2, '"Y - Código da Cidade do Local de Prestação de Serviço" maior que o permitido [7 caracteres].');
    end;

    if (Length(ListaNotasExcelSigISS.Items[I].Z_DesricaoServico) > 0) and (Length(ListaNotasExcelSigISS.Items[I].Z_DesricaoServico) > 1000) then
      AddAdvertencia(I + 2, '"Z - Descrição do Serviço" maior que o permitido [1000 caracteres].');
  end;

  if ListaAdvertenciasNotas.Count > 0 then
  begin
    FrmAdvertencias := TFrmAdvertencias.Create(Self, 'Advertências', ListaAdvertenciasNotas);
    FrmAdvertencias.showmodal;
    FrmAdvertencias.Free;
    Abort;
  end;
end;

procedure TFrmPrincipal.ValidarDados;
begin
  if IndiceToTipoSistema(Cmb_Sistema.ItemIndex) = tsSigISS then
    ValidarDados_SigISS;
end;

procedure TFrmPrincipal.ImportarDadosExcel;
begin
  if IndiceToTipoSistema(Cmb_Sistema.ItemIndex) = tsSigISS then
    ImportarDadosExcel_SigISS;
end;

procedure TFrmPrincipal.ImportarDadosExcel_SigISS;
var
  ArqExcel: String;
  Excel: OleVariant;
  Linha: Integer;
begin
  ArqExcel := Edt_Arquivo.Text;
  try
    // Cria o Objeto Excel
    Excel := CreateOleObject('Excel.Application');
    // Esconde o Excel
    Excel.Visible := False;
    // Abre o Workbook
    Excel.WorkBooks.Open(ArqExcel);
    // Abre a primeira aba
    Excel.WorkSheets[1].Activate;
  except
    ShowMessage('Ocorreu um erro ao abrir o arquivo.');
  end;

  // Realiza a leitura do arquivo, começando da segunda linha, já que a primeira é o cabeçalho
  Linha := 2;
  if Assigned(ListaNotasExcelSigISS) then
    ListaNotasExcelSigISS.Destroy;
  ListaNotasExcelSigISS := TNotaExcelListaSigISS.Create;
  try
    while VarToStrDef(Excel.Cells[Linha, 1].Value, '') <> '' do
    begin
      with ListaNotasExcelSigISS.New do
      begin
        A_CPFCNPJ := StringReplace(Trim(Excel.Cells[Linha, 1]), #13#10, ' ', [rfReplaceAll]);
        B_ValorServico := StringReplace(Trim(Excel.Cells[Linha, 2]), #13#10, ' ', [rfReplaceAll]);
        C_Email := StringReplace(Trim(Excel.Cells[Linha, 3]), #13#10, ' ', [rfReplaceAll]);
        D_CodigoServico := StringReplace(Trim(Excel.Cells[Linha, 4]), #13#10, ' ', [rfReplaceAll]);
        E_AliqSimplesNacional := StringReplace(Trim(Excel.Cells[Linha, 5]), #13#10, ' ', [rfReplaceAll]);
        F_SituacaoNotaFiscal := StringReplace(Trim(Excel.Cells[Linha, 6]), #13#10, ' ', [rfReplaceAll]);
        G_ValorRetencaoISS := StringReplace(Trim(Excel.Cells[Linha, 7]), #13#10, ' ', [rfReplaceAll]);
        H_ValorRetencaoINSS := StringReplace(Trim(Excel.Cells[Linha, 8]), #13#10, ' ', [rfReplaceAll]);
        I_ValorRetencaoCOFINS := StringReplace(Trim(Excel.Cells[Linha, 9]), #13#10, ' ', [rfReplaceAll]);
        J_ValorRetencaoPIS := StringReplace(Trim(Excel.Cells[Linha, 10]), #13#10, ' ', [rfReplaceAll]);
        K_ValorRetencaoIR := StringReplace(Trim(Excel.Cells[Linha, 11]), #13#10, ' ', [rfReplaceAll]);
        L_ValorRetencaoCSLL := StringReplace(Trim(Excel.Cells[Linha, 12]), #13#10, ' ', [rfReplaceAll]);
        M_ValorAproximadoTributos := StringReplace(Trim(Excel.Cells[Linha, 13]), #13#10, ' ', [rfReplaceAll]);
        N_IdentificadorSistema_Legado := StringReplace(Trim(Excel.Cells[Linha, 14]), #13#10, ' ', [rfReplaceAll]);
        O_InscricaoMunicipalTomador := StringReplace(Trim(Excel.Cells[Linha, 15]), #13#10, ' ', [rfReplaceAll]);
        P_InscricaoEstadualTomador := StringReplace(Trim(Excel.Cells[Linha, 16]), #13#10, ' ', [rfReplaceAll]);
        Q_NomeTomador := StringReplace(Trim(Excel.Cells[Linha, 17]), #13#10, ' ', [rfReplaceAll]);
        R_EnderecoTomador := StringReplace(Trim(Excel.Cells[Linha, 18]), #13#10, ' ', [rfReplaceAll]);
        S_NumeroEndereco := StringReplace(Trim(Excel.Cells[Linha, 19]), #13#10, ' ', [rfReplaceAll]);
        T_ComplementoEndereco := StringReplace(Trim(Excel.Cells[Linha, 20]), #13#10, ' ', [rfReplaceAll]);
        U_BairroTomador := StringReplace(Trim(Excel.Cells[Linha, 21]), #13#10, ' ', [rfReplaceAll]);
        V_CodCidadeTomador := StringReplace(Trim(Excel.Cells[Linha, 22]), #13#10, ' ', [rfReplaceAll]);
        W_UnidadeFederalTomador := StringReplace(Trim(Excel.Cells[Linha, 23]), #13#10, ' ', [rfReplaceAll]);
        X_CEPTomador := StringReplace(Trim(Excel.Cells[Linha, 24]), #13#10, ' ', [rfReplaceAll]);
        Y_CodCidPrestServ := StringReplace(Trim(Excel.Cells[Linha, 25]), #13#10, ' ', [rfReplaceAll]);
        Z_DesricaoServico := StringReplace(Trim(Excel.Cells[Linha, 26]), #13#10, ' ', [rfReplaceAll]);
      end;
      Linha := Linha + 1;
    end;

    if Linha = 2 then
      Application.MessageBox('Não foram encontrados dados na planilha', 'Aviso', MB_ICONWARNING + MB_OK);
  finally
    // Fecha o Excel
    if not VarIsEmpty(Excel) then
    begin
      Excel.Quit;
      Excel := Unassigned;
    end;
  end;

end;

function TFrmPrincipal.SigISS_IndiceSitNotaToStr(Valor: Integer): String;
begin
  case Valor of
    0:
      Result := 'T'; // Tributada
    1:
      Result := 'R'; // Retida
    2:
      Result := 'C'; // Cancelada
    3:
      Result := 'I'; // Isenta
    4:
      Result := 'N'; // Não Tributada
    5:
      Result := 'U'; // Imune
  end;
end;

function TFrmPrincipal.IndiceToTipoSistema(Indice: Integer): TTipoSistema;
begin
  case Indice of
    0:
      Result := tsSigISS;
    1:
      Result := tsEL;
  else
    Result := tsNaoSelecionado;
  end;
end;

procedure TFrmPrincipal.GerarArquivoMagnetico;
begin
  if IndiceToTipoSistema(Cmb_Sistema.ItemIndex) = tsSigISS then
  begin
    GerarArquivoMagnetico_SigISS;
    SaveConfig(tsSigISS);
  end;
end;

procedure TFrmPrincipal.GerarArquivoMagnetico_SigISS;
var
  I: Integer;
  vTamArq: Double;
begin
  ArqMag := TNFSe_ArqMagSigISS.Create;
  Self.Enabled := False;
  try
    ArqMag.CabTipoRegistro := '1';
    ArqMag.CabTipoArquivo := 'NFE_LOTE';
    ArqMag.CabInscPrestador := Trim(Edt_SigISS_InscPrestador.Text);
    ArqMag.CabVersaoArquivo := '030';
    ArqMag.CabDataArquivo := Date;

    ArqMag.ListaNotas := TNotaFiscalListaSigISS.Create;
    for I := 0 to ListaNotasExcelSigISS.Count - 1 do
    begin
      with ArqMag.ListaNotas.New do
      begin
        DetTipoRegistro := '2';
        DetIdentSistema := ListaNotasExcelSigISS[I].N_IdentificadorSistema_Legado;
        DetTipoCodificacao := 1;
        DetCodServico := IfThen(StrToIntDef(ListaNotasExcelSigISS[I].D_CodigoServico, 0) = 0, StrToIntDef(Trim(Edt_SigISS_CodigoServico.Text), 0), StrToIntDef(ListaNotasExcelSigISS[I].D_CodigoServico, 0));
        if Length(Trim(ListaNotasExcelSigISS[I].F_SituacaoNotaFiscal)) = 0 then
          DetSitNota := ArqMag.ListaNotas.StrToSituacaoNota(SigISS_IndiceSitNotaToStr(Cmb_SigISS_SituacaoNota.ItemIndex))
        else
          DetSitNota := ArqMag.ListaNotas.StrToSituacaoNota(Trim(ListaNotasExcelSigISS[I].F_SituacaoNotaFiscal));
        DetValorServico := StrToFloatDef(ListaNotasExcelSigISS[I].B_ValorServico, 0);
        DetValorBaseServico := DetValorServico;
        DetAliqSimplesNacional := IfThen(StrToFloatDef(ListaNotasExcelSigISS[I].E_AliqSimplesNacional, 0) = 0, StrToFloatDef(Trim(Edt_SigISS_AliqSimplesNacional.Text), 0), StrToFloatDef(ListaNotasExcelSigISS[I].E_AliqSimplesNacional, 0));
        DetValorRetencaoISS := StrToFloatDef(ListaNotasExcelSigISS[I].G_ValorRetencaoISS, 0);
        DetValorRetencaoINSS := StrToFloatDef(ListaNotasExcelSigISS[I].H_ValorRetencaoINSS, 0);
        DetValorRetencaoCOFINS := StrToFloatDef(ListaNotasExcelSigISS[I].I_ValorRetencaoCOFINS, 0);
        DetValorRetencaoPIS := StrToFloatDef(ListaNotasExcelSigISS[I].J_ValorRetencaoPIS, 0);
        DetValorRetencaoIR := StrToFloatDef(ListaNotasExcelSigISS[I].K_ValorRetencaoIR, 0);
        DetValorRetencaoCSLL := StrToFloatDef(ListaNotasExcelSigISS[I].L_ValorRetencaoCSLL, 0);
        DetValorAproximadoTributos := StrToFloatDef(ListaNotasExcelSigISS[I].M_ValorAproximadoTributos, 0);
        DetTomadorCPF_CNPJ := ListaNotasExcelSigISS[I].A_CPFCNPJ;
        DetTomadorInscMunicipal := ListaNotasExcelSigISS[I].O_InscricaoMunicipalTomador;
        DetTomadorInscEstadual := ListaNotasExcelSigISS[I].P_InscricaoEstadualTomador;
        DetTomadorNome := ListaNotasExcelSigISS[I].Q_NomeTomador;
        DetTomadorEndLogradouro := ListaNotasExcelSigISS[I].R_EnderecoTomador;
        DetTomadorEndNumero := ListaNotasExcelSigISS[I].S_NumeroEndereco;
        DetTomadorEndComplemento := ListaNotasExcelSigISS[I].T_ComplementoEndereco;
        DetTomadorEndBairro := ListaNotasExcelSigISS[I].U_BairroTomador;
        DetTomadorEndCodCidade := StrToIntDef(ListaNotasExcelSigISS[I].V_CodCidadeTomador, 0);
        DetTomadorEndUF := ListaNotasExcelSigISS[I].W_UnidadeFederalTomador;
        DetTomadorEndCEP := ListaNotasExcelSigISS[I].X_CEPTomador;
        DetTomadorEmail := ListaNotasExcelSigISS[I].C_Email;
        DetCodCidadePrestServ := StrToIntDef(ListaNotasExcelSigISS[I].Y_CodCidPrestServ, 0);
        DetDescriminacaoServ := IfThen(Length(ListaNotasExcelSigISS[I].Z_DesricaoServico) > 0, ListaNotasExcelSigISS[I].Z_DesricaoServico, StringReplace(Trim(Mem_SigISS_DescricaoAtividade.Text), #13#10, ' ', [rfReplaceAll]));
      end;
    end;

    ArqMag.RodTipoRegistro := '9';
    if SaveText.Execute then
    begin
      ArqMag.SaveToFile(SaveText.FileName);
      vTamArq := VerifTamArqGerado(SaveText.FileName);
      Application.MessageBox(PWideChar('Arquivo gerado com sucesso em "' + SaveText.FileName + '".' //
        + IfThen(vTamArq >= 4.99, ' Arquivos com tamanho igual ou superior a 5MB serão ignorados. Tamanho do arquivo: ' + FloatToStr(vTamArq) + 'MB.', '')), 'Informação', MB_ICONINFORMATION + MB_OK);
    end;
  finally
    Self.Enabled := True;
    ArqMag.ListaNotas.Destroy;
    ArqMag.Destroy;
  end;
end;

function TFrmPrincipal.VerifTamArqGerado(Arquivo: String): Double;
begin
  with TFileStream.Create(Arquivo, fmOpenRead or fmShareExclusive) do
  begin
    try
      Result := Size; // Valor retornado em Bytes
      Result := (Result / 1024) / 1024; // Tem que dividir duas vezes para chegar em MB
    finally
      Free;
    end;
  end;
end;

end.
