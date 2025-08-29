unit ULibrary;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  ExtCtrls, Forms, Dialogs, StdCtrls, DB, Buttons, Menus, Ora, OraSmart,
  OraError,
  // DBCtrls,  DBLookup,	DBGrids, DBClient,
  LzExpand, Windows, Math, Variants,
  Mask, Grids, ComCtrls, PRINTERS, RichEdit, ComObj,
  Activex, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, cxControls, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxCalendar, cxLookAndFeelPainters, cxStyles, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxDBData, cxHint,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, cxButtons, cxPC, cxDBFilterControl,
  cxGroupBox, cxRadioGroup, cxCheckBox, cxCurrencyEdit, cxGridStrs,
  dxPSGlbl, dxPSUtl, dxPSEngn, dxPrnPg, dxBkgnd, cxNavigator,
  dxWrap, dxPrnDev, dxPSCompsProvider, dxPSFillPatterns, dxPSEdgePatterns,
  dxPSCore, dxPScxCommon, SqlExpr, DateUtils, ClipBrd,
  // dxPScxGridLnk,
  cxLabel, cxImageComboBox, StrUtils,
  dxmdaset, cxCheckGroup, cxLookAndFeels,
  cxLocalization, cxEditConsts, cxExtEditConsts, cxDataConsts, cxLibraryStrs,
  cxFilterConsts, cxGridPopupMenuConsts, cxExportStrs, dxPSRes, IniFiles,
  ShellAPI,

  // Jhonny Oliveira - 05/09/2014
  // Necessários para o envio de e-mails
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP, IdMessage, IdAttachment, IdAttachmentFile

  // Jhonny Oliveira - 01/12/2014
  // Necessários para salvar o layout do grid em arquivo
    , cxGridBandedTableView, cxGridDBBandedTableView

  // Jhonny Oliveira - 15/01/2014
  // Necessário para exportar o conteúdo dos grids
  // da dev para Excel
    , cxGridExportLink

    , MidasLib

    ;
// UFrmLogin,
// UUsuariosWinthor;
// cxCustomPivotGrid, cxDBPivotGrid, dxPScxPivotGridLnk,cxPivotGridStrs
// cxExportGrid4Link, cxExportPivotGridLink,

// #############################################################################
// Funções
// #############################################################################

Function Pot(base, expoente: real): real; // Potenciação
Function TrocaVirgPPto(Valor: string): string; { Troca virgula por ponto }
Function TrocaVirgulaporPonto(aValue: Double): string; { Troca virgula por ponto (FUNCAO 2) }
Function ProximoDiaUtil(dData: TDateTime): TDateTime;
Function FormJaExiste(PForm: TForm): Boolean; { Verifica se o Form ja Existe }
Function AcertaData(strData: string; intSQLServer: integer): Variant; { Funcao que Acerta a Data }
Function AcertaData1(strData: string; intSQLServer: integer): Variant; { Funcao que Acerta a Data }
Function Retorna_Dia_Semana(dtdata: TDateTime): string; { Retorna o dia da Semana }
Function Encripta(Senha: string): string; { Encripta Senhas }
Function Desencripta(Senha: string): string; { Desencripta Senhas }
Function CalculaAnos(Data_de_Nascimento, Data_Base: TDateTime): integer; { Calcula a idade de acordo com a data de nascimento digitada }
Function Gerapercentual(Valor: real; Percent: real): real; { Retorna a porcentagem de um valor }
Function Maiuscula(Texto: string): string; { Converte a primeira letra do texto especificado para maiuscula e as restantes para minuscula }
Function CalculaCnpjCpf(Numero: string): Boolean;
Function ApenasNumerosStr(pStr: string): string;
Function TestaCampoBranco(DataSet: TDataSet): Boolean; { Testa se o Campo esta em Branco }
Function StringAsPChar(var S: string): PChar; { Transforma string em pchar }
Function ToString(Value: Variant): String; { tranforma qualquer coisa em string }
Function Retorna_Versao: string; { Pega a Versao do Sistema }
Function DifHora(Inicio, Fim: string): string; { Retorna A Diferença entre Duas Horas }
Function IdadeAtual(Nasc: TDate): integer;
Function GetOrCreateObject(const ClassName: string): IDispatch;
Function Justifica(mCad: string; mMAx: integer): string;
Function PegaTamanhoPapel(dmPaperSize: word): string;
Function Letra_linha(ln: integer): string;
Function soma(v1, v2: real): real;
Function MULTI(v1, v2: real): real;
Function DIVE(v1, v2: real): real;
Function DIMI(v1, v2: real): real;
Function NetSend(dest, source, msg: string): longint;
Function PrinterOnLine: Boolean;
Function SerialNum(Unidade: PChar): string;
Function SysComputerName: string;
Function StrIsDate(const S: string): Boolean;
Function AjustaStr(str: string; tam: integer): string;
Function Padr(S: string; n: integer): string;
Function MAQTOINT(VMAQ: string): longint;
Function StrToPChar(const str: string): PChar;
Function StrToChar(str: string): Char;
Function StrZero(Zeros: string; Quant: integer): string;
Function StrEspaco(Zeros: string; Quant: integer): string;
Function ReplStr(const S: string; const Len: integer): string;
Function ReplChar(const Ch: Char; const Len: integer): string;
Function Dac11(const D11_Number: string; const Operador: integer): string;
Function Dac10(const D10_Number: string; const Operador: integer): string;
Function IIf(pCond: Boolean; pTrue, pFalse: Variant): Variant;
Function Tiraponto(pStr: string): string;
Function VerifLpt(lptStr: string): string;
Function Arredondar(Valor: Double; Dec: integer): Double;
Function VlrStr(Numero: string; tam: integer): string;
Function ValidaCarteira(Numero: string; Parametro: string): Boolean; { Validacao da Carteirinha dos convenios }
Function ValidarCNPJ(CNPJ: string; var aMessage: string): string;
Function RemoveChar(Const Texto: string): string;
Function ValidaData(const S: string): Boolean;
Function AnoBiSexto(Ayear: integer): Boolean;
Function DiasPorMes(Ayear, AMonth: integer): integer;
Function FirstDayOfMonth(Data: TDateTime; lSabDom: Boolean): TDateTime;
Function LastDayOfMonth(Data: TDateTime; lSabDom: Boolean): TDateTime;
Function EspacoStr(Zeros: string; Quant: integer): string;
Function WordsCount(S: string): integer;
Function ARound(Value: Extended; Decimals: integer): Extended;
Function RemoveAcentos(str: string): string;
function RemoveEspeciais(str: String): String;
Function MensagemDlg(txtMsg: String): Boolean;
Function InputBoxPass(const ACaption, APrompt, ADefault: string): string;
Function InputSenha(const ACaption, APrompt: string; var Value: string): Boolean;
Function GetAveCharSize(Canvas: TCanvas): TPoint;
Function Esconde(DADOin: string): string;
Function Desvenda(DADOin: string): string;
Function BuscaTroca(Text, Busca, Troca: string): string; { Substitui um caractere dentro da string }
Function TrocaNaPosicao(Text: string; Posicao: integer; Troca: string): string; { Substitui um caractere dentro da string na posicao pedida }
procedure CopiarArquivo(Origem, Destino: String);

function ConverteListaEmStringParaComandoSQL(aLista: TStringList; aIncluirAspas: Boolean = true): String;
function DiferencaEntreDatasEmHoras(ADataInicial, ADataFinal: TDateTime): String;
function ObtemConfiguracao(ACodfilial: String; ACodConfiguracao: Double): String;
function ObtemConfiguracaoFloat(ACodfilial: String; ACodConfiguracao: Double): double;

function AbrirManualUsuario(ACodigoRotina: String; ANumeroVersao: String): Boolean;

function ValidadordeVersao(ACodigoRotina: string; ANumeroVersao: String): Boolean;
function DiaUtil(ACodigoFilial: string; AData: TDateTime): Boolean;

// funcoes para tratamento de codigo de barras
procedure GerarCodigo(Codigo: String; Canvas: TCanvas);
function calcula_linha(barra: string): string;
function calcula_barra(linha: string): string;
function Modulo10(Valor: String): string;
// function Modulo11(Valor: String; Base: Integer = 9; Resto : boolean = false) : string;
function Modulo11(n: string): integer;
function SoLetraeNumero(Const Texto: string): String;

// function LoginColetor() : TUsuario;

// #############################################################################
// Procedures
// #############################################################################

Procedure Verifica_Label(F: TForm); { Funcao VericaLabel (Altera Propriedades do Label para Negrito) }
Procedure FechaQuery(F: TForm); { Fecha a Qry }
Procedure Limpa_Tela(F: TForm); { Limpa a Tela }
Procedure AbreQUERY(F: TForm); { Abre a Qry }
Procedure Erro; { Erro }
Procedure Tratamento_Erro(DataSet: TDataSet; E: EOraError; var Action: TDataAction); { Trata os Erros }
Procedure VerificaFORM(F: TForm); { Verifica Form }
// asn     Procedure GridSoLeitura(F:TForm);{Grid so Leitura}
Procedure EntreDatas(DataFinal, DataInicial: TDate; var Anos, Meses, Dias: integer);
Procedure MudaTamPapel(PaperSize, Comp, Alt: integer);
Procedure CriaCodigo(Cod: string; Imagem: TCanvas);
Procedure EscondeTaskBar(Visible: Boolean);
Procedure PrintRichEdit(const Caption: string; const RichEdt: TRichEdit);
Procedure MouseParaControle(Controle: TControl);
Procedure APrint(const S: Variant; const tip: string; const Len: integer; const wl: integer; const esp: integer);
Procedure MPrint(const S: Variant; const tip: string; const Len: integer; const wl: integer; const esp: integer);
Procedure PrintMat(const ln: integer; const cl: integer; const S: Variant; const tip: string; const esp: integer; const tc: string);
Procedure GravaLog(const usrlog: Variant; const dtlog: TDateTime; const modlog: Variant; const oplog: Variant; const doclog: Variant;
  const idlog: Variant);
// asn     Procedure abrirConexaoBDE;
// Procedure abrirConexaoODAC(ASessaoPadrao: TOraSession = nil);
Procedure abrirConexaoODAC(servidor: string = ''; usuario: string = ''; senha: string = '');
// asn     Procedure abrirConexaoBDESemparametros(pnomebase:string;pnomeusuario:string;psenha:string);
// asn     Procedure abrirConexaoODACSemparametros(pnomebase:string;pnomeusuario:string;psenha:string);
procedure AtribuiDbName(F: TForm; DbName: TOraSession);
procedure AtribuiSessionForm(F: TForm; SessionName: TOraSession);
procedure AtribuiSessionDmd(F: TDataModule; SessionName: TOraSession);
procedure CentralizarPanelNoForm(APanel: TPanel; AForm: TForm);
procedure AbrirConexaoDBExpress(ANomeServidor, AUsuarioBD, SenhaBD: String; ACompomenteConnection: TSQLConnection);

// asn     procedure ConectarComBoinghiCNC();
// asn     procedure ConectarComCNC(base:String);
procedure ConectarMySQLComCNC();

function PadLeft(AStringAtual: String; ACaracterACompletar: Char; ATamanhoTotal: integer): String;

function PadRight(AStringAtual: String; ACaracterACompletar: Char; ATamanhoTotal: integer): String;

function EnviarEmail(AEnderecoHost, ANomeUsuario, ASenha, AEmailRemetente, ANomeRemetente, AEmailParaResposta, ACorpoEmail, AAssunto: String;
  ADestinatarios: TStringList; var ARespostaServidor: String): Boolean;

procedure SalvarLayoutDosGridsDoForm(AFormulario: TForm);
procedure RestaurarLayoutDosGridsDoForm(AFormulario: TForm);
function VerificaVersao(ACodigoRotina: Double; AExibirMensagens: Boolean = true): Boolean;
function IniciaOS(AnumOS: Double; Amatricula: Double; AColetor: Boolean = false): Boolean;
function FechaOS(AnumOS: Double; AtipoOS: integer { 1-Separação, 2-Recebimento, 3-Transferência }
  ; Amatricula: Double): Boolean;

function ExportarExcel(AGrid: TcxGrid; ANomeArquivo: String = ''; AExibirMensagem: Boolean = true): Boolean;

Function PrimeiroNome(Nome: string): string;
Function UltimoNome(Nome: string): string;
function NomeAbreviado(Nome: String; TamanhoMaximo: integer): String;
function ObterMaiorNome(Nomes: TStringList): integer;
function ReduzNome(Nomes: TStringList; TamanhoMaximo: integer): String;
function SeparaNomes(Nome: String; n: TStringList): TStringList;
function VerTamanhoNome(Nomes: TStringList): integer;

// Debug - Jhonny Oliveira -> 08/10/15
procedure debug(ATexto: String);

/// Marcos Pereira 05/02/2016
Function SeparaNumeroEmbalagem(pCaixasdaEmbalagem: String): Double;

function MinutosToStr(const Minutes: Cardinal; Reduzido: Boolean = false): string;

Function AbrirProgramaExternoModal(FileName: String; Params: String = ''; Visibility: integer = SW_SHOWNORMAL): DWORD;

function CalcularDigitoVerificadorEAN(Numero: string): string;

var
  fpausa: Boolean;
  linhaatual: integer = 1; // define a linha de impressao atual Aprint-PrintMat.
  linha: string;
  pag: integer;
  FIMP: TextFile;
  // BdMatriz: TFDConnection; // referente a procedure abrirConexaoBDE()
  // BdFilial: TFDConnection; // referente a procedure abrirConexaoBDEsemparametros()

  ODACSessionGlobal: TOraSession; // referente a procedure abrirConexaoOdac()
  CxFilial: TOraSession; // referente a procedure abrirConexaoOdacsemparametros()

  // Debug - Jhonny 08/10/2015
  FrmDebug: TForm; // Form aque apresenta as mensagens de debug
  habilitarDebug: Boolean; // Indica se o debug está habilitado ou não
  exibirHoraDebug: Boolean; // Indica se o debug vai adicionar ou não a data e hora corrente do sistema

implementation

type
  TChars = set of Char;

Procedure abrirConexaoODAC(servidor: string = ''; usuario: string = ''; senha: string = '');

  procedure DefineParametrosConexao(AServidor, AUsuario, ASenha: string);
  const
    constConnStr = 'Direct=True;Host=&host;Service Name=&servicename;User ID=&user;Password=&password;Login Prompt=False';
  var // variaveis usadas na conexao direta
    vHost: string; // no host pode conter host:port
    vService: string;
    vConnStr: string;
  begin
    ODACSessionGlobal.ConnectString := ''; // limpa
    ODACSessionGlobal.Server := '';
    ODACSessionGlobal.Password := '';

    if Pos('/', AServidor) > 0 then // Usa conexao direct (nao precisa de oraclehome)
    begin
      ODACSessionGlobal.Options.Direct := true;
      vHost := Copy(AServidor, 1, Pos('/', AServidor) - 1);
      vService := Copy(AServidor, Pos('/', AServidor) + 1, Length(AServidor));
      vConnStr := StringReplace(constConnStr, '&host', vHost, []);
      vConnStr := StringReplace(vConnStr, '&servicename', vService, []);
      vConnStr := StringReplace(vConnStr, '&user', AUsuario, []);
      vConnStr := StringReplace(vConnStr, '&password', ASenha, []);

      ODACSessionGlobal.ConnectString := vConnStr;
    end
    else // forma nao direta (depende de oraclehome)
    begin
      ODACSessionGlobal.Options.Direct := false;
      ODACSessionGlobal.Server := AServidor;
      ODACSessionGlobal.Username := AUsuario;
      ODACSessionGlobal.Password := ASenha;
    end;
  end;

begin
  if not Assigned(ODACSessionGlobal) then
  begin
    ODACSessionGlobal := TOraSession.Create(nil);
  end;

  try
    begin

      if (ParamCount = 6) and (ParamStr(6) = 'DEBUG') then
      begin

        ODACSessionGlobal.ConnectString := 'Direct=True;Host=10.0.1.204;Service Name=WINT;User ID=ESPERANCA;Password=TESTEESPERANCA;Login Prompt=False';
        ODACSessionGlobal.Connected := true;
      end
      else
      begin

        ODACSessionGlobal.Connected := false;
        ODACSessionGlobal.LoginPrompt := false;
        ODACSessionGlobal.Name := 'ORAConnection';

        if servidor = '' then
        begin
          servidor := ParamStr(3);
        end;

        if usuario = '' then
        begin
          usuario := ParamStr(4);
        end;

        if senha = '' then
        begin
          senha := ParamStr(2);
        end;

        DefineParametrosConexao(
          servidor, // tela de Login do Winthor = Loja
          usuario, // tela de Login do Winthor = Empresa
          senha // associado a chave do Winthor.ini
          );

        ODACSessionGlobal.Connected := true;
      end;
    end;
  except
    on E: Exception do
    begin
      debug(' Exception');
      ODACSessionGlobal.Close;
      raise Exception.Create(E.Message);
    end;
  end;
end;

Function ARound(Value: Extended; Decimals: integer): Extended;
var
  Factor, Fraction: Extended;
begin
  Factor := IntPower(10, Decimals);
  { A conversão para string e depois para float evita
    erros de arredondamentos indesejáveis. }
  Value := StrToFloat(FloatToStr(Value * Factor));
  Result := Int(Value);
  Fraction := Frac(Value);
  if Fraction > 0.5 then
    Result := Result + 1
  else if Fraction <= -0.5 then
    Result := Result - 1;
  Result := Result / Factor;
end;

function ValidarCNPJ(CNPJ: string; var aMessage: string): string;

  function LimparCNPJ(CNPJ: string): string;
  begin
    Result := StringReplace(CNPJ, '.', '', [rfReplaceAll]);
    Result := StringReplace(Result, '-', '', [rfReplaceAll]);
    Result := StringReplace(Result, '/', '', [rfReplaceAll]);
  end;

var
  i, soma, mult: integer;
  aCNPJ: string;
begin
  Result := '';

  aCNPJ := LimparCNPJ(CNPJ);

  Result := aCNPJ;

  if Length(aCNPJ) <> 14 then
  begin
    aMessage := 'CNPJ Inválido';
    Exit;
  end;

  soma := 0;
  mult := 2;

  for i := 12 downto 1 do
  begin
    soma := soma + StrToInt(aCNPJ[i]) * mult;
    mult := mult + 1;
    if mult > 9 then
      mult := 2;
  end;

  mult := soma mod 11;

  if mult <= 1 then
    mult := 0
  else
    mult := 11 - mult;

  if mult <> StrToInt(aCNPJ[13]) then
  begin
    aMessage := 'CNPJ Inválido';
    Exit;
  end;

  soma := 0;
  mult := 2;

  for i := 13 downto 1 do
  begin
    soma := soma + StrToInt(aCNPJ[i]) * mult;
    mult := mult + 1;

    if mult > 9 then
      mult := 2;
  end;

  mult := soma mod 11;
  if mult <= 1 then
    mult := 0
  else
    mult := 11 - mult;

  if mult = StrToInt(aCNPJ[14]) then
    Result := aCNPJ
  else
    aMessage := 'CNPJ Inválido';
end;

Function RemoveAcentos(str: string): string;
var
  i: integer;
begin
  for i := 1 to Length(str) do
    case str[i] of
      'á':
        str[i] := 'a';
      'é':
        str[i] := 'e';
      'í':
        str[i] := 'i';
      'ó':
        str[i] := 'o';
      'ú':
        str[i] := 'u';
      'à':
        str[i] := 'a';
      'è':
        str[i] := 'e';
      'ì':
        str[i] := 'i';
      'ò':
        str[i] := 'o';
      'ù':
        str[i] := 'u';
      'â':
        str[i] := 'a';
      'ê':
        str[i] := 'e';
      'î':
        str[i] := 'i';
      'ô':
        str[i] := 'o';
      'û':
        str[i] := 'u';
      'ä':
        str[i] := 'a';
      'ë':
        str[i] := 'e';
      'ï':
        str[i] := 'i';
      'ö':
        str[i] := 'o';
      'ü':
        str[i] := 'u';
      'ã':
        str[i] := 'a';
      'õ':
        str[i] := 'o';
      'ñ':
        str[i] := 'n';
      'ç':
        str[i] := 'c';
      'Á':
        str[i] := 'A';
      'É':
        str[i] := 'E';
      'Í':
        str[i] := 'I';
      'Ó':
        str[i] := 'O';
      'Ú':
        str[i] := 'U';
      'À':
        str[i] := 'A';
      'È':
        str[i] := 'E';
      'Ì':
        str[i] := 'I';
      'Ò':
        str[i] := 'O';
      'Ù':
        str[i] := 'U';
      'Â':
        str[i] := 'A';
      'Ê':
        str[i] := 'E';
      'Î':
        str[i] := 'I';
      'Ô':
        str[i] := 'O';
      'Û':
        str[i] := 'U';
      'Ä':
        str[i] := 'A';
      'Ë':
        str[i] := 'E';
      'Ï':
        str[i] := 'I';
      'Ö':
        str[i] := 'O';
      'Ü':
        str[i] := 'U';
      'Ã':
        str[i] := 'A';
      'Õ':
        str[i] := 'O';
      'Ñ':
        str[i] := 'N';
      'Ç':
        str[i] := 'C';
    end;
  Result := str;
end;

function RemoveEspeciais(str: String): String;
var
  i: integer;
begin
  for i := 1 to Length(str) do
  begin
    // if (char(str[i]) < 32) and (char(str[i]) > 125) then
    if (ord(str[i]) > 31) and (ord(str[i]) < 126) then
      str[1] := str[1]
    else
      str[i] := ' ';
  end;
  Result := str;
end;

Function WordsCount(S: string): integer;
var
  ps: PChar;
  nSpaces, n, o: integer;

begin
  // total de palavras
  n := 0;
  // total de letras
  o := 0;
  S := S + #0;
  ps := @S[1];

  while (#0 <> ps^) do
  begin
    while ((' ' = ps^) and (#0 <> ps^)) do
    begin
      inc(ps);
      // conta total de letras
      inc(o);
    end;
    nSpaces := 0;
    while ((' ' <> ps^) and (#0 <> ps^)) do
    begin
      inc(nSpaces);
      inc(ps);
      // conta total de letras
      inc(o);
    end;
    if (nSpaces > 0) then
    begin
      inc(n);
    end;
  end;
  // recebe o total de letras contadas incluindo os espacos
  Result := o;
end;

Procedure GravaLog(const usrlog: Variant; const dtlog: TDateTime; const modlog: Variant; const oplog: Variant; const doclog: Variant;
  const idlog: Variant);
begin
end;

Function AnoBiSexto(Ayear: integer): Boolean;
begin
  // Verifica se o ano é Bi-Sexto
  Result := (Ayear mod 4 = 0) and ((Ayear mod 100 <> 0) or (Ayear mod 400 = 0));
end;

Function EspacoStr(Zeros: string; Quant: integer): string;
{ Insere espaços atraz de uma string }
var
  i, Tamanho: integer;
  aux: string;
begin
  aux := Zeros;
  Tamanho := Length(Zeros);
  if Tamanho > Quant then
    EspacoStr := Copy(aux, 1, Quant)
  else
    EspacoStr := Copy(aux, 1, Tamanho) + ReplStr(' ', Quant - Tamanho);
end;

Function ProximoDiaUtil(dData: TDateTime): TDateTime;
begin
  if DayOfWeek(dData) = 7 then
    dData := dData + 2
  else if DayOfWeek(dData) = 1 then
    dData := dData + 1;
  ProximoDiaUtil := dData;
end;

Function FirstDayOfMonth(Data: TDateTime; lSabDom: Boolean): TDateTime;
var
  Ano, Mes, Dia: word;
  DiaDaSemana: integer;
begin
  DecodeDate(Data, Ano, Mes, Dia);
  Dia := 1;

  if lSabDom then
  begin
    DiaDaSemana := DayOfWeek(Data);
    if DiaDaSemana = 1 then
      Dia := 2
    else if DiaDaSemana = 7 then
      Dia := 3;
  end;
  FirstDayOfMonth := EncodeDate(Ano, Mes, Dia);
end;

Function LastDayOfMonth(Data: TDateTime; lSabDom: Boolean): TDateTime;
var
  Ano, Mes, Dia: word;
  AuxData: TDateTime;
  DiaDaSemana: integer;
begin
  AuxData := IncMonth(date, 1);
  AuxData := FirstDayOfMonth(AuxData, false) - 1;

  if lSabDom then
  begin
    DecodeDate(AuxData, Ano, Mes, Dia);
    DiaDaSemana := DayOfWeek(AuxData);

    if DiaDaSemana = 1 then
      Dia := Dia - 2
    else if DiaDaSemana = 7 then
      Dec(Dia);
    AuxData := EncodeDate(Ano, Mes, Dia);
  end;

  LastDayOfMonth := AuxData;
end;

Function DiasPorMes(Ayear, AMonth: integer): integer;
const
  DaysInMonth: array [1 .. 12] of integer = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
begin
  Result := DaysInMonth[AMonth];
  if (AMonth = 2) and AnoBiSexto(Ayear) then
    inc(Result);
end;

Function RemoveChar(Const Texto: string): string;
//
// Remove caracteres de uma string deixando apenas numeros
//
var
  i: integer;
  S: string;
begin
  S := '';
  for i := 1 To Length(Texto) do
  begin
    if (Texto[i] in ['0' .. '9']) then
    begin
      S := S + Copy(Texto, i, 1);
    end;
  end;
  Result := S;
end;

Function Pot(base, expoente: real): real; // Potenciação
begin
  { utiliza a função de exponencial e de logaritmo }
  Result := Exp((expoente * ln(base)));
end;

Function Arredondar(Valor: Double; Dec: integer): Double;
var
  Valor1, Numero1, Numero2, Numero3: Double;
begin
  Valor1 := Exp(ln(10) * (Dec + 1));
  Numero1 := Int(Valor * Valor1);
  Numero2 := (Numero1 / 10);
  Numero3 := Round(Numero2);
  Result := (Numero3 / (Exp(ln(10) * Dec)));
end;

Function VlrStr(Numero: string; tam: integer): string;
{ Insere Zeros e decimais à frente de uma string }
var
  NSTR, Zeros: string;
  X: integer;
begin
  Zeros := ReplChar('0', tam);
  for X := 1 to Length(Numero) do
  begin
    if (Copy(Numero, X, 1) = '.') or (Copy(Numero, X, 1) = ',') then
    begin
    end
    else
    begin
      NSTR := NSTR + Copy(Numero, X, 1);
    end;
  end;
  NSTR := Copy(Zeros, 1, tam - Length(NSTR)) + NSTR;
  Result := NSTR;
end;

Function ValidaCarteira(Numero: string; Parametro: string): Boolean;
var
  i, Acumula, Resto, nCont: integer;
begin
  Resto := 0;
  Result := true;
  if Parametro = 'AMIL' then { Verificador Amil }
  begin
    nCont := 9;
    for i := 1 to Length(Numero) - 1 do
    begin
      Acumula := Acumula + StrToInt(Copy(Numero, i, 1)) * nCont;
      nCont := nCont - 1;
      if nCont = 1 then
        nCont := 9;
    end;
    Resto := 11 - (Acumula mod 11);
    Resto := IIf(Resto > 9, 0, Resto);
    if Resto = StrToInt(Copy(Numero, Length(Numero), 1)) then
      Result := true
    else
      Result := false;
  end;
end;

Function ValidaData(const S: string): Boolean;
begin
  try
    StrToDate(S);
    Result := true;
  except
    Result := false;
  end;
end;

Function VerifLpt(lptStr: string): string;
var
  portHandle: integer;
begin
  portHandle := 0;
  portHandle := CreateFile(PChar(lptStr), GENERIC_READ or GENERIC_WRITE, 0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

  if portHandle > 0 then
    Result := 'ON'
  else

    Result := 'OFF';
end;


// ##############################################################################
// Troca a virgula pelo ponto em um valor Float
// ##############################################################################

Function IIf(pCond: Boolean; pTrue, pFalse: Variant): Variant;
begin
  if pCond then
    Result := pTrue
  else
    Result := pFalse;
end;

// Retorna o digito no modulo 11 do numero passado   *
Function Dac11(const D11_Number: string; const Operador: integer): string;
var
  D11_Resto, D11_Somatoria, D11_Digito, Multiplicador, D11_X: integer;
begin
  D11_Resto := 0;
  D11_Somatoria := 0;
  D11_Digito := 0;
  Multiplicador := 0;

  // Executa a multiplicacao de cada digito e acumula o total  da somatoria

  Multiplicador := Operador;
  for D11_X := Length(D11_Number) downto 1 do
  begin
    D11_Somatoria := D11_Somatoria + (StrToInt(Copy(D11_Number, D11_X, 1)) * Multiplicador);
    if Multiplicador = 9 then
      Multiplicador := 2
    else
      inc(Multiplicador);
  end;

  D11_Resto := D11_Somatoria mod 11;
  D11_Digito := 11 - D11_Resto;
  if not((D11_Digito > 1) And (D11_Digito < 10)) then
    D11_Digito := 1;
  Result := InttoStr(D11_Digito);
end;

// DAC10 - Retorna o digito no modulo 10 do numero passado   *
Function Dac10(const D10_Number: string; const Operador: integer): string;
var
  D10_Resto, D10_Somatoria, D10_Digito, Multiplicador, Conta, D10_X: integer;
begin
  D10_Resto := 0;
  D10_Somatoria := 0;
  D10_Digito := 0;
  Multiplicador := 0;
  Conta := 0;

  // Executa a multiplicacao de cada digito e acumula o total  da somatoria *
  Multiplicador := 2; // Operador ;
  for D10_X := Length(D10_Number) Downto 1 do
  begin
    Conta := StrToInt(Copy(D10_Number, D10_X, 1)) * Multiplicador;

    if (Length(Trim(InttoStr(Conta))) = 1) then
      D10_Somatoria := D10_Somatoria + Conta
    else
      D10_Somatoria := D10_Somatoria + (StrToInt(Copy(Trim(InttoStr(Conta)), 1, 1)) + StrToInt(Copy(Trim(InttoStr(Conta)), 2, 1)));
    Conta := 0;
    Multiplicador := IIf(Multiplicador = 1, 2, 1);

  end;

  D10_Resto := D10_Somatoria mod 10;
  D10_Digito := IIf(D10_Resto = 0, 0, 10 - D10_Resto);
  Result := InttoStr(D10_Digito);
end;

Procedure APrint(const S: Variant; const tip: string; const Len: integer; const wl: integer; const esp: integer);
var
  TXT: string;
begin
  // styl = 0: Normal; 1:Negrito ; 2: Comp;

  if (tip = 'S') then
    TXT := ReplStr(S, Len);
  if (tip = 'I') then
    TXT := StrZero(S, Len);
  if (tip = 'D') then
    TXT := formatdatetime('dd/mm/yyyy', S);
  if (tip = 'F') then
    TXT := formatfloat('#.###.###.##0,00', S);
  if (tip = 'Q') then
    TXT := formatfloat('#.###.###.##0,000', S); // TRES CASAS DECIMAIS

  if wl = 0 then // 0 nao pula linha, 1 pula linha
  begin
    write(FIMP, TXT);
    if esp > 0 then
      write(FIMP, ReplStr('', esp));
  end
  else
  begin
    writeln(FIMP, TXT);
    linhaatual := linhaatual + 1;
  end;
end;

Procedure PrintMat(const ln: integer; const cl: integer; const S: Variant; const tip: string; const esp: integer; const tc: string);
var // ln=Linha  cl=Coluna   s=O que vai ser impresso
  TXT, LinhaT: string; // tip=Tipo do dados (S,I,D,F) esp=Quando 1 imprime e salta linha
  nCont, tt, tl: integer; // tc=Caracter de impressao
begin
  nCont := 1;
  tl := Length(linha);
  tt := Length(S);
  LinhaT := '';
  if (tip = 'S') then
    TXT := ReplStr(S, tt);
  if (tip = 'I') then
    TXT := StrZero(S, tt);
  if (tip = 'D') then
    TXT := formatdatetime('dd/mm/yyyy', S);
  if (tip = 'F') then
    TXT := formatfloat('#.###.###.##0,00', S);
  if (tip = 'Q') then
    TXT := formatfloat('#.###.###.##0,000', S);

  if ln > 0 then
  begin
    AssignFile(FIMP, '/LPT1');
    Rewrite(FIMP);

    // if (tc = 'A') then
    // Write(FIMP,#27+#18) ;
    // if (tc = 'B') then
    // Write(FIMP,#27+#16) ;
    // if (tc = 'C') then
    // Write(FIMP,#27+#15) ;
    writeln(FIMP, linha);
    Write(FIMP, #27 + #64);

    linha := '';
    linha := ReplStr(' ', cl);
    linha := linha + TXT;
    linhaatual := linhaatual + 1;
    for nCont := 1 to ln do
    begin
      linhaatual := linhaatual + 1;
      writeln(FIMP, ' ');
    end;
    CloseFile(FIMP);
  end
  else
  begin
    if cl > tl then
      linha := linha + ReplStr(' ', cl - tl);
    if cl < tl then
    begin
      LinhaT := linha;
      linha := Copy(LinhaT, 1, cl);
    end;
    linha := linha + TXT + Copy(LinhaT, cl + tt + 1, tl - (cl + tt));
  end;
  if esp = 1 then
  begin
    AssignFile(FIMP, '/LPT1');
    Rewrite(FIMP);

    // if (tc = 'A') then
    // Write(FIMP,#27+#18) ;
    // if (tc = 'B') then
    // Write(FIMP,#27+#16) ;
    // if (tc = 'C') then
    // Write(FIMP,#27+#15) ;
    writeln(FIMP, linha);
    Write(FIMP, #27 + #64);

    CloseFile(FIMP);
    linha := '';

  end;
end;

Procedure MPrint(const S: Variant; const tip: string; const Len: integer; const wl: integer; const esp: integer);
var
  TXT: string;
begin
  // styl = 0: Normal; 1:Negrito ; 2: Comp;

  if (tip = 'S') then
    TXT := ReplStr(S, Len);
  if (tip = 'I') then
    TXT := StrZero(S, Len);
  if (tip = 'D') then
    TXT := formatdatetime('dd/mm/yyyy', S);
  if (tip = 'F') then
  begin
    TXT := formatfloat('#,###,###,##0.00', S);
    TXT := ReplStr('', Len - Length(Trim(TXT))) + TXT;
  end;
  if (tip = 'Q') then // tres casas decimais
  begin
    TXT := formatfloat('#,###,###,##0.000', S);
    TXT := ReplStr('', Len - Length(Trim(TXT))) + TXT;
  end;
  if wl = 0 then // 0 nao pula linha, 1 pula linha
  begin
    if esp > 0 then
      write(FIMP, ReplStr('', esp));
    write(FIMP, TXT);
  end
  else
  begin
    writeln(FIMP, TXT);
    linhaatual := linhaatual + 1;
  end;
end;

Function ReplChar(const Ch: Char; const Len: integer): string;
var
  i: integer;
begin
  SetLength(Result, Len);
  for i := 1 to Len do
    Result[i] := Ch;
end;

Function ReplStr(const S: string; const Len: integer): string;
var
  espaco: string;
begin
  espaco := '                                                            ';
  if Length(S) < Len then
    Result := S + Copy(espaco, 1, Len - (Length(S)))
  else
    Result := Copy(S, 1, Len);

end;

Function PrinterOnLine: Boolean;
Const
  PrnStInt: Byte = $17;
  StRq: Byte = $02;
  PrnNum: word = 0; { 0 para LPT1, 1 para LPT2, etc. }
var
  nResult: Byte;
begin (* PrinterOnLine *)
  Asm
    mov ah,StRq;
    mov dx,PrnNum;
    Int $17;
    mov nResult,ah;
  end;
  PrinterOnLine := (nResult and $80) = $80;
end;

Procedure MouseParaControle(Controle: TControl);
var
  IrPara: TPoint;
begin
  IrPara.X := Controle.Left + (Controle.Width div 2);
  IrPara.Y := Controle.Top + (Controle.Height div 2);
  if Controle.Parent <> nil then
    IrPara := Controle.Parent.ClientToScreen(IrPara);
  SetCursorPos(IrPara.X, IrPara.Y);
end;

Function StrZero(Zeros: string; Quant: integer): string;
{ Insere Zeros à frente de uma string }
var
  i, Tamanho: integer;
  aux: string;
begin
  aux := Zeros;
  Tamanho := Length(Zeros);
  Zeros := '';
  for i := 1 to Quant - Tamanho do
    Zeros := Zeros + '0';
  aux := Zeros + aux;
  StrZero := aux;
end;

Function StrEspaco(Zeros: string; Quant: integer): string;
{ Insere Zeros à frente de uma string }
var
  i, Tamanho: integer;
  aux: string;
begin
  aux := Zeros;
  Tamanho := Length(Zeros);
  Zeros := '';
  for i := 1 to Quant - Tamanho do
    Zeros := Zeros + ' ';
  aux := Zeros + aux;
  StrEspaco := aux;
end;

Function StrToChar(str: string): Char;
var
  A: integer;
begin
  if Length(str) > 0 then
  begin
    if (str[1] = '#') and (Length(str) > 1) then
    begin
      try
        A := StrToInt(Copy(str, 2, Length(str) - 1));
      except
        A := 0;
      end;
      Result := Chr(Byte(A));
    end
    else
      Result := str[1];
  end
  else
    Result := #0;
end;

Function StrToPChar(const str: string): PChar;
{ Converte string em Pchar }
type
  TRingIndex = 0 .. 7;
var
  Ring: array [TRingIndex] of PChar;
  RingIndex: TRingIndex;
  Ptr: PChar;
begin
  Ptr := @str[Length(str)];
  inc(Ptr);
  if Ptr^ = #0 then
  begin
    Result := @str[1];
  end
  else
  begin
    Result := StrAlloc(Length(str) + 1);
    RingIndex := (RingIndex + 1) mod (High(TRingIndex) + 1);
    StrPCopy(Result, str);
    StrDispose(Ring[RingIndex]);
    Ring[RingIndex] := Result;
  end;
end;

Function AjustaStr(str: string; tam: integer): string;
begin
  while Length(str) < tam do
    str := str + ' ';
  if Length(str) > tam then
    str := Copy(str, 1, tam);
  Result := str;
end;

Function MAQTOINT(VMAQ: string): integer;
var
  i, tot: integer;
begin
  tot := 0;
  for i := 1 To Length(VMAQ) do
    tot := tot + word(PChar(Copy(VMAQ, 1, i)));
  Result := tot;
end;

Function Padr(S: string; n: integer): string;
{ alinha uma string à direita }
begin
  Result := Format('%' + InttoStr(n) + '.' + InttoStr(n) + 's', [S]);
end;

Function TrocaVirgPPto(Valor: string): string;
var
  i: integer;
begin
  if Valor <> '' then
  begin
    for i := 0 to Length(Valor) do
    begin
      if Valor[i] = ',' then
      begin
        Valor[i] := '.';
      end;
    end;
  end;
  Result := Valor;
end;

Procedure EscondeTaskBar(Visible: Boolean);
var
  wndHandle: THandle;
  wndClass: array [0 .. 50] of Char;
begin
  StrPCopy(@wndClass[0], 'Shell_TrayWnd');
  wndHandle := FindWindow(@wndClass[0], nil);
  if Visible = true then
  begin
    ShowWindow(wndHandle, SW_RESTORE); { Mostra a barra de tarefas }
  end
  else
  begin
    ShowWindow(wndHandle, SW_HIDE); { Esconde a barra de tarefas }
  end;
end;

//
// Retorna a diferença em Dias,Meses e Anos entre 2 datas
//
Procedure EntreDatas(DataFinal, DataInicial: TDate; var Anos, Meses, Dias: integer);
  Function Calcula(Periodo: integer): integer;
  var
    intCont: integer;
  begin
    intCont := 0;
    Repeat
      inc(intCont);
      DataFinal := IncMonth(DataFinal, Periodo * -1);
    Until DataFinal < DataInicial;
    DataFinal := IncMonth(DataFinal, Periodo);
    inc(intCont, -1);
    Result := intCont;
  end;

begin
  if DataFinal <= DataInicial then
  begin
    Anos := 0;
    Meses := 0;
    Dias := 0;
    Exit;
  end;
  Anos := Calcula(12);
  Meses := Calcula(1);
  Dias := Round(DataFinal - DataInicial);
  if Dias < 0 then
    Dias := 0;
end;



// IDADE ATUAL
// ##############################################################################

Function IdadeAtual(Nasc: TDate): integer;
var
  AuxIdade, Meses: string;
  MesesFloat: real;
  IdadeInc, IdadeReal: integer;
begin
  AuxIdade := Format('%0.2f', [(date - Nasc) / 365.6]);
  Meses := FloatToStr(Frac(StrToFloat(AuxIdade)));

  if AuxIdade = '0' then
  begin
    Result := 0;
    Exit;
  end;

  if Meses[1] = '-' then
  begin
    Meses := FloatToStr(StrToFloat(Meses) * -1);
  end;

  Delete(Meses, 1, 2);

  if Length(Meses) = 1 then
  begin
    Meses := Meses + '0';
  end;

  if (Meses <> '0') And (Meses <> '') then
  begin
    MesesFloat := Round(((365.6 * StrToInt(Meses)) / 100) / 30.47)
  end
  else
  begin
    MesesFloat := 0;
  end;

  if MesesFloat <> 12 then
  begin
    IdadeReal := Trunc(StrToFloat(AuxIdade)); // + MesesFloat;
  end
  else
  begin
    IdadeInc := Trunc(StrToFloat(AuxIdade));
    inc(IdadeInc);
    IdadeReal := IdadeInc;
  end;

  Result := IdadeReal;
end;

// ##############################################################################
// Troca a virgula pelo ponto ( Funcao 2 )
// ##############################################################################

Function TrocaVirgulaporPonto(aValue: Double): string;
var
  S: string;
begin
  S := FloatToStr(aValue);
  if Pos(',', S) > 0 then
  begin
    S[Pos(',', S)] := '.';
    Result := S;
  end
  else
    Result := FloatToStr(aValue);
end;


// ##############################################################################
// Funcao VerificaLabel
// ##############################################################################

Procedure Verifica_Label(F: TForm);
var
  i: integer;
begin
  for i := 0 TO F.ComponentCount - 1 do
  begin
    if (F.Components[i] IS TLabel) AND (TLabel(F.Components[i]).Tag = 0) then
    begin
      TLabel(F.Components[i]).Font.Size := 8;
      TLabel(F.Components[i]).Font.Style := [fsBold];
      TLabel(F.Components[i]).Font.Name := 'MS Sans Serif';
    end;
  end;
end;


// ##############################################################################
// Funcao FechaQuery
// ##############################################################################

Procedure FechaQuery(F: TForm);
var
  i: integer;
begin
  for i := 0 to F.ComponentCount - 1 do
  begin
    if (F.Components[i] is TOraQuery) and (TOraQuery(F.Components[i]).Active = true) then
      TOraQuery(F.Components[i]).Close;
  end;
end;

// ##############################################################################
// Funcao Atribui Databaname a Query
// ##############################################################################

procedure AtribuiDbName(F: TForm; DbName: TOraSession);
var
  i: integer;
begin
  for i := 0 to F.ComponentCount - 1 do
  begin
    if (F.Components[i] is TOraQuery) then
    begin
      if (TOraQuery(F.Components[i]).Active = true) then
        TOraQuery(F.Components[i]).Close;
      if (TOraQuery(F.Components[i]).Connection = nil) then
        TOraQuery(F.Components[i]).Connection := DbName;
    end;
  end;
end;

procedure AtribuiSessionForm(F: TForm; SessionName: TOraSession);
var
  i: integer;
begin
  for i := 0 to F.ComponentCount - 1 do
  begin
    if (F.Components[i] is TOraQuery) then
    begin
      if (TOraQuery(F.Components[i]).Active = true) then
        TOraQuery(F.Components[i]).Close;
      if (TOraQuery(F.Components[i]).Session = nil) then
        TOraQuery(F.Components[i]).Session := SessionName;
      TOraQuery(F.Components[i]).Session := SessionName;
    end;
  end;
end;

procedure AtribuiSessionDmd(F: TDataModule; SessionName: TOraSession);
var
  i: integer;
begin
  for i := 0 to F.ComponentCount - 1 do
  begin
    if (F.Components[i] is TOraQuery) then
    begin
      if (TOraQuery(F.Components[i]).Active = true) then
        TOraQuery(F.Components[i]).Close;
      if (TOraQuery(F.Components[i]).Session = nil) then
        TOraQuery(F.Components[i]).Session := SessionName;
      TOraQuery(F.Components[i]).Session := SessionName;
    end;
  end;
end;

// ##############################################################################
// Funcao Limpa_Tela
// ##############################################################################

Procedure Limpa_Tela(F: TForm);
var
  i: integer;
begin
  for i := 0 TO F.ComponentCount - 1 do
  begin
    if (F.Components[i] IS TEdit) AND (TEdit(F.Components[i]).Tag = 0) then
      TEdit(F.Components[i]).Text := '';

    if (F.Components[i] IS TMemo) AND (TMemo(F.Components[i]).Tag = 0) then
      TMemo(F.Components[i]).Text := '';

    if (F.Components[i] IS TComboBox) AND (TComboBox(F.Components[i]).Tag = 0) then
      TComboBox(F.Components[i]).Text := '';

    if (F.Components[i] IS TMaskEdit) AND (TMaskEdit(F.Components[i]).Tag = 0) then
      TMaskEdit(F.Components[i]).Text := '';
  end;
end;

// ##############################################################################
// Funcao AbreQUERY
// ##############################################################################

Procedure AbreQUERY(F: TForm);
var
  w: integer;
begin
  for w := 0 to F.ComponentCount - 1 do
  begin;
    if F.Components[w] is TOraQuery then
    begin
      if TOraQuery(F.Components[w]).Tag = 1 then
      begin
        TOraQuery(F.Components[w]).Open;
      end
    end
  end;
end;

// ##############################################################################
// Funcao AcertaData
// ##############################################################################

Function AcertaData(strData: string; intSQLServer: integer): Variant;
begin
  if intSQLServer = 1 then
  begin
    // if Sistema_Operacional = 'Windows NT' then
    // Result := strData
    // else
    Result := Copy(strData, 4, 3) + Copy(strData, 1, 3) + Copy(strData, 7, 4);
  end
  else
    Result := StrToDateTime(strData);
end;


// ##############################################################################
// Funcao AcertaData 1
// ##############################################################################

Function AcertaData1(strData: string; intSQLServer: integer): Variant;
begin
  if intSQLServer = 1 then
  begin
    // if Sistema_Operacional = 'Windows NT' then
    // Result := strData
    // else
    Result := Copy(strData, 4, 3) + Copy(strData, 1, 3) + Copy(strData, 9, 2);
  end
  else
    Result := StrToDateTime(strData);
end;

// ##############################################################################
// Funcao Retorna_Dia_Semana
// ##############################################################################

Function Retorna_Dia_Semana(dtdata: TDateTime): string;
var
  AData: TDateTime;
  days: array [1 .. 7] of string;
begin
  days[1] := 'Domingo';
  days[2] := 'Segunda';
  days[3] := 'Terça';
  days[4] := 'Quarta';
  days[5] := 'Quinta';
  days[6] := 'Sexta';
  days[7] := 'Sabado';
  AData := dtdata;
  Result := days[DayOfWeek(AData)];
end;

// ##############################################################################
// Funcao Encripta
// ##############################################################################

Function Encripta(Senha: string): string;
var
  strRetorno: string;
  intTamanho: integer;
  i: integer;
begin
  intTamanho := Length(Senha);
  strRetorno := Senha;
  for i := 1 to intTamanho do
    try
      strRetorno[i] := Chr(ord(Senha[i]) + 10 - i);
    except
      // Nada a fazer;
    end;
  Result := strRetorno;
end;


// ##############################################################################
// Funcao Desencripta
// ##############################################################################

Function Desencripta(Senha: string): string;
var
  strRetorno: string;
  intTamanho: integer;
  i: integer;
begin
  intTamanho := Length(Trim(Senha));
  strRetorno := Senha;
  for i := 1 to intTamanho do
    try
      strRetorno[i] := Chr(ord(Senha[i]) - 10 + i);
    except
      // Nada a Fazer;
    end;
  Result := strRetorno;
end;


// ##############################################################################
// Funcao CalculaAnos
// ##############################################################################

Function CalculaAnos(Data_de_Nascimento, Data_Base: TDateTime): integer;
var
  strAno: string;
  strMes: string;
  strDia: string;
  strAno1: string;
  strDia1: string;
  strMes1: string;
  intDias: integer;
  intAnos: integer;
begin
  if (Data_Base < Data_de_Nascimento) or (Data_de_Nascimento = 0) or (Data_Base = 0) then
    Result := -1
  else
  begin
    strAno := formatdatetime('yyyy', Data_de_Nascimento);
    strMes := formatdatetime('mm', Data_de_Nascimento);
    strDia := formatdatetime('dd', Data_de_Nascimento);
    strAno1 := formatdatetime('yyyy', Data_Base);
    strMes1 := formatdatetime('mm', Data_Base);
    strDia1 := formatdatetime('dd', Data_Base);
    intAnos := StrToInt(strAno1) - StrToInt(strAno);
    if strMes1 = strMes then
    begin
      if strDia1 < strDia then
        intAnos := intAnos - 1
    end
    else if strMes1 < strMes then
      intAnos := intAnos - 1;
    if strMes1 > strMes then
      intDias := Trunc(StrToDateTime(strDia1 + '/' + strMes1 + '/' + strAno1) - StrToDateTime(strDia + '/' + strMes + '/' + strAno1))
    else if strMes1 = strMes then
    begin
      if strDia1 >= strDia then
        intDias := StrToInt(strDia1) - StrToInt(strDia)
      else
        intDias := Trunc(StrToDateTime(strDia1 + '/' + strMes1 + '/' + InttoStr(StrToInt(strAno1) - 1)) -
          StrToDateTime(strDia + '/' + strMes + '/' + strAno1))
    end
    else
      intDias := Trunc(StrToDateTime(strDia1 + '/' + strMes1 + '/' + InttoStr(StrToInt(strAno1) - 1)) -
        StrToDateTime(strDia + '/' + strMes + '/' + strAno1));
    Result := intAnos;
  end;
end;


// ##############################################################################
// Funcao Gerapercentual
// ##############################################################################

Function Gerapercentual(Valor: real; Percent: real): real;
begin
  Percent := Percent / 100;
  try
    Valor := Valor * Percent;
  finally
    Result := Valor;
  end;
end;

// ##############################################################################
// Funcao Maiuscula
// ##############################################################################

Function Maiuscula(Texto: string): string;
var
  OldStart: integer;
begin
  if Texto <> '' then
  begin
    Texto := UpperCase(Copy(Texto, 1, 1)) + LowerCase(Copy(Texto, 2, Length(Texto)));
    Result := Texto;
  end;
end;


// ##############################################################################
// Funcao CGC
// ##############################################################################

Function CalculaCnpjCpf(Numero: string): Boolean;
var
  i, d, b, Digito: Byte;
  soma: integer;
  CNPJ: Boolean;
  DgPass, DgCalc: string;
begin
  Result := false;
  Numero := ApenasNumerosStr(Numero);
  // Caso o número não seja 11 (CPF) ou 14 (CNPJ), aborta
  Case Length(Numero) of
    11:
      CNPJ := false;
    14:
      CNPJ := true;
  else
    Exit;
  end;
  // Separa o número do digito
  DgCalc := '';
  DgPass := Copy(Numero, Length(Numero) - 1, 2);
  Numero := Copy(Numero, 1, Length(Numero) - 2);
  // Calcula o digito 1 e 2
  for d := 1 to 2 do
  begin
    b := IIf(d = 1, 2, 3); // BYTE
    soma := IIf(d = 1, 0, STRTOINTDEF(DgCalc, 0) * 2);
    for i := Length(Numero) downto 1 do
    begin
      soma := soma + (ord(Numero[i]) - ord('0')) * b;
      inc(b);
      if (b > 9) And CNPJ then
        b := 2;
    end;
    Digito := 11 - soma mod 11;
    if Digito >= 10 then
      Digito := 0;
    DgCalc := DgCalc + Chr(Digito + ord('0'));
  end;
  Result := DgCalc = DgPass;
end;

Function ApenasNumerosStr(pStr: string): string;
var
  i: integer;
begin
  Result := '';
  for i := 1 To Length(pStr) do
    if pStr[i] In ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'] then
      Result := Result + pStr[i];
end;



// ##############################################################################
// Funcao Erro
// ##############################################################################

Procedure Erro;
begin
  MessageDlg('Você está tentando apagar um registro que contém registros filhos', mtWarning, [mbOK], 0);
  abort;
end;



// ##############################################################################
// Funcao TestaCampoBranco
// ##############################################################################

Function Tiraponto(pStr: string): string;
var
  i: integer;
begin
  Result := '';
  for i := 1 To Length(pStr) do
    if (Copy(pStr, i, 1) = '.') or (Copy(pStr, i, 1) = '-') then
    else
      Result := Result + pStr[i];
end;

Function TestaCampoBranco(DataSet: TDataSet): Boolean;
var
  i: integer;
begin
  with DataSet do
  begin
    for i := 0 to FieldCount - 1 do
    begin
      if ((Fields[i].Tag = 1) and ((Fields[i].IsNull) or (Fields[i].AsString = '') or (Fields[i].AsString = '0'))) then
      begin
        MessageDlg('O Campo ' + Fields[i].DisplayLabel + ' não pode estar em branco !!', mtError, [mbOK], 0);
        Fields[i].FocusControl;
        abort;
      end;
    end;
  end;
end;

// ##############################################################################
// Funcao StringAsPChar ( Transforma string em pchar )
// ##############################################################################
Function StringAsPChar(var S: string): PChar;
{ This Function null-terminates a string so that it can be passed to functions }
{ that require PChar types. if string is longer than 254 chars, then it will }
{ be truncated to 254. }
begin
  if Length(S) = 255 then
    S := Copy(S, 1, 254);
  S[ord(Length(S)) + 1] := #0; { Place null at end of string }
  StringAsPChar := @S[1]; { Return "PChar'd" string }
end;



// ##############################################################################
// Funcao FormJaExiste
// ##############################################################################

Function FormJaExiste(PForm: TForm): Boolean;
var
  i: integer;
begin
  Result := false;
  for i := 0 to Screen.FormCount - 1 do
  begin
    if Screen.Forms[i] = PForm then
    begin
      Result := true;
      Break;
    end;
  end;
end;

// ##############################################################################
// Funcao Tratamento_Erro
// ##############################################################################

Procedure Tratamento_Erro(DataSet: TDataSet; E: EOraError; var Action: TDataAction);
begin
  if (E is EOraError) then
  begin
    if (E as EOraError).Errorcode = 9729 then
      MessageDlg('Registro Duplicados não permitidos', mtWarning, [mbOK], 0)

    else if (E as EOraError).Errorcode = 9730 then
      MessageDlg('Este Registro está sendo usado em outra tabela', mtWarning, [mbOK], 0)

    else if (E as EOraError).Errorcode = 9732 then
      MessageDlg('É necessário colocar algum valor no campo', mtError, [mbOK], 0)

    else if (E as EOraError).Errorcode = 9733 then
      MessageDlg('Este Registro está sendo usado em outra tabela', mtWarning, [mbOK], 0);
    abort;
  end;
end;

// ##############################################################################
// Funcao VerificaFORM
// ##############################################################################

Procedure VerificaFORM(F: TForm);
var
  E: integer;
begin
  // if f.Active = false then
  // Abort
  if F.Active = true then
    F.WindowState := wsMaximized
end;


// ##############################################################################
// Funcao Retorna_Versao
// ##############################################################################

Function Retorna_Versao: string;
var
  VerInfoSize: DWORD;
  VerInfo: Pointer;
  VerValueSize: DWORD;
  VerValue: PVSFixedFileInfo;
  Dummy: DWORD;
  v1, v2, V3, V4: word;
  Prog: string;
begin
  Prog := Application.Exename;
  VerInfoSize := GetFileVersionInfoSize(PChar(Prog), Dummy);
  GetMem(VerInfo, VerInfoSize);
  GetFileVersionInfo(PChar(Prog), 0, VerInfoSize, VerInfo);
  VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
  with VerValue^ do
  begin
    v1 := dwFileVersionMS shr 16;
    v2 := dwFileVersionMS and $FFFF;
    V3 := dwFileVersionLS shr 16;
    V4 := dwFileVersionLS and $FFFF;
  end;
  FreeMem(VerInfo, VerInfoSize);
  Result := InttoStr(StrToInt(Copy(InttoStr(100 + v1), 2, 2))) + '.' + InttoStr(StrToInt(Copy(InttoStr(100 + v2), 2, 2))) + '.' +
    InttoStr(StrToInt(Copy(InttoStr(100 + V3), 2, 2))) + '.' + InttoStr(StrToInt(Copy(InttoStr(100 + V4), 2, 2)));
end;

// ##############################################################################
// Retorna a diferença entre duas horas
// ##############################################################################

Function DifHora(Inicio, Fim: string): string;
{ Retorna a diferença entre duas horas }
var
  FIni, FFim: TDateTime;
begin
  FIni := StrTotime(Inicio);
  FFim := StrTotime(Fim);
  if (Inicio > Fim) then
  begin
    Result := TimeToStr((StrTotime('23:59:59') - FIni) + FFim)
  end
  else
  begin
    Result := TimeToStr(FFim - FIni);
  end;
end;

Function PegaTamanhoPapel(dmPaperSize: word): string;
begin
  Result := 'Desconhecido';
  // Verifica ALGUNS TAMANHOS POSSÍVEIS. Existem outros, veja DEVMODE
  case dmPaperSize of
    DMPAPER_USER:
      Result := 'Definido pelo usuário';
    DMPAPER_LETTER:
      Result := 'Letter, 8 1/2- by 11-inches';
    DMPAPER_LEGAL:
      Result := 'Legal, 8 1/2- by 14-inches';
    DMPAPER_A4:
      Result := 'A4 Sheet, 210- by 297-millimeters';
  end;
end;

Procedure MudaTamPapel(PaperSize, Comp, Alt: integer);
var
  ADevice, ADriver, APort: array [0 .. 255] of Char;
  DeviceMode: THandle;
  M: PDevMode;
  S: string;
begin
  // Força o uso de Printer. Se esta linha for removida, a primeira
  // invocação falha. Bug da VCL
  S := Printer.PRINTERS[Printer.PrinterIndex];
  // Pega dados da impressora atual
  Printer.GetPrinter(ADevice, ADriver, APort, DeviceMode);
  // Pega um ponteiro para DEVMODE
  M := GlobalLock(DeviceMode);
  try
    if M <> nil then
    begin
      // Muda tamanho do papel
      M^.dmFields := DM_PAPERSIZE;

      if PaperSize = DMPAPER_USER then
        M^.dmFields := M^.dmFields or DM_PAPERLENGTH or DM_PAPERWIDTH;
      M^.dmPaperLength := Alt;
      M^.dmPaperWidth := Comp;
      M^.dmPaperSize := PaperSize; //
      // Atualiza
      Printer.SetPrinter(ADevice, ADriver, APort, DeviceMode);
    end;

  finally
    GlobalUnlock(DeviceMode);
  end;
end;

Function GetOrCreateObject(const ClassName: string): IDispatch; // para usar o word
var
  ClassID: TGUID;
  Unknown: IUnknown;
begin
  ClassID := ProgIDToClassID(ClassName);
  if Succeeded(GetActiveObject(ClassID, nil, Unknown)) then
    OleCheck(Unknown.QueryInterface(IDispatch, Result))
  else
    Result := CreateOleObject(ClassName);
end;

Function FilterChars(const S: string; const ValidChars: TChars): string;
var
  i: integer;
begin
  Result := '';
  for i := 1 to Length(S) do
    if S[i] in ValidChars then
      Result := Result + S[i];
end;

Function Justifica(mCad: string; mMAx: integer): string;
var
  mPos, mPont, mTam, mNr, mCont: integer;
  mStr: string;
begin
  mTam := Length(mCad);
  if mTam >= mMAx then
    Result := Copy(mCad, 1, mMAx)
  else
    mStr := '';
  mCont := 0;
  mPont := 1;
  mNr := mMAx - mTam;
  while mCont < mNr do
  begin
    mPos := Pos(mStr, Copy(mCad, mPont, 100));
    if mPos = 0 then
    begin
      mStr := mStr + ' ';
      mPont := 1;
      continue;
    end
    else
    begin
      mCont := mCont + 1;
      Insert(' ', mCad, mPos + mPont);
      mPont := mPont + mPos + Length(mStr);
    end;
    Result := mCad;
  end;
end;

Function Letra_linha(ln: integer): string;
var
  nlinha: array [1 .. 92] of string;
begin
  nlinha[1] := 'AA';
  nlinha[24] := 'BA';
  nlinha[47] := 'CA';
  nlinha[70] := 'DA';
  nlinha[2] := 'AB';
  nlinha[25] := 'BB';
  nlinha[48] := 'CB';
  nlinha[71] := 'DB';
  nlinha[3] := 'AC';
  nlinha[26] := 'BC';
  nlinha[49] := 'CC';
  nlinha[72] := 'DC';
  nlinha[4] := 'AD';
  nlinha[27] := 'BD';
  nlinha[50] := 'CD';
  nlinha[73] := 'DD';
  nlinha[5] := 'AE';
  nlinha[28] := 'BE';
  nlinha[51] := 'CE';
  nlinha[74] := 'DE';
  nlinha[6] := 'AF';
  nlinha[29] := 'BF';
  nlinha[52] := 'CF';
  nlinha[75] := 'DF';
  nlinha[7] := 'AG';
  nlinha[30] := 'BG';
  nlinha[53] := 'CG';
  nlinha[76] := 'DG';
  nlinha[8] := 'AH';
  nlinha[31] := 'BH';
  nlinha[54] := 'CH';
  nlinha[77] := 'DH';
  nlinha[9] := 'AI';
  nlinha[32] := 'BI';
  nlinha[55] := 'CI';
  nlinha[78] := 'DI';
  nlinha[10] := 'AJ';
  nlinha[33] := 'BJ';
  nlinha[56] := 'CJ';
  nlinha[79] := 'DJ';
  nlinha[11] := 'AL';
  nlinha[34] := 'BL';
  nlinha[57] := 'CL';
  nlinha[80] := 'DL';
  nlinha[12] := 'AM';
  nlinha[35] := 'BM';
  nlinha[58] := 'CM';
  nlinha[81] := 'DM';
  nlinha[13] := 'AN';
  nlinha[36] := 'BN';
  nlinha[59] := 'CN';
  nlinha[82] := 'DN';
  nlinha[14] := 'AO';
  nlinha[37] := 'BO';
  nlinha[60] := 'CO';
  nlinha[83] := 'do';
  nlinha[15] := 'AP';
  nlinha[38] := 'BP';
  nlinha[61] := 'CP';
  nlinha[84] := 'DP';
  nlinha[16] := 'AQ';
  nlinha[39] := 'BQ';
  nlinha[62] := 'CQ';
  nlinha[85] := 'DQ';
  nlinha[17] := 'AR';
  nlinha[40] := 'BR';
  nlinha[63] := 'CR';
  nlinha[86] := 'DR';
  nlinha[18] := 'AS';
  nlinha[41] := 'BS';
  nlinha[64] := 'CS';
  nlinha[87] := 'DS';
  nlinha[19] := 'AT';
  nlinha[42] := 'BT';
  nlinha[65] := 'CT';
  nlinha[88] := 'DT';
  nlinha[20] := 'AU';
  nlinha[43] := 'BU';
  nlinha[66] := 'CU';
  nlinha[89] := 'DU';
  nlinha[21] := 'AV';
  nlinha[44] := 'BV';
  nlinha[67] := 'CV';
  nlinha[90] := 'DV';
  nlinha[22] := 'AX';
  nlinha[45] := 'BX';
  nlinha[68] := 'CX';
  nlinha[91] := 'DX';
  nlinha[23] := 'AZ';
  nlinha[46] := 'BZ';
  nlinha[69] := 'CZ';
  nlinha[92] := 'DZ';
  Result := nlinha[(ln)];
end;

Function soma(v1, v2: real): real;
begin
  Result := v1 + v2;
end;

Function MULTI(v1, v2: real): real;
begin
  Result := v1 * v2;
end;

Function DIVE(v1, v2: real): real;
begin
  Result := v1 / v2;
end;

Function DIMI(v1, v2: real): real;
begin
  Result := v1 - v2;
end;

Function NetSend(dest, source, msg: string): longint;
type
  TNetMessageBufferSendFunction = Function(servername, msgname, fromname: PWideChar; buf: PWideChar; buflen: Cardinal): longint; stdcall;
var
  NetMessageBufferSend: TNetMessageBufferSendFunction;
  SourceWideChar: PWideChar;
  DestWideChar: PWideChar;
  MessagetextWideChar: PWideChar;
  Handle: THandle;
begin

  Handle := LoadLibrary('NETAPI32.DLL');
  if Handle = 0 then
  begin
    Result := GetLastError;
    Exit;
  end;

  @NetMessageBufferSend := GetProcAddress(Handle, 'NetMessageBufferSend');
  if @NetMessageBufferSend = nil then
  begin
    Result := GetLastError;
    Exit;
  end;

  MessagetextWideChar := nil;
  SourceWideChar := nil;
  DestWideChar := nil;

  try
    GetMem(MessagetextWideChar, Length(msg) * SizeOf(WideChar) + 1);
    GetMem(DestWideChar, 20 * SizeOf(WideChar) + 1);
    StringToWideChar(msg, MessagetextWideChar, Length(msg) * SizeOf(WideChar) + 1);
    StringToWideChar(dest, DestWideChar, 20 * SizeOf(WideChar) + 1);
    if source = '' then
      Result := NetMessageBufferSend(nil, DestWideChar, nil, MessagetextWideChar, Length(msg) * SizeOf(WideChar) + 1)
    else
    begin
      GetMem(SourceWideChar, 20 * SizeOf(WideChar) + 1);
      StringToWideChar(source, SourceWideChar, 20 * SizeOf(WideChar) + 1);
      Result := NetMessageBufferSend(nil, DestWideChar, SourceWideChar, MessagetextWideChar, Length(msg) * SizeOf(WideChar) + 1);
      FreeMem(SourceWideChar);
    end;
  finally
    FreeMem(MessagetextWideChar);
    FreeLibrary(Handle);
  end;
end;

Procedure CriaCodigo(Cod: string; Imagem: TCanvas);

Const
  digitos: array ['0' .. '9'] of string[5] = ('00110', '10001', '01001', '11000', '00101', '10100', '01100', '00011', '10010', '01010');
var
  Numero: string;
  Cod1: Array [1 .. 1000] Of Char;
  Cod2: Array [1 .. 1000] Of Char;
  Codigo: Array [1 .. 1000] Of Char;
  Digito: string;
  c1, c2: integer;
  X, Y, z, h: longint;
  A, b, c, d: TPoint;
  i: Boolean;
begin
  Numero := Cod;
  for X := 1 to 1000 do
  begin
    Cod1[X] := #0;
    Cod2[X] := #0;
    Codigo[X] := #0;
  end;
  c1 := 1;
  c2 := 1;
  X := 1;
  for Y := 1 to Length(Numero) div 2 do
  begin
    Digito := digitos[Numero[X]];
    for z := 1 to 5 do
    begin
      Cod1[c1] := Digito[z];
      inc(c1);
    end;
    Digito := digitos[Numero[X + 1]];
    for z := 1 to 5 do
    begin
      Cod2[c2] := Digito[z];
      inc(c2);
    end;
    inc(X, 2);
  end;
  Y := 5;
  Codigo[1] := '0';
  Codigo[2] := '0';
  Codigo[3] := '0';
  Codigo[4] := '0'; { Inicio do Codigo }
  for X := 1 to c1 - 1 do
  begin
    Codigo[Y] := Cod1[X];
    inc(Y);
    Codigo[Y] := Cod2[X];
    inc(Y);
  end;
  Codigo[Y] := '1';
  inc(Y); { Final do Codigo }
  Codigo[Y] := '0';
  inc(Y);
  Codigo[Y] := '0';
  Imagem.Pen.Width := 1;
  Imagem.Brush.Color := ClWhite;
  Imagem.Pen.Color := ClWhite;
  A.X := 1;
  A.Y := 0;
  b.X := 1;
  b.Y := 79;
  c.X := 2000;
  c.Y := 79;
  d.X := 2000;
  d.Y := 0;
  Imagem.Polygon([A, b, c, d]);
  Imagem.Brush.Color := ClBlack;
  Imagem.Pen.Color := ClBlack;
  X := 0;
  i := true;
  for Y := 1 to 1000 do
  begin
    if Codigo[Y] <> #0 then
    begin
      if Codigo[Y] = '0' then
        h := 1
      else
        h := 3;
      A.X := X;
      A.Y := 0;
      b.X := X;
      b.Y := 79;
      c.X := X + h - 1;
      c.Y := 79;
      d.X := X + h - 1;
      d.Y := 0;
      if i then
        Imagem.Polygon([A, b, c, d]);
      i := Not(i);
      X := X + h;
    end;
  end;
end;

Function SerialNum(Unidade: PChar): string;
{ Retorna o Número serial da unidade especificada }
var
  VolName, SysName: AnsiString;
  SerialNo, MaxCLength, FileFlags: DWORD;
begin
  try
    SetLength(VolName, 255);
    SetLength(SysName, 255);
    GetVolumeInformation(Unidade, PChar(VolName), 255, @SerialNo, MaxCLength, FileFlags, PChar(SysName), 255);
    Result := IntToHex(SerialNo, 8);
  except
    Result := ' ';
  end;
end;

Function SysComputerName: string;
var
  i: DWORD;
begin
  i := MAX_COMPUTERNAME_LENGTH + 1;
  SetLength(Result, i);
  Windows.GetComputerName(PChar(Result), i);
  Result := string(PChar(Result));
end;

Function StrIsDate(const S: string): Boolean;
begin
  try
    StrToDate(S);
    Result := true;
  except
    Result := false;
  end;
end;

Procedure PrintRichEdit(const Caption: string; const RichEdt: TRichEdit);
// Requer a Printers e RichEdit declaradas na clausula uses da unit
var
  Range: TFormatRange;
  LastChar, MaxLen, LogX, LogY, OldMap: integer;
begin
  FillChar(Range, SizeOf(TFormatRange), 0);
  with Printer, Range do
  begin
    BeginDoc;
    hdc := Handle;
    hdcTarget := hdc;
    LogX := GetDeviceCaps(Handle, LOGPIXELSX);
    LogY := GetDeviceCaps(Handle, LOGPIXELSY);
    if IsRectEmpty(RichEdt.PageRect) then
    begin
      rc.right := PageWidth * 1440 div LogX;
      rc.bottom := PageHeight * 1440 div LogY;
    end
    else
    begin
      rc.Left := RichEdt.PageRect.Left * 1440 div LogX;
      rc.Top := RichEdt.PageRect.Top * 1440 div LogY;
      rc.right := RichEdt.PageRect.right * 1440 div LogX;
      rc.bottom := RichEdt.PageRect.bottom * 1440 div LogY;
    end;
    rcPage := rc;
    Title := Caption;
    LastChar := 0;
    MaxLen := RichEdt.GetTextLen;
    chrg.cpMax := -1;
    OldMap := SetMapMode(hdc, MM_TEXT);
    SendMessage(RichEdt.Handle, EM_FORMATRANGE, 0, 0);
    try
      repeat
        chrg.cpMin := LastChar;
        LastChar := SendMessage(RichEdt.Handle, EM_FORMATRANGE, 1, longint(@Range));
        if (LastChar < MaxLen) and (LastChar < -1) then
        begin
          NewPage;
        end;
      until (LastChar = MaxLen) or (LastChar = -1);
      EndDoc;
    finally
      SendMessage(RichEdt.Handle, EM_FORMATRANGE, 0, 0);
      SetMapMode(hdc, OldMap);
    end;
  end;
end;

function MensagemDlg(txtMsg: String): Boolean;
var
  Mensagem: TForm;
begin
  { Cria a janela de mensagem }
  Mensagem := createmessagedialog(txtMsg, MtConfirmation, [MbYes, MbNo]);
  { Trazur o titulo da mensagem }
  Mensagem.Caption := 'Confirmação';
  { Traduz os botões da caixa de mensagem }
  (Mensagem.FindComponent('Yes') as TButton).Caption := 'Confirmar';
  (Mensagem.FindComponent('No') as TButton).Caption := 'Ignorar';
  { Exibr a caixa de mensagem }
  Mensagem.ShowModal;
  { Verifica aqul botão foi pressionado }
  If Mensagem.ModalResult = MrYes then
    Result := true; { Botão Sim }
  If Mensagem.ModalResult = MrNo then
    Result := false;
  { Botão Não }
end;

function InputBoxPass(const ACaption, APrompt, ADefault: string): string;
begin
  Result := ADefault;
  InputSenha(ACaption, APrompt, Result);
end;

function GetAveCharSize(Canvas: TCanvas): TPoint;
var
  i: integer;
  Buffer: array [0 .. 51] of Char;
begin
  for i := 0 to 25 do
    Buffer[i] := Chr(i + ord('A'));
  for i := 0 to 25 do
    Buffer[i + 26] := Chr(i + ord('a'));
  GetTextExtentPoint(Canvas.Handle, Buffer, 52, TSize(Result));
  Result.X := Result.X div 52;
end;

function InputSenha(const ACaption, APrompt: string; var Value: string): Boolean;
var
  Form: TForm;
  Prompt: TLabel;
  Edit: TEdit;
  DialogUnits: TPoint;
  ButtonTop, ButtonWidth, ButtonHeight: integer;
begin
  Result := false;
  Form := TForm.Create(Application);
  with Form do
    try
      Canvas.Font := Font;
      DialogUnits := GetAveCharSize(Canvas);
      BorderStyle := bsDialog;
      Caption := ACaption;
      ClientWidth := MulDiv(180, DialogUnits.X, 4);
      ClientHeight := MulDiv(63, DialogUnits.Y, 8);
      Position := poMainformcenter;
      Prompt := TLabel.Create(Form);
      with Prompt do
      begin
        Parent := Form;
        AutoSize := true;
        Left := MulDiv(8, DialogUnits.X, 4);
        Top := MulDiv(8, DialogUnits.Y, 8);
        Caption := APrompt;
      end;
      Edit := TEdit.Create(Form);
      with Edit do
      begin
        Parent := Form;
        Left := Prompt.Left;
        Top := MulDiv(19, DialogUnits.Y, 8);
        Width := MulDiv(164, DialogUnits.X, 4);
        { } MaxLength := 20;
        { } Passwordchar := '*';
        { } Font.Color := clBlue;
        Text := Value;
        SelectAll;
      end;
      ButtonTop := MulDiv(41, DialogUnits.Y, 8);
      ButtonWidth := MulDiv(50, DialogUnits.X, 4);
      ButtonHeight := MulDiv(14, DialogUnits.Y, 8);
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := 'Ok';
        ModalResult := mrOk;
        Default := true;
        SetBounds(MulDiv(38, DialogUnits.X, 4), ButtonTop, ButtonWidth, ButtonHeight);
      end;
      with TButton.Create(Form) do
      begin
        Parent := Form;
        Caption := 'Cancelar';
        ModalResult := mrCancel;
        Cancel := true;
        SetBounds(MulDiv(92, DialogUnits.X, 4), ButtonTop, ButtonWidth, ButtonHeight);
      end;
      if ShowModal = mrOk then
      begin
        Value := Edit.Text;
        Result := true;
      end;
    finally
      Form.Free;
    end;
end;

Function Esconde(DADOin: string): string;
var
  i: integer;
  DADOout: String;
begin
  Result := DADOin;
  DADOout := '';
  for i := 1 to Length(DADOin) do
  begin
    If Copy(DADOin, i, 1) = ' ' then
      DADOout := DADOout + Chr(7)
    else
      DADOout := DADOout + Chr(ord(DADOin[i]) + 122);
  end;
  Result := DADOout;
end;

Function Desvenda(DADOin: string): string;
var
  i: integer;
  DADOout: String;
begin
  Result := DADOin;
  DADOout := '';
  for i := 1 to Length(DADOin) do
  begin
    If Copy(DADOin, i, 1) = ' ' then
      DADOout := DADOout + Chr(7)
    else
      DADOout := DADOout + Chr(ord(DADOin[i]) - 122);
  end;
  Result := DADOout;
end;

Function BuscaTroca(Text, Busca, Troca: string): string; { Substitui um caractere dentro da string }
var
  n: integer;
begin
  for n := 1 to Length(Text) do
  begin
    if Copy(Text, n, 1) = Busca then
    begin
      Delete(Text, n, 1);
      Insert(Troca, Text, n);
    end;
  end;
  Result := Text;
end;

Function TrocaNaPosicao(Text: string; Posicao: integer; Troca: string): string; { Substitui um caractere dentro da string }
var
  n: integer;
begin
  for n := 1 to Length(Text) do
  begin
    if n = Posicao then
    begin
      Delete(Text, n, 1);
      Insert(Troca, Text, n);
    end;
  end;
  Result := Text;
end;

// *********************************//
// Bibliotecas de banco de dados   //
// *********************************//

function ConverteListaEmStringParaComandoSQL(aLista: TStringList; aIncluirAspas: Boolean = true): String;
var
  Texto: String;
  i: integer;
begin
  /// Jhonny Oliveira
  /// 06/05/2013
  ///
  /// Em diversas situações é necessário utilizar objetos do tipo
  /// TStringList para armazenar filtros para instruções SQL.
  /// Porém é necessário tratar essas strings com aspas e vírgulas.
  /// Essa function faz exatamente isso.

  Texto := '';

  for i := 0 to aLista.Count - 1 do
  begin

    if (aIncluirAspas) then
    begin

      Texto := Texto + Chr(39) + aLista[i] + Chr(39);
    end
    else
    begin

      Texto := Texto + aLista[i];
    end;

    if i < (aLista.Count - 1) then
    begin

      Texto := Texto + ', ';

    end;

  end;

  Result := Texto;

end;

procedure CentralizarPanelNoForm(APanel: TPanel; AForm: TForm);
begin
  /// Jhonny Oliveira - 14/05/2013
  ///
  /// Auxilia na centralização de panel em um determinado form
  /// normalmente será usado para centralizar mensagens de aguarde e etc
  ///

  APanel.Top := Trunc((AForm.Height / 2) - (APanel.Height / 2));
  APanel.Left := Trunc((AForm.Width / 2) - (APanel.Width / 2));

end;

procedure AbrirConexaoDBExpress(ANomeServidor, AUsuarioBD, SenhaBD: String; ACompomenteConnection: TSQLConnection);
begin
  /// Jhonny Oliveira - 17/05/2013
  ///
  /// Responsável por criar uma conexão com o banco de dados
  /// utilizando os valores informados que normalmente
  /// são obtidos como parâmetros enviados do menu do Winthor
  /// ao executável

  try
    begin
      ACompomenteConnection.Connected := false;
      ACompomenteConnection.LoginPrompt := false;
      ACompomenteConnection.KeepConnection := true;
      ACompomenteConnection.DriverName := 'Oracle';
      ACompomenteConnection.GetDriverFunc := 'getSQLDriverORACLE';
      ACompomenteConnection.LibraryName := 'dbxora30.dll';
      ACompomenteConnection.VendorLib := 'oci.dll';

      ACompomenteConnection.Params.Clear;
      ACompomenteConnection.Params.Add('DataBase=' + ANomeServidor);
      ACompomenteConnection.Params.Add('User_Name=' + AUsuarioBD);
      ACompomenteConnection.Params.Add('Password=' + SenhaBD);

      ACompomenteConnection.Connected := true;

    end;
  except
    on E: Exception do
    begin
      raise Exception.Create(E.Message);
    end;
  end;

end;

function DiferencaEntreDatasEmHoras(ADataInicial, ADataFinal: TDateTime): String;
var
  apenasHoras: integer;
  difHoras: String;
  difDias: integer;

begin
  /// Jhonny Oliveira
  /// 15/07/2013
  ///
  /// Retorna a diferença entre duas data em horas,
  /// a function DifHora criada por Alberto mostra apenas
  /// um intervalo dentro de 24 horas.
  /// Essa function mostra realmente a diferença no tempo decorrido:
  ///
  /// Ex.: 113:04:34 (cento e treze horas e quatro minutos e trinta e quatro segundos)
  ///

  difDias := DaysBetween(ADataInicial, ADataFinal);

  difHoras := TimeToStr(ADataInicial - ADataFinal);

  apenasHoras := StrToInt(Copy(difHoras, 1, 2)) + (difDias * 24);

  Result := formatfloat('00', apenasHoras) + Copy(difHoras, 3, 6);

end;

procedure ConectarMySQLComCNC();
var
  arq: TextFile;
  linha: String;
  i: integer;
  Data: TIniFile;
  teste: string;
  Connection: TSQLConnection;
begin
  /// Alberto Nunes - 23/04/2015
  ///
  /// Responsável por estabelecer uma conexão com o banco de dados MSSQL
  /// utilizando o arquivo MSSQL.cnc que deve estar no mesmo diretório
  /// que o executável do programa.
  ///

  Data := TIniFile.Create('.\MYSQL.cnc');

  if not Assigned(Connection) then
    Connection := TSQLConnection.Create(nil);
  try
    Connection.Connected := false;
    Connection.DriverName := 'MySQL';
    Connection.GetDriverFunc := 'getSQLDriverMYSQL50';
    Connection.LibraryName := 'dbxopenmysql50.dll';
    Connection.VendorLib := 'libmysql.dll';
    Connection.KeepConnection := true;
    Connection.LoginPrompt := false;

    Connection.Params.Clear;
    Connection.Params.Append('Database=' + Desvenda(Data.ReadString('PARAMETROS', 'database', '')));
    Connection.Params.Append('User_Name=' + Desvenda(Data.ReadString('PARAMETROS', 'uid', '')));
    Connection.Params.Append('Password=' + Desvenda(Data.ReadString('PARAMETROS', 'pwd', '')));
    Connection.Params.Append('HostName=' + Desvenda(Data.ReadString('PARAMETROS', 'server', '')));

    Connection.Open;

    // Connection.Free;

    // BDEDatabase.Connected := true;
  except
    on E: Exception do
    begin
      Connection.Free;
      raise Exception.Create(E.Message);
    end;
  end;

end;

// function LoginColetor() : TUsuario;
// var
// Usuario : TUsuario;
// begin
/// Jhonny Oliveira - 09/09/2013
///
/// Até agora, em todo programa que criamos para execução
/// no coletor criamos forms de login.
///
/// Esta function faz esse trabalho, valida o usuário e senha
/// do usuário no sistema Winthor e retorna um objeto do tipo
/// TUsuario, neste objeto temos acesso a diversas informações
/// como login, nome, matricula, filiais que o usuário possui
/// acesso e etc.
///
/// Para o uso correto dessa function você deve importar também
/// o arquivo UUsuariosWinthor da pasta LIB
///

// Application.CreateForm(TFrmLogin, FrmLogin);
// FrmLogin.ShowModal;
//
// if FrmLogin.UsuarioValido then
// begin
//
// Usuario := FrmLogin.Usuario;
//
// end
// else
// begin
//
// Usuario := nil;
//
// end;
//
//
// Result := Usuario;
//
//
// end;

function PadLeft(AStringAtual: String; ACaracterACompletar: Char; ATamanhoTotal: integer): String;
var
  qtdCompletar: integer;
  stringSaida: String;
  caracter: Char;
  i: integer;

begin
  // Jhonny Oliveira
  // 07/04/2014

  // Consideramos apenas a parte da string que seja menor
  // ou igual ao tamanho informado
  stringSaida := AnsiMidStr(Trim(AStringAtual), 1, ATamanhoTotal);

  qtdCompletar := ATamanhoTotal - Length(stringSaida);

  caracter := ACaracterACompletar;

  for i := 0 to qtdCompletar - 1 do
  begin

    stringSaida := caracter + stringSaida;

  end;

  Result := stringSaida;

end;

function PadRight(AStringAtual: String; ACaracterACompletar: Char; ATamanhoTotal: integer): String;
var
  qtdCompletar: integer;
  stringSaida: String;
  caracter: Char;
  i: integer;

begin
  // Jhonny Oliveira
  // 07/04/2014

  // Consideramos apenas a parte da string que seja menor
  // ou igual ao tamanho informado
  stringSaida := AnsiMidStr(Trim(AStringAtual), 1, ATamanhoTotal);

  qtdCompletar := ATamanhoTotal - Length(stringSaida);

  caracter := ACaracterACompletar;

  for i := 0 to qtdCompletar - 1 do
  begin

    stringSaida := stringSaida + caracter;

  end;

  Result := stringSaida;

end;

function AbrirManualUsuario(ACodigoRotina: String; ANumeroVersao: String): Boolean;
var
  diretorio: String;
  nome_arquivo: String;
  arquivo_encontrado: Boolean;

begin
  /// Jhonny Oliveira
  /// 31/07/2014
  ///
  /// Procura e abre o manual de usuário
  /// do programa e versão informados.
  ///
  /// O nome do manual deve começar com o número da rotina
  /// seguinte de underline ( _ ) e o número da versão sem os pontos.
  ///
  /// O arquivo deve estar no formato PDF.
  ///
  ///
  ///

  nome_arquivo := ACodigoRotina + '_' + ApenasNumerosStr(ANumeroVersao) + '.pdf';
  diretorio := ObtemConfiguracao('0', 86); // Parametrização do local dos arquivos

  arquivo_encontrado := FileExists(diretorio + '\' + nome_arquivo);

  if (arquivo_encontrado) then
  begin

    try
      ShellExecute(Application.Handle, nil, StrToPChar(diretorio + '\' + nome_arquivo), nil, nil, SW_SHOWNORMAL);

    except
      on E: Exception do
        arquivo_encontrado := false;

    end;

  end;

  Result := arquivo_encontrado;

end;

function ObtemConfiguracao(ACodfilial: String; ACodConfiguracao: Double): String;
var
  qryConfigPorFilial, qryConfigGlobal: TOraQuery;
  valorConfig: string;

begin
  /// Alberto Nunes - 26/08/2019
  /// Permite obter um valor configurado rotina 9827
  /// sendo parametrizado por filial ou não.
  ///

  valorConfig := '';

  qryConfigPorFilial := TOraQuery.Create(nil);
  qryConfigPorFilial.Session := ODACSessionGlobal;

  with qryConfigPorFilial do
  begin

    Close;

    SQL.Clear;
    SQL.Add(' SELECT boconfigdetalhe.valor                 ');
    SQL.Add(' FROM boconfigdetalhe                         ');
    SQL.Add(' WHERE boconfigdetalhe.filial = :CODFILIAL    ');
    SQL.Add(' AND boconfigdetalhe.codigo = :CODCONFIG      ');

    ParamByName('CODFILIAL').AsString := ACodfilial;
    ParamByName('CODCONFIG').AsFloat := ACodConfiguracao;

    Open;

  end;

  if qryConfigPorFilial.RecordCount > 0 then
  begin

    Result := qryConfigPorFilial.FieldByName('VALOR').AsString;
    FreeAndNil(qryConfigPorFilial);
    Exit;
  end;

  FreeAndNil(qryConfigPorFilial);

  qryConfigGlobal := TOraQuery.Create(nil);
  qryConfigGlobal.Session := ODACSessionGlobal;

  with qryConfigGlobal do
  begin

    Close;

    SQL.Clear;
    SQL.Add(' SELECT boconfig.valor FROM boconfig ');
    SQL.Add(' WHERE boconfig.codigo = :CODCONFIG  ');

    ParamByName('CODCONFIG').AsFloat := ACodConfiguracao;

    Open;
  end;

  if (qryConfigGlobal.RecordCount > 0) then
  begin

    valorConfig := qryConfigGlobal.FieldByName('VALOR').AsString;
  end;

  FreeAndNil(qryConfigGlobal);

  Result := valorConfig;

end;

function ObtemConfiguracaoFloat(ACodfilial: String; ACodConfiguracao: Double): double;
var
  valor_string: string;
  valor_float: double;
begin

  valor_string := ObtemConfiguracao(ACodfilial, ACodConfiguracao);

  if TryStrToFloat(valor_string, valor_float) then
  begin

    Result := valor_float;
    Exit;
  end;

  Result := 0;
end;

function EnviarEmail(AEnderecoHost, ANomeUsuario, ASenha, AEmailRemetente, ANomeRemetente, AEmailParaResposta, ACorpoEmail, AAssunto: String;
  ADestinatarios: TStringList; var ARespostaServidor: String): Boolean;
var
  smtp: TIdSMTP; // Responsável por conectar ao servidor SMTP
  email: TIdMessage; // Encapsula os componentes de um e-mail como destinatários,
  // corpo do e-mail e etc.

  sucesso: Boolean;
  i: integer;
  anexo: TIdAttachment;

begin

  /// Jhonny Oliveira
  /// Permite o envio de e-mails usando a biblioteca Indy

  smtp := TIdSMTP.Create(nil);
  email := TIdMessage.Create(nil);
  email.date := IncHour(now, 3);

  smtp.Port := 587;
  smtp.Host := LowerCase(AEnderecoHost);
  smtp.AuthType := {$IFDEF VER330}satDefault{$ELSE}satDefault{$ENDIF};
  smtp.Username := LowerCase(ANomeUsuario);
  smtp.Password := ASenha;

  ARespostaServidor := 'E-mail enviado com sucesso';
  sucesso := true;

  try
    begin

      try
        begin

          smtp.Connect;

          if not(smtp.Authenticate) then
          begin

            sucesso := false;
            ARespostaServidor := 'Não foi possível conectar ao servidor';
            Raise Exception.Create('Não foi possível conectar ao servidor');

          end;

          email.From.Address := AEmailRemetente;
          email.From.Name := ANomeRemetente;
          email.Subject := AAssunto;
          email.Body.Text := ACorpoEmail;
          email.ReplyTo.EMailAddresses := AEmailParaResposta;
          email.ContentType := 'text/html';
          // email.ContentType   := 'multipart/mixed';

          for i := 0 to ADestinatarios.Count - 1 do
          begin

            email.Recipients.Add.Address := ADestinatarios[i];

          end;

          smtp.Send(email);
          smtp.Disconnect;

        end;
      except
        on E: Exception do
        begin

          sucesso := false;
          ARespostaServidor := 'Erro ao conectar ao servidor:'#13 + E.Message;

        end;
      end;

    end;
  finally
    begin

      FreeAndNil(email);
      FreeAndNil(smtp);

    end;

  end;

  Result := sucesso;

end;

procedure SalvarLayoutDosGridsDoForm(AFormulario: TForm);
var
  pasta_executavel, padrao_nome_layout: String;
  i: integer;
  gridComum: TcxGridDBTableView;
  gridBanded: TcxGridDBBandedTableView;
begin

  /// Jhonny Oliveira - 01/12/2014
  /// Salvando o layout de todos os grids da DevExpress do formulário indicado
  pasta_executavel := ExcludeTrailingBackslash(ExtractFilePath(Application.Exename)) + '\grids\';

  ForceDirectories(pasta_executavel);

  // Utiliza os dois primeiros números da versão para encontrar o layout.
  // Isto por que esses dois números indicam grandes alterações de layout na rotina
  // como um todo.
  padrao_nome_layout := pasta_executavel + ParamStr(5) + '_v' + Copy(Retorna_Versao, 1, 3) + '_' + ParamStr(1) + '_';

  for i := 0 to AFormulario.ComponentCount - 1 do
  begin

    if (AFormulario.Components[i] is TcxGridDBTableView) then
    begin

      gridComum := AFormulario.Components[i] as TcxGridDBTableView;
      gridComum.StoreToIniFile(padrao_nome_layout + gridComum.Name + '.ini');

    end;

    if (AFormulario.Components[i] is TcxGridDBBandedTableView) then
    begin

      gridBanded := AFormulario.Components[i] as TcxGridDBBandedTableView;
      gridBanded.StoreToIniFile(padrao_nome_layout + gridBanded.Name + '.ini');

    end;
  end;

end;

procedure RestaurarLayoutDosGridsDoForm(AFormulario: TForm);
var
  i: integer;
  gridComum: TcxGridDBTableView;
  gridBanded: TcxGridDBBandedTableView;
  pasta_executavel: String;
  padrao_nome_layout: String;

begin

  /// Jhonny Oliveira - 01/12/2014
  /// Restaurando o layout de todos os grids da DevExpress do formulário indicado
  pasta_executavel := ExcludeTrailingBackslash(ExtractFilePath(Application.Exename)) + '\grids\';

  // Utiliza os dois primeiros números da versão para encontrar o layout.
  // Isto por que esses dois números indicam grandes alterações de layout na rotina
  // como um todo.
  padrao_nome_layout := pasta_executavel + ParamStr(5) + '_v' + Copy(Retorna_Versao, 1, 3) + '_' + ParamStr(1) + '_';

  for i := 0 to AFormulario.ComponentCount - 1 do
  begin

    if (AFormulario.Components[i] is TcxGridDBTableView) then
    begin

      gridComum := AFormulario.Components[i] as TcxGridDBTableView;
      gridComum.RestoreFromIniFile(padrao_nome_layout + gridComum.Name + '.ini');
    end;

    if (AFormulario.Components[i] is TcxGridDBBandedTableView) then
    begin

      gridBanded := AFormulario.Components[i] as TcxGridDBBandedTableView;
      gridBanded.RestoreFromIniFile(padrao_nome_layout + gridBanded.Name + '.ini');
    end;

  end;
end;

function VerificaVersao(ACodigoRotina: Double; AExibirMensagens: Boolean = true): Boolean;
var
  qry: TOraQuery;
  versaoOficial, versaoRotina: String;

begin
  {
    Alberto Nunes
    09/06/2019

    Retorna se a versão atual do programa é a mesma definida
    no campo NUMVERSAOATUAL da tabela PCROTINA, este
    campo foi criado por Marcos Pereira justamente
    para impedir execuções de rotinas desatualizadas.

    O ideal é que essa verificação seja feita antes do carregamento do
    form principal e depois da chamada da função de conexão.

    Parâmetros:
    ACodigoRotina    = Código da rotina a ser testada
    AExibirMensagens = Indica se a verificação deve ser silenciosa ou não,
    ou seja, se devemos exibir mensagens de erro e alerta

  }

  Result := true;

  try
    begin

      /// Obtendo a versão atual do executável.
      ///
      /// Caso o programador não deixar marcado a opção
      /// Project/Option/Version Info/Include version information in project
      ///
      /// A function Retorna_Versao lança uma exceção, por isso tratamos
      /// dentro de um Try..Except
      versaoRotina := Retorna_Versao();

    end;
  except
    on E: Exception do
    begin

      /// Erro na leitura do número da versão do executável

      Result := false;

      if (AExibirMensagens) then
      begin

        Application.MessageBox(StrToPChar('Não é possível obter o número da versão do programa'), 'ERRO', MB_OK + MB_ICONERROR);
        Application.Terminate;
      end;

    end;
  end;

  /// Se a conexão com o banco não estiver estabelecida
  /// não é possível verificar o número da versão oficial
  ///
  if (not Assigned(ODACSessionGlobal)) then
  begin

    Result := false;

    if (AExibirMensagens) then
    begin

      Application.MessageBox(StrToPChar('Não é possível obter o número da versão do programa' + #13 + 'Verifique a conexão com o banco de dados'),
        'ERRO', MB_OK + MB_ICONERROR);

      // Finalizando a rotina
      Application.Terminate;
      Exit;
    end;

  end;

  /// Consultando a base de dados
  qry := TOraQuery.Create(nil);
  qry.Session := ODACSessionGlobal;
  qry.SQL.Text := ' SELECT NUMVERSAOATUAL FROM PCROTINA WHERE CODIGO = :CODROTINA ';
  qry.ParamByName('CODROTINA').AsFloat := ACodigoRotina;
  qry.Open;

  if (qry.RecordCount = 0) or (Trim(qry.FieldByName('NUMVERSAOATUAL').AsString) = '') then
  begin

    Result := false;

    Exit;
  end;

  versaoOficial := Trim(qry.FieldByName('NUMVERSAOATUAL').AsString);

  if (versaoRotina <> versaoOficial) then
  begin

    Result := false;

    if (AExibirMensagens) then
    begin

      Application.MessageBox(StrToPChar('A versão da rotina no seu computador é diferente da versão oficial:' + #13 + #13 +
        'Versão no seu computador: ' + versaoRotina + #13 + 'Versão oficial: ' + versaoOficial + #13 + #13 +
        'Abra um chamado no Ocomon relatando o problema'), 'ERRO', MB_OK + MB_ICONERROR);

      // Finalizando a rotina
      Application.Terminate;
      Exit;

    end;

  end;

end;

function IniciaOS(AnumOS: Double; Amatricula: Double; AColetor: Boolean = false): Boolean;
var
  qry: TOraQuery;
Begin
  {

    Marcos Pereira
    30/12/2014
    Essa função pode iniciar uma OS, seja qual for o tipo dela, com base no último
    parâmetro 'AColetor', que se for falso, que é o seu valor padrão, indica que a OS
    não será manipulada com coletor, por isso grava data de início e usuário, caso
    contrário, se for verdadeiro, grava a data de liberação, que indica que a OS pode
    ser manipulada com o coletor que é quem irá gravar a data de início e o usuário
    para cada linha da OS.

  }

  qry := TOraQuery.Create(nil);
  qry.Session := ODACSessionGlobal;

  if not AColetor then
  begin

    qry.SQL.Add('Update pcmovendpend');
    qry.SQL.Add('Set dtinicioos =SYSDATE');
    qry.SQL.Add('    ,codfuncos = :CODFUNCOS');
    qry.SQL.Add('Where numos=:NUMOS');

    qry.ParamByName('NUMOS').AsFloat := AnumOS;
    qry.ParamByName('CODFUNCOS').AsFloat := Amatricula;

  end
  else
  Begin

    qry.SQL.Add('Update pcmovendpend');
    qry.SQL.Add('Set dataliberacao =SYSDATE');
    qry.SQL.Add('Where numos=:NUMOS');

    qry.ParamByName('NUMOS').AsFloat := AnumOS;

  end;

  qry.ExecSql;

  Result := qry.RowsAffected > 0;

  FreeAndNil(qry);
end;

function FechaOS(AnumOS: Double; AtipoOS: integer { 1-Separação, 2-Recebimento, 3-Transferência }
  ; Amatricula: Double): Boolean;
var
  qry: TOraQuery;
Begin

  {
    Marcos Pereira
    30/12/2014
    Essa função pode finalizar a separação de uma OS, seja qual for o tipo dela,
    com base no parâmetro 'AtipoOS', que define se a OS é de Recebimento,
    Abastecimento ou Separação.

  }

  qry := TOraQuery.Create(nil);
  qry.Session := ODACSessionGlobal;

  if AtipoOS = 1 then // OS de separação
  Begin

    qry.SQL.Add('Update pcmovendpend             ');
    qry.SQL.Add('Set dtfimseparacao =SYSDATE     ');
    qry.SQL.Add('Where numos=:NUMOS              ');
    qry.SQL.Add('      and dtfimseparacao is null');

    qry.ParamByName('NUMOS').AsFloat := AnumOS;

  end;

  if AtipoOS = 2 then // OS de Recebimento
  Begin

    qry.SQL.Add('Update pcmovendpend              ');
    qry.SQL.Add('Set posicao = ''A''              ');
    qry.SQL.Add('    ,dtfimos = SYSDATE           ');
    qry.SQL.Add('    ,codfuncosfim = :CODFUNCOSFIM');
    qry.SQL.Add('Where numos=:NUMOS               ');

    qry.ParamByName('NUMOS').AsFloat := AnumOS;
    qry.ParamByName('CODFUNCOSFIM').AsFloat := Amatricula;

  end;

  if AtipoOS = 3 then // OS de Abastecimento
  Begin

    qry.SQL.Add('Update pcmovendpend              ');
    qry.SQL.Add('Set posicao = ''A''              ');
    qry.SQL.Add('    ,dtfimos = SYSDATE           ');
    qry.SQL.Add('    ,codfuncosfim = :CODFUNCOSFIM');
    qry.SQL.Add('Where numos=:NUMOS               ');

    qry.ParamByName('NUMOS').AsFloat := AnumOS;
    qry.ParamByName('CODFUNCOSFIM').AsFloat := Amatricula;

  end;

  qry.ExecSql;

  Result := qry.RowsAffected > 0;

  FreeAndNil(qry);

end;

function ExportarExcel(AGrid: TcxGrid; ANomeArquivo: String = ''; AExibirMensagem: Boolean = true): Boolean;
var
  save_dialog: TSaveDialog;

begin
  {
    Jhonny Oliveira
    15/01/2015

    Facilita a exportação do conteúdo dos grids
    da devExpress para Excel.

    É necessário incluir na cláusula uses
    as bibliotecas cxGrid e cxGridExportLink

    Parâmetros:
    AGrid:
    Grid com o conteúdo a ser exportado


    ANomeArquivo (Opcional):
    Nome do arquivo a ser gerado


    AExibirMensagem (Padrão=True):
    Como esta function pode ser utilizado em um loop por exemplo,
    podemos indicar se deve ou não mostrar mensagens de diálogo
    ao usuário.
    De qualquer forma a function retorna true se o arquivo
    for gerado com sucesso.
  }

  Result := false;

  save_dialog := TSaveDialog.Create(nil);
  save_dialog.DefaultExt := '.xls';
  save_dialog.Filter := 'Arquivos .xls|*.xls';
  save_dialog.FileName := ANomeArquivo;

  if (save_dialog.Execute) then
  begin

    try
      begin

        ExportGridToExcel(save_dialog.FileName, AGrid);

        if (AExibirMensagem) then
        begin

          Application.MessageBox(StrToPChar('Arquivo salvo com sucesso'), 'Sucesso', MB_OK + MB_ICONINFORMATION);

          if (Application.MessageBox(StrToPChar('Deseja abrir o arquivo?'), 'Atenção', MB_YESNO + MB_ICONQUESTION) = IDYES) then
          begin

            ShellExecute(Application.Handle, StrToPChar('open'), StrToPChar(save_dialog.FileName), nil, nil, SW_SHOWMAXIMIZED);

          end;
        end;

        Result := true;

      end;

    except
      on E: Exception do
      begin

        if (AExibirMensagem) then
        begin

          Application.MessageBox(StrToPChar('Erro ao salvar o arquivo'), 'ERRO', MB_OK + MB_ICONERROR);
        end;
      end;
    end;

  end;

  FreeAndNil(save_dialog);

end;

function Modulo10(Valor: String): string;
{
  Rotina usada para cálculo de alguns dígitos verificadores
  Pega-se cada um dos dígitos contidos no parâmetro VALOR, da direita para a
  esquerda e multiplica-se por 2121212...
  Soma-se cada um dos subprodutos. Caso algum dos subprodutos tenha mais de um
  dígito, deve-se somar cada um dos dígitos. (Exemplo: 7*2 = 14 >> 1+4 = 5)
  Divide-se a soma por 10.
  Faz-se a operação 10-Resto da divisão e devolve-se o resultado dessa operação
  como resultado da função Modulo10.
  Obs.: Caso o resultado seja maior que 9, deverá ser substituído por 0 (ZERO).
}
var
  Auxiliar: string;
  Contador, Peso: integer;
  Digito: integer;
begin
  Valor := RemoveChar(Valor);
  Auxiliar := '';
  Peso := 2;
  for Contador := Length(Valor) downto 1 do
  begin
    Auxiliar := InttoStr(StrToInt(Valor[Contador]) * Peso) + Auxiliar;
    if Peso = 1 then
      Peso := 2
    else
      Peso := 1;
  end;

  Digito := 0;
  for Contador := 1 to Length(Auxiliar) do
  begin
    Digito := Digito + StrToInt(Auxiliar[Contador]);
  end;
  Digito := 10 - (Digito mod 10);
  if (Digito > 9) then
    Digito := 0;
  Result := InttoStr(Digito);

end;

(*
  function Modulo11(Valor: String; Base: Integer = 9; Resto : boolean = false) : string;
  {
  Rotina muito usada para calcular dígitos verificadores
  Pega-se cada um dos dígitos contidos no parâmetro VALOR, da direita para a
  esquerda e multiplica-se pela seqüência de pesos 2, 3, 4 ... até BASE.
  Por exemplo: se a base for 9, os pesos serão 2,3,4,5,6,7,8,9,2,3,4,5...
  Se a base for 7, os pesos serão 2,3,4,5,6,7,2,3,4...
  Soma-se cada um dos subprodutos.
  Divide-se a soma por 11.
  Faz-se a operação 11-Resto da divisão e devolve-se o resultado dessa operação
  como resultado da função Modulo11.
  Obs.: Caso o resultado seja maior que 9, deverá ser substituído por 0 (ZERO).
  }
  var
  Soma : integer;
  Contador, Peso, Digito : integer;
  begin

  Valor := RemoveChar(Valor);
  Soma := 0;
  Peso := 2;
  for Contador := Length(Valor) downto 1 do
  begin
  Soma := Soma + (StrToInt(Valor[Contador]) * Peso);
  if Peso < Base then
  Peso := Peso + 1
  else
  Peso := 2;
  end;

  if Resto then
  Result := IntToStr(Soma mod 11)
  else
  begin
  Digito := 11 - (Soma mod 11);
  if (Digito > 9) then
  Digito := 0;
  Result := IntToStr(Digito);
  end;

  end;
*)
function Modulo11(n: string): integer;
VAR
  soma, X, Pn: integer;
begin
  X := 2;
  soma := 0;
  For Pn := Length(n) downto 1 do
  begin
    soma := soma + (StrToInt(n[Pn]) * X);
    inc(X);
    IF X = 10 Then
      X := 2;
  end;

  Result := 11 - (soma mod 11);
  IF Result > 9 Then
    IF Length(n) = 43 Then
      Result := 1
    else
      Result := 0;
end;

function calcula_barra(linha: string): string;
var
  barra, campo1, campo2, campo3, campo4, campo5, valida, lixo: string;
begin
  // var linha = form.linha.value;	// Linha Digitável
  barra := Trim(RemoveChar(linha));
  //
  // CÁLCULO DO DÍGITO DE AUTOCONFERÊNCIA (DAC)   -   5ª POSIÇÃO
  if (Modulo11('34191000000000000001753980229122525005423000') <> 1) then
    showmessage('Função "modulo11_banco" está com erro!');

  // validar contas publicas
  if Copy(barra, 1, 1) = '8' then // segundo numero = 1=prefeituras
  begin // 2=saneamento
    if Length(barra) = 48 then // 3=energia e gas
    begin // 4=telecom
      barra := Copy(barra, 1, 11) // 5=orgaos governo
        + Copy(barra, 13, 11) // 6=carnes s/cnpj
        + Copy(barra, 25, 11) // 7=multas transito
        + Copy(barra, 37, 11);
      Result := barra;
    end
    else
    begin
      showmessage('Representação Numérica de Contas Públicas deve conter 48 números');
      Result := '';
    end;
  end
  else
  begin
    if Length(barra) < 47 then
      barra := barra + StrZero('0', 47 - Length(barra)); // '00000000000'.substr(0,47-barra.length);

    if Length(barra) <> 47 then
      showmessage('A linha do código de barras está incompleta!' + InttoStr(Length(barra)));

    barra := Copy(barra, 1, 4) + Copy(barra, 33, 15) + Copy(barra, 5, 5) + Copy(barra, 11, 10) + Copy(barra, 22, 10);
    // 23791647000097090003390090001163917200002440
    valida := Copy(barra, 1, 4) + Copy(barra, 6, 39);
    // lixo:= modulo11(valida);
    // lixo:= dac11(valida,9);
    // lixo:= intToStr(FBMod11(valida));
    // if (modulo11(valida) <> Copy(barra,5,1)) then
    // if (dac11(valida,9) <> Copy(barra,5,1)) then
    if InttoStr(Modulo11(valida)) <> Copy(barra, 5, 1) then
    begin

      showmessage('Digito verificador ' + Copy(barra, 5, 1) + ', o correto é ' + InttoStr(Modulo11(valida)) + #13 +
        'O sistema não altera automaticamente o dígito correto na quinta casa!');
      Result := '';

    end
    else
      Result := barra;

  end;

end;

function calcula_linha(barra: string): string;
var
  linha, campo1, campo2, campo3, campo4, campo5, verificar: string;
begin
  // var barra = form.barra.value;	// Codigo da Barra
  linha := barra; // place(/[^0-9]/g,'');
  if Length(linha) <> 44 then
    showmessage('A linha do código de barras está incompleta!');

  verificar := Copy(linha, 1, 4) + Copy(linha, 6, 39);

  // arrecadacoes
  if Copy(linha, 1, 1) = '8' then
  begin
    { if copy(linha,2,1) = '1' then
      lbltipo.Caption := 'ARRECADAÇOES PREFEITURAS';
      if copy(linha,2,1) = '2' then
      lbltipo.Caption := 'CONTAS DE AGUA';
      if copy(linha,2,1) = '3' then
      lbltipo.Caption := 'CONTAS DE LUZ/GAS';
      if copy(linha,2,1) = '4' then
      lbltipo.Caption := 'CONTAS DE TELEFONE';
      if copy(linha,2,1) = '5' then
      lbltipo.Caption := 'ARRECADAÇOES DO GOVERNO';
      if copy(linha,2,1) = '6' then
      lbltipo.Caption := 'ARRECADAÇOES ORGAOS DIVERSOS';
      if copy(linha,2,1) = '7' then
      lbltipo.Caption := 'MULTAS DE TRANSITO';
      if copy(linha,2,1) = '9' then
      lbltipo.Caption := 'USO EXCLUSIVO DO BANCO';
    }
    campo1 := Copy(linha, 1, 11);
    campo2 := Copy(linha, 12, 11);
    campo3 := Copy(linha, 23, 11);
    campo4 := Copy(linha, 34, 11);

    linha := campo1 + Modulo10(campo1) + ' ' + campo2 + Modulo10(campo2) + ' ' + campo3 + Modulo10(campo3) + ' ' + campo4 + Modulo10(campo4);
    Result := linha;
  end
  else // boleto bancario
  begin
    // lbltipo.Caption := 'BOLETOS BANCARIO';
    //
    campo1 := Copy(linha, 1, 4) + Copy(linha, 20, 1) + '.' + Copy(linha, 21, 4);
    campo2 := Copy(linha, 25, 5) + '.' + Copy(linha, 25 + 5, 5);
    campo3 := Copy(linha, 35, 5) + '.' + Copy(linha, 35 + 5, 5);
    campo4 := Copy(linha, 5, 1); // Digito verificador
    campo5 := Copy(linha, 6, 14); // Vencimento + Valor
    //
    // showmessage(dac11(verificar,2));
    // if (  modulo11( verificar ) <> campo4 ) then
    if (Dac11(verificar, 2) <> campo4) then
    begin

      showmessage('Digito verificador ' + campo4 + ', o correto é ' + Dac11(verificar, 2) + #13 +
        'O sistema não altera automaticamente o dígito correto na quinta casa!');
      Result := '';

    end
    else
    begin

      linha := campo1 + Modulo10(campo1) + '  ' + campo2 + Modulo10(campo2) + '  ' + campo3 + Modulo10(campo3) + '  ' + campo4 + '  ' + campo5;
      Result := linha;

    end;

  end;

end;

function ValidadordeVersao(ACodigoRotina: string; ANumeroVersao: String): Boolean;
var
  qryRotinas: TOraQuery;
Begin

  qryRotinas := TOraQuery.Create(nil);
  qryRotinas.Session := ODACSessionGlobal;

  with qryRotinas do
  begin

    Close;

    SQL.Clear;
    SQL.Add(' SELECT nvl(pcrotina.NUMVERSAOATUAL,'''') as NUMVERSAOATUAL        ');
    SQL.Add(' FROM pcrotina                                                     ');
    SQL.Add(' WHERE pcrotina.codigo = :CODROTINA                                ');

    ParamByName('CODROTINA').AsFloat := StrToFloat(ACodigoRotina);

    Open;
  end;

  if qryRotinas.RecordCount > 0 then
  Begin

    Result := qryRotinas.FieldByName('NUMVERSAOATUAL').AsString = ANumeroVersao;

  end
  else
  Begin

    Result := false;

  end;

  FreeAndNil(qryRotinas);

end;

function DiaUtil(ACodigoFilial: string; AData: TDateTime): Boolean;
var
  qryDiasUteis: TOraQuery;
Begin

  qryDiasUteis := TOraQuery.Create(nil);
  qryDiasUteis.Session := ODACSessionGlobal;

  with qryDiasUteis do
  begin

    Close;

    SQL.Clear;
    SQL.Add(' SELECT DIAFINANCEIRO AS DIAUTIL       ');
    SQL.Add(' FROM PCDIASUTEIS                      ');
    SQL.Add(' WHERE PCDIASUTEIS.CODFILIAL = :FILIAL ');
    SQL.Add(' AND PCDIASUTEIS.DATA = :DIA           ');

    ParamByName('FILIAL').AsString := ACodigoFilial;
    ParamByName('DIA').AsDate := AData;

    Open;
  end;

  if qryDiasUteis.RecordCount > 0 then
  Begin

    if qryDiasUteis.FieldByName('DIAUTIL').AsString = 'S' then

      Result := true

    else

      Result := false;

  end
  else
  Begin

    Result := false;

  end;

  FreeAndNil(qryDiasUteis);

end;

Procedure GerarCodigo(Codigo: String; Canvas: TCanvas);
// =============================================================================
{ Procedimento para gerar a imagem de código de barras para ser impresso.
  Basta Colocar um componente tImage* no form e chamar o procedimento da seguinte forma:
  GerarCodigo(numero, image1.Canvas);
  GerarCodigo(LblEdtCodigo.Text, ImgCodigoBarras.Canvas);

  *Na verdade você pode desenhar no canvas de qualquer componete,
  inclusive no componente qrImage para posterior impressão. }
const
  digitos: array ['0' .. '9'] of string[5] = ('00110', '10001', '01001', '11000', '00101', '10100', '01100', '00011', '10010', '01010');
var
  S: string;
  i, j, X, t: integer;
begin

  // Gerar o valor para desenhar o código de barras
  // Caracter de início
  S := '0000';
  for i := 1 to Length(Codigo) div 2 do
    for j := 1 to 5 do
      S := S + Copy(digitos[Codigo[i * 2 - 1]], j, 1) + Copy(digitos[Codigo[i * 2]], j, 1);
  // Caracter de fim
  S := S + '100';
  // Desenhar em um objeto canvas
  // Configurar os parâmetros iniciais
  X := 0;
  // Pintar o fundo do código de branco
  Canvas.Brush.Color := ClWhite;
  Canvas.Pen.Color := ClWhite;
  Canvas.Rectangle(0, 0, 2000, 79);
  // Definir as cores da caneta
  Canvas.Brush.Color := ClBlack;
  Canvas.Pen.Color := ClBlack;
  // Escrever o código de barras no canvas
  for i := 1 to Length(S) do
  begin
    // Definir a espessura da barra
    t := StrToInt(S[i]) * 2 + 1;
    // Imprimir apenas barra sim barra não (preto/branco - intercalado);
    if i mod 2 = 1 then
      Canvas.Rectangle(X, 0, X + t, 79);
    // Passar para a próxima barra
    X := X + t;
  end;

end;

function ToString(Value: Variant): String;
begin

  case TVarData(Value).VType of
    varSmallInt, varInteger:
      Result := InttoStr(Value);
    varSingle, varDouble, varCurrency:
      Result := FloatToStr(Value);
    varDate:
      Result := formatdatetime('dd/mm/yyyy', Value);
    varBoolean:
      if Value then
        Result := 'True'
      else
        Result := 'False';
    varString:
      Result := Value;
  else
    Result := '';
  end;

end;

function SoLetraeNumero(Const Texto: string): String;
var
  i: integer;
  S: string;
begin
  S := '';
  for i := 1 to Length(Texto) do
  begin
    if (Texto[i] in ['0' .. '9']) or (Texto[i] in ['a' .. 'z']) or (Texto[i] in ['A' .. 'Z']) or (Texto[i] = ' ') then
    begin
      S := S + Copy(Texto, i, 1);
    end;
  end;
  Result := S;
end;

procedure CopiarArquivo(Origem, Destino: String);
var
  Stream1, Stream2: TFileStream;
begin
  Stream1 := TFileStream.Create(Origem, fmOpenRead);
  try
    Stream2 := TFileStream.Create(Destino, fmOpenwrite or fmCreate);
    try
      Stream2.CopyFrom(Stream1, Stream1.Size);
    finally
      Stream2.Free;
    end;
  finally
    Stream1.Free;
  end;
end;

//
// Rotina para Abreviar Nomes
// Recebe Nome e o Tamanho Máximo do nome e
// retorna o Nome Abreviado (ou truncado caso não seja possível)
//
Function PrimeiroNome(Nome: string): string;
var
  PNome: string;
begin
  PNome := '';
  if Pos(' ', Trim(Nome)) <> 0 then
    PNome := Copy(Trim(Nome), 1, Pos(' ', Trim(Nome)) - 1)
  else
    PNome := Nome;
  Result := PNome;
end;

Function UltimoNome(Nome: string): string;
var
  PNome: string;
  i: integer;
begin
  PNome := Trim(Nome);
  for i := Length(PNome) downto 1 do
  begin
    if PNome[i] = ' ' then
      Break;
  end;

  Result := Trim(Copy(PNome, i + 1, Length(PNome) - i + 1));

end;

function NomeAbreviado(Nome: String; TamanhoMaximo: integer): String;
var
  S: String;
  Nomes: TStringList;
begin
  S := Trim(Nome);
  if Length(S) > TamanhoMaximo then
  begin
    Nomes := TStringList.Create;
    try
      Nomes := SeparaNomes(S, Nomes);
      S := ReduzNome(Nomes, TamanhoMaximo);
    finally
      Nomes.Free;
    end;
  end;
  if Length(S) > TamanhoMaximo then // Trunca caso ainda ultrapasse o Tamanho Máximo
    S := Copy(S, 1, TamanhoMaximo);
  Result := S;
end;

function ObterMaiorNome(Nomes: TStringList): integer;
var
  i, IndMax, TamMax: integer;
begin
  // Ver qual dos Nomes do meio é o maior
  IndMax := 0;
  TamMax := -1;
  if Nomes.Count > 1 then
    for i := 2 to (Nomes.Count - 2) do // Poupa o Primeiro o Segundo e o Ultimo
    begin
      if not((UpperCase(Nomes.Strings[i]) = 'DA') or { }
        (UpperCase(Nomes.Strings[i]) = 'DAS') or (UpperCase(Nomes.Strings[i]) = 'DE') or (UpperCase(Nomes.Strings[i]) = 'DO') or
        (UpperCase(Nomes.Strings[i]) = 'DOS') or (UpperCase(Nomes.Strings[i]) = 'E')) then
        if Length(Nomes.Strings[i]) > TamMax then
        begin
          IndMax := i;
          TamMax := Length(Nomes.Strings[i]);
        end;
    end;
  Result := IndMax;
end;

function ReduzNome(Nomes: TStringList; TamanhoMaximo: integer): String;
var
  S: String;
  i, vezes: integer;
  cont: Boolean;
begin
  // Tenta primeiro abreviar os nomes do meio
  cont := true;
  vezes := 0;
  while (VerTamanhoNome(Nomes) > TamanhoMaximo) and (cont) do
  begin
    i := ObterMaiorNome(Nomes);
    if Length(Nomes.Strings[i]) = 2 then
      Nomes.Strings[i] := ''
    else
      Nomes.Strings[i] := Copy(Nomes.Strings[i], 1, 1) + '.';
    inc(vezes);
    // Sai da rotina caso já tenha passado por todos os nomes do meio
    cont := (vezes <= (Nomes.Count - 2));
  end;
  cont := true;
  // Retira caso necassario os nomes do meio
  while (VerTamanhoNome(Nomes) > TamanhoMaximo) and (cont) do
  begin
    i := ObterMaiorNome(Nomes);
    Nomes.Strings[i] := '';
    inc(vezes);
    // Sai da rotina caso já tenha passado por todos os nomes do meio
    cont := (vezes <= Nomes.Count);
  end;
  // Monta o nome abreviado
  for i := 0 to (Nomes.Count - 1) do
  begin
    if Length(Nomes.Strings[i]) > 0 then
      S := S + Nomes.Strings[i] + ' ';
  end;
  Result := Trim(S);
end;

function SeparaNomes(Nome: String; n: TStringList): TStringList;
var
  S: String;
  i: integer;
begin
  // Quebra o nome em varias strings
  S := Nome;
  while Length(Trim(S)) > 0 do
  begin
    i := Pos(' ', Trim(S));
    if i = 0 then
      i := Length(S);
    n.Add(Trim(Copy(S, 1, i)));
    S := Trim(Copy(S, i + 1, Length(S)));
  end;
  Result := n;
end;

function VerTamanhoNome(Nomes: TStringList): integer;
var
  i, total, espacos: integer;
begin
  // Ver o tamanho total do nome (soma das strings)
  total := 0;
  espacos := 0; // Vai somar os espaços em branco (numero de nomes - 1)
  for i := 0 to (Nomes.Count - 1) do
  begin
    total := total + Length(Trim(Nomes.Strings[i]));
    if Length(Trim(Nomes.Strings[i])) > 0 then // Só nomes com algum conteudo
      inc(espacos);
  end;
  espacos := espacos - 1; // Qts de nomes com cont. - 1
  Result := (total + espacos);
end;

procedure debug(ATexto: String);
var
  memo: TMemo;
  i: integer;
  divisorLinha: String;

begin
  {
    Jhonny Oliveira
    08/10/2015

    Permite a exibição de um form que servirá como debug,
    sempre aparecerá sobre os outros forms.

    Para isso a váriável global habilitarDebug deve estar como true.
  }

  if not habilitarDebug then
  begin

    Exit;
  end;

  if (not Assigned(FrmDebug)) then
  begin

    FrmDebug := TForm.Create(nil);
    FrmDebug.FormStyle := fsStayOnTop;
    FrmDebug.Caption := 'Debug';
    memo := TMemo.Create(FrmDebug);

    With memo do
    begin
      Parent := FrmDebug;
      Align := alClient;
      ReadOnly := true;
      ScrollBars := ssBoth;
    end;

  end
  else
  begin

    memo := FrmDebug.Components[0] as TMemo;
  end;

  if (not FrmDebug.Showing) then
  begin

    FrmDebug.Show;
  end;

  if (exibirHoraDebug) then
  begin

    memo.Lines.Add(DateTimeToStr(now));
  end;

  memo.Lines.Add(ATexto);

  for i := 0 to 20 do
  begin

    divisorLinha := divisorLinha + '-';
  end;

  memo.Lines.Add(divisorLinha);

end;

Function SeparaNumeroEmbalagem(pCaixasdaEmbalagem: String): Double;
Var
  Contador: integer;
  Valor: string;
  Digito: integer;
Begin

  {
    ///  Versão:       5.0.3.0 - Rotina:9874
    Marcos Pereira
    05/02/2016

    Retira da string com a embalagem '30 CX C/20', apenas o valor numérico, da
    quantidade de caixas, que nesse caso seria 30 e devolve como double

  }

  Valor := '';
  Contador := 1;
  while true do
  Begin

    if (Copy(pCaixasdaEmbalagem, Contador, 1) <> ' ') and (Copy(pCaixasdaEmbalagem, Contador, 1) <> '(') then
    Begin
      if TryStrToInt(Copy(pCaixasdaEmbalagem, Contador, 1), Digito) then
        /// Versão:       5.0.3.2
        Valor := Valor + InttoStr(Digito) // copy(pCaixasdaEmbalagem,contador,1)
      else
        Break;
    end;

    Contador := Contador + 1;

  end;

  try
    Result := StrToFloat(Valor);
  except
    Result := 0;
  End;

end;

function MinutosToStr(const Minutes: Cardinal; Reduzido: Boolean = false): string;
var
  d, h, M: integer;

  horas, restante_horas: integer;
  minutos: integer;

  descricao: TStringList;
  descricao_reduzida: string;

begin
  /// Jhonny Oliveira - 04/08/2016
  /// Dado um quantidade de minutos formata a saída para algo como: 4 dias e 16 horas e 4 minutos

  h := Minutes div 60;
  M := Minutes mod 60;
  d := h div 24;
  h := h mod 24;

  descricao := TStringList.Create;
  descricao.Delimiter := 'e';

  if (not Reduzido) then
  begin

    if d > 0 then
    begin

      if d > 1 then
      begin

        descricao.Add(Format(' %d dias ', [d]));
      end
      else
      begin

        descricao.Add(Format(' %d dia ', [d]));
      end;

    end;

    if h > 0 then
    begin

      if (h > 1) then
      begin

        descricao.Add(Format(' %d horas ', [h]));
      end
      else
      begin

        descricao.Add(Format(' %d hora ', [h]));
      end;

    end;

    if M > 0 then
    begin

      if (M > 1) then
      begin

        descricao.Add(Format(' %d minutos ', [M]));
      end
      else
      begin

        descricao.Add(Format(' %d minuto ', [M]));
      end;
    end;

    Result := StringReplace(descricao.DelimitedText, '"', '', [rfReplaceAll]);;

  end
  else // Reduzido
  begin

    descricao_reduzida := '';

    if (d + h + M < 1) then
    begin

      descricao_reduzida := '00:00';
    end
    else
    begin

      if d >= 1 then
      begin

        descricao_reduzida := formatfloat('00d ', d);
      end;

      if h >= 1 then
      begin

        descricao_reduzida := descricao_reduzida + formatfloat('00:', h);
      end
      else
      begin

        descricao_reduzida := descricao_reduzida + '00:';
      end;

      if M >= 1 then
      begin

        descricao_reduzida := descricao_reduzida + formatfloat('00', M);
      end
      else
      begin

        descricao_reduzida := descricao_reduzida + '00';
      end;

    end;
    Result := descricao_reduzida;

  end;

end;

Function AbrirProgramaExternoModal(FileName: String; Params: String = ''; Visibility: integer = SW_SHOWNORMAL): DWORD;
  Procedure WaitFor(processHandle: THandle);
  Var
    msg: TMsg;
    ret: DWORD;
  Begin
    Repeat
      ret := MsgWaitForMultipleObjects(1, { 1 handle to wait on }
        processHandle, { the handle }
        false, { wake on any event }
        INFINITE, { wait without timeout }
        QS_PAINT or { wake on paint messages }
        QS_SENDMESSAGE { or messages from other threads }
        );
      If ret = WAIT_FAILED Then
        Exit; { can do little here }
      If ret = (WAIT_OBJECT_0 + 1) Then
      Begin
        { Woke on a message, process paint messages only. Calling
          PeekMessage gets messages send from other threads processed. }
        While PeekMessage(msg, 0, WM_PAINT, WM_PAINT, PM_REMOVE) Do
          DispatchMessage(msg);
      End;
    Until ret = WAIT_OBJECT_0;
  End; { Waitfor }

Var { V1 by Pat Ritchey, V2 by P.Below }
  zAppName: array [0 .. 512] of Char;
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
Begin { WinExecAndWait32V2 }

  if (Params <> '') then
  begin

    FileName := FileName + ' ' + Params;
  end;

  StrPCopy(zAppName, FileName);

  FillChar(StartupInfo, SizeOf(StartupInfo), #0);
  StartupInfo.cb := SizeOf(StartupInfo);
  StartupInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartupInfo.wShowWindow := Visibility;
  If not CreateProcess(nil, zAppName, { pointer to command line string }
    nil, { pointer to process security attributes }
    nil, { pointer to thread security attributes }
    false, { handle inheritance flag }
    CREATE_NEW_CONSOLE or { creation flags }
    NORMAL_PRIORITY_CLASS, nil, { pointer to new environment block }
    nil, { pointer to current directory name }
    StartupInfo, { pointer to STARTUPINFO }
    ProcessInfo) { pointer to PROCESS_INF }
  Then
    Result := DWORD(-1) { failed, GetLastError has error code }
  Else
  Begin
    WaitFor(ProcessInfo.hProcess);
    GetExitCodeProcess(ProcessInfo.hProcess, Result);
    CloseHandle(ProcessInfo.hProcess);
    CloseHandle(ProcessInfo.hThread);
  End; { Else }
End; { WinExecAndWait32V2 }

function CalcularDigitoVerificadorEAN(Numero: string): string;
var
  i, soma: integer;
begin

  if (Trim(Numero) = '') then
  begin

    Result := '';
    Exit;
  end;

  soma := 0;
  for i := 1 to 12 do
  begin
    if (i mod 2 = 0) then
      soma := soma + StrToInt(Numero[i]) * 3
    else
      soma := soma + StrToInt(Numero[i]);
  end;
  Result := InttoStr((10 - (soma mod 10)) mod 10);
end;

end.
