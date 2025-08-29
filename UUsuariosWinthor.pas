unit UUsuariosWinthor;

interface

uses DBAccess, Ora, OraSmart, MemDS, OraError, Dialogs, Forms, Windows, Classes,
  SysUtils;

type
  TUsuario = class

  private
    FMatricula: Double;
    FNome: String;
    FLogin: String;
    FCodigoFilial: String;
    FCodigoSetor: Double;
    FFilaisUsuario: TStringList;

    /// Jhonny Oliveira - 05/09/2014
    /// Adicionado para o uso do sistema de requisição de compras
    /// rotina 9857
    FEmail: String;

    /// Jhonny Oliveira - 06/10/2014
    /// Adicionado para o uso do sistema de controle de equipes
    /// Rotina 9868
    FCodigoCracha: String;
    FAtivo: Boolean;

  protected
    function GetMatricula(): Double;
    function GetNome(): String;
    function GetLogin(): String;
    function GetCodigoFilial(): String;
    function GetCodigoSetor(): Double;
    function GetEmail(): String;
    function GetCodigoCracha(): String;
    function GetAtivo(): Boolean;

  public

  published
    property Matricula: Double read GetMatricula;
    property Nome: String read GetNome;
    property Login: String read GetLogin;
    property CodigoFilial: String read GetCodigoFilial;
    property CodigoSetor: Double read GetCodigoSetor;
    property Filiais: TStringList read FFilaisUsuario;
    property Email: String read GetEmail write FEmail;
    property CodigoCracha: String read GetCodigoCracha;
    property Ativo: Boolean read GetAtivo;

    constructor Create();

  end;

type
  TUsuarioController = class

  private
    FDatabaseName: String;
    ConsultaComum: String;
    OdacSession: TOraSession;

  protected
    procedure GetFiliaisUsuario(AUsuario: TUsuario);

  public
    function GetUsuarioPorLogin(aLoginUsuario: String): TUsuario;
    function GetUsuarioPorMatricula(aMatriculaUsuario: Double): TUsuario;
    function PossuiAcesso(AUsuario: TUsuario; ACodigoRotina: Double; ACodSubRotina: Double = 0): Boolean;
    constructor Create(aOraSession: TOraSession);
    function GetUsuarioPorCracha(aCodigoCracha: String): TUsuario;

  end;

implementation

uses ULibrary, DB;

{ TUsuario }

constructor TUsuario.Create;
begin

  Self.FFilaisUsuario := TStringList.Create;

end;

function TUsuario.GetCodigoFilial: String;
begin

  Result := Self.FCodigoFilial;

end;

function TUsuario.GetCodigoSetor: Double;
begin

  Result := Self.FCodigoSetor;

end;

function TUsuario.GetLogin: String;
begin

  Result := Self.FLogin;

end;

function TUsuario.GetMatricula: Double;
begin

  Result := Self.FMatricula;

end;

function TUsuario.GetNome: String;
begin

  Result := Self.FNome;

end;

function TUsuario.GetEmail: String;
begin

  Result := Self.FEmail;

end;

function TUsuario.GetCodigoCracha: String;
begin

  Result := Self.FCodigoCracha;

end;

function TUsuario.GetAtivo;
begin

  Result := Self.FAtivo;
end;

{ TUsuarioController }

constructor TUsuarioController.Create(aOraSession: TOraSession);
begin

  Self.OdacSession := aOraSession;
  Self.ConsultaComum := 'SELECT pcempr.matricula, pcempr.nome, ' +
    ' pcempr.nome_guerra, pcempr.codfilial , ' +
    ' pcempr.codsetor, pcempr.email , ' +
    ' nvl(pcempr.codbarra, '''') as codbarra,  ' +
    ' nvl(pcempr.situacao, ''I'') AS SITUACAO ' + ' FROM pcempr ';

end;

procedure TUsuarioController.GetFiliaisUsuario(AUsuario: TUsuario);
var
  qry: TOraQuery;

begin

  qry := TOraQuery.Create(nil);
  qry.Session := Self.OdacSession;
  qry.SQL.Add(' SELECT codigoa                    ');
  qry.SQL.Add(' FROM pclib                        ');
  qry.SQL.Add(' WHERE pclib.codfunc = :MATRICULA  ');
  qry.SQL.Add(' AND pclib.codtabela = 1           ');
  qry.SQL.Add(' ORDER BY TO_NUMBER(CODIGOA)       ');

  qry.ParamByName('MATRICULA').AsFloat := AUsuario.Matricula;

  qry.Open;

  qry.First;

  while not qry.Eof do
  begin

    AUsuario.FFilaisUsuario.Add(qry.FieldByName('CODIGOA').AsString);

    qry.Next;

  end;

end;

function TUsuarioController.GetUsuarioPorLogin(aLoginUsuario: String): TUsuario;
var
  qry: TOraQuery;
  usuario: TUsuario;

begin

  qry := TOraQuery.Create(nil);
  qry.Session := Self.OdacSession;

  qry.Close;

  qry.SQL.Clear;
  qry.SQL.Add(Self.ConsultaComum);
  qry.SQL.Add(' WHERE nome_guerra = :LOGIN ');

  qry.ParamByName('LOGIN').AsString := aLoginUsuario;

  qry.Open;

  if qry.RecordCount = 0 then
  begin

    usuario := nil;

  end
  else
  begin

    usuario := TUsuario.Create;

    usuario.FMatricula := qry.FieldByName('MATRICULA').AsFloat;
    usuario.FLogin := qry.FieldByName('NOME_GUERRA').AsString;
    usuario.FCodigoFilial := qry.FieldByName('CODFILIAL').AsString;
    usuario.FNome := qry.FieldByName('NOME').AsString;
    usuario.FCodigoSetor := qry.FieldByName('CODSETOR').AsFloat;
    usuario.FEmail := qry.FieldByName('EMAIL').AsString;
    usuario.FCodigoCracha := qry.FieldByName('CODBARRA').AsString;
    usuario.FAtivo := qry.FieldByName('SITUACAO').AsString = 'A';

    Self.GetFiliaisUsuario(usuario);

  end;

  Result := usuario;

end;

function TUsuarioController.GetUsuarioPorMatricula(aMatriculaUsuario: Double)
  : TUsuario;
var
  qry: TOraQuery;
  usuario: TUsuario;

begin

  qry := TOraQuery.Create(nil);
  qry.Session := Self.OdacSession;

  qry.Close;

  qry.SQL.Clear;
  qry.SQL.Add(Self.ConsultaComum);
  qry.SQL.Add(' WHERE matricula = :MATRICULA');

  qry.ParamByName('MATRICULA').AsFloat := aMatriculaUsuario;

  qry.Open;

  if qry.RecordCount = 0 then
  begin

    usuario := nil;

  end
  else
  begin

    usuario := TUsuario.Create;

    usuario.FMatricula := qry.FieldByName('MATRICULA').AsFloat;
    usuario.FLogin := qry.FieldByName('NOME_GUERRA').AsString;
    usuario.FCodigoFilial := qry.FieldByName('CODFILIAL').AsString;
    usuario.FNome := qry.FieldByName('NOME').AsString;
    usuario.FCodigoSetor := qry.FieldByName('CODSETOR').AsFloat;
    usuario.FEmail := qry.FieldByName('EMAIL').AsString;
    usuario.FCodigoCracha := qry.FieldByName('CODBARRA').AsString;
    usuario.FAtivo := qry.FieldByName('SITUACAO').AsString = 'A';

    Self.GetFiliaisUsuario(usuario);

  end;

  qry.Close;
  FreeAndNil(qry);

  Result := usuario;

end;

function TUsuarioController.PossuiAcesso(AUsuario: TUsuario;
  ACodigoRotina, ACodSubRotina: Double): Boolean;
var
  qry: TOraQuery;

begin

  qry := TOraQuery.Create(nil);
  qry.Session := Self.OdacSession;
  qry.SQL.Clear;

  if ACodSubRotina > 0 then
  begin

    qry.SQL.Add(' select codrotina from pccontroi ');
    qry.SQL.Add(' where codrotina = :CODROTINA    ');
    qry.SQL.Add(' and codcontrole = :CODCONTROLE  ');
    qry.SQL.Add(' and codusuario = :CODUSUARIO    ');
    qry.SQL.Add(' and acesso = ''S''              ');

    qry.ParamByName('CODCONTROLE').AsFloat := ACodSubRotina;

  end
  else
  begin

    qry.SQL.Add(' select codrotina from pccontro  ');
    qry.SQL.Add(' where codrotina = :CODROTINA    ');
    qry.SQL.Add(' and codusuario = :CODUSUARIO    ');
    qry.SQL.Add(' and acesso = ''S''              ');

  end;

  qry.ParamByName('CODROTINA').AsFloat := ACodigoRotina;
  qry.ParamByName('CODUSUARIO').AsFloat := AUsuario.Matricula;

  qry.Open;

  Result := qry.RecordCount > 0;

  qry.Close();
  FreeAndNil(qry);

end;

function TUsuarioController.GetUsuarioPorCracha(aCodigoCracha: String)
  : TUsuario;
var
  qry: TOraQuery;
  usuario: TUsuario;

begin

  {
    Jhonny Oliveira
    06/10/2014

    Permite pesquisar o usuário utilizando o código do crachá
  }

  qry := TOraQuery.Create(nil);
  qry.Session := Self.OdacSession;

  qry.Close;

  qry.SQL.Clear;
  qry.SQL.Add(Self.ConsultaComum);
  qry.SQL.Add(' WHERE codbarra = :CODBARRA');

  qry.ParamByName('CODBARRA').AsString := aCodigoCracha;

  qry.Open;

  if qry.RecordCount = 0 then
  begin

    usuario := nil;

  end
  else
  begin

    usuario := TUsuario.Create;

    usuario.FMatricula := qry.FieldByName('MATRICULA').AsFloat;
    usuario.FLogin := qry.FieldByName('NOME_GUERRA').AsString;
    usuario.FCodigoFilial := qry.FieldByName('CODFILIAL').AsString;
    usuario.FNome := qry.FieldByName('NOME').AsString;
    usuario.FCodigoSetor := qry.FieldByName('CODSETOR').AsFloat;
    usuario.FEmail := qry.FieldByName('EMAIL').AsString;
    usuario.FCodigoCracha := qry.FieldByName('CODBARRA').AsString;
    usuario.FAtivo := qry.FieldByName('SITUACAO').AsString = 'A';

    Self.GetFiliaisUsuario(usuario);

  end;

  Result := usuario;

  qry.Close;
  FreeAndNil(qry);

end;

end.
