unit UFrmConsultaWinthor;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, UConsultasWinthor, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, Menus, cxControls, cxStyles, cxCustomData, cxFilter,
  cxData, cxDataStorage, cxEdit, DB, cxDBData, cxContainer, cxTextEdit,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, cxButtons, cxLabel, ImgList,
  dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkroom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinOffice2016Colorful,
  dxSkinOffice2016Dark, dxSkinOffice2019Colorful, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringtime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinTheBezier, dxSkinsDefaultPainters, dxSkinValentine,
  dxSkinVisualStudio2013Blue, dxSkinVisualStudio2013Dark,
  dxSkinVisualStudio2013Light, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxNavigator, dxDateRanges, MidasLib;

type
  TFrmConsultaWinthor = class(TForm)
    btnBuscar: TcxButton;
    btnCancelar: TcxButton;
    btnSelecionar: TcxButton;
    grdResultadoDBTableView1: TcxGridDBTableView;
    grdResultadoLevel1: TcxGridLevel;
    grdResultado: TcxGrid;
    edtBusca: TcxTextEdit;
    cxLabel1: TcxLabel;
    dsResultado: TDataSource;
    grdResultadoDBTableView1CODIGO: TcxGridDBColumn;
    grdResultadoDBTableView1DESCRICAO: TcxGridDBColumn;
    cxLookAndFeelController1: TcxLookAndFeelController;
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure _Selecionar(Sender: TObject);
    procedure _Cancelar(Sender: TObject);
    procedure _Pesquisar(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    TipoPesquisa: TTipoConsulta; { Tipo de pesquisa que a rotina está fazendo }
    CodigoSelecionado: String; { Código do registro selecionado pelo usuário }

  end;

var
  FrmConsultaWinthor: TFrmConsultaWinthor;

implementation

{$R *.dfm}

procedure TFrmConsultaWinthor._Pesquisar(Sender: TObject);
var
  qtd_registros: Integer;

begin

  { O usuário deve informar algum texto na pesquisa }
  if (Trim(edtBusca.Text) = '') then
  begin

    { Verificando se o datasource já foi associado com a query
      de pesquisa.

      Isso acontece na primeira pesquisa realizada
    }
    if (Assigned(dsResultado.DataSet)) then
    begin

      dsResultado.DataSet.Close;
    end;
  end;

  { Realizando a pesquisa e obtendo a quantidade de registros retornados }

  qtd_registros := TConsulta.PesquisarPorDescricao(Self.TipoPesquisa,
    edtBusca.Text, dsResultado);

  if (qtd_registros = 0) then
  begin

    Application.MessageBox('Nenhum resultado para sua pesquisa', 'Atenção',
      MB_OK + MB_ICONWARNING);
    edtBusca.SetFocus;
  end
  else
  begin

    grdResultado.SetFocus;
  end;

end;

procedure TFrmConsultaWinthor._Cancelar(Sender: TObject);
begin

  { Fecha o form sem definir o CodigoSelecionado }
  Close;
end;

procedure TFrmConsultaWinthor._Selecionar(Sender: TObject);
begin

  { Verificando se a pesquisa realmente foi feita e se tem algum
    resultado
  }
  if (not Assigned(dsResultado.DataSet)) or (dsResultado.DataSet.RecordCount = 0)
  then
  begin

    edtBusca.SetFocus;
    Exit;
  end;

  { Definindo o código selecionado pelo usuário e fechando o formulário }
  CodigoSelecionado := dsResultado.DataSet.FieldByName('CODIGO').AsString;
  Close;

end;

procedure TFrmConsultaWinthor.FormKeyPress(Sender: TObject; var Key: Char);
begin

  if (Key = #13) then
  begin

    Key := #0;
    Perform(WM_NEXTDLGCTL, 0, 0);
  end;

end;

procedure TFrmConsultaWinthor.FormShow(Sender: TObject);
begin

  CodigoSelecionado := '';

end;

end.
