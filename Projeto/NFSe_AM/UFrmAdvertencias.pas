unit UFrmAdvertencias;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons,
  System.StrUtils, Vcl.ExtDlgs;

type
  TFrmAdvertencias = class(TForm)
    PnBot: TPanel;
    Btn_Fechar: TBitBtn;
    Btn_Salvar: TBitBtn;
    MemAdvertencias: TMemo;
    SaveText: TSaveTextFileDialog;
    procedure Btn_FecharClick(Sender: TObject);
    procedure Btn_SalvarClick(Sender: TObject);
  private
    FLista: TStringList;
  public
    constructor Create(AOwner: TComponent; ACaption: String; ALista: TStringList); reintroduce;
  end;

var
  FrmAdvertencias: TFrmAdvertencias;

implementation

{$R *.dfm}

procedure TFrmAdvertencias.Btn_FecharClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmAdvertencias.Btn_SalvarClick(Sender: TObject);
begin
  if SaveText.Execute then
    FLista.SaveToFile(SaveText.FileName);
end;

constructor TFrmAdvertencias.Create(AOwner: TComponent; ACaption: String; ALista: TStringList);
begin
  inherited Create(AOwner);
  Self.Caption := ACaption;
  FLista := TStringList.Create;
  FLista := ALista;
  MemAdvertencias.Lines := FLista;
end;

end.
