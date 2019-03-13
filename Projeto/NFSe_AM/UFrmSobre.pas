unit UFrmSobre;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TFrmSobre = class(TForm)
    Lb_Sobre: TLabel;
    Lb_Versao: TLabel;
    Lb_ArquivosExcel: TLabel;
    SaveArq: TSaveDialog;
    Lb_Manual: TLabel;
    procedure Lb_ArquivosExcelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Lb_ManualClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmSobre: TFrmSobre;

implementation

{$R *.dfm}

uses funcoes;

procedure TFrmSobre.FormCreate(Sender: TObject);
begin
  Lb_Versao.Caption := 'Versão: ' + VersaoExe;
end;

procedure TFrmSobre.Lb_ArquivosExcelClick(Sender: TObject);
var
  Stream: TResourceStream;
begin
  SaveArq.FilterIndex := 1;
  SaveArq.FileName := 'ArquivosExcel.zip';
  if SaveArq.Execute then
  begin
    Stream := TResourceStream.Create(hInstance, 'ArquivosExcel', RT_RCDATA);
    try
      Stream.SaveToFile(SaveArq.FileName);
      Application.MessageBox(PWideChar('Arquivo "' + ExtractFileName(SaveArq.FileName) + '" gerado com sucesso no caminho "' + ExtractFileDir(SaveArq.FileName) + '".'), 'Informação', MB_ICONINFORMATION + MB_OK);
    finally
      Stream.Free;
    end;
  end;
end;

procedure TFrmSobre.Lb_ManualClick(Sender: TObject);
var
  Stream: TResourceStream;
begin
  SaveArq.FilterIndex := 2;
  SaveArq.FileName := 'ManualNFSe_AM.pdf';
  if SaveArq.Execute then
  begin
    Stream := TResourceStream.Create(hInstance, 'ManualNFSe_AM', RT_RCDATA);
    try
      Stream.SaveToFile(SaveArq.FileName);
      Application.MessageBox(PWideChar('Arquivo "' + ExtractFileName(SaveArq.FileName) + '" gerado com sucesso no caminho "' + ExtractFileDir(SaveArq.FileName) + '".'), 'Informação', MB_ICONINFORMATION + MB_OK);
    finally
      Stream.Free;
    end;
  end;
end;

end.
