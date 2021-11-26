unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Grids, uImportExcel,
  Vcl.ComCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error,
  FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait, FireDAC.Stan.Param,
  FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Phys.FBDef, Data.DB,
  FireDAC.Phys.IBBase, FireDAC.Phys.FB, FireDAC.Comp.Client,
  FireDAC.Comp.DataSet, Datasnap.DBClient, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    ImportExcel1: TImportExcel;
    StringGrid1: TStringGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    OpenDialog1: TOpenDialog;
    ProgressBar1: TProgressBar;
    FDQuery1: TFDQuery;
    FDUpdateSQL1: TFDUpdateSQL;
    DataSource1: TDataSource;
    FDConnectionRAD: TFDConnection;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDQuery1ID: TIntegerField;
    FDQuery1DESCRICAO: TStringField;
    FDQuery1DESCRICAO_CONCATENADA: TStringField;
    FDQuery1DATA_INICIO: TStringField;
    FDQuery1DATA_FIM: TStringField;
    FDQuery1ATO_LEGAL: TStringField;
    FDQuery1NUMERO: TStringField;
    FDQuery1ANO: TStringField;
    FDQuery1CODIGO: TStringField;
    Edit1: TEdit;
    procedure Panel1Click(Sender: TObject);
    procedure Panel2Click(Sender: TObject);
  private
    { Private declarations }
    procedure CarregarExel;
    procedure ImportarExel;
    procedure ExportarExel;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ TForm1 }

function RemoveAcento(aText : string) : string;
const
  ComAcento = '‡‚ÍÙ˚„ı·ÈÌÛ˙Á¸Ò˝¿¬ ‘€√’¡…Õ”⁄«‹—›';
  SemAcento = 'aaeouaoaeioucunyAAEOUAOAEIOUCUNY';
var
  x: Cardinal;
begin;
  for x := 1 to Length(aText) do
  try
    if (Pos(aText[x], ComAcento) <> 0) then
      aText[x] := SemAcento[ Pos(aText[x], ComAcento) ];
  except on E: Exception do
    raise Exception.Create('Erro no processo.');
  end;

  Result := aText;
end;

procedure TForm1.CarregarExel;
begin
  if OpenDialog1.Execute then begin
    ImportExcel1.ExcelFile :=   OpenDialog1.FileName;
    ImportExcel1.ExcelParaStringGrid(StringGrid1, ProgressBar1);
  end;

end;

procedure TForm1.ExportarExel;
begin
  ImportExcel1.ExportarParaExcel(FDQuery1, 'export_rocket');
end;

procedure TForm1.ImportarExel;
var
i : integer;
begin
  try
       for I := 0 to Pred(StringGrid1.RowCount) do
       begin
          if I = 0 then
            Continue;
          if Trim(StringGrid1.Cells[0,I]) = '' then
            Continue;

        FDQuery1.Active := True;
        FDQuery1.Insert;
        FDQuery1CODIGO.Value                      := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[0,I]));
        FDQuery1DESCRICAO.Value                   := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[1,I]));
        FDQuery1DESCRICAO_CONCATENADA.Value       := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[3,I]));
        FDQuery1DATA_INICIO.Value                 := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[4,I]));
        FDQuery1DATA_FIM.Value                    := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[5,I]));
        FDQuery1ATO_LEGAL.Value                   := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[6,I]));
        FDQuery1NUMERO.Value                      := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[7,I]));
        FDQuery1ANO.Value                         := RemoveAcento(AnsiUpperCase(StringGrid1.Cells[8,I]));
        //FDQuery1DATA_IMPORTACAO.AsDateTime := Now;
        FDQuery1.Post;
  end;

  Showmessage('Dados importados');

  finally
     FDQuery1.Active := False;
     FDQuery1.DisposeOf;
  end;

end;

procedure TForm1.Panel1Click(Sender: TObject);
begin
   CarregarExel;
end;

procedure TForm1.Panel2Click(Sender: TObject);
begin
   ImportarExel;
end;

end.
