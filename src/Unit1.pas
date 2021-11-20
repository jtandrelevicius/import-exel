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
  FireDAC.Comp.DataSet, Datasnap.DBClient;

type
  TForm1 = class(TForm)
    ImportExcel1: TImportExcel;
    StringGrid1: TStringGrid;
    Panel1: TPanel;
    Panel2: TPanel;
    OpenDialog1: TOpenDialog;
    ProgressBar1: TProgressBar;
    FDConnection1: TFDConnection;
    FDQuery1: TFDQuery;
    FDUpdateSQL1: TFDUpdateSQL;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    DataSource1: TDataSource;
    procedure Panel1Click(Sender: TObject);
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


        FDQuery1.Insert;
        //FDQuery1DADOS.Value            := StringGrid1.Cells[0,I];
        //FDQuery1DADOS.Value            := StringGrid1.Cells[1,I];
       // FDQuery1DADOS.Value            := StringGrid1.Cells[2,I];
        //FDQuery1DADOS.Value            := StringGrid1.Cells[3,I];
        //FDQuery1DATA_IMPORTACAO.AsDateTime := Now;
        FDQuery1.Post;
  end;

  Showmessage('Dados importados');

  finally
     FDQuery1.DisposeOf;
  end;

end;

procedure TForm1.Panel1Click(Sender: TObject);
begin
   CarregarExel;
end;

end.
