object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 648
  ClientWidth = 1035
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object StringGrid1: TStringGrid
    Left = 0
    Top = 0
    Width = 1035
    Height = 565
    Align = alTop
    FixedCols = 0
    TabOrder = 0
  end
  object Panel1: TPanel
    Left = 328
    Top = 599
    Width = 185
    Height = 41
    Caption = 'Carregar Planilha'
    TabOrder = 1
    OnClick = Panel1Click
  end
  object Panel2: TPanel
    Left = 527
    Top = 599
    Width = 185
    Height = 41
    Caption = 'Panel1'
    TabOrder = 2
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 571
    Width = 241
    Height = 17
    TabOrder = 3
  end
  object ImportExcel1: TImportExcel
    Left = 992
    Top = 592
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Exel|*.xlsx'
    Left = 928
    Top = 592
  end
  object FDConnection1: TFDConnection
    Left = 40
    Top = 592
  end
  object FDQuery1: TFDQuery
    Connection = FDConnection1
    Left = 248
    Top = 592
  end
  object FDUpdateSQL1: TFDUpdateSQL
    Connection = FDConnection1
    Left = 848
    Top = 592
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    Left = 152
    Top = 592
  end
  object DataSource1: TDataSource
    DataSet = FDQuery1
    Left = 768
    Top = 592
  end
end
