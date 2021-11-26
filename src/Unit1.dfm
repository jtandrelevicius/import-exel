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
    Caption = 'Importar'
    TabOrder = 2
    OnClick = Panel2Click
  end
  object ProgressBar1: TProgressBar
    Left = 8
    Top = 571
    Width = 241
    Height = 17
    TabOrder = 3
  end
  object Edit1: TEdit
    Left = 712
    Top = 576
    Width = 121
    Height = 21
    TabOrder = 4
    Text = 'Edit1'
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
  object FDQuery1: TFDQuery
    Connection = FDConnectionRAD
    SQL.Strings = (
      'SELECT * FROM NCM')
    Left = 248
    Top = 592
    object FDQuery1ID: TIntegerField
      AutoGenerateValue = arAutoInc
      FieldName = 'ID'
      Origin = 'ID'
    end
    object FDQuery1DESCRICAO: TStringField
      FieldName = 'DESCRICAO'
      Origin = 'DESCRICAO'
      Size = 250
    end
    object FDQuery1DESCRICAO_CONCATENADA: TStringField
      FieldName = 'DESCRICAO_CONCATENADA'
      Origin = 'DESCRICAO_CONCATENADA'
      Size = 250
    end
    object FDQuery1DATA_INICIO: TStringField
      FieldName = 'DATA_INICIO'
      Origin = 'DATA_INICIO'
      Size = 50
    end
    object FDQuery1DATA_FIM: TStringField
      FieldName = 'DATA_FIM'
      Origin = 'DATA_FIM'
      Size = 50
    end
    object FDQuery1ATO_LEGAL: TStringField
      FieldName = 'ATO_LEGAL'
      Origin = 'ATO_LEGAL'
      Size = 50
    end
    object FDQuery1NUMERO: TStringField
      FieldName = 'NUMERO'
      Origin = 'NUMERO'
      Size = 50
    end
    object FDQuery1ANO: TStringField
      FieldName = 'ANO'
      Origin = 'ANO'
      Size = 50
    end
    object FDQuery1CODIGO: TStringField
      FieldName = 'CODIGO'
      Origin = 'CODIGO'
      Size = 50
    end
  end
  object FDUpdateSQL1: TFDUpdateSQL
    Left = 848
    Top = 592
  end
  object DataSource1: TDataSource
    DataSet = FDQuery1
    Left = 768
    Top = 592
  end
  object FDConnectionRAD: TFDConnection
    Params.Strings = (
      'Protocol='
      'User_Name=sysdba'
      'Password=masterkey'
      'Database=C:\RocketRetaguarda\DataBase\ROCKET_ERP2.FDB'
      'Port=3050'
      'Server=localhost'
      'DriverID=FB')
    Connected = True
    LoginPrompt = False
    Left = 40
    Top = 592
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    VendorLib = 'C:\RocketRetaguarda\fbclient.dll'
    Left = 152
    Top = 592
  end
end
