inherited CNVESA: TCNVESA
  Left = 336
  Top = 178
  Caption = 'CNVESA'
  ClientHeight = 424
  ClientWidth = 748
  ExplicitWidth = 754
  ExplicitHeight = 468
  PixelsPerInch = 96
  TextHeight = 13
  inherited toolbar: TToolBar
    Width = 748
    ExplicitWidth = 748
  end
  inherited statusbar: TStatusBar
    Top = 404
    Width = 748
    ExplicitTop = 404
    ExplicitWidth = 748
  end
  inherited tab_control: TRzPageControl
    Width = 748
    Height = 370
    ExplicitWidth = 748
    ExplicitHeight = 370
    FixedDimension = 21
    inherited tab_pagina1: TRzTabSheet
      ExplicitLeft = 1
      ExplicitTop = 22
      ExplicitWidth = 746
      ExplicitHeight = 347
      inherited pannello_elaborazione: TRzPanel
        Width = 746
        Height = 347
        ExplicitWidth = 746
        ExplicitHeight = 347
        inherited pannello_parametri: TRzPanel
          Width = 746
          Height = 309
          ExplicitWidth = 746
          ExplicitHeight = 309
          object v_tabella_01: TRzLabel
            Left = 10
            Top = 201
            Width = 105
            Height = 13
            Caption = 'tabella in elaborazione'
            Transparent = True
          end
          object v_tabella: TRzLabel
            Left = 125
            Top = 22
            Width = 321
            Height = 24
            AutoSize = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'MS Sans Serif'
            Font.Style = []
            ParentFont = False
            Transparent = True
          end
          object GroupBox1: TGroupBox
            Left = 0
            Top = 0
            Width = 746
            Height = 161
            Align = alTop
            Caption = 'archivi da convertire'
            TabOrder = 0
            object Label8: TRzLabel
              Left = 110
              Top = 110
              Width = 71
              Height = 13
              Caption = 'c/p acquisti art'
              Transparent = True
            end
            object Label7: TRzLabel
              Left = 10
              Top = 110
              Width = 70
              Height = 13
              Caption = 'c/p vendite art'
              Transparent = True
            end
            object Label9: TRzLabel
              Left = 247
              Top = 107
              Width = 110
              Height = 13
              Caption = 'articolo fuori magazzino'
              Transparent = True
            end
            object v_sottoconti: TRzCheckBox
              Left = 130
              Top = 20
              Width = 88
              Height = 15
              Caption = 'piano dei conti'
              State = cbUnchecked
              TabOrder = 1
            end
            object v_clienti: TRzCheckBox
              Left = 250
              Top = 20
              Width = 46
              Height = 15
              Caption = 'clienti'
              State = cbUnchecked
              TabOrder = 2
            end
            object v_fornitori: TRzCheckBox
              Left = 345
              Top = 20
              Width = 53
              Height = 15
              Caption = 'fornitori'
              State = cbUnchecked
              TabOrder = 3
            end
            object v_articoli: TRzCheckBox
              Left = 435
              Top = 20
              Width = 49
              Height = 15
              Caption = 'articoli'
              State = cbUnchecked
              TabOrder = 4
            end
            object v_ind_inf: TRzCheckBox
              Left = 504
              Top = 44
              Width = 105
              Height = 15
              Caption = 'indirizzi spedizione'
              State = cbUnchecked
              TabOrder = 5
            end
            object v_lsv: TRzCheckBox
              Left = 10
              Top = 45
              Width = 41
              Height = 15
              Caption = 'listini'
              State = cbUnchecked
              TabOrder = 6
            end
            object v_pnt: TRzCheckBox
              Left = 130
              Top = 45
              Width = 65
              Height = 15
              Caption = 'primanota'
              State = cbUnchecked
              TabOrder = 7
            end
            object v_scadenze: TRzCheckBox
              Left = 250
              Top = 45
              Width = 65
              Height = 15
              Caption = 'scadenze'
              State = cbUnchecked
              TabOrder = 8
            end
            object v_mov: TRzCheckBox
              Left = 345
              Top = 45
              Width = 69
              Height = 15
              Caption = 'magazzino'
              State = cbUnchecked
              TabOrder = 9
            end
            object v_tabelle: TRzCheckBox
              Left = 10
              Top = 20
              Width = 50
              Height = 15
              Caption = 'tabelle'
              State = cbUnchecked
              TabOrder = 0
            end
            object v_ordini: TRzCheckBox
              Left = 433
              Top = 45
              Width = 44
              Height = 15
              Caption = 'ordini'
              State = cbUnchecked
              TabOrder = 10
            end
            object v_codifica_clienti: TRzCheckBox
              Left = 10
              Top = 68
              Width = 128
              Height = 15
              Caption = 'mantieni codifica clienti'
              State = cbUnchecked
              TabOrder = 11
            end
            object v_codifica_fornitori: TRzCheckBox
              Left = 250
              Top = 68
              Width = 135
              Height = 15
              Caption = 'mantieni codifica fornitori'
              State = cbUnchecked
              TabOrder = 12
            end
            object v_ricavi: trzedit_go
              Left = 10
              Top = 125
              Width = 86
              Height = 21
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              ReadOnlyColor = clBtnFace
              ReadOnlyColorOnFocus = True
              TabOrder = 13
            end
            object v_acquisti: trzedit_go
              Left = 110
              Top = 125
              Width = 86
              Height = 21
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              ReadOnlyColor = clBtnFace
              ReadOnlyColorOnFocus = True
              TabOrder = 14
            end
            object v_art_codice: trzedit_go
              Left = 247
              Top = 122
              Width = 136
              Height = 21
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWindowText
              Font.Height = -11
              Font.Name = 'MS Sans Serif'
              Font.Style = [fsBold]
              ParentFont = False
              ReadOnlyColor = clBtnFace
              ReadOnlyColorOnFocus = True
              TabOrder = 15
            end
            object v_codice_aggiuntivi: TRzCheckBox
              Left = 504
              Top = 20
              Width = 95
              Height = 15
              Caption = 'codici aggiuntivi'
              State = cbUnchecked
              TabOrder = 16
            end
            object v_provvigioni: TRzCheckBox
              Left = 614
              Top = 19
              Width = 70
              Height = 15
              Caption = 'provvigioni'
              State = cbUnchecked
              TabOrder = 17
            end
          end
        end
        inherited Panel4: TRzPanel
          Top = 309
          Width = 746
          ExplicitTop = 309
          ExplicitWidth = 746
          inherited Bevel1: TBevel
            Width = 746
            ExplicitWidth = 740
          end
          inherited v_conferma: TRzBitBtn
            Left = 1
            Top = 10
            ExplicitLeft = 1
            ExplicitTop = 10
          end
          inherited v_esci: TRzBitBtn
            Left = 93
            ExplicitLeft = 93
          end
        end
      end
    end
    inherited tab_pagina2: TRzTabSheet
      ExplicitLeft = 4
      ExplicitTop = 27
      ExplicitWidth = 740
      ExplicitHeight = 339
      inherited pannello_esposizione: TRzPanel
        Width = 740
        Height = 339
        ExplicitWidth = 740
        ExplicitHeight = 339
        object v_griglia: trzdbgrid_go
          Left = 0
          Top = 0
          Width = 740
          Height = 339
          Align = alClient
          DataSource = tabella_ds
          DrawingStyle = gdsClassic
          options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Microsoft Sans Serif'
          TitleFont.Style = []
        end
      end
    end
  end
  inherited tabella: TMyQuery_go
    SQL.Strings = (
      'select *'
      'from gen')
    Left = 131
    Top = 6
  end
  inherited tabella_iva: TMyQuery_go
    Left = 141
    Top = 65518
  end
  inherited tabella_righe: TMyQuery_go
    SQL.Strings = (
      'select *'
      'from tpc')
  end
  inherited tabella_virtuale: TVirtualTable
    Data = {03000000000000000000}
  end
  object tabella_01: TMyTable
    Connection = ARC.arcdit
    BeforePost = tabella_01BeforePost
    Left = 300
    Top = 65531
  end
  object tabella_02: TMyTable
    Connection = ARC.arcdit
    BeforePost = tabella_02BeforePost
    Left = 375
    Top = 65531
  end
  object cfg: TMyTable
    TableName = 'cfg'
    Connection = ARC.arcdit
    Left = 415
    Top = 65531
  end
  object tabella_01_ds: TMyDataSource
    DataSet = tabella_01
    Left = 330
    Top = 65531
  end
  object tsm: TMyTable
    TableName = 'tsm'
    Connection = ARC.arcdit
    Left = 463
    Top = 3
  end
  object ADOEsatto: TADOConnection
    ConnectionString = 'FILE NAME=ESATTO.UDL'
    LoginPrompt = False
    Provider = 'ESATTO.UDL'
    Left = 331
    Top = 225
  end
  object tabella_esa_01: TADOQuery
    Connection = ADOEsatto
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from BANCHECLIENTI'
      'where'
      'Parte_Fissa='#39'BN'#39
      'order by Codice_Banca')
    Left = 392
    Top = 240
  end
  object tabella_clifor: TADOQuery
    Connection = ADOEsatto
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'ind_Clifor'
        Attributes = [paNullable]
        DataType = ftString
        Precision = 1
        Size = 1
        Value = Null
      end>
    SQL.Strings = (
      'select * from CLIENTIFORNITORI'
      'where'
      'Ind_ClienteFornitore=:ind_Clifor'
      'order by Codice_Cli_For')
    Left = 464
    Top = 240
  end
  object tabella_esa_02: TADOQuery
    Connection = ADOEsatto
    CursorType = ctStatic
    CommandTimeout = 60
    Parameters = <
      item
        Name = 'ind_Clifor'
        Attributes = [paNullable]
        DataType = ftString
        Precision = 1
        Size = 1
        Value = Null
      end>
    SQL.Strings = (
      'select * from CLIENTIFORNITORI'
      'where'
      'Ind_ClienteFornitore=:ind_Clifor'
      'order by Codice_Cli_For')
    Left = 384
    Top = 296
  end
  object tabella_clienti_forn: TADOQuery
    Connection = ADOEsatto
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'codice_nom'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select * from CLIENTIFORNITORI'
      'where'
      'codice_cli_for=:codice_nom'
      'order by Codice_Cli_For')
    Left = 464
    Top = 296
  end
  object tpa: TMyTable
    TableName = 'tpa'
    Connection = ARC.arcdit
    Left = 511
    Top = 3
  end
  object tco: TMyTable
    TableName = 'tco'
    Connection = ARC.arcdit
    Left = 551
    Top = 11
  end
  object tsa: TMyTable
    TableName = 'tsa'
    Connection = ARC.arcdit
    Left = 583
    Top = 11
  end
  object cpa: TMyTable
    TableName = 'cpa'
    Connection = ARC.arcdit
    Left = 623
    Top = 11
  end
  object cpv: TMyTable
    TableName = 'cpv'
    Connection = ARC.arcdit
    Left = 663
    Top = 11
  end
  object tca: TMyTable
    TableName = 'tca'
    Connection = ARC.arcdit
    Left = 511
    Top = 51
  end
  object query_02: TMyQuery_go
    Connection = ARC.arcdit
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 703
    Top = 14
  end
  object tabella_03: TMyTable
    Connection = ARC.arcdit
    BeforePost = tabella_03BeforePost
    Left = 375
    Top = 43
  end
  object tabella_esa_03: TADOQuery
    Connection = ADOEsatto
    CursorType = ctStatic
    CommandTimeout = 60
    Parameters = <
      item
        Name = 'ind_Clifor'
        Attributes = [paNullable]
        DataType = ftString
        Precision = 1
        Size = 1
        Value = Null
      end>
    SQL.Strings = (
      'select * from CLIENTIFORNITORI'
      'where'
      'Ind_ClienteFornitore=:ind_Clifor'
      'order by Codice_Cli_For')
    Left = 304
    Top = 304
  end
  object query_03: TMyQuery_go
    Connection = ARC.arcdit
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 703
    Top = 78
  end
  object dit: TMyQuery_go
    Connection = ARC.arc
    SQL.Strings = (
      'select * from dit'
      'where codice = :codice')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 643
    Top = 128
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'codice'
      end>
  end
end
