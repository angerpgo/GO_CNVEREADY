object main: Tmain
  Left = 324
  Top = 233
  Caption = 'main'
  ClientHeight = 327
  ClientWidth = 581
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TRzLabel
    Left = 5
    Top = 5
    Width = 20
    Height = 13
    Caption = 'ditta'
  end
  object Label2: TRzLabel
    Left = 5
    Top = 50
    Width = 41
    Height = 13
    Caption = 'esercizio'
  end
  object Label3: TRzLabel
    Left = 5
    Top = 95
    Width = 30
    Height = 13
    Caption = 'utente'
  end
  object Label4: TRzLabel
    Left = 135
    Top = 50
    Width = 47
    Height = 13
    Caption = 'data inizio'
  end
  object Label5: TRzLabel
    Left = 300
    Top = 50
    Width = 41
    Height = 13
    Caption = 'data fine'
  end
  object Label6: TRzLabel
    Left = 135
    Top = 5
    Width = 53
    Height = 13
    Caption = 'descrizione'
  end
  object Label7: TRzLabel
    Left = 135
    Top = 95
    Width = 53
    Height = 13
    Caption = 'descrizione'
  end
  object GroupBox2: TGroupBox
    Left = -2
    Top = 149
    Width = 575
    Height = 77
    Caption = 'CONVERSIONE DATI'
    TabOrder = 8
    object v_cnvesa: TButton
      Left = 12
      Top = 29
      Width = 556
      Height = 25
      Caption = 'CONVERSIONE DA eReady'
      TabOrder = 0
      OnClick = v_cnvesaClick
    end
  end
  object ComboEdit1: trzedit_go
    Left = 5
    Top = 20
    Width = 121
    Height = 21
    TabStop = False
    Text = 'ditta'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 0
  end
  object ComboEdit2: trzedit_go
    Left = 5
    Top = 65
    Width = 121
    Height = 21
    TabStop = False
    Text = 'esercizio'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 1
  end
  object ComboEdit3: trzedit_go
    Left = 5
    Top = 110
    Width = 121
    Height = 21
    TabStop = False
    Text = 'utente'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 2
  end
  object ComboEdit4: trzedit_go
    Left = 135
    Top = 65
    Width = 121
    Height = 21
    TabStop = False
    Text = 'data inizio'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 3
  end
  object ComboEdit5: trzedit_go
    Left = 300
    Top = 65
    Width = 121
    Height = 21
    TabStop = False
    Text = 'data fine'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 4
  end
  object ComboEdit6: trzedit_go
    Left = 135
    Top = 20
    Width = 426
    Height = 21
    TabStop = False
    Text = 'descrizione ditta'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 5
  end
  object ComboEdit7: trzedit_go
    Left = 135
    Top = 110
    Width = 426
    Height = 21
    TabStop = False
    Text = 'descrizione utente'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 6
  end
  object v_esegui_go: TRzButton
    Left = 5
    Top = 278
    Width = 561
    Height = 36
    Caption = 'esegui Gestionale Open'
    TabOrder = 7
    OnClick = v_esegui_goClick
  end
  object ggg: TMyQuery
    Connection = ARC.arc
    SQL.Strings = (
      'select *'
      'from ggg'
      'wHere codice = '#39'G'#39
      ' ')
    ReadOnly = True
    Left = 485
    Top = 60
  end
  object dit: TMyQuery
    Connection = ARC.arc
    SQL.Strings = (
      'select min(data_ora_creazione) data_ora_creazione'
      'from dit'
      'where codice <> '#39'DEMO'#39
      ' ')
    ReadOnly = True
    Left = 450
    Top = 60
  end
end
