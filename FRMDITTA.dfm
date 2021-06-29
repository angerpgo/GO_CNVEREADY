object DITTA: TDITTA
  Left = 0
  Top = 0
  Width = 427
  Height = 286
  TabOrder = 0
  object RzLabel1: TRzLabel
    Left = 5
    Top = 5
    Width = 30
    Height = 13
    Caption = 'codice'
  end
  object Label1: TRzLabel
    Left = 5
    Top = 45
    Width = 126
    Height = 13
    Caption = 'ragione sociale / cognome'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label2: TRzLabel
    Left = 5
    Top = 85
    Width = 108
    Height = 13
    Caption = 'ragione sociale / nome'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label3: TRzLabel
    Left = 5
    Top = 125
    Width = 123
    Height = 13
    Caption = 'indirizzo (via, piazza, ecc.)'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label4: TRzLabel
    Left = 370
    Top = 125
    Width = 18
    Height = 13
    Caption = 'cap'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label5: TRzLabel
    Left = 5
    Top = 165
    Width = 33
    Height = 13
    Caption = 'localit'#224
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label6: TRzLabel
    Left = 305
    Top = 165
    Width = 43
    Height = 13
    Caption = 'provincia'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label7: TRzLabel
    Left = 5
    Top = 205
    Width = 46
    Height = 13
    Caption = 'partita iva'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label8: TRzLabel
    Left = 215
    Top = 205
    Width = 65
    Height = 13
    Caption = 'codice fiscale'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object Label9: TRzLabel
    Left = 5
    Top = 245
    Width = 38
    Height = 13
    Caption = 'telefono'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
    ShowAccelChar = False
    Transparent = True
  end
  object v_codice: trzedit_go
    Left = 5
    Top = 20
    Width = 56
    Height = 21
    CharCase = ecUpperCase
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 4
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 0
  end
  object v_descrizione1: trzedit_go
    Left = 5
    Top = 60
    Width = 411
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 40
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 1
  end
  object v_descrizione2: trzedit_go
    Left = 5
    Top = 100
    Width = 411
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 40
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 2
  end
  object v_via: trzedit_go
    Left = 5
    Top = 140
    Width = 356
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 40
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 3
  end
  object v_cap: trzedit_go
    Left = 370
    Top = 140
    Width = 46
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 10
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 4
  end
  object v_localita: trzedit_go
    Left = 5
    Top = 180
    Width = 291
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 30
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 5
  end
  object v_provincia: trzedit_go
    Left = 305
    Top = 180
    Width = 56
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 2
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 6
  end
  object v_partita_iva: trzedit_go
    Left = 5
    Top = 219
    Width = 201
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 20
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 7
  end
  object v_codice_fiscale: trzedit_go
    Left = 215
    Top = 219
    Width = 201
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 20
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 8
  end
  object v_telefono: trzedit_go
    Left = 5
    Top = 259
    Width = 201
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    MaxLength = 20
    ParentFont = False
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 9
  end
end
