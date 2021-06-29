object MENUGG: TMENUGG
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'CNVEREADY 10.01'
  ClientHeight = 260
  ClientWidth = 586
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TRzLabel
    Left = 5
    Top = 5
    Width = 20
    Height = 13
    Caption = 'ditta'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label6: TRzLabel
    Left = 135
    Top = 5
    Width = 53
    Height = 13
    Caption = 'descrizione'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TRzLabel
    Left = 5
    Top = 50
    Width = 41
    Height = 13
    Caption = 'esercizio'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TRzLabel
    Left = 135
    Top = 50
    Width = 47
    Height = 13
    Caption = 'data inizio'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label5: TRzLabel
    Left = 300
    Top = 50
    Width = 41
    Height = 13
    Caption = 'data fine'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label7: TRzLabel
    Left = 135
    Top = 95
    Width = 53
    Height = 13
    Caption = 'descrizione'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TRzLabel
    Left = 5
    Top = 95
    Width = 30
    Height = 13
    Caption = 'utente'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = []
    ParentFont = False
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
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 0
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
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 1
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
    Font.Name = 'Microsoft Sans Serif'
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
    Font.Name = 'Microsoft Sans Serif'
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
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 4
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
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 5
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
    Font.Name = 'Microsoft Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ReadOnly = True
    ReadOnlyColor = clBtnFace
    ReadOnlyColorOnFocus = True
    TabOrder = 6
  end
  object RzGroupBox2: TRzGroupBox
    Left = 0
    Top = 135
    Width = 586
    Height = 87
    Align = alBottom
    Caption = 'Conversione dati'
    TabOrder = 7
    object v_esporta_anagrafiche: TButton
      Left = 154
      Top = 30
      Width = 235
      Height = 37
      Caption = 'Conversione da E - E/READY'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = v_esporta_anagraficheClick
    end
  end
  object Panel4: TRzPanel
    Left = 0
    Top = 222
    Width = 586
    Height = 38
    Align = alBottom
    BorderOuter = fsNone
    TabOrder = 8
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 586
      Height = 2
      Align = alTop
      ExplicitWidth = 984
    end
    object v_conferma: TRzBitBtn
      Left = 5
      Top = 8
      Width = 86
      Height = 26
      Hint = 'conferma l'#39'elaborazione'
      Caption = 'Conferma'
      TabOrder = 0
      OnClick = v_confermaClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        20000000000000040000C30E0000C30E00000000000000000000000000000000
        00000000000000000000002B002A044D09A9035307E8014903FF014602FF0348
        05E8033E07A60021002A00000000000000000000000000000000000000000000
        0000003B0006045E0BA904730BFE14A025FF20AA32FF31B041FF36B144FF27AA
        35FF189A24FF014F03FE033905A900180006000000000000000000000000003B
        0006066F0FD20D9C1DFE1FAB34FF2CAD3DFF69C672FF6EC675FF52BB5CFF46B6
        50FF4DB855FF3DB147FF0E8716FE023A05D10018000600000000000000000886
        16AB0FA121FE1CAB30FF24AD38FF99DAA0FFFFFBF6FFFFFBF3FFFCF8EFFF5CC0
        64FF48B652FF4FB857FF4CB754FF0F8818FD033C06A900000000077A112D0A9D
        1AFE15AB2FFF1FAD36FFADE2B4FFFFFCFAFFFFFCF7FFFFFBF6FFFFFBF3FFBCE3
        B8FF42B54DFF49B752FF51BA59FF41B44BFF035705FE0021002A0C9E21AF16AD
        31FF1BAD36FFC5EDCCFFFFFEFCFFFFFEFBFFFFFCFAFFFFFCF7FFFFFBF6FFFFFB
        F3FF63C26AFF43B64FFF4BB754FF52BA5AFF1B9E28FF034206AF13A92AEC16B0
        37FF13AB30FF9EE0ACFFFFFFFFFFFFFEFEFF8AD796FFA9E0AFFFFFFCF7FFFFFB
        F6FFC7E9C5FF3EB44BFF44B650FF4CB856FF30AF40FF025206EB1BB033FC16B1
        39FF14AD34FF17AF34FFB4E7BFFF68CC7DFF1CAC35FF38B74CFFF8FBF4FFFFFC
        F7FFFFFBF6FF6BC775FF3FB54CFF46B752FF3DB44BFF025D05F921B73CFC18B4
        3DFF16B037FF14AD34FF13AC30FF14AB2FFF18AC33FF1EAC35FF96DAA0FFFFFC
        FAFFFFFCF7FFD4EDD0FF3CB54BFF41B54EFF38B248FF036506F926BD45EC1AB6
        41FF17B23BFF16B037FF15AF34FF13AC30FF16AC31FF1AAC34FF2EB545FFF3FA
        F0FFFFFCFAFFFFFCF7FF73CB7EFF3BB449FF2BAF3FFF046E0AEB2DC54FAF1FBA
        43FF18B53FFF17B23BFF16B137FF15AF34FF13AC30FF16AC32FF1BAC35FF84D5
        93FFFFFEFBFFFFFCFAFFE0F2DEFF3AB54BFF1BA72DFF066E0DAF35C9582A34CC
        56FE1BB843FF18B53FFF18B43BFF16B137FF15AF34FF13AC31FF18AC34FF24B1
        3DFFEDFAEEFFFFFEFBFFBDE7C1FF3BB74EFF079311FE0252042A0000000041D5
        69A928C44DFD1AB742FF18B53FFF18B43BFF16B137FF15AF35FF14AC32FF18AD
        34FF58C66EFF51C266FF25B03BFF16A928FE099016A600000000000000003BC9
        660645DA6FD129C44EFE1BB743FF19B63FFF18B43BFF16B137FF15AF35FF16AD
        33FF19AD35FF1CAF37FF18AD2DFE0C9D1CCF0066180600000000000000000000
        00003BC9660650E37DA941D768FF22BD46FF1BB741FF18B53DFF18B23BFF17B2
        38FF19B237FF1BB12FFF16AB2AA7009518060000000000000000000000000000
        0000000000000000000052E7802A47DE70A63ED464E639CF59FF32C94FFF29C1
        49E624BB42A518AA2B2A00000000000000000000000000000000}
    end
    object v_esci: TRzBitBtn
      Left = 100
      Top = 8
      Width = 86
      Height = 26
      Hint = 'esci dal programma'
      Caption = 'Esci'
      TabOrder = 1
      OnClick = v_esciClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        20000000000000040000C01E0000C01E00000000000000000000000000000000
        000000000000000000000000462A00025FA9000160E8000158FF000154FF0001
        50E9000144AB00002C2D00000000000000000000000000000000000000000000
        000000006606000375A900058BFE0613B5FF1123CAFF2438D5FF2B41D9FF1B2F
        D4FF0E1CB6FF000264FE000043AB000018060000000000000000000000000000
        660600068FD2010BADFE0E1EC5FF1D30CFFF2539D5FF2D41DAFF354BE0FF3E55
        E5FF455DEAFF354DE3FF08109CFE00004BD50000180600000000000000000007
        A0A6010BB1FE0B18BDFF6B75D7FF5360D3FF1F31CFFF263AD5FF2E43DCFF364C
        E1FF7782DEFF9299DDFF4960EAFF08109DFE000046AB000000000007A32A000C
        B6FE020CB0FF5560D0FFFFFBF3FFFFF8EFFF6B77D8FF2033D0FF283CD5FF7A84
        DCFFFFEFD7FFFFEDD1FF9199DDFF3850E5FF01036AFE0000312D000CBFAC010C
        B5FF0006A7FF3A45C6FFFFFCF8FFFFFBF3FFFFF8EFFF6D77D8FF6F7AD9FFFFF3
        E0FFFFF0DCFFFFEFD7FF7582DEFF4A62EAFF0F1CBAFF00014CAF0313C6EA0008
        ABFF0006A3FF0007A7FF4F59CEFFFFFCF8FFFFFBF3FFFFF8EFFFFFF7EAFFFFF4
        E5FFFFF3E0FF7A84DCFF3A50E1FF435AE5FF2236D5FF00015DEC061ACBF90006
        A0FF00059EFF0005A3FF0108A9FF4F5ACCFFFFFCF8FFFFFBF3FFFFF8EFFFFFF7
        EAFF6F7AD8FF2B3FD7FF3448DCFF3C52E1FF2F45DCFF000269FC0920D1F90005
        9DFF000499FF00049EFF0005A3FF505ACBFFFFFFFEFFFFFCF8FFFFFBF3FFFFF8
        EFFF6F7AD8FF2537D1FF2D41D7FF3549DCFF2A3FD8FF000270FC0C26D8E70007
        A4FF000395FF000499FF4E55C5FFFFFFFFFFFFFFFFFFFFFFFEFFFFFCF8FFFFFB
        F3FFFFF8EFFF707BD8FF2638D1FF2D41D7FF192BCCFF00037BEC0D2DDEA30311
        B8FF000290FF3237B2FFFFFFFFFFFFFFFFFFFFFFFFFF4F59CAFF525DCEFFFFFC
        F8FFFFFBF3FFFFF8EFFF5B68D5FF283AD1FF0A17BBFF000382AF153BE5271130
        DDFE000395FF5155BDFFFFFFFFFFFFFFFFFF4B52C4FF0005A1FF030BA9FF535D
        CEFFFFFCF8FFFFFBF3FF6D77D8FF1628C7FF0107A9FE00047A2D00000000153A
        E6A4081FCAFE01038FFF5154BDFF2D32B0FF000399FF00049DFF0006A1FF030C
        AAFF3742C5FF636CD3FF1321BFFF0512BAFD0006A0AB0000000000000000183B
        FF061840EACF091FC9FE000395FF000290FF000395FF000398FF00049DFF0107
        A3FF040DAAFF0611B2FF0513BCFE020DB4D00000950600000000000000000000
        0000183BFF061C46EFA5183CE3FF0413BAFF0006A3FF0005A0FF0006A3FF0008
        AAFF0310BAFF091AC5FF0413C0A7000095060000000000000000000000000000
        000000000000000000002658F72A1C43EDA3173AE5E61637E1FF1332DCFF0E28
        D7E60C22D3A5071CC92A00000000000000000000000000000000}
    end
  end
  object query: TMyQuery_go
    Connection = ARC.arcdit
    SQL.Strings = (
      'select * from tdocfin'
      'where'
      'codice=:codice')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 540
    Top = 55
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'codice'
        Value = nil
      end>
  end
end
