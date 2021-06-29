object CLOUD: TCLOUD
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 222
  Width = 566
  object arc_standard: TMyConnection_go
    Database = 'arc'
    Options.UseUnicode = True
    DataTypeMap = <
      item
        DBType = 225
        FieldType = ftWideMemo
      end>
    Username = 'root'
    Server = 'localhost'
    LoginPrompt = False
    Left = 55
    Top = 10
  end
  object cod_cloud: TMyQuery_go
    Connection = arc_standard
    SQL.Strings = (
      'select substring(schema_name from 17) cliente_cloud'
      'from information_schema.schemata'
      'where schema_name like '#39'arc_ordinamento_%'#39)
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 300
    Top = 10
  end
  object copia: TMyScript_go
    Connection = arc_standard
    Left = 230
    Top = 10
  end
  object dizionario: TMyQuery_go
    Connection = arc_standard
    SQL.Strings = (
      'select table_name'
      'from information_schema.tables'
      'where lower(table_schema) = :table_schema'
      'and table_type = '#39'BASE TABLE'#39
      'order by table_name')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 300
    Top = 65
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'table_schema'
      end>
  end
  object arc_demo_standard: TMyConnection_go
    Database = 'arc_demo'
    Options.UseUnicode = True
    DataTypeMap = <
      item
        DBType = 225
        FieldType = ftWideMemo
      end>
    Username = 'root'
    Server = 'localhost'
    LoginPrompt = False
    Left = 55
    Top = 70
  end
  object arc_cloud: TMyConnection_go
    Options.UseUnicode = True
    DataTypeMap = <
      item
        DBType = 225
        FieldType = ftWideMemo
      end>
    Username = 'root'
    Server = 'localhost'
    LoginPrompt = False
    Left = 155
    Top = 10
  end
  object dit: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select * from dit;')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 360
    Top = 10
  end
  object ese: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select * from ese;')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 360
    Top = 65
  end
  object arc_dit_cloud: TMyConnection_go
    Options.UseUnicode = True
    DataTypeMap = <
      item
        DBType = 225
        FieldType = ftWideMemo
      end>
    Username = 'root'
    Server = 'localhost'
    LoginPrompt = False
    Left = 155
    Top = 70
  end
  object prg: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select codice'
      'from prg'
      'where eseguibile_menu = '#39'menu'#39)
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 300
    Top = 115
  end
  object user: TMyQuery_go
    Connection = arc_standard
    SQL.Strings = (
      'select mysql.user.user'
      'from mysql.user'
      'where mysql.user.user <> '#39'root'#39)
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 360
    Top = 115
  end
  object ese_dit: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select * from ese'
      'where dit_codice = :dit_codice')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 360
    Top = 165
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'dit_codice'
      end>
  end
  object utn: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select *'
      'from utn')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 300
    Top = 165
  end
  object mysql_user: TMyScript_go
    Connection = arc_standard
    Left = 230
    Top = 70
  end
  object grant: TMyScript_go
    Connection = arc_standard
    Left = 230
    Top = 130
  end
  object utn_cliente: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select codice, descrizione'
      'from utn')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 420
    Top = 10
  end
  object dit_cliente: TMyQuery_go
    Connection = arc_cloud
    SQL.Strings = (
      'select codice, concat('#39'arc_'#39', codice) db,'
      ' concat(descrizione1, descrizione2) descrizione'
      'from dit')
    Options.DefaultValues = True
    Options.TrimVarChar = True
    Left = 417
    Top = 70
  end
  object utn_cliente_ds: TMyDataSource
    DataSet = utn_cliente
    Left = 495
    Top = 10
  end
  object dit_cliente_ds: TMyDataSource
    DataSet = dit_cliente
    Left = 495
    Top = 70
  end
end
