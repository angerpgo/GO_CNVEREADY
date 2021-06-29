unit DMCLOUD;

interface

uses
  System.SysUtils, System.Classes, Data.DB, DBAccess, MyAccess, query_go, MemDS,
  DAScript, MyScript, rzlabel, vcl.forms, ZZDUPLICA_DITTA, FRMDITTA;

type
  TCLOUD = class(TDataModule)
    arc_standard: TMyConnection_go;
    cod_cloud: TMyQuery_go;
    copia: TMyScript_go;
    dizionario: TMyQuery_go;
    arc_demo_standard: TMyConnection_go;
    arc_cloud: TMyConnection_go;
    dit: TMyQuery_go;
    ese: TMyQuery_go;
    arc_dit_cloud: TMyConnection_go;
    prg: TMyQuery_go;
    user: TMyQuery_go;
    ese_dit: TMyQuery_go;
    utn: TMyQuery_go;
    mysql_user: TMyScript_go;
    grant: TMyScript_go;
    utn_cliente: TMyQuery_go;
    dit_cliente: TMyQuery_go;
    utn_cliente_ds: TMyDataSource;
    dit_cliente_ds: TMyDataSource;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  protected
    label_notifiche: trzlabel;
    duplica_ditta: tduplica_ditta;
    procedure duplica(database_origine, database_nuovo: string);
    procedure copia_tabella(database_origine, database_nuovo, tabella: string);
    procedure crea_record_dit(dati_ditta: tditta);
    procedure crea_record_ese(codice_ditta: string);
    procedure crea_db(nome_db: string);
  public
    lista_clienti_cloud, lista_prg_cloud, lista_dit_cloud, lista_ese_cloud: tstringlist;
    procedure connetti_arc_cloud(cliente: string);
    procedure carica_lista_clienti_cloud;
    procedure carica_lista_prg_cloud;
    procedure carica_lista_dit_cloud;
    procedure carica_lista_ese_cloud(dit_codice: string);
    procedure assegna_grant(cliente, utente: string);
    procedure riassegna_grant_utenti(cliente: string);
    procedure genera_utente(utn_codice, utn_descrizione, prg_base, dit, ese: string);
    function controllo_codice_utente(utn_codice: string): boolean;
    function esiste_ditta(codice_ditta: string): boolean;
    procedure genera_nuovo_cliente_cloud(suffisso_cliente: string; notifiche: trzlabel);
    procedure genera_nuova_ditta(suffisso_cliente, cartella_installazione: string; dati_ditta: tditta; notifiche: trzlabel);
    procedure aggiorna_utenti_ditte_cliente(cliente: string);
  end;

var
  CLOUD: TCLOUD;

implementation

{%CLASSGROUP 'System.Classes.TPersistent'}

{$R *.dfm}

procedure TCLOUD.connetti_arc_cloud(cliente: string);
begin
  arc_cloud.connected := false;
  arc_cloud.password := arc_standard.password;
  arc_cloud.server := arc_standard.server;
  arc_cloud.port := arc_standard.port;
  arc_cloud.database := 'arc_' + cliente;
  arc_cloud.connected := true;
end;

procedure TCLOUD.genera_nuova_ditta(suffisso_cliente, cartella_installazione: string;
 dati_ditta: tditta; notifiche: trzlabel);
begin
  arc_demo_standard.password := arc_standard.password;
  arc_demo_standard.port := arc_standard.port;
  arc_demo_standard.server := arc_standard.server;
  arc_demo_standard.connected := true;

  connetti_arc_cloud(suffisso_cliente);

  crea_record_dit(dati_ditta);

  // duplica database ditta e crea cartella documenti
  crea_db('arc_' + dati_ditta.v_codice.text);
  duplica_ditta := tduplica_ditta.create(arc_standard, arc_demo_standard, 'DEMO');
  duplica_ditta.v_notifiche := notifiche;
  duplica_ditta.copia_ditta_attiva(false, dati_ditta.v_codice.text);
  duplica_ditta.crea_cartella_documenti(cartella_installazione + '\exe\', dati_ditta.v_codice.text);

  // se necessario genero l'esercizio corrente per permettere di accedere alla ditta.
  crea_record_ese(dati_ditta.v_codice.text);

  notifiche.caption := 'GENERAZIONE COMPLETATA';
  application.processmessages;
end;

procedure TCLOUD.crea_record_ese(codice_ditta: string);
var
  esercizio: string;
begin
  ese.open;
  if ese.isempty then
  begin
    esercizio := formatdatetime('yyyy', date);
    ese.append;
    ese.fieldbyname('dit_codice').asstring := codice_ditta;
    ese.fieldbyname('codice').asstring := esercizio;
    ese.fieldbyname('descrizione').asstring := 'ESERCIZIO ' + esercizio;
    ese.fieldbyname('data_inizio').asdatetime := strtodatetime('01/01/' + esercizio);
    ese.fieldbyname('data_fine').asdatetime := strtodatetime('31/12/' + esercizio);
    ese.fieldbyname('data_bilancio').asdatetime := strtodatetime('31/12/' + esercizio);
    ese.fieldbyname('data_ora').asdatetime := now;
    ese.fieldbyname('data_ora_creazione').asdatetime := now;
    ese.post;
  end;
  ese.close;
end;

function TCLOUD.esiste_ditta(codice_ditta: string): boolean;
begin
  result := true;

  arc_cloud.connected := false;
  arc_cloud.password := arc_standard.password;
  arc_cloud.port := arc_standard.port;
  arc_cloud.server := arc_standard.server;
  arc_cloud.database := 'arc_' + lowercase(codice_ditta);

  try
    arc_cloud.connected := true;
  except
    result := false;
  end;
end;

procedure TCLOUD.crea_record_dit(dati_ditta: tditta);
begin
  dit.open;
  dit.append;
  dit.fieldbyname('codice').asstring := dati_ditta.v_codice.text;
  dit.fieldbyname('descrizione1').asstring := dati_ditta.v_descrizione1.text;
  dit.fieldbyname('descrizione2').asstring := dati_ditta.v_descrizione2.text;
  dit.fieldbyname('via').asstring := dati_ditta.v_via.text;
  dit.fieldbyname('cap').asstring := dati_ditta.v_cap.text;
  dit.fieldbyname('citta').asstring := dati_ditta.v_localita.text;
  dit.fieldbyname('provincia').asstring := dati_ditta.v_provincia.text;
  dit.fieldbyname('partita_iva').asstring := dati_ditta.v_partita_iva.text;
  dit.fieldbyname('codice_fiscale').asstring := dati_ditta.v_codice_fiscale.text;
  dit.fieldbyname('telefono').asstring := dati_ditta.v_telefono.text;
  dit.fieldbyname('data_ora').asdatetime := now;
  dit.fieldbyname('data_ora_creazione').asdatetime := now;
  dit.post;
  dit.close;
end;

procedure TCLOUD.genera_nuovo_cliente_cloud(suffisso_cliente: string; notifiche: trzlabel);
begin
  label_notifiche := notifiche;

  arc_standard.connected := true;
  duplica('arc', 'arc_' + suffisso_cliente);
  duplica('arc_ordinamento', 'arc_ordinamento_' + suffisso_cliente);

  label_notifiche.caption := '...';
  application.processmessages;
end;

procedure TCLOUD.crea_db(nome_db: string);
begin
  copia.sql.text := 'create database if not exists ' + lowercase(nome_db) + ' default character set utf8 collate utf8_bin;';
  copia.execute;
end;

procedure TCLOUD.duplica(database_origine, database_nuovo: string);
begin
  crea_db(database_nuovo);

  dizionario.close;
  dizionario.parambyname('table_schema').asstring := database_origine;
  dizionario.open;

  while not dizionario.eof do
  begin
    copia_tabella(database_origine, database_nuovo, trim(dizionario.fieldbyname('table_name').asstring));
    dizionario.next;
  end;
end;

procedure TCLOUD.copia_tabella(database_origine, database_nuovo, tabella: string);
begin
  // crea sql per creazione tabella
  if (lowercase(copy(tabella, 1, 3)) <> 'rb_') and (lowercase(copy(tabella, 1, 5)) <> 'vista') then
  begin
    if label_notifiche <> nil then
    begin
      label_notifiche.caption := database_nuovo + '- tabella in elaborazione: ' + tabella;
      application.processmessages;
    end;

    if (tabella = 'ggg') or (tabella = 'est') or (tabella = 'deslin') or (tabella = 'imptpc')
      or (tabella = 'impgen') or (tabella = 'tpv') or (tabella = 'tce') or (tabella = 'tbl')
      or (tabella = 'mnu') or (tabella = 'prg') or (tabella = 'cpc') or (tabella = 'comuni')
      or (tabella = 'abi') or (tabella = 'cap') then
    begin
      copia.sql.text :=
        'create table ' + database_nuovo + '.' + tabella +
        ' select * from ' + database_origine + '.' + tabella;
    end
    else if tabella = 'vis' then
    begin
      copia.sql.text :=
        'create table ' + database_nuovo + '.' + tabella +
        ' select * from ' + database_origine + '.' + tabella + ' where utn_codice = ''OPEN''';
    end
    else
    begin
      copia.sql.text :=
        'create table ' + database_nuovo + '.' + tabella + ' like ' + database_origine + '.' + tabella;
    end;

    copia.execute;
  end;
end;

procedure TCLOUD.DataModuleCreate(Sender: TObject);
begin
  lista_clienti_cloud := tstringlist.create;
  lista_prg_cloud := tstringlist.create;
  lista_dit_cloud := tstringlist.create;
  lista_ese_cloud := tstringlist.create;
end;

procedure TCLOUD.DataModuleDestroy(Sender: TObject);
begin
  freeandnil(lista_clienti_cloud);
  freeandnil(lista_prg_cloud);
  freeandnil(lista_dit_cloud);
  freeandnil(lista_ese_cloud);
  if duplica_ditta <> nil then
  begin
    duplica_ditta.free;
  end;
end;

function TCLOUD.controllo_codice_utente(utn_codice: string): boolean;
begin
  result := true;

  user.close;
  user.open;
  if user.locate('user', utn_codice, []) then
  begin
    result := false;
  end;
end;

procedure TCLOUD.carica_lista_dit_cloud;
var
  i: integer;
begin
  dit.close;
  dit.open;
  lista_dit_cloud.clear;

  while not dit.eof do
  begin
    lista_dit_cloud.add(dit.fieldbyname('codice').asstring);
    dit.next;
  end;
end;

procedure TCLOUD.assegna_grant(cliente, utente: string);
begin
  // grant sul database arc
  grant.sql.text := 'grant all on arc_' + cliente + '.* to ' + quotedstr(utente)
    + '@''%''; flush privileges;';
  grant.execute;

  // grant sul database arc_ordinamento
  grant.sql.text := 'grant all on arc_ordinamento_' + cliente + '.* to ' + quotedstr(utente)
    + '@''%''; flush privileges;';
  grant.execute;

  // grant su tutti i database del cliente
  dit.close;
  dit.open;
  while not dit.eof do
  begin
    grant.sql.text := 'grant all on arc_' + dit.fieldbyname('codice').asstring + '.* to ' + quotedstr(utente)
    + '@''%''; flush privileges;';
    grant.execute;

    dit.next;
  end;
end;

procedure TCLOUD.riassegna_grant_utenti(cliente: string);
begin
  connetti_arc_cloud(cliente);

   // grant su tutti i database del cliente
  utn.close;
  utn.open;
  while not utn.eof do
  begin
    assegna_grant(cliente, utn.fieldbyname('codice').asstring);
    utn.next;
  end;
end;

procedure TCLOUD.genera_utente(utn_codice, utn_descrizione, prg_base, dit, ese: string);
begin
  utn.close;
  utn.open;
  utn.append;
  utn.fieldbyname('codice').asstring := utn_codice;
  utn.fieldbyname('descrizione').asstring := utn_descrizione;
  utn.fieldbyname('programma_base').asstring := prg_base;
  utn.fieldbyname('dit_codice').asstring := dit;
  utn.fieldbyname('ese_codice').asstring := ese;
  utn.fieldbyname('supervisore').asstring := 'si';
  utn.fieldbyname('data_modifica_password').asdatetime := date;
  utn.fieldbyname('data_ultimo_accesso').asdatetime := date;
  utn.fieldbyname('data_ora_creazione').asdatetime := now;
  utn.fieldbyname('data_ora').asdatetime := now;
  utn.post;

  mysql_user.sql.text := 'create user ' + quotedstr(utn_codice) + '@''%'' identified by ''GO''';
  mysql_user.execute;
end;

procedure TCLOUD.carica_lista_ese_cloud(dit_codice: string);
var
  i: integer;
begin
  ese_dit.close;
  ese_dit.parambyname('dit_codice').asstring := dit_codice;
  ese_dit.open;
  lista_ese_cloud.clear;

  while not ese_dit.eof do
  begin
    lista_ese_cloud.add(ese_dit.fieldbyname('codice').asstring);
    ese_dit.next;
  end;
end;

procedure TCLOUD.carica_lista_prg_cloud;
var
  i: integer;
begin
  prg.close;
  prg.open;
  lista_prg_cloud.clear;

  while not prg.eof do
  begin
    lista_prg_cloud.add(prg.fieldbyname('codice').asstring);
    prg.next;
  end;
end;

procedure TCLOUD.carica_lista_clienti_cloud;
var
  i: integer;
begin
  cod_cloud.close;
  cod_cloud.open;
  lista_clienti_cloud.clear;

  while not cod_cloud.eof do
  begin
    lista_clienti_cloud.add(cod_cloud.fieldbyname('cliente_cloud').asstring);
    cod_cloud.next;
  end;
end;

procedure TCLOUD.aggiorna_utenti_ditte_cliente(cliente: string);
begin
  utn_cliente.close;
  dit_cliente.close;
  connetti_arc_cloud(cliente);
  utn_cliente.open;
  dit_cliente.open;
end;

end.
