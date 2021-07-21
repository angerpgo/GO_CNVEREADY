//
// Versione 11.01
//
(*
  //  programma di generazione md5
  https://emn178.github.io/online-tools/md5_checksum.html

  //  FTP Bespoken
  ftp://95.174.1.252
  utente  gestionaleop_ftp
  password  Uzke_361
*)
unit DMARC;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Vcl.Dialogs, ComCtrls, StdCtrls, Buttons, ExtCtrls, dateutils, DB, ImgList, Menus,
  StrUtils, madExcept, shellapi, query_go, MyAccess, memdata, cxLocalization,
  cxGraphics, IdMessage, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP, VirtualTable,
  DBAccess, MemDS, types, MySqlApi, Math, IniFiles, idGlobal, cxGridTableView,
  cxGridDBTableView, cxGrid, idGlobalProtocols, FMTBcd, SqlExpr, XPMan, rzLabel,
  Registry, IdHTTP, IdAttachmentFile, WinInet, ComObj, ActiveX, RzDBGrid, RzEdit,
  RzPanel, RzDBEdit, RzListVw, RzTreeVw, RzDBChk, RzRadChk, RzButton, RzSplit,
  RzCmboBx, RzPrgres, RzSpnEdt, RzShellDialogs, RzDBCmbo, raizeedit_go, cxStyles,
  cxClasses, RzTabs, ipwcore, ipwmx, MessageDigest_5, DAScript, MyScript,
  IdUserPassProvider, IdSASL, IdSASLUserPass, IdSASLLogin, IdSASLPlain, sYSTEM.zIP,
  Vcl.Grids, Vcl.DBGrids, XLSSheetData5, XLSReadWriteII5, CREncryption, sYSTEM.ImageList,
  sYSTEM.UITypes, sYSTEM.TypInfo, GGBASE, RzCommon, sYSTEM.json, Vcl.Themes, Vcl.Styles,

  // skins devexpress
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2010Black,
  dxSkinSharp, dxSkinSilver, dxSkinWhiteprint,
  // skins devexpress fine

  cxLookAndFeels, dxSkinsForm, dxSkinsCore, PsAPI, TlHelp32, ipwtypes, cxImageList;

type
  tgridhelper = class helper for trzdbgrid_go
    function columnbyname(const aname: string): tcolumn;
  end;

  intero = longint;

  scadenze = record
    data_scadenza: tdatetime;
    importo_scadenza: double;
    importo_scadenza_euro: double;
    tipo_pagamento: string;
  end;

  campi = record
    nome_campo: string;
    tipo_campo: string;
    size_campo: word;
  end;

  contropartite = record
    tiv_codice_contropartita_zucchetti: string;
    gen_codice_contropartita: string;
    importo_contropartita: double;
    importo_iva_contropartita: double;
    cambio: double;
  end;

  Pgriglia_devex = ^TcxGridDBTableView;

type
  TARC = class(TDataModule)
    pop_data_ora: TPopupMenu;
    Oggi1: TMenuItem;
    ore1: TMenuItem;
    N1: TMenuItem;
    N2: TMenuItem;
    query_ri: tmyquery_go;
    pop_arc: TPopupMenu;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    N4: TMenuItem;
    pop_arc_art: TPopupMenu;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem9: TMenuItem;
    N3: TMenuItem;
    Contropartitevendita1: TMenuItem;
    cpa1: TMenuItem;
    N1Listinidivendita1: TMenuItem;
    Listinidiacquisto1: TMenuItem;
    Codiciabarre1: TMenuItem;
    Depositi1: TMenuItem;
    Lifo1: TMenuItem;
    N5: TMenuItem;
    Schedamovimentazioni1: TMenuItem;
    N6: TMenuItem;
    Documenti1: TMenuItem;
    Foto1: TMenuItem;
    pop_arc_cli: TPopupMenu;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem_progressivi_contabili: TMenuItem;
    MenuItem_scheda_contabile: TMenuItem;
    Fido1: TMenuItem;
    N7: TMenuItem;
    Documenti2: TMenuItem;
    pop_arc_frn: TPopupMenu;
    MenuItem33: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem42: TMenuItem;
    N10: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem47: TMenuItem;
    N8: TMenuItem;
    Documenti3: TMenuItem;
    pop_arc_gen: TPopupMenu;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem34: TMenuItem;
    N9: TMenuItem;
    Documenti4: TMenuItem;
    collegamenti_archivio: tmyquery_go;
    collegamenti_archivio_arc: tmyquery_go;
    Situazionedisponibilit1: TMenuItem;
    Lotti1: TMenuItem;
    pop_arc_cms: TPopupMenu;
    MenuItem68: TMenuItem;
    MenuItem69: TMenuItem;
    MenuItem70: TMenuItem;
    MenuItem71: TMenuItem;
    MenuItem72: TMenuItem;
    MenuItem73: TMenuItem;
    MenuItem74: TMenuItem;
    Situaizonepartitario1: TMenuItem;
    Situazionepartitario1: TMenuItem;
    Situazionepartitario2: TMenuItem;
    Estrattoconto1: TMenuItem;
    pop_arc_cen: TPopupMenu;
    MenuItem75: TMenuItem;
    MenuItem76: TMenuItem;
    MenuItem77: TMenuItem;
    MenuItem78: TMenuItem;
    MenuItem79: TMenuItem;
    MenuItem80: TMenuItem;
    MenuItem81: TMenuItem;
    MenuItem83: TMenuItem;
    Insoluti1: TMenuItem;
    Ritardipagamento1: TMenuItem;
    N11: TMenuItem;
    Ordini1: TMenuItem;
    Documentidivendita1: TMenuItem;
    Preventivi2: TMenuItem;
    N12: TMenuItem;
    Ordini2: TMenuItem;
    Documentidiacquisto1: TMenuItem;
    archiviocollegato011: TMenuItem;
    archiviocollegato021: TMenuItem;
    Accessorio1: TMenuItem;
    Accessori1: TMenuItem;
    Equivalente1: TMenuItem;
    Equivalenti1: TMenuItem;
    pop_preferiti: TPopupMenu;
    Contratti1: TMenuItem;
    Contratti2: TMenuItem;
    N13: TMenuItem;
    Contrattiarticoli1: TMenuItem;
    N14: TMenuItem;
    N15: TMenuItem;
    Ultimiprezziacquisto1: TMenuItem;
    Cruscotto1: TMenuItem;
    Cruscotto2: TMenuItem;
    Analisiordini1: TMenuItem;
    Analisiordini2: TMenuItem;
    N16: TMenuItem;
    N17: TMenuItem;
    N18: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    msg: tmyquery_go;
    pop_arc_csp: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem89: TMenuItem;
    MenuItem90: TMenuItem;
    MenuItem91: TMenuItem;
    MenuItem92: TMenuItem;
    MenuItem93: TMenuItem;
    Cruscotto3: TMenuItem;
    prg: tmyquery_go;
    acp: tmyquery_go;
    arc: TMyConnection_go;
    arcsor: TMyConnection_go;
    SaveDialog: Tsavedialog;
    sor_xls: TVirtualTable;
    assnum_cnt: tmyquery_go;
    esistenza_numerazione: tmyquery_go;
    dit: tmyquery_go;
    query_pra: tmyquery_go;
    abp: tmyquery_go;
    Contatti1: TMenuItem;
    Contatti2: TMenuItem;
    lin: tmyquery_go;
    pop_arc_nom: TPopupMenu;
    MenuItem27: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem82: TMenuItem;
    MenuItem84: TMenuItem;
    MenuItem85: TMenuItem;
    MenuItem86: TMenuItem;
    MenuItem87: TMenuItem;
    MenuItem97: TMenuItem;
    utn: tmyquery_go;
    pit: tmyquery_go;
    Documenti5: TMenuItem;
    situazioneenasarco1: TMenuItem;
    documenti6: TMenuItem;
    N23: TMenuItem;
    server_smtp: TIdSMTP;
    open_ssl: TIdSSLIOHandlerSocketOpenSSL;
    v_mail: TIdMessage;
    Documentiextra1: TMenuItem;
    N24: TMenuItem;
    documenticollegati1: TMenuItem;
    N25: TMenuItem;
    Documenticollegati2: TMenuItem;
    N26: TMenuItem;
    Documenticollegati3: TMenuItem;
    N27: TMenuItem;
    Documenticollegati4: TMenuItem;
    N28: TMenuItem;
    Documenticollegati5: TMenuItem;
    N29: TMenuItem;
    Documenticollegati6: TMenuItem;
    N30: TMenuItem;
    Documenticollegati7: TMenuItem;
    N31: TMenuItem;
    Documenticollegati8: TMenuItem;
    lettore: TVirtualTable;
    lettoreart_codice: TStringField;
    lettorequantita: TFloatField;
    lettorequantita_documento: TFloatField;
    lettoreart_descrizione: TStringField;
    immagine_16: TcxImageList;
    immagine_24: TcxImageList;
    cxTraduttore: TcxLocalizer;
    archiviocollegato031: TMenuItem;
    archiviocollegato041: TMenuItem;
    ipwmx: TipwMX;
    cxStyleRepository1: TcxStyleRepository;
    st_Window: TcxStyle;
    st_Lime: TcxStyle;
    st_Aqua: TcxStyle;
    st_Yellow: TcxStyle;
    st_Fuchsia: TcxStyle;
    st_Verde: TcxStyle;
    st_Btnface: TcxStyle;
    arcdit: TMyConnection_go;
    xlswrite: TXLSReadWriteII5;
    bilanciodicommessa1: TMenuItem;
    query_mag_esercizio: tmyquery_go;
    imp_gen: tmyquery_go;
    prf_gruppo: tmyquery_go;
    esiste_trigger_mysql: tmyquery_go;
    dxSkinController1: TdxSkinController;
    cxImageList1: TcxImageList;
    tips: tmyquery_go;
    st_Link: TcxStyle;
    st_bold: TcxStyle;
    stile_01: TcxStyle;
    stile_02: TcxStyle;
    stile_03: TcxStyle;
    stile_04: TcxStyle;
    prs_generatore: TMyStoredProc;
    procedure Documenti1Click(Sender: TObject);
    procedure Foto1Click(Sender: TObject);
    procedure Contropartitevendita1Click(Sender: TObject);
    procedure cpa1Click(Sender: TObject);
    procedure N1Listinidivendita1Click(Sender: TObject);
    procedure Listinidiacquisto1Click(Sender: TObject);
    procedure Codiciabarre1Click(Sender: TObject);
    procedure Depositi1Click(Sender: TObject);
    procedure Lifo1Click(Sender: TObject);
    procedure Schedamovimentazioni1Click(Sender: TObject);
    procedure MenuItem31Click(Sender: TObject);
    procedure MenuItem_progressivi_contabiliClick(Sender: TObject);
    procedure MenuItem_scheda_contabileClick(Sender: TObject);
    procedure Fido1Click(Sender: TObject);
    procedure Documenti2Click(Sender: TObject);
    procedure MenuItem46Click(Sender: TObject);
    procedure MenuItem47Click(Sender: TObject);
    procedure Documenti3Click(Sender: TObject);
    procedure MenuItem32Click(Sender: TObject);
    procedure MenuItem34Click(Sender: TObject);
    procedure Documenti4Click(Sender: TObject);
    procedure MenuItem44Click(Sender: TObject);
    procedure MenuItem45Click(Sender: TObject);
    procedure MenuItem54Click(Sender: TObject);
    procedure MenuItem62Click(Sender: TObject);
    procedure menu_item_utn_disabilita_ditteClick(Sender: TObject);
    procedure menu_item_utn_disabilita_programmiClick(Sender: TObject);
    procedure manu_item_utn_Assegna_stampantiClick(Sender: TObject);
    procedure Situazionedisponibilit1Click(Sender: TObject);
    procedure Lotti1Click(Sender: TObject);
    procedure MenuItem_cms_tipologieClick(Sender: TObject);
    procedure Situazionepartitario2Click(Sender: TObject);
    procedure Situazionepartitario1Click(Sender: TObject);
    procedure Situaizonepartitario1Click(Sender: TObject);
    procedure Estrattoconto1Click(Sender: TObject);
    procedure MenuItem83Click(Sender: TObject);
    procedure Insoluti1Click(Sender: TObject);
    procedure Ritardipagamento1Click(Sender: TObject);
    procedure Ordini1Click(Sender: TObject);
    procedure Documentidivendita1Click(Sender: TObject);
    procedure Preventivi2Click(Sender: TObject);
    procedure Ordini2Click(Sender: TObject);
    procedure Documentidiacquisto1Click(Sender: TObject);
    procedure archiviocollegato011Click(Sender: TObject);
    procedure pop_arc_artPopup(Sender: TObject);
    procedure archiviocollegato021Click(Sender: TObject);
    procedure racciabilitdocumenti1Click(Sender: TObject);
    procedure Accessorio1Click(Sender: TObject);
    procedure Accessori1Click(Sender: TObject);
    procedure Equivalente1Click(Sender: TObject);
    procedure Equivalenti1Click(Sender: TObject);
    procedure Contratti1Click(Sender: TObject);
    procedure Contratti2Click(Sender: TObject);
    procedure Contrattiarticoli1Click(Sender: TObject);
    procedure Ultimiprezziacquisto1Click(Sender: TObject);
    procedure Cruscotto1Click(Sender: TObject);
    procedure Cruscotto2Click(Sender: TObject);
    procedure Analisiordini1Click(Sender: TObject);
    procedure Analisiordini2Click(Sender: TObject);
    procedure MenuItem93Click(Sender: TObject);
    procedure Cruscotto3Click(Sender: TObject);
    procedure DataModuleCreate(Sender: TObject);
    procedure Contatti1Click(Sender: TObject);
    procedure Contatti2Click(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure MenuItem97Click(Sender: TObject);
    procedure Documenti5Click(Sender: TObject);
    procedure situazioneenasarco1Click(Sender: TObject);
    procedure documenti6Click(Sender: TObject);
    procedure Documentiextra1Click(Sender: TObject);
    procedure documenticollegati1Click(Sender: TObject);
    procedure Documenticollegati2Click(Sender: TObject);
    procedure Documenticollegati3Click(Sender: TObject);
    procedure Documenticollegati4Click(Sender: TObject);
    procedure Documenticollegati5Click(Sender: TObject);
    procedure Documenticollegati6Click(Sender: TObject);
    procedure Documenticollegati7Click(Sender: TObject);
    procedure Documenticollegati8Click(Sender: TObject);
    procedure ConnectionLost(Sender: TObject; Component: TComponent; ConnLostCause: TConnLostCause; var RetryMode: TRetryMode);
    procedure archiviocollegato031Click(Sender: TObject);
    procedure archiviocollegato041Click(Sender: TObject);
    procedure ipwMXResponse(Sender: TObject; RequestId: Integer; const Domain, MailServer: string; Precedence, TimeToLive, StatusCode: Integer; const Description: string; Authoritative: Boolean);
    procedure bilanciodicommessa1Click(Sender: TObject);
    procedure arcAfterConnect(Sender: TObject);
    procedure arcditAfterConnect(Sender: TObject);
    procedure arcsorAfterConnect(Sender: TObject);
  protected
    codice_passato: variant;

    cartella_documenti: string;
    esistenza: string;
    server_controllo_dominio: string;
    nome_tabella, codice_tabella, codice_01_tabella, codice_02_tabella, codice_03_tabella: string;

    imponibile_trasporto, importo_trasporto, importo_trasporto_assegnato: double;
    imponibile_bollo, importo_bollo: double;
    importo_incasso, totale_documento, totale_iva: double;
    totale, percentuale_sconto, importo_sconto_calcolato, importo_sconto, importo_sconto_assegnato: double;
    importo_assegnato: double;
    quantita_medio_bilancio, importo_medio_bilancio: double;
    importo_cassa_professionisti, importo_ritenuta: double;

    tabella_scadenze: array of scadenze;
    tabella_campi: array of campi;

    numero_rate: word;

    numero_riga: Integer;
    progressivo_opt_configurazioni: Integer;

    assegnato, flag_bollo: Boolean;
    giorno_fisso_usato: Boolean;

    data_inizio_inventario: tdatetime;

    documenti: Topendialog;

    procedure AddMainContentType(SL: TStrings; img_count: Integer; allegati: Boolean);
    procedure crea_query_ri(uno: Boolean; nome_database: TMyConnection_go; nome_tabella: string; codice_tabella: string; codice1_tabella: string; codice2_tabella: string; codice3_tabella: string; valore_codice: variant; valore_codice1: variant; valore_codice2: variant; valore_codice3: variant);
    procedure esegui_visdel(archivio: string; querydel: tmyquery_go);
    function GetWinVersion: string;
    function formattazione_html_tobit(testo: string): string;

    function IsFileInUse(FileName: TFileName): Boolean;
  public
    lista_programmi_recenti, lista_personalizzati: tstringlist;

    crittografia: tmyencryptor;

    src: string;

    procedure ActiveFormChange(Sender: TObject);

    function connessione_arc(utente_connessione: string = ''): Boolean;
    procedure assegna_variabili_utente(codice_utente: string);
    procedure controllo_msg(timer: Boolean = false);
    procedure xlstxt(tabella_esportare: TVirtualTable; tipo_esportazione, nome_file: string);
    procedure esporta_xls(query_xls: tdataset; nome_esportazione: string; griglia: trzdbgrid_go; cartella_standard: Boolean = false; esegui_excel: Boolean = true); overload;
    procedure esporta_csv(query_xls: tdataset; nome_esportazione: string; griglia: trzdbgrid_go); overload;
    procedure esporta_listbox(lista: tlistbox; nome_esportazione: string);
    procedure esiste_dati_aggiuntivi_archivio(var tipo_control: Boolean; tabella: tmyquery_go; nome_tabella: string; var dati_utente_creazione: string; var dati_data_ora_creazione: tdatetime; var dati_utente: string; var dati_data_ora: tdatetime);
    function controllo_integrita_referenziale(nome_database: string; nome_tabella: string; codice_tabella: string; codice1_tabella: string; codice2_tabella: string; codice3_tabella: string; valore_codice: variant; valore_codice1: variant; valore_codice2: variant; valore_codice3: variant): Boolean;

    (*
      function setta_valore_generatore(connessione: TMyConnection_go; id_generatore: string): integer; overload;
      function setta_valore_generatore(connessione: TMyConnection_go; id_generatore, codice_ditta: string): integer; overload;
 *)
    function setta_valore_generatore(connessione: TMyConnection_go; id_generatore: string; codice_ditta_passato: string = ''): Integer;

    procedure assegna_numerazione(codice_ditta, tipo, serie: string; data: tdatetime; var data_precedente: tdatetime; var progressivo, progressivo_precedente: double; aggiorna: Boolean; avviso: Boolean = true); overload;
    procedure assegna_numerazione(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data: tdatetime; var data_precedente: tdatetime; var progressivo, progressivo_precedente: double; aggiorna: Boolean; avviso: Boolean = true); overload;
    // nuova versione
    procedure assegna_numerazione(connessione: TMyConnection_go; tipo, serie: string; data: tdatetime; var progressivo: double; aggiorna: Boolean = true; avviso: Boolean = true); overload;
    procedure storna_numerazione(codice_ditta, tipo, serie: string; data, data_precedente: tdatetime; progressivo, progressivo_precedente: double); overload;
    procedure storna_numerazione(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data, data_precedente: tdatetime; progressivo, progressivo_precedente: double); overload;
    procedure storna_numerazione(connessione: TMyConnection_go; tipo, serie: string; data: tdatetime; progressivo: double); overload;
    procedure storna_numerazione_cancellata(codice_ditta, tipo, serie: string; data: tdatetime; numero: double); overload;
    procedure storna_numerazione_cancellata(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data: tdatetime; numero: double); overload;
    function esistenza_documento(documento, serie, cfg_codice: string; data: tdatetime; numero: double; progressivo: Integer; revisione: Integer = 0): Boolean; overload;
    function esistenza_documento(documento, serie, cfg_codice: string; data: tdatetime; numero: string; progressivo: Integer; revisione: Integer = 0): Boolean; overload;
    function esistenza_documento(connessione: TMyConnection_go; documento, serie, cfg_codice: string; data: tdatetime; numero: double; progressivo: Integer; revisione: Integer = 0): Boolean; overload;
    function esistenza_documento(connessione: TMyConnection_go; documento, serie, cfg_codice: string; data: tdatetime; numero: string; progressivo: Integer; revisione: Integer = 0): Boolean; overload;
    function controllo_ora(ora: string): string;
    procedure connessione_database_ditta(database: TMyConnection_go; nome_database, nome_utente, hostname: string);
    procedure assegna_variabili_sistema(connessione: tmyconnection);
    procedure aggiorna_data_fine(nome_tabella, operazione, campo_01, codice_01, campo_02, codice_02, campo_03, codice_03, campo_04, codice_04, campo_05, codice_05: string; data_inizio: tdatetime);
    procedure spezza_descrizione(descrizione: string; var descrizione1, descrizione2: string; caratteri: word);
    procedure crea_ltm_lettore(art_codice, lot_codice, tma_codice, quantita, data_scadenza, documento_origine, esistenza, cfg_tipo, cfg_codice, serie_documento: string; progressivo, riga: Integer; numero_documento: double; data_registrazione, data_documento: tdate); overload;
    procedure chiamata_telefono(numero: string);
    procedure chiamata_skype(skype_user: string);
    procedure messaggio_skype(skype_user: string);
    procedure messaggio_whatsapp(cellulare: string);
    function Unita(k: Integer): string;
    function Decine(k: Integer): string;
    function Migliaia(k: Integer): string;
    function CalcolaLettere(Importo: double): string;
    function settimana(data: tdate): word;
    function assegna_codice_lotto_automatico(data: tdate; frn_codice: string = ''; numero_documento: double = 0; data_documento: tdate = 0; art_codice: string = ''): string; overload;
    function data_database(data: tdate): string;
    procedure controllo_prezzo_costo(art_codice: string; Importo, quantita: double);
    function puntino(numero: double; numero_decimali: Integer = 2): string;
    procedure traduzione(data_set: tmyquery_go; parola, campo_01, campo_02, campo_03, campo_04, campo_05: string);
    procedure traduzione_cinese(data_set: tmyquery_go; parola, campo: string);
    function traduci(const parola, sourcelng, destlng: string): string;
    procedure WinInet_HttpGet(const Url: string; Stream: TStream); overload;
    function WinInet_HttpGet(const Url: string): string; overload;
    function invia_messaggio(pec: Boolean; oggetto, conoscenza, messaggio_testo: string; var lista: string; allegati: tstringlist; user_host, user_id, user_password, user_mail: string; porta_smtp: Integer; num_img_html: Integer; conoscenza_ccn: string = ''; no_tls: string = '';
      protocollo_tls: string = ''): Boolean;
    procedure escludi_tco_tna_iva_sospensione(tabella: tmyquery_go; nom_codice: string);
    procedure generazione_barcode(art: tmyquery_go; forzatura: Boolean = false);
    procedure aggiorna_cnt(anno, tipo, sottotipo: string; data: tdatetime; var progressivo: double; avviso: Boolean = true);
    procedure storna_cnt(tipo, serie: string; data: tdatetime; progressivo: double);
    function controllo_dominio(dominio: string): Boolean;
    function normalizza_codice(codice: string; normalizza: Boolean = false): string;
    function normalizza_documento(numero_documento: string): string;
    function calcola_ricarico(trl_codice: string; prezzo: double; decimali: Integer; superiore: Boolean = false): double;
    function cerca_campo_csv(numero_campo: word; sorgente: string; separatore: string = ';'): string;
    function assegna_fine_mese(anno, mese: Integer): tdatetime;
    function scorporo(Importo, percentuale: double; decimali: word = 2): double; overload;
    function scorporo(Importo: double; art_codice: string; scorpora: Boolean; vendite: Boolean = true): double; overload;
    procedure edit_note(var note_out: string; note: string; data_set: tmyquery_go; modifica: Boolean = true; font: word = 14);
    procedure calcola_peso_lordo(testata_documento: tmyquery_go; colli: Integer = 0);
    function tipo_variabile(campo: tfieldtype): string;
    procedure cerca_valore_passato(valore_passato_ricerca: string; data_set: tdataset; campo_tabella: string; successivo: Boolean = false);
    procedure ricerca_griglia(griglia: trzdbgrid_go);
    procedure filtro_griglia(griglia: trzdbgrid_go; var filtro_impostato: string);
    procedure totalizza_griglia(griglia: trzdbgrid_go);
    procedure esegui_programma_msg(modulo_documento, tipo_documento: string; progressivo_documento: Integer);
    function numero_documento_alfa(tabella: tmyquery_go; campo_numero_documento, numero_documento_alfa: string): double;
    function serie_documento_alfa(tabella: tmyquery_go; campo_serie_documento, numero_documento_alfa: string): string;
    function controllo_data_utilizzo: Boolean;
    function controllo_data_licenza: Boolean;
    function mag_esercizio(art_codice, tma_codice, ese_codice: string; dalla_data, alla_data: tdate): double;
    function messaggio_nuovo(codice_messaggio: Integer; descrizione_messaggio: string; lista_opzioni: string = ''; standard: Integer = 1): Integer;
    function presenti_provvisori(ese_codice: string; dalla_data: tdate = 0; alla_data: tdate = 0): Boolean;
    function multireplace(valore, old, new: string): string;

    procedure aggiorna_database;
    procedure assegna_skin;
    procedure assegna_peso_modulo(peso: Integer);
    procedure sconti_percentuale(componente: twincontrol);
    function crea_tsm(sconto_maggiorazione: string; percentuale_01, percentuale_02, percentuale_03, percentuale_04, percentuale_05, percentuale_06, percentuale_07, percentuale_08: double): string;
    procedure assegna_monitor;
  end;

procedure disabilita_campo(campo: TObject; tabstop: Boolean = true);
function abilita_campo(campo: TObject; controllo_abilitato: Boolean = false): Boolean;
procedure colore_control(contenitore: twincontrol; codice_abilitato: Boolean);
function sconto(tsm_codice: string): double;
procedure calcola_importo_documento(quantita, prezzo, cambio, importo_sconto: double; sconto_imponibile_lordo, listino_con_iva, tum_codice, tiv_codice, tsm_codice, tsm_codice_art: string; var Importo, importo_euro, importo_iva, importo_iva_euro, importo_non_arrotondato: double;
  solo_iva: Boolean = false);
function messaggio(codice_messaggio: Integer; descrizione_messaggio: string; touch: Boolean = false): Integer;
procedure esegui(nome_file: string);
procedure esegui_effettivo(nome_file: string; parametro: string = ''; cartella: string = '');
procedure esegui_stampa_diretta(nome_file: string);
procedure esegui_collegato(nome_file: string);
procedure assegna_file_cfg;
function esegui_programma(programma_da_eseguire: string; codice_archivio: variant; modale: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean; overload;
function esegui_programma(programma_da_eseguire: string; codice_archivio: variant; modale, record_singolo: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean; overload;
function numerico(stringa: string): Boolean;
function setta_lunghezza(stringa: string; caratteri: word): string; overload;
function setta_lunghezza(stringa: string; caratteri: word; destra: Boolean; carattere: string): string; overload;
function setta_lunghezza(numero: double; caratteri, decimali: word): string; overload;
function setta_lunghezza(numero: Integer; caratteri, decimali: word): string; overload;
function setta_lunghezza(numero: double; caratteri, decimali: word; carattere: string): string; overload;
function replicate(Ch: Char; Len: Integer): string;
function arrotonda(numero: double; nrdec: word; tipoarrotondamento: word): double; overload;
function arrotonda(numero: double; nrdec: word): double; overload;
function arrotonda(numero: double): double; overload;
procedure azzera_tabella(nome_tabella: string; var tabella: tmytable); overload;
procedure azzera_tabella(nome_tabella: string); overload;
function codice_tum(art_codice: string): string;
function decimali_quantita(tum_codice: string): word;
function StrTran(InString: string; SearchString: string; SubString: string; Incremental: Boolean): string;
function assegna_parametri_lavoro: string;
function eseguire_alias_personalizzato_esterno(aprogramma_standard: string; acodice_archivio: variant): string;
function call_programma(programma_da_eseguire: string; codice_archivio: variant; modale, record_singolo: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean;
function read_tabella(data_base: TMyConnection_go; nome_archivio, nome_codice: string; valore_codice: variant): Boolean; overload;
function read_tabella(data_base: TMyConnection_go; nome_archivio, nome_codice: string; valore_codice: variant; nome_campi: string): Boolean; overload;
function read_tabella(tabella_lettura: tmyquery_go; codice_archivio: variant): Boolean; overload;
function read_tabella(tabella_lettura: tmyquery_go): Boolean; overload;
function read_tabella(data_base: TMyConnection_go; nome_tabella: string): Boolean; overload;
function cambio(codice_tva: string; data_valuta: tdatetime): double;
function create_query(data_base: TMyConnection_go; testo_sql: string): tmyquery_go;
function decimali_quantita_art(art_codice: string; tipo_tum: string = ''): word;
function decimali_prezzo(tva_codice: string): word;
function decimali_prezzo_nom(nom_codice: string): word;
function decimali_prezzo_acq(tva_codice: string): word;
function decimali_prezzo_acq_nom(nom_codice: string): word;
function decimali_importo(tva_codice: string): word;
function decimali_importo_nom(nom_codice: string): word;
function cancella_escluso(const testo: string; caratteri: tsyscharset): string;
procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string); overload;
procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; var lista_multiselezione: tstringlist); overload;
procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; obbligatorio: Boolean); overload;
procedure esegui_visarc_effettivo(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; var lista_multiselezione: tstringlist;
  obbligatorio: Boolean = false);
function EncodePwd(const valpwd: string): string;
function DecodePwd(const valpwd: string): string;
function md5print(parola: string): string;
procedure assegna_parametri_passati;
function estrai_tag(campo, tag: string): string;
function etichetta_campo(griglia_devex: Pgriglia_devex; i: Integer; nome_campo_db, etichetta: string): Boolean;
procedure rinomina_campi(griglia_devex: Pgriglia_devex; titoli_minuscoli: string = 'si');
function tabella_edit(tabella: tdataset): Boolean;
function formato_display(decimali: word): string;
function formato_display_zero(decimali: word): string;
function fuoco(componente: twincontrol): Boolean;
procedure filtra_eccezioni(const exceptintf: imeexception; var handled: Boolean);

procedure apri_transazione(cursore_normale: string = 'no');
procedure commit_transazione(testo: string = 'transazione non eseguita');
procedure rollback_transazione(testo: string = '');
procedure chiudi_transazione;
function decodifica_html(atesto_html: string): string;
function exe_in_esecuzione(nome_file_exe: string): Boolean;

var
  arc: TARC;

  go_easy: Boolean;

  monitor_attivo: Integer;

  versione_procedura: string;
  versione_aggiornamento: Integer;

  archivio, archivio_arc: tmyquery_go;
  cifre_decimali_importo: word;

  codice_messaggio: shortint;
  descrizione_messaggio: string;

  // codice attivazione procedura
  ggg_codice_attivazione: string;
  nome_licenza, via_licenza, citta_licenza: string;
  numero_massimo_utenti: word;
  numero_nominativi: Integer;

  // variabile di salvataggio della forma del curosore
  cursore: tcursor;

  // tasto funzione utilizzato
  tasto_funzione: word;
  // tasto funzione extra (shift, alt, ctrl) utilizzato
  tasto_funzione1: tshiftstate;

  // ultimo utente login
  salta_ultimo_utente_login, ultimo_utente_login: string;

  // barra delle applicazioni
  taskbarwnd: Integer;
  barra_applicazioni_visibile: Boolean;

  // variabili di gestione
  data_attivazione: tdatetime;

  utente: string;
  utente_passato_programma_collegato: string;
  supervisore_utente: Boolean;
  blocco_utilizzo_ctrl_f12: string;
  importi_archivi: string;
  importi_vendite: string;
  importi_acquisti: string;
  importi_magazzino: string;
  login_accessi_utente: string;
  risoluzione_utente: Boolean;
  user_id_utente: string;
  user_password_utente: string;
  user_host_utente: string;
  user_e_mail_utente: string;
  disabilita_suoni_utente: string;
  richiesta_conferma_cancellazione: string;
  visualizza_hint: string;
  // modalita_inserimento_documenti: string;
  campi_grassetto: string;
  filtro_visarc_iniziale: string;
  font_standard: string;
  filtro_gesarc: string;

  ditta: string;
  descrizione_ditta: string;
  descrizione2_ditta: string;
  via_ditta: string;
  cap_ditta: string;
  citta_ditta: string;
  provincia_ditta: string;
  tna_codice_ditta, nazione_ditta, codice_iso_ditta: string;
  codice_fiscale_ditta: string;
  partita_iva_ditta: string;
  divisa_di_conto: string;
  /// /
  registri_prenumerati_ditta: string;
  registro_imprese_ditta: string;
  via_fiscale_ditta: string;
  cap_fiscale_ditta: string;
  citta_fiscale_ditta: string;
  provincia_fiscale_ditta: string;
  marchio_percorso_ditta: string;
  marchio_sinistra_ditta, marchio_superiore_ditta, marchio_altezza_ditta, marchio_larghezza_ditta: Integer;
  telefono_ditta: string;
  fax_ditta: string;
  web_ditta: string;
  e_mail_ditta: string;
  cartella_stampe_ditta: string;
  capitale_sociale_ditta: double;
  /// //
  codice_nom_numerico, codice_articolo_numerico, codice_matricola_numerico, codice_cespite_numerico: string;
  password_storno_evasione_vendite, password_storno_consolidamento_vendite, password_storno_differita_vendite, password_storno_evasione_acquisti, password_storno_consolidamento_acquisti, password_storno_differita_acquisti: string;
  gestione_revisioni: string;

  esercizio: string;
  descrizione_esercizio: string;
  data_inizio: tdatetime;
  data_fine: tdatetime;
  data_bilancio: tdatetime;

  esercizio_precedente: string;
  data_inizio_precedente: tdatetime;
  data_fine_precedente: tdatetime;
  data_bilancio_precedente: tdatetime;

  esercizio_successivo: string;
  esercizio_chiuso: string;
  esercizio_chiuso_magazzino: string;
  esercizio_chiuso_precedente: string;
  esercizio_chiuso_magazzino_precedente: string;
  storico: Boolean;
  commessa_attiva: string;
  lotto_articolo_letto: string;
  help_personalizzato: string;
  inventario_fiscale, inventario_gestionale: string;

  codice_procedura: string;
  nome_procedura: string;
  valore_passato_ricerca, carattere_ricerca_griglia: string;

  listini_01, listini_02, listini_03, listini_04, listini_05, listini_06, listini_07, listini_08, listini_09, listini_10, listini_11: string;
  promozioni_01, promozioni_02, promozioni_03, promozioni_04, promozioni_05, promozioni_06, promozioni_07, promozioni_08, promozioni_09, promozioni_10, promozioni_11: string;
  listini_01_fls, listini_02_fls, listini_03_fls, listini_04_fls, listini_05_fls, listini_06_fls, listini_07_fls, listini_08_fls, listini_09_fls, listini_10_fls: string;

  // programma vendita negozio VENNEG
  tastiera_touch: Boolean;

  // loop controllo campi
  controllo_campi_ok: Boolean;

  // help in linea url base
  url_base_help: string;

  // versione SO
  versione_os: string;

  // cartella in cui è installato GG
  programma_mkt: Boolean;
  cartella_installazione, cartella_root_installazione: string;

  // path file temporanei
  cartella_temp: string;

  // path file di testo da condividere o esportare
  cartella_base_file, cartella_file: string;

  // path bitmap standard
  cartella_bitmap: string;

  // path report standard
  cartella_report: string;

  // cartella mail
  cartella_email: string;

  // cartella esportazioni xls e csv
  cartella_esporta: string;

  // cartella stampe
  cartella_stampe: string;

  // cartella filtri_vis
  cartella_filtri_vis: string;

  // cartella stili
  cartella_stili: string;

  // database server e porta per connessione
  // tipo_database, versione_database,
  tipo_server, porta_server, versione_server, utente_database: string;
  tipo_portable: Boolean;

  // perametro per ambiente cloud
  codice_cliente_cloud: string;

  // programma alias utilizzato (da visualizzare come codice del programma)
  alias_programma: string;

  // utilizza descrizione programmi personalizzati
  descrizione_programmi_personalizzata: string;

  // controllo se è visibile la task bar
  video_libero: Boolean = true;

  // voci del menu pop up archivi
  pop_archivi_item: array [1 .. 20] of TMenuItem;
  codice_dati_aggiuntivi_popup: variant;
  tipo_codice_dati_aggiuntivi_popup: string;

  // voci visarc
  visarc_codice: variant;
  // visarc_lista_campi_griglia: tstringlist;
  visarc_stampa: Boolean;
  lista_ricerca_visarc: tstringlist;
  lista_ricerca_visarc_valore: tstringlist;
  lista_ricerca_visarc_filtro: tstringlist;

  decimali_prezzo_euro, decimali_prezzo_acq_euro: word;
  decimali_importo_euro: word;

  bitmap_login, bitmap_menu, bitmap_tabulati, sito_web, forum, telefono_assistenza, mail_assistenza: string;

  // controllo abilitazione
  modulo_abilitato: array [1 .. 30] of Boolean;
  ggg_abilitazione_moduli_ditta: string;

  // codice memorizzato in gesarc (solo se composto da un campo)
  codice_gesarc: variant;

  // menu utilizzato
  menu_utilizzato: string;

  // interruzione elaborazione
  interruzione_elaborazione: Boolean;

  // blocco articoli obsoleti
  blocco_obsoleti: string;

  // ultimo programma eseguito con Ctrl+F12
  ultimo_programma_f12: string;

  // parametri extra passati al programma chiamato
  parametri_extra_programma_chiamato: array [0 .. 31] of variant;

  // parametro gesnom
  gesnom_cli_frn: string;

  // hint tasti toolbar
  hint_help, hint_visualizza_archivio, hint_notizie_archivio, hint_visualizza_archivio_collegato, hint_gestione_archivio_collegato, hint_notizie_archivio_collegato, hint_duplica, hint_memorizza, hint_cancella, hint_stampa, hint_evadi_documento, hint_saldo, hint_aggiungi_preferiti, hint_torna_menu,
    hint_barra_applicazioni, hint_esegui_programmi, hint_inserisci_record, hint_analisi_valori, hint_articoli_equivalenti, hint_inserisci_riga: string;

  // data 31/12/9999
  data_31_12_9999: string;

  // mesi validità password
  mesi_password: word;

  // lingua nominativi
  lingua_nominativi: string;

  // linguaggio interfaccia
  linguaggio_interfaccia: string;

  // ricerche extra per articoli
  ricerca_articolo_codice_fornitore: string;
  ricerca_articolo_numero_serie: string;

  // numero decimali
  decimali_max_quantita: Integer;
  formato_display_quantita: string;
  formato_display_quantita_zero: string;
  decimali_max_prezzo: Integer;
  formato_display_prezzo: string;
  formato_display_prezzo_zero: string;
  decimali_max_prezzo_acq: Integer;
  formato_display_prezzo_acq: string;
  formato_display_prezzo_acq_zero: string;
  formato_display_importo: string;
  formato_display_importo_zero: string;

  // esecuzione automatica
  dit_codice_automatico, ese_codice_automatico, prg_codice_automatico, chiudi_procedura_automatico: string;

  // bottoni evidenziati
  evidenzia_campi: string;
  evidenzia_colore, colore_bordo: tcolor;

  campo_attivo_precedente: twincontrol;
  colore_attivo_precedente: tcolor;

  letto_barcode_articolo: Boolean;
  quantita_barcode_articolo: double;

  // ridimensiona form
  rapporto_espansione_font, rapporto_espansione_control: double;

  // dimensioni anteprima video
  tipo_anteprima: string;

  // disabilita screen saver
  abilita_screen_saver: string;

  // ricerca testo nelle anteprime
  cerca_testo_stampe: string;

  // flag descrizioni articolo unite
  descrizioni_articolo_unite: string;

  archivi_multiaziendali: string;

  programma_teleassistenza: string;

  // leggi file cfg
  // utente_manutenzione: string;
  cursore_database: string;

  // access violation  bugreport
  access_violation, bugreport_nominale: string;

  // collegamento foxgo
  foxgo: string;

  // disabilita visualizzazione postit quando ricontrolla i campi alla fine della form
  disabilita_postit: Boolean;

  file_archivio_cfg: textfile;
  record_letto_cfg: string;
  tasti_errati: Boolean;
  tasti_hint_cfg: string;

  // forzatura per evitare il controllo di assegnazione numerazione
  avviso_assegna_numerazione: string;

  // pagina da attivare in entrata archivio GESARC
  tab_pagina_attiva: Integer;

  // smtp controllo accessi
  // smtp_controllo_accessi: string;

  // filtri lim...
  testo_limart, testo_limcen, testo_limcli, testo_limcms, testo_limcsp, testo_limfrn, testo_limgen, testo_limind, testo_limnom, testo_limtag, testo_limtop, testo_limtcc, testo_limtco, testo_limtda, testo_limtdo, testo_limtgm, testo_limtma, testo_limtmo, testo_limtpa, testo_limtts, testo_limtvc,
    testo_limvuo, testo_limdip, testo_limtba, testo_limtcn, testo_limmtr: string;

  // codice barre a peso
  codice_barre_quantita: double;

  // parametri passati dall'esterno
  parametro_globale: string;
  parametro_utente, parametro_salta_login, parametro_password, parametro_programma, parametro_tvm_codice, parametro_progressivo_gesven, parametro_tag_codice, parametro_vending_assistenza, parametro_sessione, parametro_vending_ordini, parametro_vending_segnalazioni, parametro_ditta,
    parametro_esercizio, parametro_negozio, ditta_parametro, esercizio_parametro, parametro_tpresho_codice, parametro_multi, parametro_personalizzato, parametro_personalizzazioni, parametro_codice_gesarc, parametro_schedulato, parametro_stampa_diretta_pdf: string;

  // password digitata dall'utente al login
  password_utente_login: string;

  // blocco data per programma personalizzato
  programma_collegato_personalizzato: Boolean;

  // nome del file per gli scontrini fiscali
  nome_file_scontrino: string;
  numero_registratore: string;
  numero_scontrino_reso: string;
  data_scontrino_reso: tdate;

  // skin attivo
  skin_utilizzato: string;

  // anno fe
  anno_fe: string;

  // modulo ritenute clienti
  ritenute_clienti: Boolean;

  // controlla il primo accesso
  primoAccesso: Boolean;

const
  // password fissa
  password_database = '!*Go2002*!';

  // display percentuali
  formato_display_percentuale = ',0.00;-,0.00;#';
  formato_display_cambio = ',0.000000;-,0.000000;#';

  // nome e versione procedura
  file_registro = 'Registro.go';

  // messaggi
  codice_inesistente = 'codice inesistente';

  valore_non_consentito = 'il valore utilizzato non è consentito';
  valore_non_consentito1 = 'il valore utilizzato';
  valore_non_consentito2 = 'non è consentito';
  conferma_memorizzazione = 'memorizzare le variazioni effettuate?';
  conferma_memorizzazione_uscita = 'i dati dell''archivio sono stati variati' + #13 + 'memorizzare le variazioni effettuate?';
  conferma_cancellazione = 'conferma cancellazione del codice';
  inizio_archivio = 'raggiunto l''inizio dell''archivio';
  fine_archivio = 'raggiunta la fine dell''archivio';
  abbandoni_modifiche = 'abbandonare le modifiche in corso?';
  record_locked = 'record utilizzato da un altro utente' + #13 + 'riprovare ad eseguire la modifica?';
  record_esistente = 'il record è già presente nella tabella' + #13 + 'l''operazione verrà annullata';
  numero_tentativi_lock_01 = 'sono stati eseguiti ';
  numero_tentativi_lock_02 = 'tentativi di leggere il record che risulta utilizzato da un altro utente' + #13 + 'è ora possibile rimuovere il blocco con possibilità di danneggiare l''integrità degli archivi' + #13 + 'indicare se si vuole procedere alla rimozione del blocco';

  utente_std = 'OPEN';

  // status bar
  visualizzazione_statusbar = 'visualizzazione';
  inserimento_statusbar = 'inserimento';
  modifica_statusbar = 'modifica';

  // stili per export Excel
  stile_stringa = 0;
  stile_numero = 1;
  stile_data = 2;
  stile_time = 3;
  stile_stringagrassetto = 4;
  stile_memo = 5;

  // const
  si = 'si';
  no = 'no';

  decimali_massimo = 'il valore massimo non può assere superiore a quello impostato in anagrafica ditta';
  password_errata = 'la password digitata è errata';
  assegnati_valori = 'sono stati assegnati valori non consentiti a uno o più tasti funzione nel file di configurazione';
  non_utilizzo = 'questo errore causerà il non corretto utilizzo della porcedura e quindi va corretto';
  duecento_primanota = 'la versione DEMO non può gestire più di 200 documenti di primanota per la ditta';
  duecento_magazzino = 'la versione DEMO non può gestire più di 200 documenti di magazzino per la ditta';
  duecento_fatture_vendita = 'la versione DEMO non può gestire più di 200 fatture di vendita per la ditta';
  duecento_fatture_acquisto = 'la versione DEMO non può gestire più di 200 fatture di acquisto per la ditta';
  identificazione_utente = 'Identificazione utente';
  licenza_rilasciata = 'La licenza d''uso è stata rilasciata a:';
  abuso_perseguibile = 'Ogni abuso è perseguibile ai sensi di legge.';
  programma_non_funziona = 'il programma non può funzionare' + #13 + 'perchè manca una tabella di sistema' + #13 + #13 + 'consultare l''assistenza tecnica';
  utente_collegato = 'l''utente risulta già collegato e non va utilizzato da un''altra postazione' + #13 + 'se l''avviso è causato da un crash di sistema si può forzare l''utilizzo dell''utente' + #13 + #13 + 'si vuole proseguire utilizzando l''utente selezionato?';
  codice_attivazione_non_corretto = 'codice attivazione della licenza d''uso non corretto' + #13 + 'vuoi eseguire ora l''attivazione?' + #13 + #13 + 'in attesa del codice di attivazione' + #13 + 'si può utilizzare l''utente "DEMO"' + #13 + 'con password "DEMO"';
  eseguire_nuova_attivazione = 'vuoi eseguire ora la nuova attivazione?';
  mancano = 'mancano ';
  giorni_alla_scadenza = ' giorni alla scadenza della licenza d''uso di Gestionale Open' + #13 + 'eseguire il rinnovo per tempo' + #13 + #13 + 'la licenza d''uso di questo programma è stata rilasciata a:' + #13;
  descrizione_hint_help = 'help contestuale del programma ';
  descrizione_hint_visualizza_archivio = 'visualizza archivio in gestione ';
  descrizione_hint_notizie_archivio = 'informazioni sull''archivio in gestione ';
  descrizione_hint_visualizza_archivio_collegato = 'visualizza archivio collegato ';
  descrizione_hint_notizie_archivio_collegato = 'informazioni sull''archivio collegato ';
  descrizione_hint_gestione_archivio_collegato = 'gestione archivio collegato ';
  descrizione_hint_duplica = 'duplica ';
  descrizione_hint_memorizza = 'memorizza record su disco ';
  descrizione_hint_inserisci_record = 'inserisce nuovo record ';
  descrizione_hint_stampa = 'stampa archivio in gestione ';
  descrizione_hint_aggiungi_preferiti = 'aggiunge programma a preferiti ';
  descrizione_hint_torna_menu = 'attiva il navigatore dell''archivio ';
  descrizione_hint_barra_applicazioni = 'mostra/nasconde barra delle applicazioni ';
  descrizione_hint_esegui_programmi = 'esegui programmi ';
  descrizione_hint_cancella = 'cancella record da disco ';
  descrizione_hint_inserisci_riga = 'inserisce nuova riga prima di quella attiva ';
  descrizione_hint_analisi_valori = 'analisi varie su valori (quantità, prezzi e importi) ';
  descrizione_hint_articoli_equivalenti = 'visualizza articoli equivalenti ';

  // messaggi
  msg_0002 = 'l''utente ha tutti i livelli di privilegio necessari per creare nuove aziende,';
  msg_0003 = 'nuovi utenti e per eseguire tutte le attività necessarie alla gestione aziendale';
  msg_0004 = 'ti è stato inviato il seguente messaggio dall''utente: ';
  msg_0005 = 'segnalazione scadenza attiva';
  msg_0006 = ' per il campo';
  msg_0007 = 'l''utente non ha il livello necessario per modificare il campo';
  msg_0008 = 'la tabella di elaborazione: ';
  msg_0009 = ' non può essere creata';
  msg_0010 = 'il programma';
  msg_0011 = 'è già in esecuzione';
  msg_0012 = 'programma in lavorazione';
  msg_0015 = 'il codice che si sta cercando di aggiornare: [';
  msg_0016 = 'non esiste nella tabella ';
  msg_0017 = 'eseguire il programma di ricreazione progressivi contabili';
  msg_0019 = 'il progressivo inserito';
  msg_0020 = 'non è superiore all''ultimo memorizzato nell''archivio contatori';
  msg_0021 = 'si vuole assegnare il primo numero disponibile?';
  msg_0022 = 'non è successivo all''ultimo memorizzato nell''archivio contatori';
  msg_0023 = 'confermi il valore inserito?';
  msg_0024 = 'non è stato possibile ripristinare la numerazione fiscale incrementata';
  msg_0025 = 'perchè è già stato assegnato un nuovo progressivo con valore: ';
  msg_0026 = 'il documento è già stato inserito al progressivo: ';
  msg_0027 = 'il numero è già stato utilizzato nell''anno';
  msg_0028 = 'nel documento con progressivo: ';
  msg_0029 = 'i valori di prezzo e sconto proposti sono quelli';
  msg_0030 = 'del fornitore standard dell''articolo';
  msg_0031 = 'in anagrafica ditta mancano i dati per il giroconto';
  msg_0032 = 'delle autofatture intrastat';
  msg_0033 = 'il sottoconto o la causale per il giroconto';
  msg_0034 = 'del pagamento immediato non sono compilati correttamente';
  msg_0035 = 'il codice non può essere cancellato perchè è utilizzato';
  msg_0036 = 'nella tabella';
  msg_0037 = 'visualizzare il dettaglio dei record interessati?';
  msg_0038 = 'codice dell''errore:';
  msg_0039 = '- tabella:';
  msg_0040 = '- codice:';
  msg_0041 = 'la tabella è da ricostruire perchè rovinata';
  msg_0042 = 'uscire dalla procedura e contattare l''assistenza tecnica';
  msg_0043 = 'l''area dati della tabella è da ricostruire perchè rovinata';
  msg_0044 = 'il carattere di controllo del codice fiscale deve essere:';
  msg_0046 = 'non è stata indicata la divisa di conto in anagrafica ditta';
  msg_0047 = 'la licenza della procedura è scaduta';
  msg_0048 = 'è necessario richiedere la nuova attivazione';
  msg_0050 = 'il nuovo cliente verrà generato con il codice:';
  msg_0053 = 'record utilizzato da un altro utente';
  msg_0054 = 'l''archivio degli indice della tabella è da ricostruire perchè rovinato';
  msg_0055 = 'per ulteriori chiarimenti contattare l''assistenza tecnica';
  msg_0056 = 'il codice non può essere cancellato perchè ha collegamenti in altre tabelle';
  msg_0057 = 'la tabella non può essere aperta probabilmente perchè';
  msg_0058 = 'utilizzata in modo esclusivo da un altro programma';
  msg_0059 = 'il codice è già presente nella tabella di riferimento';
  msg_0060 = 'l''operazione verrà annullata';
  msg_0061 = 'il codice non esiste nella tabella di riferimento';
  msg_0062 = 'Query di aggiornamento fallita per record bloccati (Errore Nativo:';
  msg_0063 = 'errore di sistema del database';
  msg_0064 = 'essendo abilitato il collegamento zucchetti o sicom o team system non verrà eseguito';
  msg_0065 = 'il giroconto automatico della ritenuta d''acconto che andrà effettuato manualmente';
  msg_0066 = 'il movimento di iva con esigibilità differita è movimentato';
  msg_0067 = 'vanno prima cancellati i documenti collegati';
  msg_0068 = 'la partita aperta dalla riga di primanota è movimentata o presente su una distinta di pagamento o di sollecito';
  msg_0069 = 'per eseguire l''operazione desiderata';
  msg_0070 = 'utente';
  msg_0071 = 'esercizio storico';
  msg_0072 = 'non è settata la ditta di lavoro';
  msg_0073 = 'non è settato l''esercizio di lavoro';
  msg_0074 = '[storico]';
  msg_0075 = 'ci sono programmi in esecuzione verificarli dalla taskbar di Window';
  msg_0076 = 'provvedere alla loro chiusura prima di terminare l''esecuzione';
  msg_0078 = 'per poter variare i parametri operativi è necessario';
  msg_0079 = 'accesso non consentito';
  msg_0080 = 'perchè il programma possa funzionare deve essere definito';
  msg_0081 = 'un drive logico (es. G:) per la condivisione';
  msg_0082 = 'della cartella di installazione';
  msg_0083 = 'non è possibile accedere al numero seriale del disco su cui è installato il programma';
  msg_0084 = 'il modulo non è abilitato all''utilizzo';
  msg_0085 = 'la licenza del programma è scaduta';
  msg_0086 = 'mancano ';
  msg_0087 = ' giorni alla scadenza della licenza d''uso del programma';
  msg_0088 = 'richiedere l''attivazione di prova gratuita della durata di 3 mesi';
  msg_0089 = 'oppure quella definitiva a pagamento per 12, 24 o 36 mesi';
  msg_0090 = 'tramite l''apposita pagina del sito';
  msg_0091 = 'nella richiesta vanno indicati:';
  msg_0092 = 'il nome del modulo da abilitare [';
  msg_0093 = '  esercizio: ';
  msg_0094 = ']';
  msg_0095 = 'il numero di serie del programma [';
  msg_0096 = 'la durata dell''abilitazione [demo 3 mesi, 12 mesi, 24 mesi, 36 mesi]';
  msg_0097 = 'il programma è già presente nella lista dei preferiti';
  msg_0098 = 'la data ed il codice di attivazione che verranno assegnati sono da caricare';
  msg_0099 = 'utilizzando il programma [GESGGG] che si può attivare utilizzando i tasti [Ctrl+F12]';
  msg_0100 = 'oppure tramite la voce del menu [Sistema | Gestione archivi di sistema | Abilitazione moduli]';

  // hint
  hint_001 = 'visualizza tutti i tipi scadenza';
  hint_002 = 'visualizza tutte le scadenze indipendentemente dal tipo';
  hint_003 = 'visualizza solo le scadenze del tipo: ';
  hint_004 = 'programma ';
  hint_005 = 'saldo automatico ';

  boundary_text_and_html = '----=_TEXT_AND_HTML_BOUNDARY_';
  boundary_message_and_pictures = '----=_MESSAGE_AND_PICS_BOUNDARY_';
  boundary_main_and_attachments = '----=_MAIN_AND_ATTACHMENTS_BOUNDARY_';

implementation

{$R *.dfm}

uses ZZCALLPRG, GGFORMBASE, GGVISDEL, GGIMPALF, GGGESARC, GGVISINH, GGVIS20INH, ZZSETTA_GENERATORE,
  ZZARROTONDAMENTO, ZZCHECKDIGIT, GGMESSAGGIO, GGNOTE, GGGESDCT, ZZVERSIONE_PROCEDURA,
  xmlxform, xmldom, XMLDoc, XMLIntf, ZZVALIDA_EMAIL, GGSCONTIPERC, GGAVVISO, sYSTEM.Character;

// ******************************************************************************
// errori sul data base
// ******************************************************************************

procedure TARC.Documenti1Click(Sender: TObject);
var
  continua: Boolean;
  codice_articolo: string;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_articolo := normalizza_codice(collegamenti_archivio.fieldbyname('codice').asstring);

    if directoryexists(cartella_file + '\articoli\' + codice_articolo) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := cartella_file + '\articoli\' + codice_articolo;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Foto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('VISART', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Contropartitevendita1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select tca_codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select tca_codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('tca_codice').asstring, '']);
    esegui_programma('GESCPV', codice_passato, true);
  end;
end;

procedure TARC.cpa1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select taq_codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select taq_codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('taq_codice').asstring, '']);
    esegui_programma('GESCPA', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.N1Listinidivendita1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '', datetostr(0)]);
    esegui_programma('GESLSV', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Listinidiacquisto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '', datetostr(0)]);
    esegui_programma('GESLSA', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Codiciabarre1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESBAR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Depositi1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESMAG', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Lifo1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESLIF', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Schedamovimentazioni1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('SCHART', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Situazionedisponibilit1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('DISART', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.situazioneenasarco1Click(Sender: TObject);
var
  frn_tag: tmyquery_go;
begin
  frn_tag := tmyquery_go.create(nil);
  frn_tag.connection := arcdit;
  frn_tag.sql.text := 'select codice from tag where frn_codice = :frn_codice';

  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    frn_tag.Close;
    frn_tag.parambyname('frn_codice').asstring := collegamenti_archivio.fieldbyname('codice').asstring;
    frn_tag.open;
    if not frn_tag.isempty then
    begin
      codice_passato := vararrayof([frn_tag.fieldbyname('codice').asstring, 0, '', 0]);
      esegui_programma('GESENACO', codice_passato, true);
    end
    else
    begin
      messaggio(200, 'il fornitore non è un agente');
    end;
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
  frn_tag.free;
end;

procedure TARC.MenuItem31Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESINDINH', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem_progressivi_contabiliClick(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('VISCFG', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem_scheda_contabileClick(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('SCHCON', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Fido1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('SITFID', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

function TARC.formattazione_html_tobit(testo: string): string;
var
  temp_str: string;
begin
  temp_str := testo;
  temp_str := stringreplace(temp_str, '&', '&amp;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '"', '&quot;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '<', '&lt;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '>', '&gt;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, ' ', '&nbsp;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¡', '&iexcl;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¢', '&cent;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '£', '&pound;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¤', '&curren;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¥', '&yen;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¦', '&brvbar;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '§', '&sect;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¨', '&uml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '©', '&copy;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ª', '&ordf;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '«', '&laquo;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¬', '&not;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '­', '&shy;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '®', '&reg;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¯', '&macr;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '°', '&deg;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '±', '&plusmn;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '²', '&sup2;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '³', '&sup3;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '´', '&acute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'µ', '&micro;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¶', '&para;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '·', '&middot;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¸', '&cedil;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¹', '&sup1;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'º', '&ordm;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '»', '&raquo;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¼', '&frac14;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '½', '&frac12;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¾', '&frac34;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '¿', '&iquest;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'À', '&Agrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Á', '&Aacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Â', '&Acirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ã', '&Atilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ä', '&Auml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Å', '&Aring;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Æ', '&AElig;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ç', '&Ccedil;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'È', '&Egrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'É', '&Eacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ê', '&Ecirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ë', '&Euml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ì', '&Igrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Í', '&Iacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Î', '&Icirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ï', '&Iuml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ð', '&ETH;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ñ', '&Ntilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ò', '&Ograve;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ó', '&Oacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ô', '&Ocirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Õ', '&Otilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ö', '&Ouml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '×', '&times;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ø', '&Oslash;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ù', '&Ugrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ú', '&Uacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Û', '&Ucirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ü', '&Uuml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Ý', '&Yacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'Þ', '&THORN;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ß', '&szlig;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'à', '&agrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'á', '&aacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'â', '&acirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ã', '&atilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ä', '&auml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'å', '&aring;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'æ', '&aelig;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ç', '&ccedil;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'è', '&egrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'é', '&eacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ê', '&ecirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ë', '&euml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ì', '&igrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'í', '&iacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'î', '&icirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ï', '&iuml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ð', '&eth;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ñ', '&ntilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ò', '&ograve;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ó', '&oacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ô', '&ocirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'õ', '&otilde;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ö', '&ouml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, '÷', '&divide;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ø', '&oslash;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ù', '&ugrave;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ú', '&uacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'û', '&ucirc;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ü', '&uuml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ý', '&yacute;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'þ', '&thorn;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, 'ÿ', '&yuml;', [rfreplaceall]);
  temp_str := stringreplace(temp_str, #13, '<br>', [rfreplaceall]);
  result := '<div><font color=#000000 face=arial>' + temp_str + '</div>'
end;

procedure TARC.Documenti2Click(Sender: TObject);
var
  continua: Boolean;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if directoryexists(cartella_file + '\clienti\' + collegamenti_archivio.fieldbyname('codice').asstring) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := cartella_file + '\clienti\' + collegamenti_archivio.fieldbyname('codice').asstring;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem46Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('VISCFG', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem47Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('SCHCON', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documenti3Click(Sender: TObject);
var
  continua: Boolean;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if directoryexists(cartella_file + '\fornitori\' + collegamenti_archivio.fieldbyname('codice').asstring) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := cartella_file + '\fornitori\' + collegamenti_archivio.fieldbyname('codice').asstring;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem32Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from gen where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from gen where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['G', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('VISCFG', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il sottoconto non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem34Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from gen where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from gen where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['G', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('SCHCON', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il sottoconto non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documenti4Click(Sender: TObject);
var
  continua: Boolean;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from gen where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from gen where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if directoryexists(cartella_file + '\sottoconti\' + collegamenti_archivio.fieldbyname('codice').asstring) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := cartella_file + '\sottoconti\' + collegamenti_archivio.fieldbyname('codice').asstring;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'il sottoconto non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documenti5Click(Sender: TObject);
var
  continua: Boolean;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cms where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cms where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if directoryexists(cartella_file + '\commesse\' + collegamenti_archivio.fieldbyname('codice').asstring) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := cartella_file + '\commesse\' + collegamenti_archivio.fieldbyname('codice').asstring;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'la commessa non è stata ancora memorizzata');
  end;
end;

procedure TARC.documenti6Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', documenti6.caption);
  fine := pos('id:', documenti6.caption);
  tabella := trim(copy(documenti6.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(documenti6.caption, fine + 3, length(documenti6.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.documenticollegati1Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', documenticollegati1.caption);
  fine := pos('id:', documenticollegati1.caption);
  tabella := trim(copy(documenticollegati1.caption, inizio + 8, fine - (inizio + 8)));
  if (inizio <> 0) and (fine <> 0) then
  begin
    id := strtoint(trim(copy(documenticollegati1.caption, fine + 3, length(documenticollegati1.caption))));

    pr := tgesdct.create(nil);
    pr.codice := vararrayof([tabella, id, 0]);
    pr.showmodal;
    freeandnil(pr);
  end;
end;

procedure TARC.MenuItem44Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from tva where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from tva where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if collegamenti_archivio.fieldbyname('codice').asstring = divisa_di_conto then
    begin
      messaggio(000, 'l''archivio storico fixing  può essere gestito' + #13 + 'solo se la valuta non è quella di conto');
    end
    else
    begin
      codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, datetostr(now - 2)]);
      esegui_programma('GESTVF', codice_passato, true);
    end;
  end
  else
  begin
    messaggio(000, 'la valuta non è stata ancora memorizzata');
  end;
end;

procedure TARC.MenuItem45Click(Sender: TObject);
begin
  codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, datetostr(now - 2)]);
  esegui_programma('ELAFIX', codice_passato, true);
end;

procedure TARC.MenuItem54Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from tts where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from tts where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '', datetostr(0)]);
    esegui_programma('GESTSI', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il tipo scadenza non è stata ancora memorizzata');
  end;
end;

procedure TARC.MenuItem62Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from tst where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from tst where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '', datetostr(0)]);
    esegui_programma('GESTTI', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il recupero spese trasporto non è stato ancora memorizzato');
  end;
end;

procedure TARC.menu_item_utn_disabilita_ditteClick(Sender: TObject);
begin
  collegamenti_archivio_arc.Close;
  collegamenti_archivio_arc.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from utn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from utn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio_arc.open;
  if not collegamenti_archivio_arc.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESABD', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''utente non è stato ancora memorizzato');
  end;
end;

procedure TARC.menu_item_utn_disabilita_programmiClick(Sender: TObject);
begin
  collegamenti_archivio_arc.Close;
  collegamenti_archivio_arc.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from utn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from utn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio_arc.open;
  if not collegamenti_archivio_arc.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESABP', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''utente non è stato ancora memorizzato');
  end;
end;

procedure TARC.manu_item_utn_Assegna_stampantiClick(Sender: TObject);
begin
  collegamenti_archivio_arc.Close;
  collegamenti_archivio_arc.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from utn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from utn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio_arc.open;
  if not collegamenti_archivio_arc.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESUTP', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''utente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Lotti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, lotti from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, lotti from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if collegamenti_archivio.fieldbyname('lotti').asstring = 'si' then
    begin
      codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
      esegui_programma('GESLOT', codice_passato, true);
    end
    else
    begin
      messaggio(000, 'l''articolo non prevede la gestitone dei lotti');
    end;
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem_cms_tipologieClick(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cms where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cms where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESCMT', codice_passato, true);
  end;
end;

procedure TARC.Situazionepartitario2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('STAPAR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Situazionepartitario1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('STAPAR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.Situaizonepartitario1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from gen where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from gen where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['G', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('STAPAR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il sottoconto non è stato ancora memorizzato');
  end;
end;

procedure TARC.Estrattoconto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('STAEST', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.MenuItem83Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cen where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cen where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('SCHCEM', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il centro di costo e ricavo non è stato ancora memorizzato');
  end;

end;

procedure TARC.Insoluti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('STAINS', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Ritardipagamento1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RITPAG', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Ordini1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RIEORDV', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documentidivendita1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['C', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RIEDOCV', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documentiextra1Click(Sender: TObject);
var
  continua: Boolean;
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select cartella_documenti from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select cartella_documenti from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    if directoryexists(collegamenti_archivio.fieldbyname('cartella_documenti').asstring) then
    begin
      documenti := Topendialog.create(nil);

      cartella_documenti := collegamenti_archivio.fieldbyname('cartella_documenti').asstring;
      documenti.initialdir := cartella_documenti;
      continua := true;
      while continua do
      begin
        if documenti.execute then
        begin
          esegui(documenti.FileName);
        end;
        if messaggio(300, 'prosegui con l''analisi dei documenti') <> 1 then
        begin
          continua := false;
        end;
      end;

      documenti.free;
    end
    else
    begin
      messaggio(200, 'non esiste la cartella documenti extra per il codice selezionato');
    end;
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Preventivi2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RIEPREA', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.Ordini2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RIEORDA', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.Documenticollegati2Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati2.caption);
  fine := pos('id:', Documenticollegati2.caption);
  tabella := trim(copy(Documenticollegati2.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati2.caption, fine + 3, length(Documenticollegati2.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati3Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati3.caption);
  fine := pos('id:', Documenticollegati3.caption);
  tabella := trim(copy(Documenticollegati3.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati3.caption, fine + 3, length(Documenticollegati3.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati4Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati4.caption);
  fine := pos('id:', Documenticollegati4.caption);
  tabella := trim(copy(Documenticollegati4.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati4.caption, fine + 3, length(Documenticollegati4.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati5Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati5.caption);
  fine := pos('id:', Documenticollegati5.caption);
  tabella := trim(copy(Documenticollegati5.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati5.caption, fine + 3, length(Documenticollegati5.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati6Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati6.caption);
  fine := pos('id:', Documenticollegati6.caption);
  tabella := trim(copy(Documenticollegati6.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati6.caption, fine + 3, length(Documenticollegati6.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati7Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati7.caption);
  fine := pos('id:', Documenticollegati7.caption);
  tabella := trim(copy(Documenticollegati7.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati7.caption, fine + 3, length(Documenticollegati7.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documenticollegati8Click(Sender: TObject);
var
  pr: tgesdct;

  tabella: string;
  id: Integer;
  inizio, fine: word;
begin
  inizio := pos('tabella:', Documenticollegati8.caption);
  fine := pos('id:', Documenticollegati8.caption);
  tabella := trim(copy(Documenticollegati8.caption, inizio + 8, fine - (inizio + 8)));
  id := strtoint(trim(copy(Documenticollegati8.caption, fine + 3, length(Documenticollegati8.caption))));

  pr := tgesdct.create(nil);
  pr.codice := vararrayof([tabella, id, 0]);
  pr.showmodal;
  freeandnil(pr);
end;

procedure TARC.Documentidiacquisto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['F', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('RIEDOCA', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.ConnectionLost(Sender: TObject; Component: TComponent; ConnLostCause: TConnLostCause; var RetryMode: TRetryMode);
begin
  RetryMode := rmReconnectExecute;
end;

procedure TARC.archiviocollegato011Click(Sender: TObject);
var
  cartella: string;
begin
  cartella := dit.fieldbyname('cartella_documenti_collegati').asstring;
  if cartella <> '' then
  begin
    cartella := cartella + '\';
  end;

  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_01 from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_01 from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    esegui(cartella + collegamenti_archivio.fieldbyname('archivio_collegato_01').asstring);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.pop_arc_artPopup(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, descrizione_archivio_collega_01, descrizione_archivio_collega_02,');
    collegamenti_archivio.sql.add('descrizione_archivio_collegato_03, descrizione_archivio_collegato_04 from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, descrizione_archivio_collega_01, descrizione_archivio_collega_02,');
    collegamenti_archivio.sql.add('descrizione_archivio_collegato_03, descrizione_archivio_collegato_04 from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if collegamenti_archivio.eof then
  begin
    archiviocollegato011.visible := false;
    archiviocollegato021.visible := false;
  end
  else
  begin
    if collegamenti_archivio.fieldbyname('descrizione_archivio_collega_01').asstring = '' then
    begin
      archiviocollegato011.visible := false;
    end
    else
    begin
      archiviocollegato011.visible := true;
      archiviocollegato011.caption := collegamenti_archivio.fieldbyname('descrizione_archivio_collega_01').asstring;
    end;
    if collegamenti_archivio.fieldbyname('descrizione_archivio_collega_02').asstring = '' then
    begin
      archiviocollegato021.visible := false;
    end
    else
    begin
      archiviocollegato021.visible := true;
      archiviocollegato021.caption := collegamenti_archivio.fieldbyname('descrizione_archivio_collega_02').asstring;
    end;
    if collegamenti_archivio.fieldbyname('descrizione_archivio_collegato_03').asstring = '' then
    begin
      archiviocollegato031.visible := false;
    end
    else
    begin
      archiviocollegato031.visible := true;
      archiviocollegato031.caption := collegamenti_archivio.fieldbyname('descrizione_archivio_collegato_03').asstring;
    end;
    if collegamenti_archivio.fieldbyname('descrizione_archivio_collegato_04').asstring = '' then
    begin
      archiviocollegato041.visible := false;
    end
    else
    begin
      archiviocollegato041.visible := true;
      archiviocollegato041.caption := collegamenti_archivio.fieldbyname('descrizione_archivio_collegato_04').asstring;
    end;
  end;
end;

procedure TARC.archiviocollegato021Click(Sender: TObject);
var
  cartella: string;
begin
  cartella := dit.fieldbyname('cartella_documenti_collegati').asstring;
  if cartella <> '' then
  begin
    cartella := cartella + '\';
  end;
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_02 from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_02 from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    esegui(cartella + collegamenti_archivio.fieldbyname('archivio_collegato_02').asstring);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.archiviocollegato031Click(Sender: TObject);
var
  cartella: string;
begin
  cartella := dit.fieldbyname('cartella_documenti_collegati').asstring;
  if cartella <> '' then
  begin
    cartella := cartella + '\';
  end;
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_03 from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_03 from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    esegui(cartella + collegamenti_archivio.fieldbyname('archivio_collegato_03').asstring);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.archiviocollegato041Click(Sender: TObject);
var
  cartella: string;
begin
  cartella := dit.fieldbyname('cartella_documenti_collegati').asstring;
  if cartella <> '' then
  begin
    cartella := cartella + '\';
  end;

  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_04 from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, archivio_collegato_04 from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    esegui(cartella + collegamenti_archivio.fieldbyname('archivio_collegato_04').asstring);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.racciabilitdocumenti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof(['cliente', collegamenti_archivio.fieldbyname('codice').asstring]);
    esegui_programma('TRADOCV', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Accessorio1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESACC', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Accessori1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESACCPR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Equivalente1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESEQU', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Equivalenti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, '']);
    esegui_programma('GESEQUPR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Contratti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CNTCLI', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Contratti2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CNTFRN', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.Contrattiarticoli1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CNTART', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Ultimiprezziacquisto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    parametri_extra_programma_chiamato[0] := 'articoli';
    parametri_extra_programma_chiamato[1] := '';
    parametri_extra_programma_chiamato[2] := '';
    parametri_extra_programma_chiamato[3] := collegamenti_archivio.fieldbyname('codice').asstring;
    parametri_extra_programma_chiamato[4] := 0;
    esegui_programma('ULTPRZ', '', true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Cruscotto1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CRUCLI', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Cruscotto2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from art where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from art where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CRUART', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'l''articolo non è stato ancora memorizzato');
  end;
end;

procedure TARC.Analisiordini1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('ANAORDCL', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Analisiordini2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('ANAORDFR', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

// messaggi inviati da altri utenti*********************************************

procedure TARC.controllo_msg(timer: Boolean = false);
var
  stringa: string;
  esegui: Boolean;

  dit_codice, tipo_documento: string;
begin
  msg.Close;
  msg.params[0].asstring := utente;
  msg.open;
  while not msg.eof do
  begin
    stringa := msg_0004 + msg.fieldbyname('mittente').asstring + slinebreak + 'messaggio originale' + slinebreak + msg.fieldbyname('m_descrizione').asstring;
    if msg.fieldbyname('riga').asinteger <> 1 then
    begin
      stringa := stringa + slinebreak + slinebreak + 'messaggio' + slinebreak + msg.fieldbyname('mr_descrizione').asstring;
    end;

    if msg.fieldbyname('modulo_documento').asstring = '' then
    begin
      if tabella_edit(msg) then
      begin
        msg.fieldbyname('letto').asstring := 'si';
        msg.post;
      end;

      messaggio(100, stringa);

      if timer then
      begin
        break;
      end;
    end
    else
    begin
      esegui := false;

      stringa := stringa + slinebreak + slinebreak + 'il messaggio è stato eseguito dalla gestione documenti' + slinebreak + msg.fieldbyname('modulo_documento').asstring + '[' + msg.fieldbyname('tipo_documento').asstring + ']' + slinebreak + slinebreak +
        'conferma per eseguire la gestione del documento';

      if tabella_edit(msg) then
      begin
        msg.fieldbyname('letto').asstring := 'xx';
        msg.post;
      end;

      if not timer then
      begin
        if messaggio(300, stringa) = 1 then
        begin
          esegui := true;
        end;
      end
      else
      begin
        messaggio(100, stringa);
        esegui := false;
        break;
      end;

      if esegui and (msg.active) then
      begin
        tipo_documento := msg.fieldbyname('tipo_documento').asstring;
        dit_codice := '';
        if (pos('[', msg.fieldbyname('tipo_documento').asstring) > 0) and (pos(']', msg.fieldbyname('tipo_documento').asstring) > 0) then
        begin
          dit_codice := copy(msg.fieldbyname('tipo_documento').asstring, pos('[', msg.fieldbyname('tipo_documento').asstring) + 1, length(msg.fieldbyname('tipo_documento').asstring));
          dit_codice := copy(dit_codice, 1, length(dit_codice) - 1);
          tipo_documento := trim(copy(tipo_documento, 1, pos('[', tipo_documento) - 1));
        end;

        if (dit_codice = '') or (dit_codice = ditta) then
        begin
          if tabella_edit(msg) then
          begin
            msg.fieldbyname('letto').asstring := 'si';
            msg.post;
          end;

          esegui_programma_msg(msg.fieldbyname('modulo_documento').asstring, msg.fieldbyname('tipo_documento').asstring, msg.fieldbyname('progressivo_documento').asinteger);
        end
        else
        begin
          messaggio(200, 'la ditta di riferimento [' + dit_codice + '] del documento è diversa da quella attiva [' + ditta + ']');
        end;
      end;
    end;

    if not msg.active then
    begin
      msg.open;
    end
    else
    begin
      msg.next;
    end;
  end;

  if msg.active then
  begin
    msg.Close;
  end;
end;

procedure TARC.esegui_programma_msg(modulo_documento, tipo_documento: string; progressivo_documento: Integer);
begin
  if modulo_documento = 'produzione' then
  begin
    esegui_programma('GESORDP', progressivo_documento, true);
  end
  else if modulo_documento = 'acquisto' then
  begin
    if tipo_documento = 'ddt' then
    begin
      esegui_programma('GESDDTA', progressivo_documento, true);
    end
    else if tipo_documento = 'ddt clienti' then
    begin
      esegui_programma('GESDDTC', progressivo_documento, true);
    end
    else if tipo_documento = 'fattura' then
    begin
      esegui_programma('GESFATA', progressivo_documento, true);
    end
    else if tipo_documento = 'fattura differita' then
    begin
      esegui_programma('GESFADA', progressivo_documento, true);
    end
    else if tipo_documento = 'nota credito' then
    begin
      esegui_programma('GESNOCA', progressivo_documento, true);
    end
    else if tipo_documento = 'ordine' then
    begin
      esegui_programma('GESORDA', progressivo_documento, true);
    end
    else if tipo_documento = 'preventivo' then
    begin
      esegui_programma('GESPREA', progressivo_documento, true);
    end;
  end
  else if modulo_documento = 'vendita' then
  begin
    if tipo_documento = 'bolla' then
    begin
      esegui_programma('GESBOLV', progressivo_documento, true);
    end
    else if tipo_documento = 'corrispettivo' then
    begin
      esegui_programma('GESCORV', progressivo_documento, true);
    end
    else if tipo_documento = 'ddt' then
    begin
      esegui_programma('GESDDTV', progressivo_documento, true);
    end
    else if tipo_documento = 'ddt fornitori' then
    begin
      esegui_programma('GESDDTF', progressivo_documento, true);
    end
    else if tipo_documento = 'fattura accompagnatoria' then
    begin
      esegui_programma('GESFAAV', progressivo_documento, true);
    end
    else if tipo_documento = 'fattura differita' then
    begin
      esegui_programma('GESFADV', progressivo_documento, true);
    end
    else if tipo_documento = 'fattura immediata' then
    begin
      esegui_programma('GESFAIV', progressivo_documento, true);
    end
    else if tipo_documento = 'nota credito' then
    begin
      esegui_programma('GESNOCV', progressivo_documento, true);
    end
    else if tipo_documento = 'ordine' then
    begin
      esegui_programma('GESORDV', progressivo_documento, true);
    end
    else if tipo_documento = 'preventivo' then
    begin
      esegui_programma('GESPREV', progressivo_documento, true);
    end
    else if tipo_documento = 'preventivo nominativi' then
    begin
      esegui_programma('GESPREVNOM', progressivo_documento, true);
    end
    else if tipo_documento = 'documento web' then
    begin
      esegui_programma('GESDOCW', progressivo_documento, true);
    end;
  end;
end;

// fine messaggi inviati da altri utenti***************************************

// controllo scadenze utente **************************************************

// fine controllo scadenze utente *********************************************

procedure TARC.MenuItem93Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from csp where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from csp where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('SCHCSP', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cespite non è stato ancora memorizzato');
  end;
end;

procedure TARC.Cruscotto3Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    codice_passato := collegamenti_archivio.fieldbyname('codice').asstring;
    esegui_programma('CRUFRN', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.assegna_variabili_utente(codice_utente: string);
begin
  utn.Close;
  utn.parambyname('codice').asstring := codice_utente;
  utn.open;

  utente := codice_utente;

  tips.Close;
  tips.parambyname('utn_codice').asstring := codice_utente;
  tips.open;

  if utn['supervisore'] = 'si' then
  begin
    supervisore_utente := true;
  end
  else
  begin
    supervisore_utente := false;
  end;
  blocco_utilizzo_ctrl_f12 := utn['blocco_utilizzo_ctrl_f12'];
  importi_archivi := utn['importi_archivi'];
  importi_vendite := utn['importi_vendite'];
  importi_acquisti := utn['importi_acquisti'];
  importi_magazzino := utn['importi_magazzino'];
  login_accessi_utente := utn['login_accessi'];
  disabilita_suoni_utente := utn['disabilita_suoni'];
  visualizza_hint := utn['visualizza_hint'];
  // modalita_inserimento_documenti := utn['modalita_inserimento_documenti'];
  evidenzia_campi := utn['evidenzia_campi'];
  tipo_anteprima := utn['tipo_anteprima'];
  cerca_testo_stampe := utn['cerca_testo_stampe'];
  campi_grassetto := utn['campi_grassetto'];
  filtro_visarc_iniziale := utn['filtro_visarc_iniziale'];
  abilita_screen_saver := utn['abilita_screen_saver'];
  font_standard := utn['font_standard'];

  if evidenzia_campi = 'si' then
  begin
    if utn['evidenzia_colore'] = 'azzurro' then
    begin
      evidenzia_colore := claqua;
      colore_bordo := claqua;
    end
    else if utn['evidenzia_colore'] = 'giallo' then
    begin
      evidenzia_colore := clyellow;
      colore_bordo := clyellow;
    end
    else if utn['evidenzia_colore'] = 'rosso' then
    begin
      evidenzia_colore := clred;
      colore_bordo := clred;
    end
    else if utn['evidenzia_colore'] = 'verde' then
    begin
      evidenzia_colore := cllime;
      colore_bordo := cllime;
    end
    else if utn['evidenzia_colore'] = 'fucsia' then
    begin
      evidenzia_colore := clfuchsia;
      colore_bordo := clfuchsia;
    end
    else if utn['evidenzia_colore'] = 'nessuno' then
    begin
      evidenzia_colore := clwindow;
      colore_bordo := clbtnshadow;
    end
    else if utn['evidenzia_colore'] = 'blu' then
    begin
      evidenzia_colore := clblue;
      colore_bordo := clblue;
    end
    else if utn['evidenzia_colore'] = 'nero' then
    begin
      evidenzia_colore := clblack;
      colore_bordo := clblack;
    end
    else if utn['evidenzia_colore'] = 'rosso' then
    begin
      evidenzia_colore := clred;
      colore_bordo := clred;
    end;
  end;

  user_id_utente := utn.fieldbyname('user_id').asstring;
  user_password_utente := utn.fieldbyname('user_password').asstring;
  user_host_utente := utn.fieldbyname('user_host').asstring;
  user_e_mail_utente := utn.fieldbyname('user_e_mail').asstring;

  dit_codice_automatico := utn.fieldbyname('dit_codice_automatico').asstring;
  ese_codice_automatico := utn.fieldbyname('ese_codice_automatico').asstring;
  prg_codice_automatico := utn.fieldbyname('prg_codice_automatico').asstring;
  chiudi_procedura_automatico := utn.fieldbyname('chiudi_procedura_automatico').asstring;
  if parametro_programma <> '' then
  begin
    prg_codice_automatico := parametro_programma;
    chiudi_procedura_automatico := 'si';
  end;

  ditta := utn.fieldbyname('dit_codice').asstring;
  if ditta_parametro <> '' then
  begin
    ditta := ditta_parametro;
  end;

  if dit_codice_automatico <> '' then
  begin
    ditta := dit_codice_automatico;
  end;

  esercizio := utn.fieldbyname('ese_codice').asstring;
  if esercizio_parametro <> '' then
  begin
    esercizio := esercizio_parametro;
  end;
  if ese_codice_automatico <> '' then
  begin
    esercizio := ese_codice_automatico;
  end;

  storico := utn.fieldbyname('storico').asboolean;
end;

function TARC.connessione_arc(utente_connessione: string = ''): Boolean;
begin
  result := true;

  arc.username := utente_database;
  arc.password := password_database;

  arcsor.username := utente_database;
  arcsor.password := password_database;

  arc.server := tipo_server;
  arcsor.server := tipo_server;
  arcdit.server := tipo_server;

  if codice_cliente_cloud = '' then
  begin
    arc.database := 'arc';
  end
  else
  begin
    arc.database := 'arc_' + codice_cliente_cloud;
  end;

  if porta_server <> '' then
  begin
    try
      arc.port := strtoint(porta_server);
      arcsor.port := strtoint(porta_server);
      arcdit.port := strtoint(porta_server);
    except
    end;
  end;

  // arc
  if arc.connected then
  begin
    arc.connected := false;
  end;

  try
    arc.connected := true;
  except
    on e: exception do
    begin
      messaggio(000, 'non è stata eseguita la connessione al database [arc]' + slinebreak + 'verificare utente e password oppure' + slinebreak + 'annotare il messaggio seguente e comunicarlo all''assistenza tecnica' + slinebreak + e.message);
      result := false;
      exit;
    end;
  end;

  // arcsor
  arcsor.connected := false;

  if codice_cliente_cloud = '' then
  begin
    arcsor.database := 'arc_ordinamento';
  end
  else
  begin
    arcsor.database := 'arc_ordinamento_' + codice_cliente_cloud;
  end;

  try
    arcsor.connected := true;
  except
    on e: exception do
    begin
      messaggio(000, 'non è stata eseguita la connessione al database [arc_ordinamento]' + slinebreak + 'verificare utente e password oppure' + slinebreak + 'annotare il messaggio seguente e comunicarlo all''assistenza tecnica' + slinebreak + e.message);
      result := false;
      exit;
    end;
  end;

  lin.Close;
  lin.open;
end;

procedure TARC.esporta_xls(query_xls: tdataset; nome_esportazione: string; griglia: trzdbgrid_go; cartella_standard: Boolean = false; esegui_excel: Boolean = true);
var
  i, j: word;
  nome_file: string;
  prosegui: Boolean;
begin
  if cartella_standard and (nome_esportazione <> '') then
  begin
    prosegui := true;
    nome_file := cartella_esporta + '\' + nome_esportazione + '_' + utente + formatdatetime('_yyyymmdd', now) + '.xlsx';
  end
  else
  begin
    SaveDialog.defaultext := 'xlsx';
    SaveDialog.FileName := normalizza_codice(nome_esportazione);
    SaveDialog.filter := 'file Excel (*.xlsx)|*.xlsx';
    SaveDialog.initialdir := cartella_esporta;
    if SaveDialog.execute then
    begin
      prosegui := true;
      nome_file := SaveDialog.FileName;
    end
    else
    begin
      prosegui := false;
    end;
  end;

  if prosegui then
  begin
    sor_xls.deletefields;
    if griglia = nil then
    begin
      for i := 0 to query_xls.fields.count - 1 do
      begin
        if query_xls.fields[i].datatype in [ftstring, ftwidestring] then
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, query_xls.fields[i].datatype, query_xls.fields[i].datasize, false);
        end
        else
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, query_xls.fields[i].datatype, 0, false);
        end;

      end;
    end
    else
    begin
      for i := 0 to griglia.datasource.dataset.fields.count - 1 do
      begin
        if griglia.datasource.dataset.fields[i].datatype in [ftstring, ftwidestring] then
        begin
          sor_xls.addfield(griglia.datasource.dataset.fields[i].FieldName, griglia.datasource.dataset.fields[i].datatype, query_xls.fields[i].datasize, false);
        end
        else if (griglia.datasource.dataset.fields[i].datatype = ftmemo) or (griglia.datasource.dataset.fields[i].datatype = ftblob) then
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, ftmemo, 0, false);
        end
        else
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, griglia.datasource.dataset.fields[i].datatype, 0, false);
        end;

      end;
    end;

    sor_xls.open;

    query_xls.first;
    while not query_xls.eof do
    begin
      sor_xls.append;
      for i := 0 to query_xls.fields.count - 1 do
      begin
        for j := 0 to sor_xls.fields.count - 1 do
        begin
          if query_xls.fields[i].FieldName = sor_xls.fields[j].FieldName then
          begin
            sor_xls.fields[j].value := query_xls.fields[i].value;
          end;
        end;
      end;
      sor_xls.post;

      query_xls.next;
    end;
    xlstxt(sor_xls, 'xls', nome_file);
    sor_xls.Close;

    if esegui_excel then
    begin
      esegui(nome_file);
    end;
  end;
end;

procedure TARC.esporta_csv(query_xls: tdataset; nome_esportazione: string; griglia: trzdbgrid_go);
var
  i, j: word;
begin
  SaveDialog.defaultext := 'csv';
  SaveDialog.FileName := normalizza_codice(nome_esportazione);
  SaveDialog.filter := 'file Excel (*.csv)|*.csv';
  SaveDialog.initialdir := cartella_esporta;
  if SaveDialog.execute then
  begin
    sor_xls.deletefields;
    if griglia = nil then
    begin
      for i := 0 to query_xls.fields.count - 1 do
      begin
        if query_xls.fields[i].datatype in [ftstring, ftwidestring] then
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, query_xls.fields[i].datatype, query_xls.fields[i].datasize, false);
        end
        else
        begin
          sor_xls.addfield(query_xls.fields[i].FieldName, query_xls.fields[i].datatype, 0, false);
        end;
      end;
    end
    else
    begin
      for i := 0 to griglia.datasource.dataset.fields.count - 1 do
      begin

        begin
          if griglia.datasource.dataset.fields[i].datatype in [ftstring, ftwidestring] then
          begin
            sor_xls.addfield(griglia.datasource.dataset.fields[i].FieldName, griglia.datasource.dataset.fields[i].datatype, query_xls.fields[i].datasize, false);
          end
          else if (griglia.datasource.dataset.fields[i].datatype = ftmemo) or (griglia.datasource.dataset.fields[i].datatype = ftblob) then
          begin
            sor_xls.addfield(query_xls.fields[i].FieldName, ftmemo, 0, false);
          end
          else
          begin
            sor_xls.addfield(query_xls.fields[i].FieldName, griglia.datasource.dataset.fields[i].datatype, 0, false);
          end;
        end;
      end;
    end;

    sor_xls.open;

    query_xls.first;
    while not query_xls.eof do
    begin
      sor_xls.append;
      for i := 0 to query_xls.fields.count - 1 do
      begin
        for j := 0 to sor_xls.fields.count - 1 do
        begin
          if query_xls.fields[i].FieldName = sor_xls.fields[j].FieldName then
          begin
            sor_xls.fields[j].value := query_xls.fields[i].value;
          end;
        end;
      end;

      sor_xls.post;

      query_xls.next;
    end;
    xlstxt(sor_xls, 'txt', SaveDialog.FileName);
    sor_xls.Close;

    esegui(SaveDialog.FileName);
  end;
end;

procedure TARC.xlstxt(tabella_esportare: TVirtualTable; tipo_esportazione, nome_file: string);
var
  i: word;
  nRiga: Integer;

  F: textfile;
  stringa: string;
const
  stile_stringa = 0;
  stile_numero = 1;
  stile_data = 2;
  stile_time = 3;
  stile_stringagrassetto = 4;
  stile_memo = 5;
begin
  if tipo_esportazione = 'xls' then
  begin
    nRiga := 0;

    xlswrite.FileName := nome_file;
    xlswrite.clear;
    xlswrite.add;
    xlswrite.Sheets[0].name := codice_procedura;

    // ----------------------------------------------------------------------------
    // Intestazione campi
    // ----------------------------------------------------------------------------
    for i := 0 to tabella_esportare.FieldCount - 1 do
    begin
      xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].FieldName;
    end;
    nRiga := nRiga + 1;

    tabella_esportare.first;
    while not tabella_esportare.eof do
    begin
      for i := 0 to tabella_esportare.FieldCount - 1 do
      begin
        if tabella_esportare.fields[i].datatype in [ftstring, ftwidestring] then
        begin
          xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].asstring;
        end
        else if (tabella_esportare.fields[i].datatype = ftSmallInt) or (tabella_esportare.fields[i].datatype = ftInteger) or (tabella_esportare.fields[i].datatype = ftFloat) or (tabella_esportare.fields[i].datatype = ftWord) then
        begin
          xlswrite.Sheets[0].asfloat[i, nRiga] := tabella_esportare.fields[i].asfloat;
        end
        else if (tabella_esportare.fields[i].datatype = ftDate) then
        begin
          xlswrite.Sheets[0].asdatetime[i, nRiga] := tabella_esportare.fields[i].asdatetime;
        end
        else if (tabella_esportare.fields[i].datatype = ftTime) then
        begin
          xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].asstring;
        end
        else if (tabella_esportare.fields[i].datatype = ftBoolean) then
        begin
          xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].asstring;
        end
        else if (tabella_esportare.fields[i].datatype = ftDatetime) then
        begin
          xlswrite.Sheets[0].asdatetime[i, nRiga] := tabella_esportare.fields[i].asdatetime;
        end
        else if (tabella_esportare.fields[i].datatype = ftmemo) or (tabella_esportare.fields[i].datatype = ftblob) then
        begin
          xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].asstring;
        end
        else
        begin
          try
            xlswrite.Sheets[0].asstring[i, nRiga] := tabella_esportare.fields[i].asstring;
          except
          end;
        end
      end;
      tabella_esportare.next;
      nRiga := nRiga + 1;
    end;

    try
      xlswrite.write;
    except
      messaggio(000, 'errore nella scrittura del file Excel' + #13 + 'verificare che non sia aperto da un altro programma');
    end;
  end
  else if tipo_esportazione = 'txt' then
  begin
    try
      try
        AssignFile(F, nome_file);
        Rewrite(F);

        // ----------------------------------------------------------------------------
        // Intestazione campi
        // ----------------------------------------------------------------------------
        stringa := '';
        for i := 0 to tabella_esportare.FieldCount - 1 do
        begin
          stringa := stringa + tabella_esportare.fields[i].FieldName + ';';
        end;
        copy(stringa, 1, length(stringa) - 1);
        WriteLn(F, stringa);
        tabella_esportare.first;
        while not(tabella_esportare.eof) do
        begin
          stringa := '';
          for i := 0 to tabella_esportare.FieldCount - 1 do
          begin
            if tabella_esportare.fields[i].datatype in [ftstring, ftwidestring] then
            begin
              stringa := stringa + tabella_esportare.fields[i].asstring + ';';
            end
            else if (tabella_esportare.fields[i].datatype = ftSmallInt) or (tabella_esportare.fields[i].datatype = ftInteger) or (tabella_esportare.fields[i].datatype = ftFloat) or (tabella_esportare.fields[i].datatype = ftWord) then
            begin
              stringa := stringa + tabella_esportare.fields[i].asstring + ';';
            end
            else if (tabella_esportare.fields[i].datatype = ftDate) then
            begin
              stringa := stringa + tabella_esportare.fields[i].asstring + ';';
            end
            else if (tabella_esportare.fields[i].datatype = ftTime) then
            begin
              stringa := stringa + tabella_esportare.fields[i].asstring + ';';
            end
            else if (tabella_esportare.fields[i].datatype = ftBoolean) then
            begin
              stringa := stringa + tabella_esportare.fields[i].asstring + ';';
            end

          end;
          stringa := copy(stringa, 1, length(stringa) - 1);
          WriteLn(F, stringa);

          tabella_esportare.next;
        end;

      except
        on e: exception do
        begin
          messaggio(000, e.message);
        end;
      end;
    finally
      closefile(F);
    end;
  end;
end;

procedure TARC.esporta_listbox(lista: tlistbox; nome_esportazione: string);
var
  i: word;
begin
  SaveDialog.defaultext := 'xls';
  SaveDialog.FileName := nome_esportazione;
  SaveDialog.filter := 'file Excel (*.xlsx)|*.xlsx';
  SaveDialog.initialdir := cartella_esporta;
  if SaveDialog.execute then
  begin
    screen.Cursor := crhourglass;

    if lista.count > 0 then
    begin
      sor_xls.deletefields;
      sor_xls.addfield('tabella', ftstring, 60, false);
      sor_xls.open;

      for i := 0 to lista.count - 1 do
      begin
        sor_xls.append;
        sor_xls.fieldbyname('tabella').asstring := lista.items[i];
        sor_xls.post;
      end;
      xlstxt(sor_xls, 'xls', SaveDialog.FileName);
      sor_xls.Close;
    end;

    screen.Cursor := cursore;
    esegui(SaveDialog.FileName);
  end;
end;

(*
  function TARC.setta_valore_generatore(connessione: TMyConnection_go; id_generatore: string): integer;
  begin
  result := setta_generatore(connessione, arc, id_generatore);
  end;

  function TARC.setta_valore_generatore(connessione: TMyConnection_go; id_generatore, codice_ditta: string): integer;
  begin
  result := setta_generatore(connessione, arc, id_generatore, codice_ditta);
  end;
*)

function TARC.setta_valore_generatore(connessione: TMyConnection_go; id_generatore: string; codice_ditta_passato: string = ''): Integer;
const
  PROCEDURE_NAME = 'P_OTTIENI_VALORE_GENERATORE';
  PARAM_I_CODICE = 'i_codice';
  PARAM_I_CODICE_DITTA = 'i_codice_ditta';
  PARAM_O_VALORE = 'o_valore';
var
  _id_generatore: string;
  _codice_ditta: string;
  _stored_procedure: TMyStoredProc;
begin
  _id_generatore := 'S_' + uppercase(id_generatore);
  if lowercase(connessione.database) = 'arc' then
  begin
    _codice_ditta := '';
  end
  else
  begin
    if codice_ditta_passato = '' then
    begin
      _codice_ditta := ditta;
    end
    else
    begin
      _codice_ditta := codice_ditta_passato;
    end;
  end;

  (*
    _stored_procedure := TMyStoredProc.Create(nil);
    try
    _stored_procedure.Connection := arcdit;
    _stored_procedure.StoredProcName := PROCEDURE_NAME;
    _stored_procedure.PrepareSQL;
    _stored_procedure.ParamByName(PARAM_I_CODICE).AsString := _id_generatore;
    _stored_procedure.ParamByName(PARAM_I_CODICE_DITTA).AsString := _codice_ditta;
    _stored_procedure.execute;
    valore := _stored_procedure.ParamByName(PARAM_O_VALORE).AsInteger;
    finally
    _stored_procedure.Free;
    end;
 *)

  prs_generatore.parambyname(PARAM_I_CODICE).asstring := _id_generatore;
  prs_generatore.parambyname(PARAM_I_CODICE_DITTA).asstring := _codice_ditta;
  prs_generatore.execute;

  result := prs_generatore.parambyname(PARAM_O_VALORE).asinteger;
end;

procedure TARC.esiste_dati_aggiuntivi_archivio(var tipo_control: Boolean; tabella: tmyquery_go; nome_tabella: string; var dati_utente_creazione: string; var dati_data_ora_creazione: tdatetime; var dati_utente: string; var dati_data_ora: tdatetime);
var
  query_dati_aggiuntivi: tmyquery_go;
begin
  query_dati_aggiuntivi := tmyquery_go.create(nil);
  query_dati_aggiuntivi.connection := TMyConnection_go(tabella.connection);
  query_dati_aggiuntivi.sql.text := 'select utente_creazione, data_ora_creazione, utente, data_ora ' + 'from ' + nome_tabella + ' where id = ' + inttostr(tabella.fieldbyname('id').asinteger);

  tipo_control := false;

  query_dati_aggiuntivi.open;
  if not query_dati_aggiuntivi.isempty then
  begin
    tipo_control := true;
    dati_utente_creazione := query_dati_aggiuntivi.fieldbyname('utente_creazione').asstring;
    dati_data_ora_creazione := query_dati_aggiuntivi.fieldbyname('data_ora_creazione').asdatetime;
    dati_utente := query_dati_aggiuntivi.fieldbyname('utente').asstring;
    dati_data_ora := query_dati_aggiuntivi.fieldbyname('data_ora').asdatetime;
  end;

  query_dati_aggiuntivi.free;
end;

function TARC.controllo_integrita_referenziale(nome_database: string; nome_tabella: string; codice_tabella: string; codice1_tabella: string; codice2_tabella: string; codice3_tabella: string; valore_codice: variant; valore_codice1: variant; valore_codice2: variant; valore_codice3: variant): Boolean;
var
  connessione: TMyConnection_go;
begin
  result := false;

  if nome_database = 'arc' then
  begin
    connessione := arc;
  end
  else
  begin
    connessione := arcdit;
  end;

  crea_query_ri(true, connessione, nome_tabella, codice_tabella, codice1_tabella, codice2_tabella, codice3_tabella, valore_codice, valore_codice1, valore_codice2, valore_codice3);

  query_ri.Close;
  query_ri.open;
  if not query_ri.eof then
  begin
    if messaggio(300, msg_0035 + #13 + msg_0036 + ' ' + nome_tabella + '''' + #13 + #13 + msg_0037) = 1 then
    begin
      crea_query_ri(false, connessione, nome_tabella, codice_tabella, codice1_tabella, codice2_tabella, codice3_tabella, valore_codice, valore_codice1, valore_codice2, valore_codice3);

      query_ri.open;
      esegui_visdel(nome_tabella, query_ri);
    end;
    result := true;
  end;
end;

procedure TARC.crea_query_ri(uno: Boolean; nome_database: TMyConnection_go; nome_tabella: string; codice_tabella: string; codice1_tabella: string; codice2_tabella: string; codice3_tabella: string; valore_codice: variant; valore_codice1: variant; valore_codice2: variant; valore_codice3: variant);
begin
  query_ri.Close;
  query_ri.connection := nome_database;
  query_ri.sql.clear;
  if uno then
  begin
    query_ri.sql.add('select id from ' + nome_tabella);
  end
  else
  begin
    query_ri.sql.add('select * from ' + nome_tabella);
  end;
  query_ri.sql.add('where ' + codice_tabella + '= :codice');
  if codice1_tabella <> '' then
  begin
    query_ri.sql.add('and ' + codice1_tabella + '= :codice1');
  end;
  if codice2_tabella <> '' then
  begin
    query_ri.sql.add('and ' + codice2_tabella + '= :codice2');
  end;
  if codice3_tabella <> '' then
  begin
    query_ri.sql.add('and ' + codice3_tabella + '= :codice3');
  end;
  if (nome_tabella = 'dit') or (nome_tabella = 'dit01') or (nome_tabella = 'dit02') or (nome_tabella = 'dit03') or (nome_tabella = 'dit04') or (nome_tabella = 'dit05') then
  begin
    query_ri.sql.add('and codice = ' + quotedstr(ditta));
  end;
  if uno then
  begin
    query_ri.sql.add('limit 1');
  end;

  query_ri.params[0].value := valore_codice;

  if codice1_tabella <> '' then
  begin
    query_ri.params[1].value := valore_codice1;
  end;

  if codice2_tabella <> '' then
  begin
    query_ri.params[2].value := valore_codice2;
  end;

  if codice3_tabella <> '' then
  begin
    query_ri.params[3].value := valore_codice3;
  end;
end;

procedure TARC.esegui_visdel(archivio: string; querydel: tmyquery_go);
var
  pr: tvisdel;
begin
  pr := tvisdel.create(nil);
  pr.query.connection := querydel.connection;
  pr.query.sql.clear;
  pr.query.sql.add(querydel.sql.text);
  pr.query.params := querydel.params;
  pr.archivio := archivio;
  pr.showmodal;
  pr.free;
end;

// assnum**********************************************************************

procedure TARC.assegna_numerazione(codice_ditta, tipo, serie: string; data: tdatetime; var data_precedente: tdatetime; var progressivo, progressivo_precedente: double; aggiorna: Boolean; avviso: Boolean = true);
begin
  assegna_numerazione(arcdit, codice_ditta, tipo, serie, data, data_precedente, progressivo, progressivo_precedente, aggiorna, avviso);
end;

procedure TARC.assegna_numerazione(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data: tdatetime; var data_precedente: tdatetime; var progressivo, progressivo_precedente: double; aggiorna: Boolean; avviso: Boolean = true);
var
  anno, mese, giorno: word;
begin
  decodedate(data, anno, mese, giorno);

  assnum_cnt.Close;
  assnum_cnt.connection := connessione;

  if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
  begin
    assnum_cnt.parambyname('anno').asstring := '';
  end
  else
  begin
    assnum_cnt.parambyname('anno').asstring := inttostr(anno);
  end;
  assnum_cnt.parambyname('tipo').asstring := tipo;
  assnum_cnt.parambyname('sottotipo').asstring := serie;
  assnum_cnt.open;
  if assnum_cnt.isempty then
  begin
    assnum_cnt.append;
    assnum_cnt.fieldbyname('tipo').asstring := tipo;
    if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
    begin
      assnum_cnt.fieldbyname('anno').asstring := '';
    end
    else
    begin
      assnum_cnt.fieldbyname('anno').asstring := inttostr(anno);
    end;
    assnum_cnt.fieldbyname('sottotipo').asstring := serie;
    assnum_cnt.post;
  end;

  if assnum_cnt.fieldbyname('data_aggiornamento').asdatetime > data then
  begin
    if avviso then
    begin
      messaggio(000, 'l''ultimo progressivo è stato assegnato in data superiore rispetto a quella attuale');
    end;
  end;

  assnum_cnt.edit;

  progressivo_precedente := strtoint(floattostr(assnum_cnt.fieldbyname('progressivo').asfloat));
  data_precedente := assnum_cnt.fieldbyname('data_aggiornamento').asdatetime;

  if progressivo = 0 then
  begin
    assnum_cnt.fieldbyname('progressivo').asfloat := arrotonda(assnum_cnt.fieldbyname('progressivo').asfloat + 1, 0);
    assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
    progressivo := strtoint(floattostr(assnum_cnt.fieldbyname('progressivo').asfloat));
  end
  else
  begin
    if not avviso then
    begin
      assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
      assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
    end
    else
    begin
      if assnum_cnt.fieldbyname('progressivo').asinteger >= progressivo then
      begin
        if messaggio(304, msg_0019 + ' (' + floattostr(progressivo) + ')' + #13 + msg_0020 + ' (' + floattostr(assnum_cnt.fieldbyname('progressivo').asinteger) + ')' + #13 + #13 + msg_0021) = 1 then
        begin
          progressivo := trunc(assnum_cnt.fieldbyname('progressivo').asfloat + 1);
          assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
          assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
        end;
      end
      else if assnum_cnt.fieldbyname('progressivo').asfloat + 1 <> progressivo then
      begin
        if messaggio(304, msg_0019 + ' (' + floattostr(progressivo) + ')' + #13 + msg_0022 + ' (' + floattostr(assnum_cnt.fieldbyname('progressivo').asfloat) + ')' + #13 + #13 + msg_0023) = 1 then
        begin
          assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
          assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
          progressivo := assnum_cnt.fieldbyname('progressivo').asfloat;
        end
        else
        begin
          progressivo := assnum_cnt.fieldbyname('progressivo').asfloat + 1;
          assnum_cnt.fieldbyname('progressivo').asfloat := assnum_cnt.fieldbyname('progressivo').asfloat + 1;
          assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
        end;
      end
      else
      begin
        assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
        assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
      end;
    end;
  end;

  if (aggiorna) and (progressivo <> 0) and (assnum_cnt.state <> dsbrowse) then
  begin
    assnum_cnt.post;
  end
  else
  begin
    assnum_cnt.cancel;
  end;
  assnum_cnt.Close;
  assnum_cnt.connection := arcdit;
end;

procedure TARC.assegna_numerazione(connessione: TMyConnection_go; tipo, serie: string; data: tdatetime; var progressivo: double; aggiorna: Boolean = true; avviso: Boolean = true);
var
  prosegui: Boolean;
  anno, mese, giorno: word;
  anno_stringa: string;
begin
  prosegui := true;
  if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
  begin
    anno_stringa := '';
  end
  else
  begin
    if data = 0 then
    begin
      prosegui := false;
    end
    else
    begin
      decodedate(data, anno, mese, giorno);
      anno_stringa := inttostr(anno);
    end;
  end;

  if prosegui then
  begin
    assnum_cnt.Close;
    assnum_cnt.connection := connessione;

    // assnum_cnt.lockmode := lmpessimistic;

    assnum_cnt.parambyname('anno').asstring := anno_stringa;
    assnum_cnt.parambyname('tipo').asstring := tipo;
    assnum_cnt.parambyname('sottotipo').asstring := serie;

    if aggiorna then
    begin
      aggiorna_cnt(anno_stringa, tipo, serie, data, progressivo, avviso);
      assnum_cnt.post;
    end
    else
    begin
      assnum_cnt.open;
      progressivo := assnum_cnt.fieldbyname('progressivo').asinteger + 1;
    end;

    assnum_cnt.Close;
    assnum_cnt.connection := arcdit;
  end;
end;

procedure TARC.aggiorna_cnt(anno, tipo, sottotipo: string; data: tdatetime; var progressivo: double; avviso: Boolean = true);
begin
  assnum_cnt.Close;
  assnum_cnt.open;
  try
    if assnum_cnt.isempty then
    begin
      assnum_cnt.append;
      assnum_cnt.fieldbyname('anno').asstring := anno;
      assnum_cnt.fieldbyname('tipo').asstring := tipo;
      assnum_cnt.fieldbyname('sottotipo').asstring := sottotipo;
      assnum_cnt.post;
      assnum_cnt.refresh;
    end;

    if avviso then
    begin
      if assnum_cnt.fieldbyname('data_aggiornamento').asdatetime > data then
      begin
        messaggio(200, 'l''ultima data di aggiornamento del contatore ' + slinebreak + '[' + tipo + ' ' + sottotipo + ' ' + formatdatetime('dd/mm/yyyy', assnum_cnt.fieldbyname('data_aggiornamento').asdatetime) + ']' + slinebreak +
          'è maggiore della data con cui si sta eseguendo l''aggiornamento attuale ' + slinebreak + '[' + formatdatetime('dd/mm/yyyy', data) + ']');
      end;
    end;

    assnum_cnt.edit;
    if progressivo = 0 then
    begin
      assnum_cnt.fieldbyname('progressivo').asfloat := assnum_cnt.fieldbyname('progressivo').asfloat + 1;
      assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
      progressivo := assnum_cnt.fieldbyname('progressivo').asfloat;
    end
    else
    begin
      if assnum_cnt.fieldbyname('progressivo').asfloat + 1 <> progressivo then
      begin
        // if not avviso or (avviso and (messaggio(304, 'il progressivo inserito [' + floattostr(progressivo) + ' ' + sottotipo + ']' + #13 +
        if avviso and (avviso_assegna_numerazione <> 'no') and (messaggio(304, 'il progressivo inserito [' + floattostr(progressivo) + ' ' + sottotipo + ']' + #13 + 'non è successivo all''ultimo memorizzato nell''archivio contatori [' + assnum_cnt.fieldbyname('progressivo').asstring + ' ' +
          sottotipo + ']' + #13 + 'assegna il primo numero disponibile') = 1) then
        begin
          assnum_cnt.fieldbyname('progressivo').asfloat := assnum_cnt.fieldbyname('progressivo').asfloat + 1;
          assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
          progressivo := assnum_cnt.fieldbyname('progressivo').asfloat;
        end
        else if assnum_cnt.fieldbyname('progressivo').asfloat + 1 < progressivo then
        begin
          assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
          assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
        end;

        if avviso_assegna_numerazione <> '' then
        begin
          if messaggio(304, 'conferma per non visualizzare più il messaggio di avviso incongruenza numerazione') = 1 then
          begin
            avviso_assegna_numerazione := 'no';
          end
          else
          begin
            avviso_assegna_numerazione := 'eseguito';
          end;
        end;
      end
      else
      begin
        assnum_cnt.fieldbyname('progressivo').asfloat := progressivo;
        assnum_cnt.fieldbyname('data_aggiornamento').asdatetime := data;
      end;
    end;
  except
    aggiorna_cnt(anno, tipo, sottotipo, data, progressivo);
  end;
end;

procedure TARC.storna_numerazione(codice_ditta, tipo, serie: string; data, data_precedente: tdatetime; progressivo, progressivo_precedente: double);
begin
  storna_numerazione(arcdit, codice_ditta, tipo, serie, data, data_precedente, progressivo, progressivo_precedente);
end;

procedure TARC.storna_numerazione(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data, data_precedente: tdatetime; progressivo, progressivo_precedente: double);
var
  anno, mese, giorno: word;
begin
  decodedate(data, anno, mese, giorno);

  assnum_cnt.Close;
  assnum_cnt.connection := connessione;

  if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
  begin
    assnum_cnt.parambyname('anno').asstring := '';
  end
  else
  begin
    assnum_cnt.parambyname('anno').asstring := inttostr(anno);
  end;
  assnum_cnt.parambyname('tipo').asstring := tipo;
  assnum_cnt.parambyname('sottotipo').asstring := serie;
  assnum_cnt.open;

  if assnum_cnt.fieldbyname('progressivo').asfloat = progressivo then
  begin
    assnum_cnt.edit;
    assnum_cnt.fieldbyname('progressivo').asfloat := progressivo_precedente;
    assnum_cnt.fieldbyname('data_aggiornamento').asfloat := data_precedente;
    assnum_cnt.post;
  end
  else
  begin
    messaggio(200, 'non è stato possibile ripristinare la numerazione annullata' + #13 + tipo + ' numero ' + floattostr(progressivo) + ' ' + serie + '  data ' + datetostr(data) + #13 + 'perché è già stato assegnato il progressivo: ' + assnum_cnt.fieldbyname('progressivo').asstring);
  end;

  assnum_cnt.Close;
  assnum_cnt.connection := arcdit;
end;

procedure TARC.storna_numerazione(connessione: TMyConnection_go; tipo, serie: string; data: tdatetime; progressivo: double);
var
  anno, mese, giorno: word;
begin
  if progressivo <> 0 then
  begin
    decodedate(data, anno, mese, giorno);

    assnum_cnt.Close;
    assnum_cnt.connection := connessione;

    if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
    begin
      assnum_cnt.parambyname('anno').asstring := '';
    end
    else
    begin
      assnum_cnt.parambyname('anno').asstring := inttostr(anno);
    end;
    assnum_cnt.parambyname('tipo').asstring := tipo;
    assnum_cnt.parambyname('sottotipo').asstring := serie;

    storna_cnt(tipo, serie, data, progressivo);

    if assnum_cnt.state = dsedit then
    begin
      assnum_cnt.post;
    end;

    assnum_cnt.Close;
    assnum_cnt.connection := arcdit;
  end;
end;

procedure TARC.storna_cnt(tipo, serie: string; data: tdatetime; progressivo: double);
begin
  assnum_cnt.Close;
  assnum_cnt.open;
  try
    if assnum_cnt.fieldbyname('progressivo').asfloat = progressivo then
    begin
      assnum_cnt.edit;
      assnum_cnt.fieldbyname('progressivo').asfloat := progressivo - 1;
    end
    else
    begin
      messaggio(200, 'non è stato possibile ripristinare la numerazione annullata' + #13 + tipo + ' numero ' + floattostr(progressivo) + ' ' + serie + '  data ' + datetostr(data) + #13 + 'perché è già stato assegnato il progressivo: ' + assnum_cnt.fieldbyname('progressivo').asstring + #13 +
        'diverso da quello che si vuole stornare: ' + floattostr(progressivo));
    end;
  except
    storna_cnt(tipo, serie, data, progressivo);
  end;
end;

procedure TARC.storna_numerazione_cancellata(codice_ditta, tipo, serie: string; data: tdatetime; numero: double);
begin
  storna_numerazione_cancellata(arcdit, codice_ditta, tipo, serie, data, numero);
end;

procedure TARC.storna_numerazione_cancellata(connessione: TMyConnection_go; codice_ditta, tipo, serie: string; data: tdatetime; numero: double);
var
  anno, mese, giorno: word;
begin
  decodedate(data, anno, mese, giorno);

  assnum_cnt.Close;
  assnum_cnt.connection := connessione;
  if (tipo = 'DICHIARAZIONI INTRASTAT') or (tipo = 'CONFIGURAZIONE') or (tipo = 'FATTURAZIONE ELETTRONICA PA') or (copy(tipo, 1, 7) = 'CFGART-') or (tipo = 'REGISTRO COMMERCIALIZZAZIONE') or (copy(tipo, 1, 3) = 'SDA') then
  begin
    assnum_cnt.parambyname('anno').asstring := '';
  end
  else
  begin
    assnum_cnt.parambyname('anno').asstring := inttostr(anno);
  end;
  assnum_cnt.parambyname('tipo').asstring := tipo;
  assnum_cnt.parambyname('sottotipo').asstring := serie;
  assnum_cnt.open;

  if assnum_cnt.fieldbyname('progressivo').asfloat = numero then
  begin
    assnum_cnt.edit;

    assnum_cnt.fieldbyname('progressivo').asfloat := numero - 1;
    assnum_cnt.post;
  end
  else
  begin
    messaggio(200, 'non è stato possibile ripristinare la numerazione annullata' + #13 + tipo + ' numero ' + floattostr(numero) + ' ' + serie + '  data ' + datetostr(data) + #13 + 'perché è già stato assegnato il progressivo: ' + assnum_cnt.fieldbyname('progressivo').asstring);
  end;

  assnum_cnt.Close;
  assnum_cnt.connection := arcdit;
end;

function TARC.esistenza_documento(documento, serie, cfg_codice: string; data: tdatetime; numero: string; progressivo: Integer; revisione: Integer = 0): Boolean;
begin
end;

function TARC.esistenza_documento(connessione: TMyConnection_go; documento, serie, cfg_codice: string; data: tdatetime; numero: string; progressivo: Integer; revisione: Integer = 0): Boolean;
begin
end;

function TARC.esistenza_documento(documento, serie, cfg_codice: string; data: tdatetime; numero: double; progressivo: Integer; revisione: Integer = 0): Boolean;
begin
  result := esistenza_documento(arcdit, documento, serie, cfg_codice, data, numero, progressivo, revisione);
end;

function TARC.esistenza_documento(connessione: TMyConnection_go; documento, serie, cfg_codice: string; data: tdatetime; numero: double; progressivo: Integer; revisione: Integer = 0): Boolean;
var
  anno, mese, giorno: word;
begin
  result := false;
  esistenza_numerazione.Close;
  esistenza_numerazione.connection := connessione;
  if numero <> 0 then
  begin
    decodedate(data, anno, mese, giorno);
    if (documento = 'dat') or (documento = 'fat') then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from ' + documento);
      esistenza_numerazione.sql.add('where frn_codice = :frn_codice');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asstring := cfg_codice;
      esistenza_numerazione.params[1].asstring := serie;
      esistenza_numerazione.params[2].asfloat := numero;
      esistenza_numerazione.params[3].asfloat := progressivo;
      esistenza_numerazione.params[4].asinteger := anno;
    end
    else if documento = 'pnt_acquisti' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from pnt');
      esistenza_numerazione.sql.add('where cfg_tipo = ''F'' and cfg_codice = :cfg_codice');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asstring := cfg_codice;
      esistenza_numerazione.params[1].asstring := serie;
      esistenza_numerazione.params[2].asfloat := numero;
      esistenza_numerazione.params[3].asfloat := progressivo;
      esistenza_numerazione.params[4].asinteger := anno;
    end
    else if documento = 'pnt_acquisti_protocollo' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_registrazione, protocollo from pnt');
      esistenza_numerazione.sql.add('where protocollo = :protocollo');
      esistenza_numerazione.sql.add('and serie_protocollo = :serie_protocollo');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_registrazione) = :anno_data_registrazione');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asstring := serie;
      esistenza_numerazione.params[2].asfloat := progressivo;
      esistenza_numerazione.params[3].asinteger := anno;
    end
    else if documento = 'pnt_vendite' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from pnt');
      esistenza_numerazione.sql.add('where cfg_tipo = ''C'' and tipo_documento_iva = ''vendite''');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asstring := serie;
      esistenza_numerazione.params[1].asfloat := numero;
      esistenza_numerazione.params[2].asfloat := progressivo;
      esistenza_numerazione.params[3].asinteger := anno;
    end
    else if documento = 'pnt_corrispettivi' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from pnt');
      esistenza_numerazione.sql.add('where tipo_documento_iva = ''corrispettivi''');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asstring := serie;
      esistenza_numerazione.params[1].asfloat := numero;
      esistenza_numerazione.params[2].asfloat := progressivo;
      esistenza_numerazione.params[3].asinteger := anno;
    end
    else if documento = 'opt' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, serie_documento, numero_documento from opt');
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento and serie_documento = :serie_documento');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
      esistenza_numerazione.params[3].asstring := serie;
    end
    else if documento = 'olt' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from olt');
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
    end
    else if (documento = 'bvt') or (documento = 'cvt') or (documento = 'dvt') or (documento = 'fvt') or (documento = 'ovt') or (documento = 'pvt') or (documento = 'prat') then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.sql.add('and revisione = :revisione');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asstring := serie;
      esistenza_numerazione.params[2].asfloat := progressivo;
      esistenza_numerazione.params[3].asinteger := anno;
      esistenza_numerazione.params[4].asinteger := revisione;
    end
    else if (documento = 'rat') or (documento = 'oat') then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and serie_documento = :serie_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.sql.add('and revisione = :revisione');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asstring := serie;
      esistenza_numerazione.params[2].asfloat := progressivo;
      esistenza_numerazione.params[3].asinteger := anno;
      esistenza_numerazione.params[4].asinteger := revisione;
    end
    else if documento = 'oatfg' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
    end
    else if documento = 'spd' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
    end
    else if documento = 'lti' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_registrazione, numero_registrazione from ' + documento);
      esistenza_numerazione.sql.add('where numero_registrazione = :numero_registrazione');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_registrazione) = :anno_data_registrazione');
      esistenza_numerazione.sql.add('and cfg_tipo = :serie');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
      esistenza_numerazione.params[3].asstring := serie;
    end
    else if documento = 'ripct' then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo, data_documento, numero_documento from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.sql.add('and tipo_chiamata = :tipo_chiamata');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
      esistenza_numerazione.params[3].asstring := serie;

    end
    else if (documento = 'bvcmrt') or (documento = 'dvcmrt') or (documento = 'fvcmrt') then
    begin
      esistenza_numerazione.sql.clear;
      esistenza_numerazione.sql.add('select progressivo from ' + documento);
      esistenza_numerazione.sql.add('where numero_documento = :numero_documento');
      esistenza_numerazione.sql.add('and progressivo <> :progressivo');
      esistenza_numerazione.sql.add('and year(data_documento) = :anno_data_documento');
      esistenza_numerazione.params[0].asfloat := numero;
      esistenza_numerazione.params[1].asfloat := progressivo;
      esistenza_numerazione.params[2].asinteger := anno;
    end;

    esistenza_numerazione.open;
    if not esistenza_numerazione.eof then
    begin
      messaggio(000, 'il numero' + ' [' + floattostr(numero) + '] è già stato utilizzato nell''anno' + #13 + 'nel documento con progressivo: ' + esistenza_numerazione.fieldbyname('progressivo').asstring);
      result := true;
    end;
  end;
end;
// fine assnum*****************************************************************

function TARC.controllo_ora(ora: string): string;
var
  i: word;
  intero, decimale: string;
begin
  result := '00' + formatsettings.timeseparator + '00';

  if trim(ora) = '' then
  begin
    result := '00' + formatsettings.timeseparator + '00';
  end
  else
  begin
    for i := 1 to length(ora) do
    begin
      if (not numerico(copy(ora, i, 1))) and (copy(ora, i, 1) <> formatsettings.timeseparator) then
      begin
        result := '00' + formatsettings.timeseparator + '00';
      end;
    end;
    if pos(formatsettings.timeseparator, ora) <> 0 then
    begin
      intero := setta_lunghezza(strtoint(copy(ora, 1, pos(formatsettings.timeseparator, ora) - 1)), 2, 0);
      decimale := setta_lunghezza(strtoint(copy(ora, pos(formatsettings.timeseparator, ora) + 1, length(ora) - pos(formatsettings.timeseparator, ora))), 2, 0);
    end
    else
    begin
      if length(ora) <= 2 then
      begin
        intero := setta_lunghezza(strtoint(copy(ora, 1, length(ora))), 2, 0);
        decimale := '00';
      end
      else
      begin
        intero := copy(ora, 1, 2);
        if length(ora) = 4 then
        begin
          decimale := copy(ora, 3, 2);
        end
        else
        begin
          decimale := copy(ora, 3, 1) + '0';
        end;
      end;
    end;
    if intero > '23' then
    begin
      intero := '23';
    end;
    if decimale > '59' then
    begin
      decimale := '59';
    end;

    result := intero + formatsettings.timeseparator + decimale;
  end;
end;

function TARC.Unita(k: Integer): string;
var
  lettere: array [0 .. 20] of string;
begin

  lettere[0] := '';
  lettere[1] := 'uno';
  lettere[2] := 'due';
  lettere[3] := 'tre';
  lettere[4] := 'quattro';
  lettere[5] := 'cinque';
  lettere[6] := 'sei';
  lettere[7] := 'sette';
  lettere[8] := 'otto';
  lettere[9] := 'nove';
  lettere[10] := 'dieci';
  lettere[11] := 'undici';
  lettere[12] := 'dodici';
  lettere[13] := 'tredici';
  lettere[14] := 'quattordici';
  lettere[15] := 'quindici';
  lettere[16] := 'sedici';
  lettere[17] := 'diciassette';
  lettere[18] := 'diciotto';
  lettere[19] := 'diciannove';

  if (k < 0) or (k > high(lettere)) then
    result := ''
  else
    result := lettere[k];
end;

function TARC.Decine(k: Integer): string;
var
  lettere: array [0 .. 9] of string;
begin
  lettere[0] := '';
  lettere[1] := 'dieci';
  lettere[2] := 'venti';
  lettere[3] := 'trenta';
  lettere[4] := 'quaranta';
  lettere[5] := 'cinquanta';
  lettere[6] := 'sessanta';
  lettere[7] := 'settanta';
  lettere[8] := 'ottanta';
  lettere[9] := 'novanta';

  if (k < 0) or (k > high(lettere)) then
  begin
    result := ''
  end
  else
  begin
    result := lettere[k];
  end;
end;

function TARC.Migliaia(k: Integer): string;
var
  lettere: array [0 .. 10] of string;
begin
  lettere[0] := '';
  lettere[1] := 'mille';
  lettere[2] := 'unmilione';
  lettere[3] := 'unmiliardo';
  lettere[4] := 'millemiliardi';
  lettere[5] := 'mila';
  lettere[6] := 'milioni';
  lettere[7] := 'miliardi';
  lettere[8] := 'milamiliardi';
  lettere[9] := 'milamiliardi';
  lettere[10] := 'migliaiadimiliardi';
  if (k < 0) or (k > high(lettere)) then
    result := ''
  else
    result := lettere[k];
end;

function TARC.CalcolaLettere(Importo: double): string;
var
  intero, parziale, tripla, resto, s: string;
  tv, td, tc, k, mille, x, y: Integer;
begin
  result := '';

  intero := Formatfloat('0.00', Importo);
  resto := ' / ' + copy(intero, length(intero) - 1, 2);
  intero := copy(intero, 1, length(intero) - 3);
  if copy(intero, 1, 1) = '-' then
  begin
    intero := copy(intero, 2, length(intero));
  end;

  if Importo = 0 then
  begin
    CalcolaLettere := 'zero / 00 ';
    exit
  end;

  mille := -1;
  k := length(intero) mod 3;
  if not(k = 0) then
  begin
    intero := replicate('0', 3 - k) + intero;
  end;

  while not(intero = '') do
  begin
    mille := mille + 1;
    parziale := '';
    tripla := copy(intero, length(intero) - 2, 3);
    s := '';
    intero := copy(intero, 1, length(intero) - 3);

    tv := strtoint(tripla);

    td := tv mod 100;

    tc := (tv - td) div 100;
    if not(tc = 0) then
    begin
      parziale := 'cento';
      if tc > 1 then
      begin
        parziale := Unita(tc) + parziale;
      end;
    end;

    if td < 20 then
      parziale := parziale + Unita(td)
    else
    begin
      x := td mod 10;
      y := (td - x) div 10;
      parziale := parziale + Decine(y);
      s := Unita(x);
      if (pos(copy(s, 1, 1), 'uo') > 0) and (s <> '') and (not(y = 0)) then
        parziale := copy(parziale, 1, length(parziale) - 1);

      parziale := parziale + s;
    end;
    s := Migliaia(mille);
    if (mille > 0) and (not(parziale = '')) then
    begin
      k := mille;
      if not(parziale = 'uno') then
      begin
        k := k + 4;

        s := Migliaia(k);
        if copy(parziale, length(parziale) - 2, 3) = 'uno' then
        begin
          parziale := copy(parziale, 1, length(parziale) - 1)
        end
      end
      else
        parziale := '';

      parziale := parziale + s;
    end; // if

    result := parziale + result;
  end; // while
  if Importo < 0 then
  begin
    result := 'meno' + result;
  end;

  result := result + resto;
end;

function sconto(tsm_codice: string): double;
var
  tsm: tmyquery_go;
begin
  result := 100;

  tsm := tmyquery_go.create(nil);
  tsm.connection := arc.arcdit;
  tsm.sql.text := 'select percentuale_totale from tsm where codice = :codice';
  tsm.parambyname('codice').asstring := tsm_codice;
  tsm.open;
  if not tsm.isempty then
  begin
    result := tsm.fieldbyname('percentuale_totale').asfloat;
  end;

  tsm.free;
end;

procedure calcola_importo_documento(quantita, prezzo, cambio, importo_sconto: double; sconto_imponibile_lordo, listino_con_iva, tum_codice, tiv_codice, tsm_codice, tsm_codice_art: string; var Importo, importo_euro, importo_iva, importo_iva_euro, importo_non_arrotondato: double;
  solo_iva: Boolean = false);
var
  imponibile: double;
  decimali: double;
  interi: Integer;

  tum, tiv: tmyquery_go;
begin
  tum := tmyquery_go.create(nil);
  tum.connection := arc.arcdit;
  tum.sql.text := 'select gestione_minuti from tum where codice = :codice';

  tiv := tmyquery_go.create(nil);
  tiv.connection := arc.arcdit;
  tiv.sql.text := 'select percentuale from tiv where codice = :codice';

  if not solo_iva then
  begin
    if not((quantita = 0) or (prezzo = 0)) then
    begin
      tum.Close;
      tum.parambyname('codice').asstring := tum_codice;
      tum.open;
      if tum.fieldbyname('gestione_minuti').asstring = 'si' then
      begin
        interi := trunc(quantita);
        decimali := quantita - interi;
        if sconto_imponibile_lordo = 'si' then
        begin
          importo_non_arrotondato := (interi + decimali * 100 / 60) * prezzo * sconto(tsm_codice) / 100;
          importo_non_arrotondato := importo_non_arrotondato - ((interi + decimali * 100 / 60) * prezzo * (1 - sconto(tsm_codice_art) / 100));

          Importo := arrotonda((interi + decimali * 100 / 60) * prezzo * sconto(tsm_codice) / 100);
          Importo := arrotonda(Importo - ((interi + decimali * 100 / 60) * prezzo * (1 - sconto(tsm_codice_art) / 100)));
        end
        else
        begin
          importo_non_arrotondato := (interi + decimali * 100 / 60) * prezzo * sconto(tsm_codice) * sconto(tsm_codice_art) / 10000;

          Importo := arrotonda((interi + decimali * 100 / 60) * prezzo * sconto(tsm_codice) * sconto(tsm_codice_art) / 10000);
        end;
      end
      else
      begin
        if sconto_imponibile_lordo = 'si' then
        begin
          importo_non_arrotondato := quantita * prezzo * sconto(tsm_codice) / 100;
          importo_non_arrotondato := importo_non_arrotondato - (quantita * prezzo * (1 - sconto(tsm_codice_art) / 100));

          Importo := arrotonda(quantita * prezzo * sconto(tsm_codice) / 100);
          Importo := arrotonda(Importo - (quantita * prezzo * (1 - sconto(tsm_codice_art) / 100)));
        end
        else
        begin
          importo_non_arrotondato := quantita * prezzo * sconto(tsm_codice) * sconto(tsm_codice_art) / 10000;

          Importo := arrotonda(quantita * prezzo * sconto(tsm_codice) * sconto(tsm_codice_art) / 10000);
        end;
      end;
      importo_non_arrotondato := importo_non_arrotondato - importo_sconto;

      Importo := arrotonda(Importo - importo_sconto);
    end;
    importo_euro := arrotonda(Importo / cambio);
  end;

  tiv.Close;
  tiv.parambyname('codice').asstring := tiv_codice;
  tiv.open;
  if not tiv.isempty then
  begin
    if listino_con_iva = 'no' then
    begin
      importo_iva := arrotonda(Importo * tiv.fieldbyname('percentuale').asfloat / 100);
      importo_iva_euro := arrotonda(importo_iva / cambio);
    end
    else
    begin
      imponibile := arc.scorporo(Importo, tiv.fieldbyname('percentuale').asfloat);
      importo_iva := arrotonda(Importo - imponibile);
      importo_iva_euro := arrotonda(importo_iva / cambio);
    end;
  end;

  tum.free;
  tiv.free;
end;

function TARC.scorporo(Importo, percentuale: double; decimali: word = 2): double;
var
  negativo: Boolean;
  imponibile, iva: currency;
  n: extended;
begin
  if imponibile < 0 then
  begin
    negativo := true;
  end
  else
  begin
    negativo := false;
  end;

  imponibile := arrotonda(Importo / (1 + percentuale / 100), decimali);
  iva := arrotonda(imponibile * percentuale / 100, decimali);
  result := imponibile;

  if negativo then
  begin
    result := result * -1;
  end;
end;

function TARC.scorporo(Importo: double; art_codice: string; scorpora: Boolean; vendite: Boolean = true): double;
var
  query: tmyquery_go;
begin
  result := Importo;

  query := tmyquery_go.create(nil);
  query.connection := arcdit;
  if vendite then
  begin
    query.sql.text := 'select tiv.percentuale from art inner join tiv on tiv.codice = art.tiv_codice_vendite ' + 'where art.codice = ' + quotedstr(art_codice);
  end
  else
  begin
    query.sql.text := 'select tiv.percentuale from art inner join tiv on tiv.codice = art.tiv_codice_acquisti ' + 'where art.codice = ' + quotedstr(art_codice);
  end;
  query.open;
  if not query.eof then
  begin
    if scorpora then
    begin
      result := scorporo(Importo, query.fieldbyname('percentuale').asfloat, decimali_max_prezzo)
    end
    else
    begin
      result := Importo * (1 + query.fieldbyname('percentuale').asfloat / 100);
    end;
  end;
  query.free;
end;

function assegna_parametri_lavoro: string;
var
  cartella_archivi_ditta: string;
  log_inserimento, log_modifica, log_cancellazione: tstringlist;

  ese, tva, tna: tmyquery_go;
begin
  result := '';

  screen.Cursor := crhourglass;

  try
    log_inserimento := tstringlist.create;
    log_modifica := tstringlist.create;
    log_cancellazione := tstringlist.create;

    arc.dit.Close;
    arc.dit.parambyname('codice').asstring := ditta;
    arc.dit.open;
    result := arc.dit.fieldbyname('descrizione1').asstring;

    ese := tmyquery_go.create(nil);
    ese.connection := arc.arc;
    ese.sql.text := 'select ese.*, e.data_inizio data_inizio_precedente, ' + 'e.data_fine data_fine_precedente, e.data_bilancio data_bilancio_precedente, ' + 'coalesce(e.esercizio_chiuso, ''si'') esercizio_chiuso_precedente, ' +
      'coalesce(e.esercizio_chiuso_magazzino, ''si'') esercizio_chiuso_magazzino_precedente ' + 'from ese left join ese e on e.dit_codice = ese.dit_codice and e.codice = ese.esercizio_precedente ' + 'where ese.dit_codice = :dit_codice and ese.codice = :codice';
    ese.parambyname('dit_codice').asstring := ditta;
    ese.parambyname('codice').asstring := esercizio;
    ese.open;
    result := result + msg_0093 + ese.fieldbyname('descrizione').asstring;

    if storico then
    begin
      result := msg_0074 + ' ' + result;
    end;

    // aggiorna parametri variabili esterne ditta
    descrizione_ditta := arc.dit.fieldbyname('descrizione1').asstring;
    descrizione2_ditta := arc.dit.fieldbyname('descrizione2').asstring;
    via_ditta := arc.dit.fieldbyname('via').asstring;
    cap_ditta := arc.dit.fieldbyname('cap').asstring;
    citta_ditta := arc.dit.fieldbyname('citta').asstring;
    provincia_ditta := arc.dit.fieldbyname('provincia').asstring;
    tna_codice_ditta := arc.dit.fieldbyname('tna_codice').asstring;
    codice_fiscale_ditta := arc.dit.fieldbyname('codice_fiscale').asstring;
    partita_iva_ditta := arc.dit.fieldbyname('partita_iva').asstring;

    registri_prenumerati_ditta := arc.dit.fieldbyname('registri_prenumerati').asstring;
    registro_imprese_ditta := arc.dit.fieldbyname('registro_imprese').asstring;
    via_fiscale_ditta := arc.dit.fieldbyname('via_fiscale').asstring;
    cap_fiscale_ditta := arc.dit.fieldbyname('cap_fiscale').asstring;
    citta_fiscale_ditta := arc.dit.fieldbyname('citta_fiscale').asstring;
    provincia_fiscale_ditta := arc.dit.fieldbyname('provincia_fiscale').asstring;
    marchio_percorso_ditta := arc.dit.fieldbyname('marchio_percorso').asstring;
    marchio_sinistra_ditta := arc.dit.fieldbyname('marchio_sinistra').asinteger;
    marchio_superiore_ditta := arc.dit.fieldbyname('marchio_superiore').asinteger;
    marchio_altezza_ditta := arc.dit.fieldbyname('marchio_altezza').asinteger;
    marchio_larghezza_ditta := arc.dit.fieldbyname('marchio_larghezza').asinteger;
    telefono_ditta := arc.dit.fieldbyname('telefono').asstring;
    fax_ditta := arc.dit.fieldbyname('fax').asstring;
    web_ditta := arc.dit.fieldbyname('web').asstring;
    e_mail_ditta := arc.dit.fieldbyname('e_mail').asstring;
    cartella_stampe_ditta := arc.dit.fieldbyname('cartella_stampe').asstring;
    capitale_sociale_ditta := arc.dit.fieldbyname('capitale_sociale').asfloat;

    divisa_di_conto := arc.dit.fieldbyname('tva_codice').asstring;
    codice_nom_numerico := arc.dit.fieldbyname('codice_nom_numerico').asstring;
    codice_articolo_numerico := arc.dit.fieldbyname('codice_articolo_numerico').asstring;
    codice_cespite_numerico := arc.dit.fieldbyname('codice_cespite_numerico').asstring;
    codice_matricola_numerico := arc.dit.fieldbyname('codice_matricola_numerico').asstring;
    password_storno_evasione_vendite := arc.dit.fieldbyname('pswd_storno_evasione_vendite').asstring;
    password_storno_consolidamento_vendite := arc.dit.fieldbyname('pswd_storno_consolida_vendite').asstring;
    password_storno_differita_vendite := arc.dit.fieldbyname('pswd_storno_differita_vendite').asstring;
    password_storno_evasione_acquisti := arc.dit.fieldbyname('pswd_storno_evasione_acquisti').asstring;
    password_storno_consolidamento_acquisti := arc.dit.fieldbyname('pswd_storno_consolida_acquisti').asstring;
    password_storno_differita_acquisti := arc.dit.fieldbyname('pswd_storno_differita_acquisti').asstring;
    blocco_obsoleti := arc.dit.fieldbyname('blocco_obsoleti').asstring;
    lingua_nominativi := arc.dit.fieldbyname('lingua_nominativi').asstring;

    ricerca_articolo_codice_fornitore := arc.dit.fieldbyname('ricerca_articolo_codice_fornito').asstring;
    ricerca_articolo_numero_serie := arc.dit.fieldbyname('ricerca_articolo_numero_serie').asstring;
    descrizioni_articolo_unite := arc.dit.fieldbyname('descrizioni_articolo_unite').asstring;
    help_personalizzato := arc.dit.fieldbyname('help_personalizzato').asstring;
    inventario_fiscale := arc.dit.fieldbyname('tipo_inventario').asstring;
    inventario_gestionale := arc.dit.fieldbyname('valorizzazione_gestionale').asstring;
    gestione_revisioni := arc.dit.fieldbyname('gestione_revisioni').asstring;
    cartella_file := cartella_base_file + 'documenti_' + ditta;

    listini_01 := arc.dit.fieldbyname('listini_01').asstring;
    listini_02 := arc.dit.fieldbyname('listini_02').asstring;
    listini_03 := arc.dit.fieldbyname('listini_03').asstring;
    listini_04 := arc.dit.fieldbyname('listini_04').asstring;
    listini_05 := arc.dit.fieldbyname('listini_05').asstring;
    listini_06 := arc.dit.fieldbyname('listini_06').asstring;
    listini_07 := arc.dit.fieldbyname('listini_07').asstring;
    listini_08 := arc.dit.fieldbyname('listini_08').asstring;
    listini_09 := arc.dit.fieldbyname('listini_09').asstring;
    listini_10 := arc.dit.fieldbyname('listini_10').asstring;
    listini_11 := arc.dit.fieldbyname('listini_11').asstring;

    promozioni_01 := arc.dit.fieldbyname('promozioni_01').asstring;
    promozioni_02 := arc.dit.fieldbyname('promozioni_02').asstring;
    promozioni_03 := arc.dit.fieldbyname('promozioni_03').asstring;
    promozioni_04 := arc.dit.fieldbyname('promozioni_04').asstring;
    promozioni_05 := arc.dit.fieldbyname('promozioni_05').asstring;
    promozioni_06 := arc.dit.fieldbyname('promozioni_06').asstring;
    promozioni_07 := arc.dit.fieldbyname('promozioni_07').asstring;
    promozioni_08 := arc.dit.fieldbyname('promozioni_08').asstring;
    promozioni_09 := arc.dit.fieldbyname('promozioni_09').asstring;
    promozioni_10 := arc.dit.fieldbyname('promozioni_10').asstring;
    promozioni_11 := arc.dit.fieldbyname('promozioni_11').asstring;

    listini_01_fls := arc.dit.fieldbyname('listini_frn_01').asstring;
    listini_02_fls := arc.dit.fieldbyname('listini_frn_02').asstring;
    listini_03_fls := arc.dit.fieldbyname('listini_frn_03').asstring;
    listini_04_fls := arc.dit.fieldbyname('listini_frn_04').asstring;
    listini_05_fls := arc.dit.fieldbyname('listini_frn_05').asstring;
    listini_06_fls := arc.dit.fieldbyname('listini_frn_06').asstring;
    listini_07_fls := arc.dit.fieldbyname('listini_frn_07').asstring;
    listini_08_fls := arc.dit.fieldbyname('listini_frn_08').asstring;
    listini_09_fls := arc.dit.fieldbyname('listini_frn_09').asstring;
    listini_10_fls := arc.dit.fieldbyname('listini_frn_10').asstring;

    decimali_max_quantita := arc.dit.fieldbyname('decimali_max_quantita').asinteger;
    formato_display_quantita := formato_display(decimali_max_quantita);
    formato_display_quantita_zero := formato_display_zero(decimali_max_quantita);

    decimali_max_prezzo := arc.dit.fieldbyname('decimali_max_prezzo').asinteger;
    formato_display_prezzo := formato_display(decimali_max_prezzo);
    formato_display_prezzo_zero := formato_display_zero(decimali_max_prezzo);

    decimali_max_prezzo_acq := arc.dit.fieldbyname('decimali_max_prezzo_acq').asinteger;
    formato_display_prezzo_acq := formato_display(decimali_max_prezzo_acq);
    formato_display_prezzo_acq_zero := formato_display_zero(decimali_max_prezzo_acq);

    formato_display_importo := ',0.00;-,0.00;0.00';
    formato_display_importo_zero := ',0.00;-,0.00;#';

    // aggiorna parametri variabili esterne esercizio
    descrizione_esercizio := ese.fieldbyname('descrizione').asstring;
    data_inizio := ese.fieldbyname('data_inizio').asdatetime;
    data_fine := ese.fieldbyname('data_fine').asdatetime;
    data_bilancio := ese.fieldbyname('data_bilancio').asdatetime;
    esercizio_chiuso := ese.fieldbyname('esercizio_chiuso').asstring;
    esercizio_chiuso_magazzino := ese.fieldbyname('esercizio_chiuso_magazzino').asstring;

    esercizio_precedente := ese.fieldbyname('esercizio_precedente').asstring;
    data_inizio_precedente := ese.fieldbyname('data_inizio_precedente').asdatetime;
    data_fine_precedente := ese.fieldbyname('data_fine_precedente').asdatetime;
    data_bilancio_precedente := ese.fieldbyname('data_bilancio_precedente').asdatetime;
    esercizio_chiuso_precedente := ese.fieldbyname('esercizio_chiuso_precedente').asstring;
    esercizio_chiuso_magazzino_precedente := ese.fieldbyname('esercizio_chiuso_magazzino_precedente').asstring;

    esercizio_successivo := ese.fieldbyname('esercizio_successivo').asstring;

    // connessione database ditta
    cartella_archivi_ditta := '';
    if cartella_archivi_ditta = '' then
    begin
      if storico then
      begin
        cartella_archivi_ditta := 'arc_' + lowercase(ditta) + '_storico';
      end
      else
      begin
        cartella_archivi_ditta := 'arc_' + lowercase(ditta);
      end;
    end;

    arc.connessione_database_ditta(arc.arcdit, cartella_archivi_ditta, utente, tipo_server);

    if trim(divisa_di_conto) <> '' then
    begin
      tva := tmyquery_go.create(nil);
      tva.connection := arc.arcdit;
      tva.sql.text := 'select decimali_prezzo, decimali_prezzo_acq, decimali_importo from tva where codice = :codice';
      tva.parambyname('codice').asstring := divisa_di_conto;
      try
        tva.open;
      except
        messaggio(000, 'manca in archivio valute la divisa di conto: ' + divisa_di_conto);
      end;

      if not tva.isempty then
      begin
        decimali_prezzo_euro := tva.fieldbyname('decimali_prezzo').asinteger;
        decimali_prezzo_acq_euro := tva.fieldbyname('decimali_prezzo_acq').asinteger;
        decimali_importo_euro := tva.fieldbyname('decimali_importo').asinteger;
      end;
      tva.free;
    end;

    if arc.arcdit.connected then
    begin
      tna := tmyquery_go.create(nil);
      tna.Close;
      tna.connection := arc.arcdit;
      tna.sql.text := 'select descrizione, codice_iso from tna where codice = :codice';
      tna.parambyname('codice').asstring := tna_codice_ditta;
      tna.open;
      nazione_ditta := tna.fieldbyname('descrizione').asstring;
      codice_iso_ditta := tna.fieldbyname('codice_iso').asstring;
      tna.free;
    end;

    freeandnil(log_inserimento);
    freeandnil(log_modifica);
    freeandnil(log_cancellazione);

    freeandnil(ese);
  finally
    screen.Cursor := cursore;
  end;
end;

procedure TARC.connessione_database_ditta(database: TMyConnection_go; nome_database, nome_utente, hostname: string);
begin
  if (database.connected) and (database.intransaction) then
  begin
    database.rollback;
  end;
  database.connected := false;

  database.database := lowercase(nome_database);

  if hostname <> '' then
  begin
    database.server := hostname;
  end;

  if porta_server <> '' then
  begin
    database.port := strtoint(porta_server);
  end;

  // assicura grant a tutti gli utenti
  database.username := utente_database;
  database.password := password_database;

  database.username := nome_utente;
  database.password := password_database;

  try
    database.connected := true;
  except
    on e: exception do
    begin
      messaggio(000, 'non è stata eseguita la connessione al database della ditta: ' + slinebreak + nome_database + slinebreak + 'annotare il messaggio seguente e comunicarlo all''assistenza tecnica' + slinebreak + e.message)
    end;
  end;

  if dit.fieldbyname('tva_codice').asstring = '' then
  begin
    messaggio(200, 'non è stato assegnato il codice valuta alla ditta [' + ditta + ']');
  end;
  if dit.fieldbyname('tna_codice').asstring = '' then
  begin
    messaggio(200, 'non è stato assegnato il codice nazione alla ditta [' + ditta + ']');
  end;

  // assegna valori da ditta
  cifre_decimali_importo := 2;

  // controllo esistenza trigger
  if (extractfilename(lowercase(application.exename)) <> 'go_conversioni.exe') and (extractfilename(lowercase(application.exename)) <> 'gestionale_bovini.exe') then
  begin
    if not isdebuggerpresent then
    begin
      esiste_trigger_mysql.Close;
      esiste_trigger_mysql.parambyname('trigger_schema').asstring := lowercase('arc_' + ditta);
      esiste_trigger_mysql.open;
      if esiste_trigger_mysql.isempty then
      begin
        if messaggio(304, 'non sono presenti i trigger del database' + slinebreak + 'la loro mancanza può causare errori nell''esecuzione dei programmi' + slinebreak + slinebreak + 'per crearli eseguire il programma PRELEASE,' + slinebreak + 'senza indicare il codice d''accesso' + slinebreak +
          slinebreak + 'conferma per proseguire comunque') <> 1 then
        begin
          esiste_trigger_mysql.Close;
          application.terminate;
        end
        else
        begin
          esiste_trigger_mysql.Close;
        end;
      end
      else
      begin
        esiste_trigger_mysql.Close;
      end;
    end;
  end;
end;

procedure TARC.assegna_variabili_sistema(connessione: tmyconnection);
begin
  // solo versione MySQL 5.7
  connessione.execsql('set session sql_mode="STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION"');
  // MYSQL 8
  // connessione.execsql('set session sql_mode="STRICT_TRANS_TABLES,NO_ENGINE_SUBSTITUTION"');

  connessione.execsql('set @utn_codice = ' + quotedstr(utente));
  connessione.execsql('set @dit_codice = ' + quotedstr(ditta));
  connessione.execsql('set @ese_codice = ' + quotedstr(esercizio));
  connessione.execsql('set @data_inizio = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_inizio)));
  connessione.execsql('set @data_fine = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_fine)));
  connessione.execsql('set @data_bilancio = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_bilancio)));
  connessione.execsql('set @ese_codice_precedente = ' + quotedstr(esercizio_precedente));
  connessione.execsql('set @data_inizio_precedente = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_inizio_precedente)));
  connessione.execsql('set @data_fine_precedente = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_fine_precedente)));
  connessione.execsql('set @data_bilancio_precedente = ' + quotedstr(formatdatetime('yyyy-mm-dd', data_bilancio_precedente)));
  connessione.execsql('set @archivi_multiaziendali = ' + quotedstr(archivi_multiaziendali));

  if connessione.name = 'arcdit' then
  begin
    connessione.execsql('set @decimali_quantita = ' + dit.fieldbyname('decimali_max_quantita').asstring);
    connessione.execsql('set @decimali_prezzo_ven = ' + dit.fieldbyname('decimali_max_prezzo').asstring);
    connessione.execsql('set @decimali_prezzo_acq = ' + dit.fieldbyname('decimali_max_prezzo_acq').asstring);
  end;
end;

procedure TARC.arcAfterConnect(Sender: TObject);
begin
  assegna_variabili_sistema(arc);
end;

procedure TARC.arcsorAfterConnect(Sender: TObject);
begin
  assegna_variabili_sistema(arcsor);
end;

procedure TARC.arcditAfterConnect(Sender: TObject);
begin
  assegna_variabili_sistema(arc);
  assegna_variabili_sistema(arcsor);
  assegna_variabili_sistema(arcdit);
end;

function messaggio(codice_messaggio: Integer; descrizione_messaggio: string; touch: Boolean = false): Integer;
var
  risultato: Integer;

  pr: tmessaggio;

  tipo_cursore: tcursor;
begin
  // 0-99 messaggi di errore con solo tasto OK senza result di ritorno
  // 100-199 messaggi info
  // 200-299 messaggi avviso con solo tasto OK senza result di ritorno
  // 300-399 messaggi conferma

  tipo_cursore := screen.Cursor;
  screen.Cursor := crdefault;

  result := 0;

  // controllo versione Windows e utilizzo skin per nuovi messaggi
  if (versione_os = 'Windows 8 o superiore') then
  begin
    case codice_messaggio of
      000:
        // messaggio senza codice (solo descrizione passata dal chiamante)
        begin
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      001:
        // codice inesistente (per ricerca in archivi con programma lookup)
        begin
          descrizione_messaggio := codice_inesistente + #13 + descrizione_messaggio;
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      002:
        // valore non consentito
        begin
          descrizione_messaggio := valore_non_consentito1 + msg_0006 + #13 + '[' + descrizione_messaggio + ']' + #13 + valore_non_consentito2;
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      003:
        // manca permesso di modifica sugli importi
        begin
          descrizione_messaggio := msg_0007 + #13 + descrizione_messaggio;
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      004:
        // valore non corretto
        begin
          descrizione_messaggio := valore_non_consentito1 + msg_0006 + #13 + '[' + descrizione_messaggio + ']' + #13 + 'non è corretto';
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      100:
        // messaggio senza codice (solo descrizione passata dal chiamante)
        begin
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      200:
        begin
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      201:
        // inizio file
        begin
          descrizione_messaggio := inizio_archivio;
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      202:
        // fine file
        begin
          descrizione_messaggio := fine_archivio;
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio);
          end;
        end;
      300:
        // richiesta conferma senza codice (solo descrizione passata dal chiamante)
        begin
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            if arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio) = mryes then
            begin
              result := 1;
            end;
          end;
        end;
      301:
        // conferma cancellazione
        begin
          if trim(descrizione_messaggio) = '' then
          begin
            descrizione_messaggio := conferma_cancellazione + slinebreak + descrizione_messaggio;
          end
          else
          begin
            descrizione_messaggio := conferma_cancellazione + slinebreak + slinebreak + descrizione_messaggio;
          end;

          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            if arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio) = mryes then
            begin
              result := 1;
            end;
          end;
        end;
      303:
        // richiesta memorizzazione dopo modifica
        begin
          descrizione_messaggio := conferma_memorizzazione;

          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            if arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio) = mryes then
            begin
              result := 1;
            end;
          end;
        end;
      304:
        // info con messaggio passato dal programma chiamante e proposta mbno
        begin
          if tastiera_touch or touch then
          begin
            pr := tmessaggio.create(nil);
            pr.v_descrizione.text := descrizione_messaggio;
            pr.tipo := codice_messaggio;
            pr.showmodal;
            result := pr.risultato;
            pr.free;
          end
          else
          begin
            if arc.messaggio_nuovo(codice_messaggio, descrizione_messaggio) = mryes then
            begin
              result := 1;
            end;
          end;
        end;
      400:
        // info con messaggio passato dal programma chiamante e proposta mbno
        begin
          risultato := messagedlg(descrizione_messaggio, mtconfirmation, [mbyes, mbno, mbyestoall], 0, mbyes);
          if risultato = mryes then
          begin
            result := 1;
          end
          else if risultato = mryestoall then
          begin
            result := 2;
          end;
        end;
    end;
  end
  else
  begin
    case codice_messaggio of
      000:
        // messaggio senza codice (solo descrizione passata dal chiamante)
        begin
          risultato := messagedlg(descrizione_messaggio, mtConfirmation, [mbok], 0, mbok);
        end;
      001:
        // codice inesistente (per ricerca in archivi con programma lookup)
        begin
          descrizione_messaggio := codice_inesistente + #13 + descrizione_messaggio;
          messagedlg(descrizione_messaggio, mtWarning, [mbOk], 0);
        end;
      002:
        // valore non consentito
        begin
          descrizione_messaggio := valore_non_consentito1 + msg_0006 + #13 + descrizione_messaggio + #13 + valore_non_consentito2;
          messagedlg(descrizione_messaggio, mtWarning, [mbOk], 0);
        end;
      003:
        // manca permesso di modifica sugli importi
        begin
          descrizione_messaggio := msg_0007 + #13 + descrizione_messaggio;
          messagedlg(descrizione_messaggio, TMsgDlgType.mtError, [mbOk], 0);
        end;
      004:
        // valore non corretto
        begin
          descrizione_messaggio := msg_0007 + #13 + descrizione_messaggio;
          messagedlg(descrizione_messaggio, TMsgDlgType.mterror, [mbOk], 0);
        end;
      100:
        // messaggio senza codice (solo descrizione passata dal chiamante)
        begin
          messagedlg(descrizione_messaggio, TMsgDlgType.mtinformation, [mbOk], 0);
        end;
      200:
        begin
          messagedlg(descrizione_messaggio, TMsgDlgType.mtwarning, [mbOk], 0);
        end;
      201:
        // inizio file
        begin
          descrizione_messaggio := inizio_archivio;
          messagedlg(descrizione_messaggio, TMsgDlgType.mtwarning, [mbOk], 0);
        end;
      202:
        // fine file
        begin
          descrizione_messaggio := fine_archivio;
          messagedlg(descrizione_messaggio, mtwarning, [mbOk], 0);
        end;
      300:
        // richiesta conferma senza codice (solo descrizione passata dal chiamante)
        begin
          if messagedlg(descrizione_messaggio, mtconfirmation, [mbyes, mbno], 0) = mryes then
          begin
            result := 1;
          end;
        end;
      301:
        // conferma cancellazione
        begin
          if trim(descrizione_messaggio) = '' then
          begin
            descrizione_messaggio := conferma_cancellazione + #13 + descrizione_messaggio;
          end
          else
          begin
            descrizione_messaggio := conferma_cancellazione + #13 + #13 + descrizione_messaggio;
          end;

          if messagedlg(descrizione_messaggio, mtconfirmation, [mbyes, mbno], 0, mbno) = mryes then
          begin
            result := 1;
          end;
        end;
      303:
        // richiesta memorizzazione dopo modifica
        begin
          descrizione_messaggio := conferma_memorizzazione;
          if messagedlg(descrizione_messaggio, mtconfirmation, [mbyes, mbno], 0) = mryes then
          begin
            result := 1;
          end;
        end;
      304:
        // info con messaggio passato dal programma chiamante e proposta mbno
        begin
          if messagedlg(descrizione_messaggio, mtconfirmation, [mbyes, mbno], 0, mbno) = mryes then
          begin
            result := 1;
          end;
        end;
    end;
  end;

  screen.Cursor := tipo_cursore;
end;

function TARC.messaggio_nuovo(codice_messaggio: Integer; descrizione_messaggio: string; lista_opzioni: string = ''; standard: Integer = 1): Integer;
var
  stringa: string;
  taskdialog: ttaskdialog;
begin
  result := 0;
  taskdialog := ttaskdialog.create(nil);

  taskdialog.flags := [tfusehiconmain, tfAllowDialogCancellation];
  taskdialog.custommainicon := application.icon;

  if lista_opzioni <> '' then
  begin
    taskdialog.flags := [tfusehiconmain];

    taskdialog.caption := 'Selezione tipo operazione';
    taskdialog.text := descrizione_messaggio;
    taskdialog.commonbuttons := [tcbok];

    stringa := lista_opzioni;
    while pos(';', stringa) > 0 do
    begin
      taskdialog.radiobuttons.add.caption := copy(stringa, 1, pos(';', stringa) - 1);
      stringa := trim(copy(stringa, pos(';', stringa) + 1, length(stringa)));
    end;
    if stringa <> '' then
    begin
      taskdialog.radiobuttons.add.caption := stringa;
    end;

    taskdialog.radiobuttons.items[standard - 1].default := true;

    taskdialog.footericon := 0;
    taskdialog.footertext := '';
    taskdialog.expandbuttoncaption := '';
    taskdialog.expandedtext := '';

    taskdialog.execute;

    if taskdialog.modalresult = mrok then
    begin
      result := taskdialog.radiobutton.index + 1;
    end
    else
    begin
      result := 0;
    end;
  end
  else
  begin
    if (codice_messaggio = 0) or (codice_messaggio = 1) or (codice_messaggio = 2) or (codice_messaggio = 4) then
    begin
      taskdialog.caption := 'Errore';
      taskdialog.commonbuttons := [tcbok];
      taskdialog.footericon := 2;
      taskdialog.footertext := 'Messaggio di errore';
      taskdialog.text := descrizione_messaggio;
      taskdialog.expandbuttoncaption := 'Errore bloccante';
      taskdialog.expandedtext := 'Prima di proseguire nell''esecuzione del programma' + slinebreak + 'indagare a fondo sulla causa dell''errore' + slinebreak + 'che potrebbe portare a pesanti conseguenze' + slinebreak + 'sulla coerenza dei dati inseriti nella procedura';

      taskdialog.execute;
    end
    else if (codice_messaggio = 100) or (codice_messaggio = 201) or (codice_messaggio = 202) then
    begin
      taskdialog.caption := 'Informazione';
      taskdialog.commonbuttons := [tcbok];
      taskdialog.footericon := 3;
      taskdialog.footertext := 'Messaggio informativo';
      taskdialog.text := descrizione_messaggio;

      taskdialog.execute;
    end
    else if (codice_messaggio = 3) or (codice_messaggio = 200) then
    begin
      taskdialog.caption := 'Attenzione';
      taskdialog.commonbuttons := [tcbok];
      taskdialog.footericon := 1;
      taskdialog.footertext := 'Messaggio di allerta';
      taskdialog.text := descrizione_messaggio;
      taskdialog.expandbuttoncaption := 'Messaggio esteso';
      taskdialog.expandedtext := 'Prestare particolare attenzione a questi tipi di messaggio' + slinebreak + 'perché la loro sottovalutazione potrebbe portare a gravi' + slinebreak + 'conseguenze in fasi successive della procedura';

      taskdialog.execute;
    end
    else if (codice_messaggio = 300) or (codice_messaggio = 301) or (codice_messaggio = 303) or (codice_messaggio = 304) then
    begin
      taskdialog.caption := 'Conferma';
      taskdialog.commonbuttons := [tcbYes, tcbNo];
      if (codice_messaggio = 300) or (codice_messaggio = 303) then
      begin
        taskdialog.defaultbutton := tcbYes;
      end
      else
      begin
        taskdialog.defaultbutton := tcbNo;
      end;
      taskdialog.text := descrizione_messaggio;

      taskdialog.footericon := 0;
      taskdialog.footertext := '';
      taskdialog.expandbuttoncaption := '';
      taskdialog.expandedtext := '';

      taskdialog.execute;
      result := taskdialog.modalresult;
    end;
  end;

  freeandnil(taskdialog);
end;

procedure esegui(nome_file: string);
begin
  esegui_effettivo(nome_file);
end;

procedure esegui_effettivo(nome_file: string; parametro: string = ''; cartella: string = '');
var
  cartella_attuale: string;
  ret: word;
begin
  cartella_attuale := cartella_installazione;
  ret := shellexecute(application.handle, pchar('open'), pchar(nome_file), pchar(parametro), pchar(cartella), SW_SHOWNORMAL);
  if ret <= 32 then
  begin
    messaggio(000, syserrormessage(getlasterror) + slinebreak + 'file: ' + nome_file);
  end;
  setcurrentdir(cartella_attuale);
end;

procedure esegui_stampa_diretta(nome_file: string);
var
  cartella: string;
  ret: word;
begin
  cartella := cartella_installazione;
  ret := shellexecute(application.handle, pchar('print'), pchar(nome_file), pchar(''), nil, SW_HIDE);
  if ret <= 32 then
  begin
    messaggio(000, syserrormessage(getlasterror));
  end;
  setcurrentdir(cartella);
end;

procedure esegui_collegato(nome_file: string);
var
  ret: word;
begin
  (*
    ret := shellexecute(application.handle, pchar('open'), pchar(nome_file),
    pchar('<utente>' + utente + '</utente> <salta_login>si</salta_login> <multi>si</multi> <password>' +
    password_utente_login + '</password>' + ' ' + parametro_globale),
    nil, SW_SHOWNORMAL);
 *)

  ret := shellexecute(application.handle, pchar('open'), pchar(nome_file), pchar('<utente>' + utente + '</utente> <salta_login>si</salta_login> <multi>si</multi> ' + parametro_globale), nil, SW_SHOWNORMAL);

  if ret <= 32 then
  begin
    messaggio(000, syserrormessage(getlasterror));
  end;
  setcurrentdir(cartella_installazione);
end;

function esegui_programma(programma_da_eseguire: string; codice_archivio: variant; modale: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean;
begin
  result := esegui_programma(programma_da_eseguire, codice_archivio, modale, false, parametri_prg_codice_diretto, cartella_lavoro, programma_chiamante);
end;

function esegui_programma(programma_da_eseguire: string; codice_archivio: variant; modale, record_singolo: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean;
var
  i: word;
begin
  result := true;
  if trim(programma_da_eseguire) = '' then
  begin
    abort;
  end
  else
  begin
    if call_programma(trim(programma_da_eseguire), codice_archivio, modale, record_singolo, parametri_prg_codice_diretto, cartella_lavoro, programma_chiamante) then
    begin
      if arc.lista_programmi_recenti.indexof(trim(programma_da_eseguire)) = -1 then
      begin
        if arc.lista_programmi_recenti.count = 21 then
        begin
          for i := 1 to 20 do
          begin
            arc.lista_programmi_recenti[i - 1] := arc.lista_programmi_recenti[i];
          end;
          arc.lista_programmi_recenti.delete(20);
        end;
        arc.lista_programmi_recenti.add(trim(programma_da_eseguire));
      end;
      result := true;
    end
    else
    begin
      result := false;
    end;
  end;
end;

function numerico(stringa: string): Boolean;
var
  i: word;
begin
  result := true;
  if length(stringa) = 0 then
  begin
    result := false;
  end
  else
  begin
    for i := 1 to length(stringa) do
    begin
      if ((not isnumeric(stringa[i])) or (stringa[i] = ' ')) and (stringa[i] <> '+') and (stringa[i] <> '-') and (stringa[i] <> ',') and (stringa[i] <> '.') then
      begin
        result := false;
        break;
      end;
    end;
  end;
end;

function setta_lunghezza(stringa: string; caratteri: word): string;
var
  i, j: word;
begin
  result := stringa;

  if length(stringa) >= caratteri then
  begin
    result := copy(stringa, 1, caratteri);
  end
  else
  begin
    j := (length(result) + 1);
    for i := j to caratteri do
    begin
      result := result + ' ';
    end;
  end;
end;

function setta_lunghezza(stringa: string; caratteri: word; destra: Boolean; carattere: string): string;
var
  i, j, n: word;
begin
  result := stringa;
  if carattere = '' then
  begin
    carattere := ' ';
  end;

  if length(stringa) >= caratteri then
  begin
    result := copy(stringa, 1, caratteri);
  end
  else
  begin
    result := '';
    n := 0;
    j := length(stringa);
    for i := 1 to caratteri do
    begin
      if i <= (caratteri - j) then
      begin
        result := result + carattere;
      end
      else
      begin
        n := n + 1;
        result := result + stringa[n];
      end;
    end;
  end;
end;

function setta_lunghezza(numero: double; caratteri, decimali: word): string;
var
  i, j: word;
  numero_stringa: string;
begin
  result := '';

  if decimali = 0 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1, 0));
  end
  else if decimali = 1 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10, 0));
  end
  else if decimali = 2 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 100, 0));
  end
  else if decimali = 3 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1000, 0));
  end
  else if decimali = 4 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10000, 0));
  end;

  j := (length(numero_stringa));
  j := caratteri - j;

  for i := 1 to j do
  begin
    result := result + '0';
  end;
  result := result + numero_stringa;
end;

function setta_lunghezza(numero: double; caratteri, decimali: word; carattere: string): string;
var
  i, j: word;
  numero_stringa: string;
begin
  result := '';

  if decimali = 0 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1, 0));
  end
  else if decimali = 1 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10, 0));
  end
  else if decimali = 2 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 100, 0));
  end
  else if decimali = 3 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1000, 0));
  end
  else if decimali = 4 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10000, 0));
  end;

  j := (length(numero_stringa));
  j := caratteri - j;

  for i := 1 to j do
  begin
    result := result + carattere;
  end;
  result := result + numero_stringa;
end;

function setta_lunghezza(numero: Integer; caratteri, decimali: word): string;
var
  i, j: word;
  numero_stringa: string;
begin
  result := '';

  if decimali = 0 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1, 0));
  end
  else if decimali = 1 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10, 0));
  end
  else if decimali = 2 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 100, 0));
  end
  else if decimali = 3 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 1000, 0));
  end
  else if decimali = 4 then
  begin
    numero_stringa := floattostr(arrotonda(numero * 10000, 0));
  end;

  j := (length(numero_stringa));
  j := caratteri - j;

  for i := 1 to j do
  begin
    result := result + '0';
  end;
  result := result + numero_stringa;
end;

function arrotonda(numero: double; nrdec: word; tipoarrotondamento: word): double;
begin
  result := numero;

  case tipoarrotondamento of
    0:
      // tronca
      result := decimalrounddbl(numero, nrdec, drrnddown);
    1:
      // eccesso/difetto
      if numero > 0 then
      begin
        result := decimalrounddbl(numero, nrdec, drhalfpos);
      end
      else if numero < 0 then
      begin
        result := decimalrounddbl(numero, nrdec, drhalfneg);
      end;
    2:
      // sempre al valore superiore
      result := decimalrounddbl(numero, nrdec, drrndup);
  end;
end;

function arrotonda(numero: double; nrdec: word): double;
begin
  result := arrotonda(numero, nrdec, 1);
end;

function arrotonda(numero: double): double;
begin
  result := arrotonda(numero, decimali_importo_euro, 1);
end;

function replicate(Ch: Char; Len: Integer): string;
begin
  result := stringofchar(Ch, Len);
end;

function controllo_ora(ora: string): string;
var
  i: word;
  intero, decimale: string;
begin
  result := '00.00';

  if trim(ora) = '' then
  begin
    result := '00.00';
  end
  else
  begin
    for i := 1 to length(ora) do
    begin
      if (not numerico(copy(ora, i, 1))) and (copy(ora, i, 1) <> '.') then
      begin
        result := '00.00';
      end;
    end;
    if pos('.', ora) <> 0 then
    begin
      intero := setta_lunghezza(strtoint(copy(ora, 1, pos('.', ora) - 1)), 2, 0);
      decimale := setta_lunghezza(strtoint(copy(ora, pos('.', ora) + 1, length(ora) - pos('.', ora))), 2, 0);
    end
    else
    begin
      if length(ora) <= 2 then
      begin
        intero := setta_lunghezza(strtoint(copy(ora, 1, length(ora))), 2, 0);
        decimale := '00';
      end
      else
      begin
        intero := copy(ora, 1, 2);
        if length(ora) = 4 then
        begin
          decimale := copy(ora, 3, 2);
        end
        else
        begin
          decimale := copy(ora, 3, 1) + '0';
        end;
      end;
    end;
    if intero > '23' then
    begin
      intero := '23';
    end;
    if decimale > '59' then
    begin
      decimale := '59';
    end;

    result := intero + '.' + decimale;
  end;
end;

procedure azzera_tabella(nome_tabella: string; var tabella: tmytable);
var
  query: tmyquery_go;
begin
  query := tmyquery_go.create(nil);
  query.connection := arc.arcsor;
  query.Close;
  query.sql.text := 'delete from ' + nome_tabella + ' where utn_codice = ' + quotedstr(utente);
  query.execsql;

  query.free;

  if tabella.active then
  begin
    tabella.Close;
  end;
  tabella.connection := arc.arcsor;
  tabella.tablename := nome_tabella;
  tabella.filtered := true;
  tabella.filter := 'utn_codice = ' + quotedstr(utente);
  tabella.open;
end;

procedure azzera_tabella(nome_tabella: string);
var
  query: tmyquery_go;
begin
  query := tmyquery_go.create(nil);
  query.connection := arc.arcsor;
  query.Close;
  query.sql.text := 'delete from ' + nome_tabella + ' where utn_codice = ' + quotedstr(utente);
  query.execsql;

  query.free;
end;

procedure assegna_file_cfg;
var
  i: word;
  icona: string;

  file_ini: tinifile;
  stringa: string;
begin
  mesi_password := 6;
  linguaggio_interfaccia := 'italiano';
  data_31_12_9999 := '31' + formatsettings.dateseparator + '12' + formatsettings.dateseparator + '2999';

  for i := 0 to 31 do
  begin
    parametri_extra_programma_chiamato[i] := null;
  end;

  tasti_errati := false;

  cursore := screen.Cursor;

  cartella_installazione := extractfilepath(application.exename);

  for i := (length(cartella_installazione) - 1) downto 1 do
  begin
    if copy(cartella_installazione, i, 1) = '\' then
    begin
      // cartella di installazione programma (per accedere alle altre cartelle di GO)
      cartella_root_installazione := copy(cartella_installazione, 1, i);
      break;
    end;
  end;
  decimali_importo_euro := 2;
  decimali_prezzo_euro := 6;
  decimali_prezzo_acq_euro := 6;

  application.title := '';
  icona := '';
  descrizione_programmi_personalizzata := 'no';
  salta_ultimo_utente_login := 'no';
  ultimo_utente_login := 'no';

  menu_utilizzato := '';

  // utente_manutenzione := '';

  file_ini := tinifile.create(cartella_installazione + '\go.cfg');
  cartella_temp := cartella_root_installazione + 'temp\';

  codice_cliente_cloud := file_ini.readstring('gestione archivi', 'codice_cliente_cloud', '');

  if codice_cliente_cloud <> '' then
  begin
    arc.arc.suffisso := codice_cliente_cloud;
    arc.arcdit.suffisso := codice_cliente_cloud;
    arc.arcsor.suffisso := codice_cliente_cloud;
  end;
  url_base_help := file_ini.readstring('funzioni varie', 'url_base_help', '');

  cartella_bitmap := file_ini.readstring('gestione archivi', 'cartella_bitmap', '');
  if cartella_bitmap = '' then
  begin
    cartella_bitmap := cartella_root_installazione + 'bmp\';
  end
  else
  begin
    if cartella_bitmap[length(cartella_bitmap)] <> '\' then
    begin
      cartella_bitmap := cartella_bitmap + '\';
    end;
  end;

  cartella_report := file_ini.readstring('gestione archivi', 'cartella_report', '');
  if cartella_report = '' then
  begin
    cartella_report := cartella_root_installazione;
  end
  else
  begin
    if cartella_report[length(cartella_report)] <> '\' then
    begin
      cartella_report := cartella_report + '\';
    end;
  end;

  cartella_base_file := file_ini.readstring('gestione archivi', 'cartella_documenti', '');

  if cartella_base_file = '' then
  begin
    cartella_base_file := cartella_root_installazione;
  end
  else
  begin
    if cartella_base_file[length(cartella_base_file)] <> '\' then
    begin
      cartella_base_file := cartella_base_file + '\';
    end;
  end;

  cartella_email := file_ini.readstring('gestione archivi', 'cartella_email', '');
  if cartella_email = '' then
  begin
    cartella_email := cartella_root_installazione + 'email\';
  end
  else
  begin
    if cartella_email[length(cartella_email)] <> '\' then
    begin
      cartella_email := cartella_email + '\';
    end;
  end;

  cartella_esporta := file_ini.readstring('gestione archivi', 'cartella_esporta', '');
  if cartella_esporta = '' then
  begin
    cartella_esporta := cartella_root_installazione + 'esporta';
  end
  else
  begin
    if cartella_esporta[length(cartella_esporta)] = '\' then
    begin
      cartella_esporta := copy(cartella_esporta, 1, length(cartella_esporta) - 1);
    end;
  end;
  if pos('%USERPROFILE%', cartella_esporta) > 0 then
  begin
    cartella_esporta := stringreplace(cartella_esporta, '%USERPROFILE%', GetEnvironmentVariable('USERPROFILE'), []);
  end;
  if not directoryexists(cartella_esporta) then
  begin
    createdir(cartella_esporta);
  end;

  cartella_stampe := file_ini.readstring('gestione archivi', 'cartella_stampe', '');
  if cartella_stampe = '' then
  begin
    cartella_stampe := cartella_root_installazione + 'stampe';
  end
  else
  begin
    if cartella_stampe[length(cartella_stampe)] = '\' then
    begin
      cartella_stampe := copy(cartella_stampe, 1, length(cartella_stampe) - 1);
    end;
  end;
  if pos('%USERPROFILE%', cartella_stampe) > 0 then
  begin
    cartella_stampe := stringreplace(cartella_stampe, '%USERPROFILE%', GetEnvironmentVariable('USERPROFILE'), []);
  end;
  if not directoryexists(cartella_stampe) then
  begin
    createdir(cartella_stampe);
  end;
  if not directoryexists(cartella_root_installazione + 'stampe') then
  begin
    createdir(cartella_root_installazione + 'stampe');
  end;

  cartella_filtri_vis := file_ini.readstring('gestione archivi', 'cartella_filtri_vis', '');
  if cartella_filtri_vis = '' then
  begin
    cartella_filtri_vis := cartella_root_installazione + 'filtri_vis';
  end
  else
  begin
    if cartella_filtri_vis[length(cartella_filtri_vis)] = '\' then
    begin
      cartella_filtri_vis := copy(cartella_filtri_vis, 1, length(cartella_filtri_vis) - 1);
    end;
  end;

  cartella_stili := file_ini.readstring('gestione archivi', 'cartella_stili', '');
  if cartella_stili = '' then
  begin
    cartella_stili := cartella_root_installazione + 'stili\';
  end
  else
  begin
    if cartella_stili[length(cartella_stili)] <> '\' then
    begin
      cartella_stili := cartella_stili + '\';
    end;
  end;

  codice_procedura := file_ini.readstring('personalizzazione procedura', 'codice_procedura', 'go');
  if trim(codice_procedura) = '' then
  begin
    codice_procedura := 'go';
  end;

  nome_procedura := file_ini.readstring('personalizzazione procedura', 'nome_procedura', 'Gestionale Open');
  if trim(nome_procedura) = '' then
  begin
    nome_procedura := 'Gestionale Open';
  end;
  application.title := nome_procedura;

  icona := file_ini.readstring('personalizzazione procedura', 'icona', '');
  if icona <> '' then
  begin
    application.icon.loadfromfile(icona);
  end;

  bitmap_login := file_ini.readstring('personalizzazione procedura', 'bitmap_login', '');
  if trim(bitmap_login) = '' then
  begin
    bitmap_login := cartella_bitmap + codice_procedura + '.jpg';
  end;

  bitmap_menu := file_ini.readstring('personalizzazione procedura', 'bitmap_menu', '');
  if trim(bitmap_menu) = '' then
  begin
    bitmap_menu := cartella_bitmap + codice_procedura + '_menu.jpg';
  end;

  bitmap_tabulati := file_ini.readstring('personalizzazione procedura', 'bitmap_tabulati', '');
  if trim(bitmap_tabulati) = '' then
  begin
    bitmap_tabulati := cartella_bitmap + codice_procedura + '_tabulati.jpg';
  end;

  sito_web := file_ini.readstring('personalizzazione procedura', 'sito_web', 'http://www.gestionaleopen.org');
  if trim(sito_web) = '' then
  begin
    sito_web := 'http://www.gestionaleopen.org';
  end;

  forum := file_ini.readstring('personalizzazione procedura', 'forum', 'http://www.gestionaleopen.org/forum/');
  if trim(forum) = '' then
  begin
    forum := 'http://www.gestionaleopen.org/forum/';
  end;

  telefono_assistenza := file_ini.readstring('personalizzazione procedura', 'telefono_assistenza', '035.0521150');
  if trim(telefono_assistenza) = '' then
  begin
    telefono_assistenza := '035.0521150 / int. 2';
  end;

  mail_assistenza := file_ini.readstring('personalizzazione procedura', 'mail_assistenza', 'assistenza@gestionaleopen.org');
  if trim(mail_assistenza) = '' then
  begin
    mail_assistenza := 'assistenza@gestionaleopen.org';
  end;

  tipo_server := file_ini.readstring('gestione archivi', 'server', 'localhost');

  if tipo_server = 'portable' then
  begin
    tipo_portable := true;
    tipo_server := 'localhost';
  end
  else
  begin
    tipo_portable := false;
  end;

  porta_server := file_ini.readstring('gestione archivi', 'server_porta', '0');

  mesi_password := file_ini.readinteger('funzioni varie', 'mesi_password', 6);
  if mesi_password = 0 then
  begin
    mesi_password := 6;
  end;

  salta_ultimo_utente_login := file_ini.readstring('funzioni varie', 'salta_ultimo_utente_login', 'no');
  if salta_ultimo_utente_login = '' then
  begin
    salta_ultimo_utente_login := 'no';
  end;

  if salta_ultimo_utente_login = 'si' then
  begin
    ultimo_utente_login := 'no';
  end
  else
  begin
    ultimo_utente_login := file_ini.readstring('funzioni varie', 'ultimo_utente_login', 'no');
    if ultimo_utente_login = '' then
    begin
      ultimo_utente_login := 'no';
    end;
  end;

  // utente_manutenzione := file_ini.readstring('funzioni varie', 'utente_manutenzione', '');

  cursore_database := file_ini.readstring('funzioni varie', 'cursore_database', 'no');
  if cursore_database = '' then
  begin
    cursore_database := 'no';
  end;

  archivi_multiaziendali := file_ini.readstring('funzioni varie', 'tabelle_multiaziendali', 'no');
  if archivi_multiaziendali = '' then
  begin
    archivi_multiaziendali := 'no';
  end;

  programma_teleassistenza := file_ini.readstring('funzioni varie', 'programma_teleassistenza', '');
  if programma_teleassistenza = '' then
  begin
    programma_teleassistenza := 'TeamViewer_GestionaleOpen.exe';
  end;

  // smtp_controllo_accessi := file_ini.readstring('funzioni varie', 'smtp_controllo_accessi', '');

  access_violation := file_ini.readstring('funzioni varie', 'access_violation', 'si');
  if access_violation = '' then
  begin
    access_violation := 'si';
  end;

  bugreport_nominale := file_ini.readstring('funzioni varie', 'bugreport_nominale', 'no');
  if bugreport_nominale = '' then
  begin
    bugreport_nominale := 'no';
  end;

  foxgo := file_ini.readstring('funzioni varie', 'foxgo', 'no');
  if foxgo = '' then
  begin
    foxgo := 'no';
  end;

  file_ini.free;
end;

function codice_tum(art_codice: string): string;
var
  query: tmyquery_go;
begin
  query := tmyquery_go.create(nil);
  query.connection := arc.arcdit;
  query.sql.text := 'select art.tum_codice from art where codice = :codice';
  query.params[0].asstring := art_codice;
  query.open;
  if query.isempty then
  begin
    result := '';
  end
  else
  begin
    result := query.fieldbyname('tum_codice').asstring;
  end;

  query.free;
end;

function decimali_quantita(tum_codice: string): word;
var
  query: tmyquery_go;
begin
  result := 4;

  query := tmyquery_go.create(nil);
  query.connection := arc.arcdit;
  query.sql.text := 'select decimali from tum where codice = :codice';
  query.parambyname('codice').asstring := tum_codice;
  query.open;
  if not query.eof then
  begin
    result := query.fieldbyname('decimali').asinteger;
  end;

  query.free;
end;

function formato_display(decimali: word): string;
begin
  result := '';

  if decimali = 0 then
  begin
    result := ',0;-,0;0';
  end
  else if decimali = 1 then
  begin
    result := ',0.0;-,0.0;0.0';
  end
  else if decimali = 2 then
  begin
    result := ',0.00;-,0.00;0.00';
  end
  else if decimali = 3 then
  begin
    result := ',0.000;-,0.000;0.000';
  end
  else if decimali = 4 then
  begin
    result := ',0.0000;-,0.0000;0.0000';
  end
  else if decimali = 5 then
  begin
    result := ',0.00000;-,0.00000;0.00000';
  end
  else if decimali = 6 then
  begin
    result := ',0.000000;-,0.000000;0.000000';
  end;
end;

function fuoco(componente: twincontrol): Boolean;
begin
  result := false;
  if componente.canfocus then
  begin
    result := true;
    componente.setfocus;
  end;
end;

function formato_display_zero(decimali: word): string;
begin
  result := '';

  if decimali = 0 then
  begin
    result := ',0;-,0;#';
  end
  else if decimali = 1 then
  begin
    result := ',0.0;-,0.0;#';
  end
  else if decimali = 2 then
  begin
    result := ',0.00;-,0.00;#';
  end
  else if decimali = 3 then
  begin
    result := ',0.000;-,0.000;#';
  end
  else if decimali = 4 then
  begin
    result := ',0.0000;-,0.0000;#';
  end
  else if decimali = 5 then
  begin
    result := ',0.00000;-,0.00000;#';
  end
  else if decimali = 6 then
  begin
    result := ',0.000000;-,0.000000;#';
  end;
end;

function StrTran(InString: string; SearchString: string; SubString: string; Incremental: Boolean): string;
var
  lStringa, lNewStringa: string;
begin
  lStringa := InString;
  InString := '';
  if pos(SearchString, lStringa) > 0 then
  begin
    while pos(SearchString, lStringa) > 0 do
    begin
      if Incremental then
      begin
        lNewStringa := copy(lStringa, 1, pos(SearchString, lStringa) - 1) + SubString;
        lStringa := copy(lStringa, pos(SearchString, lStringa) + length(SearchString), length(lStringa));
        InString := InString + lNewStringa;
      end
      else
      begin
        lNewStringa := lStringa;
        lStringa := copy(lNewStringa, 1, pos(SearchString, lNewStringa) - 1);
        lStringa := lStringa + SubString;
        lStringa := lStringa + copy(lNewStringa, pos(SearchString, lNewStringa) + length(SearchString), length(lNewStringa));
        InString := lStringa;
      end;
    end;

    if Incremental then
      InString := InString + lStringa
  end
  else
    InString := lStringa;

  result := InString;
end;

function eseguire_alias_personalizzato_esterno(aprogramma_standard: string; acodice_archivio: variant): string;
var
  ret: word;
  tag_parametri_personalizzato: string;
begin
  // se il programma è in questo elenco che va mantenuto ed esteso passo come parametro a go.exe il progressivo_gesven
  // altrimenti in tutti gli altri casi passo come parametro codice_gesarc
  if (aprogramma_standard = 'GESPRI') or (aprogramma_standard = 'GESFAIV') or (aprogramma_standard = 'GESFADV') or (aprogramma_standard = 'GESFAAV') or (aprogramma_standard = 'GESFATA') or (aprogramma_standard = 'GESFADA') or (aprogramma_standard = 'GESNOCA') or (aprogramma_standard = 'GESNOCV') or
    (aprogramma_standard = 'GESFAIVP') or (aprogramma_standard = 'GESNOCVP') or (aprogramma_standard = 'GESDDTA') then
  begin
    tag_parametri_personalizzato := 'progressivo_gesven';
  end
  else
  begin
    tag_parametri_personalizzato := 'codice_gesarc';
  end;

  ret := shellexecute(application.handle, pchar('open'), pchar(codice_procedura + '.exe'), pchar('<utente>' + utente + '</utente> <salta_login>si</salta_login> <multi>si</multi> ' +
    // '<password>' + password_database + '</password>' +
    '<programma>' + aprogramma_standard + '</programma> ' + '<' + tag_parametri_personalizzato + '>' + vartostr(acodice_archivio) + '</' + tag_parametri_personalizzato + '>'), nil, SW_SHOWNORMAL);

  // se non riesco ad aprire l'eseguibile con il personalizzato cerco di aprire lo standard.
  result := ifthen(ret > 32, '', aprogramma_standard);
end;

function call_programma(programma_da_eseguire: string; codice_archivio: variant; modale, record_singolo: Boolean; parametri_prg_codice_diretto: string = ''; cartella_lavoro: string = ''; programma_chiamante: string = ''): Boolean;
var
  non_eseguire, programma_personalizzato: Boolean;
  j: Integer;

  // variabile per esecuzione programmi
  pr: tbase;

  baseclass: tbaseclass;

  // assegna cartella lavoro e parametri per programmi esterni
  programma_esterno, cartella_lavoro_esterno, parametro_esterno: string;
begin

  // PER COMPILAZIONE 64BIT
  pr := nil;

  if uppercase(programma_da_eseguire) <> 'EVENTIUTN' then
  begin
    arc.controllo_msg;
  end;

  non_eseguire := false;

  arc.query_pra.Close;
  arc.query_pra.params[0].asstring := uppercase(programma_da_eseguire);
  arc.query_pra.open;
  if not arc.query_pra.eof then
  begin
    programma_da_eseguire := uppercase(arc.query_pra.fieldbyname('prg_codice').asstring);
    alias_programma := arc.query_pra.fieldbyname('codice').asstring;

    // se sto eseguendo un modulo aggiuntivo diverso da GO ed è stato richiamato un programma con un alias
    // devo lanciare l'eseguibile di GO aprendo direttamente il programma customizzato
    if not(extractfilename(application.exename.tolower) = codice_procedura.tolower + '.exe') then
    begin
      programma_da_eseguire := eseguire_alias_personalizzato_esterno(arc.query_pra.params[0].asstring, codice_archivio);
    end;
  end;

  if programma_da_eseguire <> '' then
  begin
    arc.prg.Close;
    arc.prg.parambyname('codice').asstring := uppercase(programma_da_eseguire);
    arc.prg.open;

    if (arc.prg.fieldbyname('eseguibile').asstring = 'no') and not((arc.prg.fieldbyname('eseguibile_menu').asstring = 'esterno') or (arc.prg.fieldbyname('eseguibile_menu').asstring = 'collegato') or (arc.prg.fieldbyname('eseguibile_menu').asstring = 'query')) then
    begin
      messaggio(000, 'il codice digitato non è eseguibile direttamente');
    end
    else
    begin
      if arc.prg.fieldbyname('eseguibile_menu').asstring = 'esterno' then
      begin
        result := true;
        if pos(' ', trim(arc.prg.fieldbyname('programma_esterno').asstring)) = 0 then
        begin
          programma_esterno := arc.prg.fieldbyname('programma_esterno').asstring;
          cartella_lavoro_esterno := '';
          parametro_esterno := parametri_prg_codice_diretto;
        end
        else
        begin
          programma_esterno := copy(arc.prg.fieldbyname('programma_esterno').asstring, 1, pos(' ', arc.prg.fieldbyname('programma_esterno').asstring) - 1);

          cartella_lavoro_esterno := extractfiledir(arc.prg.fieldbyname('programma_esterno').asstring);

          if parametri_prg_codice_diretto <> '' then
          begin
            parametro_esterno := parametri_prg_codice_diretto;
          end
          else
          begin
            // originale GO
            // parametro_esterno := trim(copy(arc.prg.fieldbyname('programma_esterno').asstring,
            // originale GO fine

            // modifica SISTED
            parametro_esterno := arc.arcdit.server + ' ' + inttostr(arc.arcdit.port) + ' ' + arc.arc.database + ' ' + arc.arcdit.database + ' ' + utente + ' ' + password_utente_login + ' ' + cartella_file + ' ' + ditta + ' ' +
              trim(copy(arc.prg.fieldbyname('programma_esterno').asstring, pos(' ', arc.prg.fieldbyname('programma_esterno').asstring) + 1, length(arc.prg.fieldbyname('programma_esterno').asstring)));
            // modifica SISTED fine
          end;
        end;

        esegui_effettivo(programma_esterno, parametro_esterno, cartella_lavoro_esterno);
      end
      else if arc.prg.fieldbyname('eseguibile_menu').asstring = 'collegato' then
      begin
        result := true;
        if (arc.utn.fieldbyname('modalita_abilitazione_programmi').asstring = 'disabilita') and (read_tabella(arc.abp, vararrayof([utente, arc.prg.fieldbyname('codice').asstring]))) then
        begin
          messaggio(000, 'l''utente non è autorizzato ad accedere al programma ' + arc.prg.fieldbyname('codice').asstring);
        end
        else
        begin
          esegui_collegato(arc.prg.fieldbyname('programma_esterno').asstring);
        end;
      end
      else if arc.prg.fieldbyname('eseguibile_menu').asstring = 'query' then
      begin
        result := true;
        esegui_programma('ESEQUERY', arc.prg.fieldbyname('programma_esterno').asstring, true);
      end
      else
      begin
        j := -1;
        programma_da_eseguire := uppercase(programma_da_eseguire);

        result := true;
        if j = -1 then
        begin
          // controlli accesso
          arc.prg.Close;

          if not non_eseguire then
          begin
            screen.Cursor := crhourglass;

            assegna_programma_personalizzato(programma_da_eseguire, codice_archivio, programma_personalizzato);
            if not programma_personalizzato then
            begin
              if vartype(codice_archivio) = varustring then
              begin

                if (trim(programma_da_eseguire) = 'GESESE') or (trim(programma_da_eseguire) = 'GESESEINH') or (trim(programma_da_eseguire) = 'GESVIS') or (trim(programma_da_eseguire) = 'GESVISINH') or (trim(programma_da_eseguire) = 'GESABI') or (trim(programma_da_eseguire) = 'GESABIINH') or
                  (trim(programma_da_eseguire) = 'GESTVF') or (trim(programma_da_eseguire) = 'GESTVFINH') or (trim(programma_da_eseguire) = 'GESIND') or (trim(programma_da_eseguire) = 'GESINDINH') or (trim(programma_da_eseguire) = 'GESTDOCOLL') or (trim(programma_da_eseguire) = 'GESTDOCOLLINH') or
                  (trim(programma_da_eseguire) = 'GESMAG') or (trim(programma_da_eseguire) = 'GESMAGINH') or (trim(programma_da_eseguire) = 'GESTBO') or (trim(programma_da_eseguire) = 'GESTBOINH') or (trim(programma_da_eseguire) = 'GESSCT') or (trim(programma_da_eseguire) = 'GESSCTINH') or
                  (trim(programma_da_eseguire) = 'GESTCZ') or (trim(programma_da_eseguire) = 'GESTCZINH') or (trim(programma_da_eseguire) = 'GESGNG') or (trim(programma_da_eseguire) = 'GESGNGINH') or (trim(programma_da_eseguire) = 'GOOGLEMAPPA') or (trim(programma_da_eseguire) = 'GOOGLEMAPPAINH') or
                  (trim(programma_da_eseguire) = 'GESNMD') or (trim(programma_da_eseguire) = 'GESNMDINH') or (trim(programma_da_eseguire) = 'GESCPV') or (trim(programma_da_eseguire) = 'GESCPVINH') or (trim(programma_da_eseguire) = 'GESIMGE') or (trim(programma_da_eseguire) = 'GESIMGEINH') or
                  (trim(programma_da_eseguire) = 'VISCMSOV') or (trim(programma_da_eseguire) = 'VISCMSOVINH') or (trim(programma_da_eseguire) = 'GESILF') or (trim(programma_da_eseguire) = 'GESILFINH') or (trim(programma_da_eseguire) = 'GESTNCAC') or (trim(programma_da_eseguire) = 'GESTNCACINH') or
                  (trim(programma_da_eseguire) = 'GESSOA') or (trim(programma_da_eseguire) = 'GESSOAINH') or (trim(programma_da_eseguire) = 'SCHCON') or (trim(programma_da_eseguire) = 'SCHCONINH') or (trim(programma_da_eseguire) = 'SCHCON01') or (trim(programma_da_eseguire) = 'SCHCON01INH') or
                  (trim(programma_da_eseguire) = 'STAPAR') or (trim(programma_da_eseguire) = 'STAPARINH') or (trim(programma_da_eseguire) = 'STAPAR01') or (trim(programma_da_eseguire) = 'STAPAR01INH') or (trim(programma_da_eseguire) = 'STAEST') or (trim(programma_da_eseguire) = 'STAESTINH') or
                  (trim(programma_da_eseguire) = 'SITFID') or (trim(programma_da_eseguire) = 'SITFIDINH') or (trim(programma_da_eseguire) = 'RIEORDV') or (trim(programma_da_eseguire) = 'RIEORDVINH') or (trim(programma_da_eseguire) = 'RIEDOCV') or (trim(programma_da_eseguire) = 'RIEDOCVINH') or
                  (trim(programma_da_eseguire) = 'TRADOCV') or (trim(programma_da_eseguire) = 'TRADOCVINH') or (trim(programma_da_eseguire) = 'GESPRP') or (trim(programma_da_eseguire) = 'GESPRPINH') or (trim(programma_da_eseguire) = 'RIEPREA') or (trim(programma_da_eseguire) = 'RIEPREAINH') or
                  (trim(programma_da_eseguire) = 'RIEORDA') or (trim(programma_da_eseguire) = 'RIEORDAINH') or (trim(programma_da_eseguire) = 'GESUTP') or (trim(programma_da_eseguire) = 'GESUTPINH') or (trim(programma_da_eseguire) = 'GESUTQ') or (trim(programma_da_eseguire) = 'GESUTQINH') or
                  (trim(programma_da_eseguire) = 'GESABD') or (trim(programma_da_eseguire) = 'GESABDINH') or (trim(programma_da_eseguire) = 'GESSCF') or (trim(programma_da_eseguire) = 'GESSCFINH') or (trim(programma_da_eseguire) = 'GESABP') or (trim(programma_da_eseguire) = 'GESABPINH') or
                  (trim(programma_da_eseguire) = 'GESCPA') or (trim(programma_da_eseguire) = 'GESCPAINH') or (trim(programma_da_eseguire) = 'GESPRGSUB') or (trim(programma_da_eseguire) = 'GESPRGSUBINH') or (trim(programma_da_eseguire) = 'GESPGPSUB') or (trim(programma_da_eseguire) = 'GESPGPSUBINH')
                  or (trim(programma_da_eseguire) = 'GESARC01') or (trim(programma_da_eseguire) = 'GESARC01INH') or (trim(programma_da_eseguire) = 'GESARC02') or (trim(programma_da_eseguire) = 'GESARC02INH') or (trim(programma_da_eseguire) = 'GESTR2') or (trim(programma_da_eseguire) = 'GESTR2INH')
                  or (trim(programma_da_eseguire) = 'GESTR3') or (trim(programma_da_eseguire) = 'GESTR3INH') or (trim(programma_da_eseguire) = 'GESCLD') or (trim(programma_da_eseguire) = 'GESCLDINH') or (trim(programma_da_eseguire) = 'GESFRD') or (trim(programma_da_eseguire) = 'GESFRDINH') or
                  (trim(programma_da_eseguire) = 'GESARA') or (trim(programma_da_eseguire) = 'GESARAINH') or (trim(programma_da_eseguire) = 'GESARF01') or (trim(programma_da_eseguire) = 'GESARF01INH') or (trim(programma_da_eseguire) = 'GESARF02') or (trim(programma_da_eseguire) = 'GESARF02INH') or
                  (trim(programma_da_eseguire) = 'GESBAR') or (trim(programma_da_eseguire) = 'GESBARINH') or (trim(programma_da_eseguire) = 'GESTCPECDES') or (trim(programma_da_eseguire) = 'GESTCPECDESINH') or (trim(programma_da_eseguire) = 'GESVDC') or (trim(programma_da_eseguire) = 'GESVDCINH') or
                  (trim(programma_da_eseguire) = 'GESVDA') or (trim(programma_da_eseguire) = 'GESVDAINH') or (trim(programma_da_eseguire) = 'GESACC') or (trim(programma_da_eseguire) = 'GESACCINH') or (trim(programma_da_eseguire) = 'GESEQU') or (trim(programma_da_eseguire) = 'GESEQUINH') or
                  (trim(programma_da_eseguire) = 'GESACCPR') or (trim(programma_da_eseguire) = 'GESACCPRINH') or (trim(programma_da_eseguire) = 'GESEQUPR') or (trim(programma_da_eseguire) = 'GESEQUPRINH') or (trim(programma_da_eseguire) = 'GESLOT') or (trim(programma_da_eseguire) = 'GESLOTINH') or
                  (trim(programma_da_eseguire) = 'GESUBI') or (trim(programma_da_eseguire) = 'GESUBIINH') or (trim(programma_da_eseguire) = 'GESUBI01') or (trim(programma_da_eseguire) = 'GESUBI01INH') or (trim(programma_da_eseguire) = 'GESLIF') or (trim(programma_da_eseguire) = 'GESLIFINH') or
                  (trim(programma_da_eseguire) = 'GESSAF') or (trim(programma_da_eseguire) = 'GESSAFINH') or (trim(programma_da_eseguire) = 'VISCFG') or (trim(programma_da_eseguire) = 'VISCFGINH') or (trim(programma_da_eseguire) = 'STAINS') or (trim(programma_da_eseguire) = 'STAINSINH') or
                  (trim(programma_da_eseguire) = 'RITPAG') or (trim(programma_da_eseguire) = 'RITPAGINH') or (trim(programma_da_eseguire) = 'RIEPREV') or (trim(programma_da_eseguire) = 'RIEPREVINH') or (trim(programma_da_eseguire) = 'INVCIC01') or (trim(programma_da_eseguire) = 'INVCIC01INH') or
                  (trim(programma_da_eseguire) = 'GESSOA') or (trim(programma_da_eseguire) = 'GESSOAINH') or (trim(programma_da_eseguire) = 'GESINF') or (trim(programma_da_eseguire) = 'GESINFINH') or (trim(programma_da_eseguire) = 'GESABM') or (trim(programma_da_eseguire) = 'GESABMINH') or
                  (trim(programma_da_eseguire) = 'GESFLT') or (trim(programma_da_eseguire) = 'GESFLTINH') or (trim(programma_da_eseguire) = 'GESPCS01') or (trim(programma_da_eseguire) = 'GESPCS01INH') or (trim(programma_da_eseguire) = 'GESBDG') or (trim(programma_da_eseguire) = 'GESBDGINH') or
                  (trim(programma_da_eseguire) = 'GESCUM') or (trim(programma_da_eseguire) = 'GESCUMINH') or (trim(programma_da_eseguire) = 'GESLDA') or (trim(programma_da_eseguire) = 'GESLDAINH') or (trim(programma_da_eseguire) = 'GESCAC') or (trim(programma_da_eseguire) = 'GESCACINH') or
                  (trim(programma_da_eseguire) = 'RIEDOCA') or (trim(programma_da_eseguire) = 'RIEDOCAINH') or (trim(programma_da_eseguire) = 'GESCMT') or (trim(programma_da_eseguire) = 'GESCMTINH') or (trim(programma_da_eseguire) = 'GESCMC') or (trim(programma_da_eseguire) = 'GESCMCINH') or
                  (trim(programma_da_eseguire) = 'GESPRD') or (trim(programma_da_eseguire) = 'GESPRDINH') or (trim(programma_da_eseguire) = 'GESTRA') or (trim(programma_da_eseguire) = 'GESTRAINH') or (trim(programma_da_eseguire) = 'GESTDC') or (trim(programma_da_eseguire) = 'GESTDCINH') or
                  (trim(programma_da_eseguire) = 'GESTCU') or (trim(programma_da_eseguire) = 'GESTCUINH') or (trim(programma_da_eseguire) = 'GESPCL') or (trim(programma_da_eseguire) = 'GESPCLINH') or (trim(programma_da_eseguire) = 'GESPCS') or (trim(programma_da_eseguire) = 'GESPCSINH') then
                begin
                  codice_archivio := vararrayof(['', '']);
                end
                else if (trim(programma_da_eseguire) = 'GESSTM') or (trim(programma_da_eseguire) = 'GESSTMINH') or (trim(programma_da_eseguire) = 'GESCNT') or (trim(programma_da_eseguire) = 'GESCNTINH') or (trim(programma_da_eseguire) = 'GESTTI') or (trim(programma_da_eseguire) = 'GESTTIINH') or
                  (trim(programma_da_eseguire) = 'GESTRF') or (trim(programma_da_eseguire) = 'GESTRFINH') or (trim(programma_da_eseguire) = 'GESNML') or (trim(programma_da_eseguire) = 'GESNMLINH') or (trim(programma_da_eseguire) = 'GESLSV') or (trim(programma_da_eseguire) = 'GESLSVINH') or
                  (trim(programma_da_eseguire) = 'GESCLA') or (trim(programma_da_eseguire) = 'GESCLAINH') or (trim(programma_da_eseguire) = 'GESPARSEL') or (trim(programma_da_eseguire) = 'GESPARSELINH') or (trim(programma_da_eseguire) = 'GESVDD') or (trim(programma_da_eseguire) = 'GESVDDINH') or
                  (trim(programma_da_eseguire) = 'GESCGC') or (trim(programma_da_eseguire) = 'GESCGCINH') then
                begin
                  codice_archivio := vararrayof(['', '', '']);
                end
                else if (trim(programma_da_eseguire) = 'GESTPC') or (trim(programma_da_eseguire) = 'GESTPCINH') or (trim(programma_da_eseguire) = 'GESTCS') or (trim(programma_da_eseguire) = 'GESTCSINH') then
                begin
                  codice_archivio := vararrayof(['', '', '', '']);
                end
                else if (trim(programma_da_eseguire) = 'GESCTC') or (trim(programma_da_eseguire) = 'GESCTCINH') or (trim(programma_da_eseguire) = 'GESPRO') or (trim(programma_da_eseguire) = 'GESPROINH') then
                begin
                  codice_archivio := vararrayof(['', '', '', 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESTSI') or (trim(programma_da_eseguire) = 'GESTSIINH') or (trim(programma_da_eseguire) = 'GESLSA') or (trim(programma_da_eseguire) = 'GESLSAINH') or (trim(programma_da_eseguire) = 'GESLAF') or (trim(programma_da_eseguire) = 'GESLAFINH') then
                begin
                  codice_archivio := vararrayof(['', '', 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESDIPCOS') or (trim(programma_da_eseguire) = 'GESDIPCOSINH') or (trim(programma_da_eseguire) = 'GESPVVOLD') or (trim(programma_da_eseguire) = 'GESPVVOLDINH') or (trim(programma_da_eseguire) = 'GESRAP') or
                  (trim(programma_da_eseguire) = 'GESRAPINH') or (trim(programma_da_eseguire) = 'GENSDA') or (trim(programma_da_eseguire) = 'GENSDAINH') or (trim(programma_da_eseguire) = 'STAPKL') or (trim(programma_da_eseguire) = 'STAPKLINH') or (trim(programma_da_eseguire) = 'GESNSL') or
                  (trim(programma_da_eseguire) = 'GESNSLINH') or (trim(programma_da_eseguire) = 'GESCCF') or (trim(programma_da_eseguire) = 'GESCCFINH') or (trim(programma_da_eseguire) = 'GESCPP') or (trim(programma_da_eseguire) = 'GESCPPINH') or (trim(programma_da_eseguire) = 'GESKIT') or
                  (trim(programma_da_eseguire) = 'GESKITINH') or (trim(programma_da_eseguire) = 'GESPTI') or (trim(programma_da_eseguire) = 'GESPTIINH') or (trim(programma_da_eseguire) = 'GESMCS') or (trim(programma_da_eseguire) = 'GESMCSINH') or (trim(programma_da_eseguire) = 'VARAMM') or
                  (trim(programma_da_eseguire) = 'VARAMMINH') or (trim(programma_da_eseguire) = 'GENBARTO') or (trim(programma_da_eseguire) = 'GENBARTOINH') then
                begin
                  codice_archivio := vararrayof(['', 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESVARDSB') or (trim(programma_da_eseguire) = 'GESVARDSBINH') then
                begin
                  codice_archivio := vararrayof(['', 0, 0, '', '']);
                end
                else if (trim(programma_da_eseguire) = 'STOEVAV') or (trim(programma_da_eseguire) = 'STOEVAVINH') or (trim(programma_da_eseguire) = 'STOCONV') or (trim(programma_da_eseguire) = 'STOCONVINH') or (trim(programma_da_eseguire) = 'STOFADV') or (trim(programma_da_eseguire) = 'STOFADVINH')
                  or (trim(programma_da_eseguire) = 'STOEVAA') or (trim(programma_da_eseguire) = 'STOEVAAINH') or (trim(programma_da_eseguire) = 'STOFADA') or (trim(programma_da_eseguire) = 'STOFADAINH') or (trim(programma_da_eseguire) = 'STOCONA') or (trim(programma_da_eseguire) = 'STOCONAINH')
                then
                begin
                  codice_archivio := vararrayof([0, '']);
                end
                else if (trim(programma_da_eseguire) = 'CARPROD') or (trim(programma_da_eseguire) = 'CARPRODINH') then
                begin
                  codice_archivio := vararrayof([utente, '']);
                end
                else if (trim(programma_da_eseguire) = 'CONDOCV') or (trim(programma_da_eseguire) = 'CONDOCVINH') or (trim(programma_da_eseguire) = 'GESMNU') or (trim(programma_da_eseguire) = 'GESMNUINH') or (trim(programma_da_eseguire) = 'GESDSB') or (trim(programma_da_eseguire) = 'GESDSBINH') or
                  (trim(programma_da_eseguire) = 'GESSAE') or (trim(programma_da_eseguire) = 'GESSAEINH') or (trim(programma_da_eseguire) = 'GESCFD') or (trim(programma_da_eseguire) = 'GESCFDINH') then
                begin
                  codice_archivio := vararrayof(['', '0', '']);
                end
                else if (trim(programma_da_eseguire) = 'ASSAPPCL') or (trim(programma_da_eseguire) = 'ASSAPPCLINH') or (trim(programma_da_eseguire) = 'CRESOS') or (trim(programma_da_eseguire) = 'CRESOSINH') or (trim(programma_da_eseguire) = 'ASSAPP') or (trim(programma_da_eseguire) = 'ASSAPPINH') or
                  (trim(programma_da_eseguire) = 'ELAORDP') or (trim(programma_da_eseguire) = 'ELAORDPINH') or (trim(programma_da_eseguire) = 'AGGAST') or (trim(programma_da_eseguire) = 'AGGASTINH') or (trim(programma_da_eseguire) = 'AGGAST01') or (trim(programma_da_eseguire) = 'ETIBARTO') or
                  (trim(programma_da_eseguire) = 'ETIBARTOINH') then
                begin
                  codice_archivio := vararrayof([0, 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESMACCHFRN') or (trim(programma_da_eseguire) = 'GESMACCHFRNINH') or (trim(programma_da_eseguire) = 'GESMACCOSFRN') or (trim(programma_da_eseguire) = 'GESMACCOSFRNINH') then
                begin
                  codice_archivio := vararrayof(['', 0, 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESENACO') or (trim(programma_da_eseguire) = 'GESENACOINH') then
                begin
                  codice_archivio := vararrayof(['', 0, '', 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESSTV') or (trim(programma_da_eseguire) = 'GESSTVINH') then
                begin
                  codice_archivio := vararrayof(['', '', '', 0, 0]);
                end
                else if (trim(programma_da_eseguire) = 'GESPRC') or (trim(programma_da_eseguire) = 'GESPRCINH') then
                begin
                  codice_archivio := vararrayof(['', '', '', '', '', '']);
                end
                else if (trim(programma_da_eseguire) = 'GESSTA') or (trim(programma_da_eseguire) = 'GESSTAINH') then
                begin
                  codice_archivio := vararrayof(['', '', 0, 0]);
                end;
              end;
            end;

            // ESECUZIONE
            baseclass := tbaseclass(getclass('t' + programma_da_eseguire));
            if assigned(baseclass) then
            begin
              if not assigned(pr) then
              begin
                pr := baseclass.create(application);
                pr.codice := codice_archivio;
                pr.record_singolo := record_singolo;
                if modale or (arc.utn.fieldbyname('esegui_modale').asstring = 'si') then
                begin
                  if pr.esegui_form then
                  begin
                    pr.showmodal;
                  end;
                  pr.free;
                end
                else
                begin
                  if pr.esegui_form then
                  begin
                    pr.show;
                  end;
                end;
              end
              else
              begin
                messaggio(000, 'programma non eseguito');
              end;
            end
            else
            begin
              // ******************************************************************************
              // il programma passato non è nella lista precedente
              // ******************************************************************************
              result := false;
              if programma_chiamante <> 'MENUGG' then
              begin
                messaggio(000, 'il programma selezionato [' + programma_da_eseguire + '] non esiste');
              end;
            end;

            screen.Cursor := cursore;
          end;
        end;
      end;
    end;
  end;
end;

procedure TARC.DataModuleCreate(Sender: TObject);
begin
  if extractfilename(lowercase(application.exename)) = 'go_easy.exe' then
  begin
    go_easy := true;
  end
  else
  begin
    go_easy := false;
  end;

  // disabilita stili per sDialogs
  TStyleManager.SystemHooks := [shToolTips];
  // abilita stili per sDialogs
  // TStyleManager.SystemHooks := [shMenus, shDialogs, shToolTips];

  // versione_procedura := GetExeVersion(application.exename);
  versione_procedura := '11.01.02';
  versione_aggiornamento := strtoint(copy(versione_procedura, 7, 2));

  mysqlembdisableeventlog := true;

  versione_os := GetWinVersion;

  // dns google
  ipwmx.dnsserver := '8.8.8.8';

  // per gestire ActiveControlChange
  screen.OnActiveFormChange := ActiveFormChange;
  // fine per gestire ActiveControlChange

  // lista programmi recenti
  lista_programmi_recenti := tstringlist.create;

  // lista programmi personalizzati
  lista_personalizzati := tstringlist.create;

  // archivio e archivio_arc
  archivio_arc := tmyquery_go.create(nil);
  archivio_arc.connection := arc;

  archivio := tmyquery_go.create(nil);
  archivio.connection := arcdit;

  // Attiva traduzione griglie devexpress
  if fileexists(extractfilepath(application.exename) + 'grid6_ita.ini') then
  begin
    cxTraduttore.FileName := extractfilepath(application.exename) + 'grid6_ita.ini';
    cxTraduttore.active := true;
    cxTraduttore.Locale := 1040;
  end;

  crittografia := tmyencryptor.create(nil);
  crittografia.dataheader := ehNone;
  crittografia.encryptionalgorithm := eaTripleDES;
  crittografia.password := 'GESTIONALEOPEN';
end;

procedure TARC.assegna_skin;
var
  style: tstyleinfo;
begin
  skin_utilizzato := 'Windows';
  try
    if (utn.fieldbyname('skin_utilizzato').asstring <> '') and (utn.fieldbyname('skin_utilizzato').asstring <> 'nessuno') then
    begin
      skin_utilizzato := utn.fieldbyname('skin_utilizzato').asstring;
    end;
  except
  end;

  if fileexists(cartella_stili + skin_utilizzato + '.vsf') and TStyleManager.isvalidstyle(cartella_stili + skin_utilizzato + '.vsf', style) then
  begin
    TStyleManager.loadfromfile(cartella_stili + skin_utilizzato + '.vsf');
    try
      TStyleManager.trysetstyle(style.name);
    except
    end;
  end;

  try
    if utn.fieldbyname('skin_utilizzato_devexpress').asstring <> '' then
    begin
      dxSkinController1.skinname := utn.fieldbyname('skin_utilizzato_devexpress').asstring;
      dxSkinController1.useskins := true;
    end
    else
    begin
      dxSkinController1.skinname := '';
      dxSkinController1.useskins := false;
    end;
  except
  end;
end;

procedure TARC.Contatti1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from cli where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from cli where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    parametri_extra_programma_chiamato[0] := 'C';
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, 0]);
    esegui_programma('GESCCF', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il cliente non è stato ancora memorizzato');
  end;
end;

procedure TARC.Contatti2Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from frn where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from frn where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    parametri_extra_programma_chiamato[0] := 'F';
    codice_passato := vararrayof([collegamenti_archivio.fieldbyname('codice').asstring, 0]);
    esegui_programma('GESCCF', codice_passato, true);
  end
  else
  begin
    messaggio(000, 'il fornitore non è stato ancora memorizzato');
  end;
end;

procedure TARC.DataModuleDestroy(Sender: TObject);
begin
  // lista programmi recenti
  lista_programmi_recenti.free;

  // lista programmi personalizzati
  lista_personalizzati.free;

  // archivio e archivio_arc
  archivio.free;
  archivio_arc.free;

  crittografia.free;
end;

procedure TARC.ActiveFormChange(Sender: TObject);
begin
  if (screen.ActiveCustomForm is TGESARC) then
  begin
    screen.OnActiveControlChange := TGESARC(screen.ActiveCustomForm).ActiveControlChange;
  end
  else if (screen.ActiveCustomForm is TFORMBASE) then
  begin
    screen.OnActiveControlChange := TFORMBASE(screen.ActiveCustomForm).ActiveControlChange;
  end
  else if (screen.ActiveCustomForm is tbase) then
  begin
    screen.OnActiveControlChange := tbase(screen.ActiveCustomForm).ActiveControlChange;
  end
  else
  begin
    screen.OnActiveControlChange := nil;
  end;

  assegna_monitor;
end;

procedure TARC.assegna_monitor;
var
  i: word;
begin
  if assigned(screen.activeform) then
  begin
    for i := 0 to screen.monitorcount - 1 do
    begin
      try
        if screen.monitorfromwindow(screen.activeform.handle) = screen.monitors[i] then
        begin
          monitor_attivo := i;
        end;
      except
      end;
    end;
  end;
end;

procedure TARC.aggiorna_data_fine(nome_tabella, operazione, campo_01, codice_01, campo_02, codice_02, campo_03, codice_03, campo_04, codice_04, campo_05, codice_05: string; data_inizio: tdatetime);
//
// aggiornamento data fine validità
//
var
  query, query_controllo: tmyquery_go;
  data_fine: tdatetime;
begin
  query := tmyquery_go.create(nil);
  query_controllo := tmyquery_go.create(nil);
  query.connection := arcdit;
  query_controllo.connection := arcdit;

  if operazione = 'A' then
  begin
    query.Close;
    query.sql.clear;
    query.sql.add('select * from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
    if campo_02 <> '' then
    begin
      query.sql.add('and ' + campo_02 + ' = :campo_02');
    end;
    if campo_03 <> '' then
    begin
      query.sql.add('and ' + campo_03 + ' = :campo_03');
    end;
    if campo_04 <> '' then
    begin
      query.sql.add('and ' + campo_04 + ' = :campo_04');
    end;
    if campo_05 <> '' then
    begin
      query.sql.add('and ' + campo_05 + ' = :campo_05');
    end;
    query.sql.add('and data_inizio = :data_inizio');

    query.params[0].asstring := codice_01;
    if campo_02 <> '' then
    begin
      query.params[1].asstring := codice_02;
    end;
    if campo_03 <> '' then
    begin
      query.params[2].asstring := codice_03;
    end;
    if campo_04 <> '' then
    begin
      query.params[3].asstring := codice_04;
    end;
    if campo_05 <> '' then
    begin
      query.params[4].asstring := codice_05;
    end;
    query.parambyname('data_inizio').asdate := data_inizio;
    query.open;

    // gestione data_fine record attivo
    query_controllo.Close;
    query_controllo.sql.clear;
    query_controllo.sql.add('select id from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
    if campo_02 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
    end;
    if campo_04 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
    end;
    if campo_05 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
    end;
    query_controllo.sql.add('and data_inizio > :data_inizio');

    query_controllo.params[0].asstring := codice_01;
    if campo_02 <> '' then
    begin
      query_controllo.params[1].asstring := codice_02;
    end;
    if campo_03 <> '' then
    begin
      query_controllo.params[2].asstring := codice_03;
    end;
    if campo_04 <> '' then
    begin
      query_controllo.params[3].asstring := codice_04;
    end;
    if campo_05 <> '' then
    begin
      query_controllo.params[4].asstring := codice_05;
    end;
    query_controllo.parambyname('data_inizio').asdate := data_inizio;
    query_controllo.open;

    if query_controllo.isempty then
    begin
      // non esiste successivo e assegno come data_fine 31/12/9999
      query.edit;
      query.fieldbyname('data_fine').asstring := data_31_12_9999;
      query.post;
    end
    else
    begin
      // esiste successivo e assegno come data_fine quella inizio_successivo - 1
      query_controllo.Close;
      query_controllo.sql.clear;
      query_controllo.sql.add('select data_inizio from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
      if campo_02 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
      end;
      query_controllo.sql.add('and data_inizio > :data_inizio');
      query_controllo.sql.add('order by ' + campo_01);
      if campo_02 <> '' then
      begin
        query_controllo.sql.add(',' + campo_02);
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add(',' + campo_03);
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add(',' + campo_04);
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add(',' + campo_05);
      end;
      query_controllo.sql.add(',data_inizio limit 1');

      query_controllo.params[0].asstring := codice_01;
      if campo_02 <> '' then
      begin
        query_controllo.params[1].asstring := codice_02;
      end;
      if campo_03 <> '' then
      begin
        query_controllo.params[2].asstring := codice_03;
      end;
      if campo_04 <> '' then
      begin
        query_controllo.params[3].asstring := codice_04;
      end;
      if campo_05 <> '' then
      begin
        query_controllo.params[4].asstring := codice_05;
      end;
      query_controllo.parambyname('data_inizio').asdate := data_inizio;
      query_controllo.open;

      query.edit;
      query.fieldbyname('data_fine').asdatetime := query_controllo.fieldbyname('data_inizio').asdatetime - 1;
      query.post;
    end;

    // gestione precedente e data_fine precedente
    query_controllo.Close;
    query_controllo.sql.clear;
    query_controllo.sql.add('select id from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
    if campo_02 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
    end;
    if campo_04 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
    end;
    if campo_05 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
    end;
    query_controllo.sql.add('and data_inizio < :data_inizio');

    query_controllo.params[0].asstring := codice_01;
    if campo_02 <> '' then
    begin
      query_controllo.params[1].asstring := codice_02;
    end;
    if campo_03 <> '' then
    begin
      query_controllo.params[2].asstring := codice_03;
    end;
    if campo_04 <> '' then
    begin
      query_controllo.params[3].asstring := codice_04;
    end;
    if campo_05 <> '' then
    begin
      query_controllo.params[4].asstring := codice_05;
    end;
    query_controllo.parambyname('data_inizio').asdate := data_inizio;
    query_controllo.open;

    if not query_controllo.isempty then
    begin
      query_controllo.Close;
      query_controllo.sql.clear;
      query_controllo.sql.add('select data_inizio from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
      if campo_02 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
      end;
      query_controllo.sql.add('and data_inizio < :data_inizio');
      query_controllo.sql.add('order by ' + campo_01);
      if campo_02 <> '' then
      begin
        query_controllo.sql.add(',' + campo_02);
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add(',' + campo_03);
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add(',' + campo_04);
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add(',' + campo_05);
      end;
      query_controllo.sql.add(',data_inizio desc limit 1');

      query_controllo.params[0].asstring := codice_01;
      if campo_02 <> '' then
      begin
        query_controllo.params[1].asstring := codice_02;
      end;
      if campo_03 <> '' then
      begin
        query_controllo.params[2].asstring := codice_03;
      end;
      if campo_04 <> '' then
      begin
        query_controllo.params[3].asstring := codice_04;
      end;
      if campo_05 <> '' then
      begin
        query_controllo.params[4].asstring := codice_05;
      end;
      query_controllo.parambyname('data_inizio').asdate := data_inizio;
      query_controllo.open;

      query.Close;
      query.sql.clear;
      query.sql.add('select * from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
      if campo_02 <> '' then
      begin
        query.sql.add('and ' + campo_02 + ' = :campo_02');
      end;
      if campo_03 <> '' then
      begin
        query.sql.add('and ' + campo_03 + ' = :campo_03');
      end;
      if campo_04 <> '' then
      begin
        query.sql.add('and ' + campo_04 + ' = :campo_04');
      end;
      if campo_05 <> '' then
      begin
        query.sql.add('and ' + campo_05 + ' = :campo_05');
      end;
      query.sql.add('and data_inizio = :data_inizio');

      query.params[0].asstring := codice_01;
      if campo_02 <> '' then
      begin
        query.params[1].asstring := codice_02;
      end;
      if campo_03 <> '' then
      begin
        query.params[2].asstring := codice_03;
      end;
      if campo_04 <> '' then
      begin
        query.params[3].asstring := codice_04;
      end;
      if campo_05 <> '' then
      begin
        query.params[4].asstring := codice_05;
      end;
      query.parambyname('data_inizio').asdate := query_controllo.fieldbyname('data_inizio').asdatetime;
      query.open;

      query.edit;
      query.fieldbyname('data_fine').asdatetime := data_inizio - 1;
      query.post;
    end;
  end;

  if operazione = 'D' then
  begin
    query_controllo.Close;
    query_controllo.sql.clear;
    query_controllo.sql.add('select id from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
    if campo_02 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
    end;
    if campo_04 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
    end;
    if campo_05 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
    end;
    query_controllo.sql.add('and data_inizio > :data_inizio');

    query_controllo.params[0].asstring := codice_01;
    if campo_02 <> '' then
    begin
      query_controllo.params[1].asstring := codice_02;
    end;
    if campo_03 <> '' then
    begin
      query_controllo.params[2].asstring := codice_03;
    end;
    if campo_04 <> '' then
    begin
      query_controllo.params[3].asstring := codice_04;
    end;
    if campo_05 <> '' then
    begin
      query_controllo.params[4].asstring := codice_05;
    end;
    query_controllo.parambyname('data_inizio').asdate := data_inizio;
    query_controllo.open;

    if query_controllo.isempty then
    begin
      // non esiste successivo e assegno come data_fine 31/12/9999
      data_fine := strtodate(data_31_12_9999);
    end
    else
    begin
      // esiste successivo e assegno come data_fine quella inizio_successivo - 1
      query_controllo.Close;
      query_controllo.sql.clear;
      query_controllo.sql.add('select data_inizio from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
      if campo_02 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
      end;
      query_controllo.sql.add('and data_inizio > :data_inizio');

      query_controllo.sql.add('order by ' + campo_01);
      if campo_02 <> '' then
      begin
        query_controllo.sql.add(',' + campo_02);
      end;
      if campo_03 <> '' then
      begin
        query_controllo.sql.add(',' + campo_03);
      end;
      if campo_04 <> '' then
      begin
        query_controllo.sql.add(',' + campo_04);
      end;
      if campo_05 <> '' then
      begin
        query_controllo.sql.add(',' + campo_05);
      end;
      query_controllo.sql.add(',data_inizio limit 1');

      query_controllo.params[0].asstring := codice_01;
      if campo_02 <> '' then
      begin
        query_controllo.params[1].asstring := codice_02;
      end;
      if campo_03 <> '' then
      begin
        query_controllo.params[2].asstring := codice_03;
      end;
      if campo_04 <> '' then
      begin
        query_controllo.params[3].asstring := codice_04;
      end;
      if campo_05 <> '' then
      begin
        query_controllo.params[4].asstring := codice_05;
      end;
      query_controllo.parambyname('data_inizio').asdate := data_inizio;
      query_controllo.open;

      data_fine := query_controllo.fieldbyname('data_inizio').asdatetime - 1;
    end;

    // gestione precedente e data_fine precedente
    query_controllo.Close;
    query_controllo.sql.clear;
    query_controllo.sql.add('select data_inizio from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');

    if campo_02 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_02 + ' = :campo_02');
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_03 + ' = :campo_03');
    end;
    if campo_04 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_04 + ' = :campo_04');
    end;
    if campo_05 <> '' then
    begin
      query_controllo.sql.add('and ' + campo_05 + ' = :campo_05');
    end;
    query_controllo.sql.add('and data_inizio < :data_inizio');

    query_controllo.sql.add('order by ' + campo_01);
    if campo_02 <> '' then
    begin
      query_controllo.sql.add(',' + campo_02);
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add(',' + campo_03);
    end;
    if campo_03 <> '' then
    begin
      query_controllo.sql.add(',' + campo_03);
    end;
    if campo_04 <> '' then
    begin
      query_controllo.sql.add(',' + campo_04);
    end;
    if campo_05 <> '' then
    begin
      query_controllo.sql.add(',' + campo_05);
    end;
    query_controllo.sql.add(',data_inizio desc limit 1');

    query_controllo.params[0].asstring := codice_01;
    if campo_02 <> '' then
    begin
      query_controllo.params[1].asstring := codice_02;
    end;
    if campo_03 <> '' then
    begin
      query_controllo.params[2].asstring := codice_03;
    end;
    if campo_04 <> '' then
    begin
      query_controllo.params[3].asstring := codice_04;
    end;
    if campo_05 <> '' then
    begin
      query_controllo.params[4].asstring := codice_05;
    end;
    query_controllo.parambyname('data_inizio').asdate := data_inizio;
    query_controllo.open;

    if not query_controllo.isempty then
    begin
      query.Close;
      query.sql.clear;
      query.sql.add('select * from ' + nome_tabella + ' where ' + campo_01 + ' = :campo_01');
      if campo_02 <> '' then
      begin
        query.sql.add('and ' + campo_02 + ' = :campo_02');
      end;
      if campo_03 <> '' then
      begin
        query.sql.add('and ' + campo_03 + ' = :campo_03');
      end;
      if campo_04 <> '' then
      begin
        query.sql.add('and ' + campo_04 + ' = :campo_04');
      end;
      if campo_05 <> '' then
      begin
        query.sql.add('and ' + campo_05 + ' = :campo_05');
      end;
      query.sql.add('and data_inizio = :data_inizio');
      query.params[0].asstring := codice_01;
      if campo_02 <> '' then
      begin
        query.params[1].asstring := codice_02;
      end;
      if campo_03 <> '' then
      begin
        query.params[2].asstring := codice_03;
      end;
      if campo_04 <> '' then
      begin
        query.params[3].asstring := codice_04;
      end;
      if campo_05 <> '' then
      begin
        query.params[4].asstring := codice_05;
      end;
      query.parambyname('data_inizio').asdate := query_controllo.fieldbyname('data_inizio').asdatetime;
      query.open;

      query.edit;
      query.fieldbyname('data_fine').asdatetime := data_fine;
      query.post;
    end;
  end;

  query.free;
  query_controllo.free;
end;

procedure TARC.spezza_descrizione(descrizione: string; var descrizione1, descrizione2: string; caratteri: word);
var
  i, j: word;
begin
  j := 0;
  descrizione1 := '';
  descrizione2 := '';

  if length(descrizione) > caratteri then
  begin
    for i := caratteri downto 1 do
    begin
      if descrizione[i] = ' ' then
      begin
        j := i;
        descrizione2 := copy(descrizione, j + 1, length(descrizione) - j);
        break;
      end;
    end;
  end
  else
  begin
    j := length(descrizione);
  end;
  descrizione1 := trim(copy(descrizione, 1, j));
  descrizione2 := trim(descrizione2);
end;

procedure TARC.MenuItem97Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice from nom where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice from nom where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    documenti := Topendialog.create(nil);

    cartella_documenti := cartella_file + '\nominativi\' + collegamenti_archivio.fieldbyname('codice').asstring;
    documenti.initialdir := cartella_documenti;
    if documenti.execute then
    begin
      esegui(documenti.FileName);
    end;

    documenti.free;
  end
  else
  begin
    messaggio(000, 'il sottoconto non è stato ancora memorizzato');
  end;
end;

// ******************************************************************************

procedure TARC.bilanciodicommessa1Click(Sender: TObject);
begin
  collegamenti_archivio.Close;
  collegamenti_archivio.sql.clear;
  if tipo_codice_dati_aggiuntivi_popup = 'id' then
  begin
    collegamenti_archivio.sql.add('select codice, cli_codice from cms where id = :id');
    collegamenti_archivio.params[0].asinteger := codice_dati_aggiuntivi_popup;
  end
  else
  begin
    collegamenti_archivio.sql.add('select codice, cli_codice from cms where codice = :codice_archivio');
    collegamenti_archivio.params[0].asstring := codice_dati_aggiuntivi_popup;
  end;
  collegamenti_archivio.open;
  if not collegamenti_archivio.eof then
  begin
    parametri_extra_programma_chiamato[0] := collegamenti_archivio.fieldbyname('codice').asstring;
    parametri_extra_programma_chiamato[1] := collegamenti_archivio.fieldbyname('cli_codice').asstring;
    esegui_programma('BILCMM', '', true);
  end
  else
  begin
    messaggio(000, 'la commessa non è stata ancora memorizzata');
  end;
end;

function TARC.GetWinVersion: string;
var
  osVerInfo: TOSVersionInfo;
  majorVersion, minorVersion: Integer;
begin
  result := 'Sconosciuto';
  osVerInfo.dwOSVersionInfoSize := SizeOf(TOSVersionInfo);
  if GetVersionEx(osVerInfo) then
  begin
    minorVersion := osVerInfo.dwMinorVersion;
    majorVersion := osVerInfo.dwMajorVersion;

    case osVerInfo.dwPlatformId of
      VER_PLATFORM_WIN32_NT:
        begin
          if (majorVersion <= 5) then
          begin
            result := 'antecedente Windows XP';
          end
          else if (majorVersion = 5) and (minorVersion = 1) then
          begin
            result := 'Windows XP';
          end
          else if (majorVersion = 6) and (minorVersion = 1) then
          begin
            result := 'Windows 7';
          end
          else if (majorVersion = 6) and (minorVersion = 2) then
          begin
            result := 'Windows 8 o superiore';
          end
          else if (majorVersion = 6) then
          begin
            result := 'Windows Vista / Windows 7';
          end;
        end;
      VER_PLATFORM_WIN32_WINDOWS:
        begin
          result := 'antecedente Windows XP';
        end;
    end;
  end;
end;

function TARC.settimana(data: tdate): word;
var
  giorno, mese, anno: word;
  firstofyear: tdatetime;
begin
  decodedate(data, anno, mese, giorno);
  firstofyear := encodedate(anno, 1, 1);
  result := trunc(data - firstofyear) div 7 + 1;
end;

function TARC.assegna_codice_lotto_automatico(data: tdate; frn_codice: string = ''; numero_documento: double = 0; data_documento: tdate = 0; art_codice: string = ''): string;
var
  query: tmyquery_go;
  numero, numero_opt: Integer;
  anno, mese, giorno: word;
const
  mesi = 'ABCDEFGHILMN';
begin
  if dit.fieldbyname('codice_lotto_automatico').asstring = 'si' then
  begin
    if (dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.settimana.giorno') or (dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.settimana.giorno.progressivo') then
    begin
      result := formatdatetime('yyyy', data) + '.' + setta_lunghezza(settimana(data), 2, 0) + '.' + setta_lunghezza(dayoftheweek(data), 2, 0);
      if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.settimana.giorno.progressivo' then
      begin
        query := tmyquery_go.create(nil);
        query.connection := arcdit;

        query.Close;
        query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 10) = ' + quotedstr(result);
        query.open;
        if query.fieldbyname('lotto').value = null then
        begin
          numero := 0;
        end
        else
        begin
          try
            numero := strtoint(copy(query.fieldbyname('lotto').asstring, 12, 4));
          except
            numero := 0;
          end;
        end;

        query.Close;
        query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 10) = ' + quotedstr(result);
        query.open;
        try
          numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 12, 4));
        except
          numero_opt := 0;
        end;
        if numero_opt > numero then
        begin
          numero := numero_opt;
        end;

        inc(numero);
        result := result + '.' + setta_lunghezza(numero, 4, 0);

        query.free;
      end;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.giorno.progressivo' then
    begin
      result := formatdatetime('yyyy', data) + '.' + setta_lunghezza(dayoftheyear(data), 3, 0);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 8) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 10, 4));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 8) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 10, 4));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + '.' + setta_lunghezza(numero, 4, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'aamprogressivo' then
    begin
      decodedate(data, anno, mese, giorno);
      result := formatdatetime('yy', data);
      result := result + mesi[mese];

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(mid(lotto, 4, 27)) lotto from lot where left(lotto, 3) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(query.fieldbyname('lotto').asstring);
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(mid(lot_codice, 4, 27)) lotto from opt where left(lot_codice, 3) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(query.fieldbyname('lotto').asstring);
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 5, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'aaMprogressivo' then
    begin
      decodedate(data, anno, mese, giorno);
      result := formatdatetime('yy', data);
      result := result + mesi[mese];

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(mid(lotto, 4, 27)) lotto from lot where left(lotto, 2) = ' + quotedstr(copy(result, 1, 2));
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(query.fieldbyname('lotto').asstring);
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(mid(lot_codice, 4, 27)) lotto from opt where left(lot_codice, 2) = ' + quotedstr(copy(result, 1, 2));
      query.open;
      try
        numero_opt := strtoint(query.fieldbyname('lotto').asstring);
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 5, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.progressivo' then
    begin
      result := formatdatetime('yyyy', data);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 5) = ' + quotedstr(result + '.');
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 8));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 5) = ' + quotedstr(result + '.');
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 8));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + '.' + setta_lunghezza(numero, 8, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.prg' then
    begin
      result := formatdatetime('yyyy', data);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 5) = ' + quotedstr(result + '.');
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 3));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 5) = ' + quotedstr(result + '.');
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 3));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + '.' + setta_lunghezza(numero, 3, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'fornitore.documento.anno' then
    begin
      result := frn_codice + '.' + floattostr(numero_documento) + '.' + formatdatetime('yyyy', data_documento);
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'anno.mese' then
    begin
      result := formatdatetime('yyyy', data) + '.' + formatdatetime('mm', data);
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'cm__aaprogressivo' then
    begin
      read_tabella(arcdit, 'art', 'codice', art_codice, 'tcm_codice');
      result := stringreplace(copy(setta_lunghezza(archivio.fieldbyname('tcm_codice').asstring, 4) + formatdatetime('yy', data), 1, 6), ' ', '_', [rfreplaceall]);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 6) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 6) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 8, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'gm__aaprogressivo' then
    begin
      read_tabella(arcdit, 'art', 'codice', art_codice, 'tgm_codice');
      result := stringreplace(copy(setta_lunghezza(archivio.fieldbyname('tgm_codice').asstring, 4) + formatdatetime('yy', data), 1, 6), ' ', '_', [rfreplaceall]);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 6) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 6) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 8, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'sa__aaprogressivo' then
    begin
      read_tabella(arcdit, 'art', 'codice', art_codice, 'tsa_codice');
      result := stringreplace(copy(setta_lunghezza(archivio.fieldbyname('tsa_codice').asstring, 4) + formatdatetime('yy', data), 1, 6), ' ', '_', [rfreplaceall]);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 6) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 6) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 7, 8));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 8, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'inaammprogressivo' then
    begin
      read_tabella(arcdit, 'art', 'codice', art_codice, 'tin_codice');
      result := stringreplace(copy(setta_lunghezza(archivio.fieldbyname('tin_codice').asstring, 4) + formatdatetime('yymm', data), 1, 8), ' ', '_', [rfreplaceall]);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 8) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 9, 6));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 8) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 9, 6));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 6, 0);

      query.free;
    end
    else if dit.fieldbyname('tipo_codice_lotto_automatico').asstring = 'tin-progressivo' then
    begin
      read_tabella(arcdit, 'art', 'codice', art_codice, 'tin_codice');
      result := stringreplace(copy(setta_lunghezza(archivio.fieldbyname('tin_codice').asstring, 4) + '-', 1, 5), ' ', '_', [rfreplaceall]);

      query := tmyquery_go.create(nil);
      query.connection := arcdit;
      query.sql.text := 'select max(lotto) lotto from lot where left(lotto, 5) = ' + quotedstr(result);
      query.open;
      if query.fieldbyname('lotto').value = null then
      begin
        numero := 0;
      end
      else
      begin
        try
          numero := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 6));
        except
          numero := 0;
        end;
      end;

      query.Close;
      query.sql.text := 'select max(lot_codice) lotto from opt where left(lot_codice, 5) = ' + quotedstr(result);
      query.open;
      try
        numero_opt := strtoint(copy(query.fieldbyname('lotto').asstring, 6, 6));
      except
        numero_opt := 0;
      end;
      if numero_opt > numero then
      begin
        numero := numero_opt;
      end;

      inc(numero);
      result := result + setta_lunghezza(numero, 6, 0);

      query.free;
    end;

    if result <> '' then
    begin
      if dit.fieldbyname('case_lotti').asstring = 'maiuscolo' then
      begin
        result := uppercase(result);
      end
      else if dit.fieldbyname('case_lotti').asstring = 'minuscolo' then
      begin
        result := lowercase(result);
      end;
    end;
  end;
end;

function TARC.data_database(data: tdate): string;
begin
  result := quotedstr(formatdatetime('yyyy/mm/dd', data));
end;

procedure TARC.crea_ltm_lettore(art_codice, lot_codice, tma_codice, quantita, data_scadenza, documento_origine, esistenza, cfg_tipo, cfg_codice, serie_documento: string; progressivo, riga: Integer; numero_documento: double; data_registrazione, data_documento: tdate);
var
  aggiornato: Boolean;
  lot, ltm: tmyquery_go;
begin
  lot := tmyquery_go.create(nil);
  lot.connection := arcdit;
  lot.sql.text := 'select * from lot where art_codice = ' + quotedstr(art_codice) + ' and lotto = ' + quotedstr(lot_codice);
  lot.open;

  ltm := tmyquery_go.create(nil);
  ltm.connection := arcdit;
  ltm.sql.add('insert into ltm (');
  ltm.sql.add('progressivo, art_codice, lotto, tma_codice, doc_progressivo_origine,');
  ltm.sql.add('data_registrazione, quantita, esistenza, documento_origine, doc_riga_origine,');
  ltm.sql.add('quantita_entrate, quantita_uscite, cfg_tipo, cfg_codice, data_documento,');
  ltm.sql.add('serie_documento, numero_documento');
  ltm.sql.add(') values (');
  ltm.sql.add(':progressivo, :art_codice, :lotto, :tma_codice, :doc_progressivo_origine,');
  ltm.sql.add(':data_registrazione, :quantita, :esistenza, :documento_origine, :doc_riga_origine,');
  ltm.sql.add(':quantita_entrate, :quantita_uscite, :cfg_tipo, :cfg_codice, :data_documento,');
  ltm.sql.add(':serie_documento, :numero_documento');
  ltm.sql.add(')');

  ltm.parambyname('progressivo').asinteger := setta_valore_generatore(arcdit, 'ltm_progressivo');
  ltm.parambyname('art_codice').asstring := art_codice;
  ltm.parambyname('lotto').asstring := lot_codice;
  ltm.parambyname('tma_codice').asstring := tma_codice;
  ltm.parambyname('documento_origine').asstring := documento_origine;
  ltm.parambyname('doc_progressivo_origine').asinteger := progressivo;
  ltm.parambyname('doc_riga_origine').asinteger := riga;
  ltm.parambyname('data_registrazione').asdatetime := data_registrazione;
  ltm.parambyname('quantita').asfloat := strtofloat(quantita);
  ltm.parambyname('esistenza').asstring := esistenza;
  ltm.parambyname('cfg_tipo').asstring := cfg_tipo;
  ltm.parambyname('cfg_codice').asstring := cfg_codice;
  ltm.parambyname('data_documento').asdatetime := data_documento;
  ltm.parambyname('serie_documento').asstring := serie_documento;
  ltm.parambyname('numero_documento').asfloat := numero_documento;

  ltm.execsql;

  if lot.eof and (data_scadenza <> '') then
  begin
    aggiornato := false;
    while not aggiornato do
    begin
      lot.Close;
      lot.open;
      if not lot.isempty then
      begin
        lot.edit;
        lot.fieldbyname('data_scadenza').asdatetime := strtodate(data_scadenza);
        lot.post;
        aggiornato := true;
      end;
    end;
  end;

  ltm.free;
  lot.free;
end;

procedure TARC.controllo_prezzo_costo(art_codice: string; Importo, quantita: double);
var
  query: tmyquery_go;
  costo: double;
begin
  query := tmyquery_go.create(nil);
  query.connection := arcdit;
  query.sql.text := 'select mag.prezzo_carico from mag where mag.art_codice = ' + quotedstr(art_codice) + ' ' + 'order by mag.data_carico desc limit 1';
  query.open;
  costo := query.fieldbyname('prezzo_carico').asfloat;
  try
    if (costo <> 0) and (Importo / quantita < costo) then
    begin
      messaggio(200, 'il prezzo netto di vendita dell''articolo [' + art_codice + ']: ' + Formatfloat(formato_display_prezzo, Importo / quantita) + #13 + 'è inferiore all''ultimo prezzo netto di acquisto: ' + Formatfloat(formato_display_prezzo_acq, costo));
    end;
  except
  end;
  query.free;
end;

function TARC.puntino(numero: double; numero_decimali: Integer = 2): string;
var
  i: word;
  stringa: string;
begin
  result := '';

  // stringa := floattostr(numero);
  if numero_decimali = 0 then
  begin
    stringa := Formatfloat('0;-0;0', numero);
  end
  else if numero_decimali = 1 then
  begin
    stringa := Formatfloat('0.0;-0.0;0.0', numero);
  end
  else if numero_decimali = 2 then
  begin
    stringa := Formatfloat('0.00;-0.00;0.00', numero);
  end
  else if numero_decimali = 3 then
  begin
    stringa := Formatfloat('0.000;-0.000;0.000', numero);
  end
  else if numero_decimali = 4 then
  begin
    stringa := Formatfloat('0.0000;-0.0000;0.0000', numero);
  end
  else if numero_decimali = 5 then
  begin
    stringa := Formatfloat('0.00000;-0.00000;0.00000', numero);
  end
  else if numero_decimali = 6 then
  begin
    stringa := Formatfloat('0.000000;-0.000000;0.000000', numero);
  end;

  for i := 1 to length(stringa) do
  begin
    if stringa[i] = ',' then
    begin
      result := result + '.';
    end
    else
    begin
      result := result + stringa[i];
    end;
  end;
end;

function URLEncode(const s: AnsiString): AnsiString;
const
  NoConversion = ['A' .. 'Z', 'a' .. 'z', 'á', '*', '@', '.', '_', '-', '/', ':', '=', '?'];
var
  i, idx, Len: Integer;

  function DigitToHex(Digit: Integer): AnsiChar;
  begin
    case Digit of
      0 .. 9:
        result := AnsiChar(Chr(Digit + Ord('0')));
      10 .. 15:
        result := AnsiChar(Chr(Digit - 10 + Ord('A')));
    else
      result := '0';
    end;
  end;

begin
  Len := 0;
  for i := 1 to length(s) do
    if s[i] in NoConversion then
      Len := Len + 1
    else
      Len := Len + 3;
  SetLength(result, Len);
  idx := 1;
  for i := 1 to length(s) do
    if s[i] in NoConversion then
    begin
      result[idx] := s[i];
      idx := idx + 1;
    end
    else
    begin
      result[idx] := '%';
      result[idx + 1] := DigitToHex(Ord(s[i]) div 16);
      result[idx + 2] := DigitToHex(Ord(s[i]) mod 16);
      idx := idx + 3;
    end;
end;

(*
  function TARC.traduci(const parola, sourcelng, destlng: string): string;
  var
  risposta: string;
  doc, traduzioni: TJSONArray;
  elemento: TJSONValue;
  http: tidhttp;
  json_da_inviare: TStringStream;
  const
  // chiave azure gestibile da http://portal.azure.com
  chiave_azure = 'f8c7fefc670048e5a486aedce5551723';
  request_url = 'https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=%s&to=%s';
  begin
  result := '';
  http := tidhttp.create(nil);

  json_da_inviare := TStringStream.Create('[{"Text": "' + parola.replace('\', '\u005c').replace('"', '\u201c')
  + '"}]', TEncoding.UTF8);

  try
  http.Request.useragent := 'Mozilla/5.0';
  http.Request.accept := 'application/json';
  http.Request.ContentType := 'application/json';
  http.Request.CustomHeaders.AddValue('Ocp-Apim-Subscription-Key', chiave_azure);

  risposta := http.post(format(request_url, [sourcelng, destlng]), json_da_inviare);
  doc := TJSONObject.ParseJSONValue(risposta) as TJSONArray;
  for elemento in doc do
  begin
  traduzioni := elemento.getvalue<TJSONArray>('translations');
  end;
  for elemento in traduzioni do
  begin
  result := elemento.GetValue<string>('text');
  end;

  finally
  http.free;
  json_da_inviare.free;
  end;
  end;
*)

function TARC.traduci(const parola, sourcelng, destlng: string): string;
var
  risposta: string;
  doc, traduzioni: TJSONArray;
  elemento: TJSONValue;
  http: tidhttp;
  json_da_inviare: TStringStream;
  handleropenssl: TIdSSLIOHandlerSocketOpenSSL;
const
  // chiave azure gestibile da http://portal.azure.com
  chiave_azure = 'f8c7fefc670048e5a486aedce5551723';
  request_url = 'https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&from=%s&to=%s';
begin
  result := '';
  http := tidhttp.create(nil);
  handleropenssl := TIdSSLIOHandlerSocketOpenSSL.create;
  handleropenssl.ssloptions.sslversions := [sslvsslv23, sslvtlsv1, sslvtlsv1_1, sslvtlsv1_2];
  http.iohandler := handleropenssl;

  json_da_inviare := TStringStream.create('[{"Text": "' + parola.replace('\', '\u005c').replace('"', '\u201c') + '"}]', TEncoding.UTF8);

  try
    http.Request.useragent := 'Mozilla/5.0';
    http.Request.accept := 'application/json';
    http.Request.ContentType := 'application/json';
    http.Request.CustomHeaders.AddValue('Ocp-Apim-Subscription-Key', chiave_azure);

    risposta := http.post(format(request_url, [sourcelng, destlng]), json_da_inviare);
    doc := TJSONObject.ParseJSONValue(risposta) as TJSONArray;
    for elemento in doc do
    begin
      traduzioni := elemento.getvalue<TJSONArray>('translations');
    end;
    for elemento in traduzioni do
    begin
      result := elemento.getvalue<string>('text');
    end;

  finally
    http.free;
    json_da_inviare.free;
    handleropenssl.free;
  end;
end;

procedure TARC.WinInet_HttpGet(const Url: string; Stream: TStream);
const
  BuffSize = 1024 * 1024;
var
  hInter: HINTERNET;
  UrlHandle: HINTERNET;
  BytesRead: DWORD;
  Buffer: Pointer;
begin
  hInter := InternetOpen('', INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);
  if assigned(hInter) then
    try
      Stream.Seek(0, 0);
      GetMem(Buffer, BuffSize);
      try
        UrlHandle := InternetOpenUrl(hInter, pchar(Url), nil, 0, INTERNET_FLAG_RELOAD, 0);
        if assigned(UrlHandle) then
        begin
          repeat
            InternetReadFile(UrlHandle, Buffer, BuffSize, BytesRead);
            if BytesRead > 0 then
              Stream.WriteBuffer(Buffer^, BytesRead);
          until BytesRead = 0;
          InternetCloseHandle(UrlHandle);
        end;
      finally
        FreeMem(Buffer);
      end;
    finally
      InternetCloseHandle(hInter);
    end;
end;

function TARC.WinInet_HttpGet(const Url: string): string;
var
  StringStream: TStringStream;
begin
  result := '';
  StringStream := TStringStream.create('');
  try
    WinInet_HttpGet(Url, StringStream);
    if StringStream.Size > 0 then
    begin
      StringStream.Seek(0, 0);
      result := StringStream.readstring(StringStream.Size);
    end;
  finally
    StringStream.free;
  end;
end;

procedure TARC.traduzione(data_set: tmyquery_go; parola, campo_01, campo_02, campo_03, campo_04, campo_05: string);
var
  testo: string;
begin
  if (lin.fieldbyname('lingua_01_google').asstring <> '') and (parola <> '') then
  begin
    testo := traduci(parola, 'it', lin.fieldbyname('lingua_01_google').asstring);
    data_set.fieldbyname(campo_01).asstring := testo;
  end;
  if (lin.fieldbyname('lingua_02_google').asstring <> '') and (parola <> '') then
  begin
    testo := traduci(parola, 'it', lin.fieldbyname('lingua_02_google').asstring);
    data_set.fieldbyname(campo_02).asstring := testo;
  end;
  if (lin.fieldbyname('lingua_03_google').asstring <> '') and (parola <> '') then
  begin
    testo := traduci(parola, 'it', lin.fieldbyname('lingua_03_google').asstring);
    data_set.fieldbyname(campo_03).asstring := testo;
  end;
  if (lin.fieldbyname('lingua_04_google').asstring <> '') and (parola <> '') then
  begin
    testo := traduci(parola, 'it', lin.fieldbyname('lingua_04_google').asstring);
    data_set.fieldbyname(campo_04).asstring := testo;
  end;
  if (lin.fieldbyname('lingua_05_google').asstring <> '') and (parola <> '') then
  begin
    testo := traduci(parola, 'it', lin.fieldbyname('lingua_05_google').asstring);
    data_set.fieldbyname(campo_05).asstring := testo;
  end;
end;

procedure TARC.traduzione_cinese(data_set: tmyquery_go; parola, campo: string);
begin
  if parola <> '' then
  begin
    tabella_edit(data_set);
    data_set.fieldbyname(campo).asstring := traduci(parola, 'it', 'ZH');
  end;
end;

function TARC.assegna_fine_mese(anno, mese: Integer): tdatetime;
var
  giorno: word;
begin
  giorno := 31;
  if not tryencodedate(anno, mese, giorno, result) then
  begin
    giorno := 30;
    if not tryencodedate(anno, mese, giorno, result) then
    begin
      giorno := 29;
      if not tryencodedate(anno, mese, giorno, result) then
      begin
        giorno := 28;
        tryencodedate(anno, mese, giorno, result);
      end;
    end;
  end;
end;

function TARC.normalizza_codice(codice: string; normalizza: Boolean = false): string;
begin
  result := codice;
  if normalizza or (dit.fieldbyname('normalizza_art_codice').asstring = 'si') then
  begin
    result := stringreplace(result, '/', '_', [rfreplaceall]);
    result := stringreplace(result, '\', '_', [rfreplaceall]);
    result := stringreplace(result, ':', '_', [rfreplaceall]);
    result := stringreplace(result, '*', '_', [rfreplaceall]);
    result := stringreplace(result, '?', '_', [rfreplaceall]);
    result := stringreplace(result, '"', '_', [rfreplaceall]);
    result := stringreplace(result, '>', '_', [rfreplaceall]);
    result := stringreplace(result, '<', '_', [rfreplaceall]);
    result := stringreplace(result, '|', '_', [rfreplaceall]);
  end;
end;

function TARC.normalizza_documento(numero_documento: string): string;
var
  i: word;
begin
  result := '';
  for i := 1 to length(numero_documento) do
  begin
    if (numero_documento[i] >= '0') and (numero_documento[i] <= '9') then
    begin
      result := result + numero_documento[i];
    end;
  end;
  if length(result) > 6 then
  begin
    result := copy(result, 1, 6);
  end
  else
  begin
    result := setta_lunghezza(result, 6, true, '0');
  end;
end;

function TARC.cerca_campo_csv(numero_campo: word; sorgente: string; separatore: string = ';'): string;
var
  ind: word;
begin
  result := '';
  ind := 0;
  while pos(separatore, sorgente) > 0 do
  begin
    inc(ind);
    if ind = numero_campo then
    begin
      result := trim(copy(sorgente, 1, pos(separatore, sorgente) - 1));
      break;
    end
    else
    begin
      sorgente := copy(sorgente, pos(separatore, sorgente) + 1, length(sorgente));
    end;
  end;
  if (result = '') and (trim(sorgente) <> '') then
  begin
    inc(ind);
    if ind = numero_campo then
    begin
      result := trim(sorgente);
    end;
  end;
end;

function TARC.invia_messaggio(pec: Boolean; oggetto, conoscenza, messaggio_testo: string; var lista: string; allegati: tstringlist; user_host, user_id, user_password, user_mail: string; porta_smtp: Integer; num_img_html: Integer; conoscenza_ccn: string = ''; no_tls: string = '';
  protocollo_tls: string = ''): Boolean;
var
  i: word;
  file_tobit: textfile;
  nome_file_tobit, testo_tobit, stringa, str: string;
  prosegui: Boolean;
begin
  result := false;
  stringa := '';

  if dit.fieldbyname('collegamento_tobit').asstring = 'si' then
  begin
    if not directoryexists(cartella_file + '\tobit') then
    begin
      createdir(cartella_file + '\tobit');
    end;

    nome_file_tobit := cartella_file + '\tobit\' + normalizza_codice(oggetto, true) + formatdatetime('_yyyy_mm_dd_hh_mm_ss', now) + '.EML';

    AssignFile(file_tobit, nome_file_tobit);
    Rewrite(file_tobit);

    WriteLn(file_tobit, '@@HTML@@');

    WriteLn(file_tobit, formattazione_html_tobit(messaggio_testo));

    if (utn.fieldbyname('firma_email').asstring <> '') then
    begin
      if (containsstr(utn.fieldbyname('firma_email').asstring, '<div>')) or (containsstr(utn.fieldbyname('firma_email').asstring, '<DIV>')) then
      begin
        WriteLn(file_tobit, utn.fieldbyname('firma_email').asstring);
      end
      else
      begin
        WriteLn(file_tobit, formattazione_html_tobit(utn.fieldbyname('firma_email').asstring));
      end;
    end;

    WriteLn(file_tobit, '@@ANSI@@');
    WriteLn(file_tobit, '@@EMAIL@@');
    // writeln(file_tobit, messaggio_testo);

    if pec then
    begin
      if user_host = '' then
      begin
        testo_tobit := utn.fieldbyname('user_e_mail_pec').asstring;
      end
      else
      begin
        testo_tobit := user_mail;
      end;
    end
    else
    begin
      if user_host = '' then
      begin
        testo_tobit := utn.fieldbyname('user_e_mail').asstring;
      end
      else
      begin
        testo_tobit := user_mail;
      end;
    end;
    WriteLn(file_tobit, '@@FROM ' + testo_tobit + '@@');

    stringa := lista;
    while pos(';', stringa) > 0 do
    begin
      testo_tobit := copy(stringa, 1, pos(';', stringa) - 1);
      WriteLn(file_tobit, '@@NUMBERLIST ' + testo_tobit + '@@');
      // writeln(file_tobit, '@@NUMBERLIST ;@@');
      stringa := trim(copy(stringa, pos(';', stringa) + 1, length(stringa)));
    end;
    if stringa <> '' then
    begin
      WriteLn(file_tobit, '@@NUMBERLIST ' + stringa + '@@');
      // writeln(file_tobit, '@@NUMBERLIST ;@@');
    end;

    if conoscenza <> '' then
    begin
      WriteLn(file_tobit, '@@NUMBERLIST ' + conoscenza + '@@');
      // writeln(file_tobit, '@@NUMBERLIST ;@@');
    end;

    if conoscenza_ccn <> '' then
    begin
      WriteLn(file_tobit, '@@NUMBERLIST ' + conoscenza_ccn + '@@');
      // writeln(file_tobit, '@@NUMBERLIST ;@@');
    end;

    if oggetto <> '' then
    begin
      WriteLn(file_tobit, '@@SUBJECT ' + oggetto + '@@');
    end;

    if allegati.count > 0 then
    begin
      for i := 0 to allegati.count - 1 do
      begin
        if fileexists(allegati[i]) then
        begin
          WriteLn(file_tobit, '@@ATTACH ' + allegati[i] + ', Documento allegato, del@@');
        end
        else
        begin
          messaggio(200, 'l''allegato [' + allegati[i] + '] non esiste');
        end;
      end;
    end;

    closefile(file_tobit);
  end
  else
  begin
    v_mail.clear;
    v_mail.charset := 'utf-8';

    // setup smtp
    if pec then
    begin
      if user_host = '' then
      begin
        server_smtp.host := utn.fieldbyname('user_host_pec').asstring;
        server_smtp.usetls := utUseImplicitTLS;
        server_smtp.port := utn.fieldbyname('porta_smtp_pec').asinteger;
        server_smtp.username := utn.fieldbyname('user_id_pec').asstring;
        server_smtp.password := utn.fieldbyname('user_password_pec').asstring;
      end
      else
      begin
        server_smtp.host := user_host;
        server_smtp.usetls := utUseImplicitTLS;
        server_smtp.port := porta_smtp;
        server_smtp.username := user_id;
        server_smtp.password := user_password;
      end;
    end
    else
    begin
      if user_host = '' then
      begin
        server_smtp.host := utn.fieldbyname('user_host').asstring;
        server_smtp.usetls := utuseexplicittls;
        if utn.fieldbyname('porta_smtp').asinteger > 0 then
        begin
          server_smtp.port := utn.fieldbyname('porta_smtp').asinteger;
        end
        else
        begin
          server_smtp.port := 25;
        end;
        server_smtp.username := utn.fieldbyname('user_id').asstring;
        server_smtp.password := utn.fieldbyname('user_password').asstring;
      end
      else
      begin
        server_smtp.host := user_host;
        server_smtp.usetls := utuseexplicittls;
        if porta_smtp > 0 then
        begin
          server_smtp.port := porta_smtp;
        end
        else
        begin
          server_smtp.port := 25;
        end;
        server_smtp.username := user_id;
        server_smtp.password := user_password;
      end;
    end;

    if server_smtp.port = 465 then
    begin
      server_smtp.usetls := utUseImplicitTLS;
    end;

    // setup ssl
    open_ssl.port := server_smtp.port;
    if pec then
    begin
      if user_host = '' then
      begin
        open_ssl.host := utn.fieldbyname('user_host_pec').asstring;
      end
      else
      begin
        open_ssl.host := user_host;
      end;
      open_ssl.ssloptions.Method := sslvsslv23;
    end
    else
    begin
      if user_host = '' then
      begin
        open_ssl.host := utn.fieldbyname('user_host').asstring;
      end
      else
      begin
        open_ssl.host := user_host;
      end;

      (*
        if no_tls = '' then
        begin
        no_tls := utn.fieldbyname('no_tls').asstring;
        end;
        if no_tls = 'si' then
        begin
        server_smtp.usetls := utNoTLSSupport;
        end;

        if usa_tlsv1 = '' then
        begin
        usa_tlsv1 := utn.fieldbyname('usa_tlsv1').asstring;
        end;
        if usa_tlsv1 = 'si' then
        begin
        open_ssl.SSLOptions.Method := sslvTLSv1;
        end
        else
        begin
        open_ssl.SSLOptions.Method := sslvSSLv3;
        end;
 *)

      // controllo se passati valori vecchi di usa_tlsv1, ma non dovrebbe mai succedere
      if protocollo_tls = 'no' then
      begin
        protocollo_tls := 'v3';
      end;
      if protocollo_tls = 'si' then
      begin
        protocollo_tls := 'v1';
      end;
      if no_tls = 'si' then
      begin
        protocollo_tls := 'nessuno';
      end;

      // valore non passato
      if protocollo_tls = '' then
      begin
        protocollo_tls := utn.fieldbyname('protocollo_tls').asstring;
      end;

      if protocollo_tls = 'nessuno' then
      begin
        server_smtp.usetls := utNoTLSSupport;
      end
      else if protocollo_tls = 'v1' then
      begin
        open_ssl.ssloptions.Method := sslvtlsv1;
      end
      else if protocollo_tls = 'v1_1' then
      begin
        open_ssl.ssloptions.Method := sslvtlsv1_1;
      end
      else if protocollo_tls = 'v1_2' then
      begin
        open_ssl.ssloptions.Method := sslvtlsv1_2;
      end
      else if protocollo_tls = 'v2' then
      begin
        open_ssl.ssloptions.Method := sslvSSLv2;
      end
      else if protocollo_tls = 'v23' then
      begin
        open_ssl.ssloptions.Method := sslvsslv23;
      end
      else if protocollo_tls = 'v3' then
      begin
        open_ssl.ssloptions.Method := sslvSSLv3;
      end;
    end;

    // mittente
    if pec then
    begin
      if user_host = '' then
      begin
        v_mail.from.address := utn.fieldbyname('user_e_mail_pec').asstring;
        v_mail.Sender.address := utn.fieldbyname('user_e_mail_pec').asstring;
      end
      else
      begin
        v_mail.from.address := user_mail;
        v_mail.Sender.address := user_mail;
      end;
    end
    else
    begin
      if user_host = '' then
      begin
        v_mail.from.address := utn.fieldbyname('user_e_mail').asstring;
        if utn.fieldbyname('descrizione_utente_mail').asstring = 'si' then
        begin
          v_mail.from.address := '"' + utn.fieldbyname('descrizione').asstring + '" <' + v_mail.from.address + '>';
        end;
        v_mail.Sender.address := utn.fieldbyname('user_e_mail').asstring;
      end
      else
      begin
        v_mail.from.address := user_mail;
        v_mail.Sender.address := user_mail;
      end;
    end;

    // verifica correttezza email
    stringa := lista;
    lista := '';
    while pos(';', stringa) > 0 do
    begin
      str := copy(stringa, 1, pos(';', stringa) - 1);
      if not valida_email(str) then
      begin
        result := true;
        messaggio(200, 'mail [' + str + '] non corretta');
      end
      else
      begin
        if lista <> '' then
        begin
          lista := lista + ';' + str;
        end
        else
        begin
          lista := str;
        end;
      end;
      stringa := trim(copy(stringa, pos(';', stringa) + 1, length(stringa)));
    end;
    if stringa <> '' then
    begin
      str := stringa;
      if not valida_email(str) then
      begin
        result := true;
        messaggio(200, 'mail [' + str + '] non corretta');
      end
      else
      begin
        if lista <> '' then
        begin
          lista := lista + ';' + str;
        end
        else
        begin
          lista := str;
        end;
      end;
    end;

    if lista <> '' then
    begin
      // destinatario
      if pec then
      begin
        v_mail.recipients.emailaddresses := lista;
        v_mail.replyto.emailaddresses := '';
      end
      else
      begin
        if pos(',', lista) = 0 then
        begin
          v_mail.recipients.emailaddresses := lista;
          if user_host = '' then
          begin
            v_mail.replyto.emailaddresses := utn.fieldbyname('user_e_mail').asstring;
            if utn.fieldbyname('notifica_mail_inviate').asstring = 'si' then
            begin
              v_mail.bcclist.emailaddresses := utn.fieldbyname('user_e_mail').asstring;
            end;
          end
          else
          begin
            v_mail.replyto.emailaddresses := user_mail;
            if utn.fieldbyname('notifica_mail_inviate').asstring = 'si' then
            begin
              v_mail.bcclist.emailaddresses := user_mail;
            end;
          end;
        end
        else
        begin
          v_mail.bcclist.emailaddresses := lista;
          if user_host = '' then
          begin
            v_mail.recipients.emailaddresses := utn.fieldbyname('user_e_mail').asstring;
            v_mail.replyto.emailaddresses := utn.fieldbyname('user_e_mail').asstring;
          end
          else
          begin
            v_mail.recipients.emailaddresses := user_mail;
            v_mail.replyto.emailaddresses := user_mail;
          end;
        end;
      end;

      // conoscenza
      if pec or (conoscenza = '') then
      begin
        //
      end
      else
      begin
        v_mail.cclist.emailaddresses := conoscenza;
      end;

      // conoscenza bccl
      if pec or (conoscenza_ccn = '') then
      begin
        //
      end
      else
      begin
        if v_mail.bcclist.emailaddresses = '' then
        begin
          v_mail.bcclist.emailaddresses := conoscenza_ccn;
        end
        else
        begin
          v_mail.bcclist.emailaddresses := v_mail.bcclist.emailaddresses + ';' + conoscenza_ccn;
        end;
      end;

      // messaggio
      v_mail.subject := oggetto;
      v_mail.body.text := messaggio_testo;

      if ContainsText(v_mail.body.text, 'This is a multi-part message in MIME format') then
      begin
        v_mail.ExtraHeaders.add('MIME-Version: 1.0');
        // gestione allegati per messaggi HTML
        if allegati.count > 0 then
        begin
          AddMainContentType(v_mail.ExtraHeaders, num_img_html, true);
        end
        else
        begin
          AddMainContentType(v_mail.ExtraHeaders, num_img_html, false);
        end;
      end
      else
      begin
        // allegati per messaggi "normali"
        if allegati.count > 0 then
        begin
          for i := 0 to allegati.count - 1 do
          begin
            if fileexists(allegati[i]) then
            begin
              prosegui := false;
              while not prosegui do
              begin
                if IsFileInUse(allegati[i]) then
                begin
                  if messaggio(304, 'il file [' + allegati[i] + '] è in uso da un altro programma' + slinebreak + slinebreak + 'escludi il file e prosegui (oppure libera il file)') = 1 then
                  begin
                    prosegui := true;
                  end;
                end
                else
                begin
                  tidattachmentfile.create(v_mail.messageparts, allegati[i]);
                  prosegui := true;
                end;
              end;
            end
            else
            begin
              messaggio(200, 'l''allegato [' + allegati[i] + '] non esiste');
            end;
          end;
        end;
      end;

      // invio
      try
        try
          server_smtp.connect;
          server_smtp.send(v_mail);
        except
          on e: exception do
          begin
            result := true;
            messaggio(200, 'mail all''indirizzo: ' + lista + ' non inviata' + slinebreak + slinebreak + 'causa del problema' + slinebreak + e.message);
          end;
        end;
      finally
        server_smtp.disconnect;
      end;
    end;
    v_mail.messageparts.clear;
  end;
end;

function TARC.IsFileInUse(FileName: TFileName): Boolean;
var
  HFileRes: HFILE;
begin
  result := false;
  if not fileexists(FileName) then
    exit;
  HFileRes := CreateFile(pchar(FileName), GENERIC_READ or GENERIC_WRITE, 0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
  result := (HFileRes = INVALID_HANDLE_VALUE);
  if not result then
    CloseHandle(HFileRes);
end;

function TARC.calcola_ricarico(trl_codice: string; prezzo: double; decimali: Integer; superiore: Boolean = false): double;
var
  trl: tmyquery_go;
begin
  trl := tmyquery_go.create(nil);
  trl.connection := arcdit;
  trl.sql.text := 'select * from trl where codice = ' + quotedstr(trl_codice);
  trl.open;
  if not trl.isempty then
  begin
    prezzo := prezzo * (1 + trl.fieldbyname('ricarico').asfloat / 100);

    if prezzo <= trl.fieldbyname('limite_01').asfloat then
    begin
      if trl.fieldbyname('valore_01').asfloat <> 0 then
      begin
        if superiore then
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_01').asfloat, 0, 2);
        end
        else
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_01').asfloat, 0);
        end;
        prezzo := arrotonda(prezzo * trl.fieldbyname('valore_01').asfloat, decimali);
      end;
    end
    else if prezzo <= trl.fieldbyname('limite_02').asfloat then
    begin
      if trl.fieldbyname('valore_02').asfloat <> 0 then
      begin
        if superiore then
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_02').asfloat, 0, 2);
        end
        else
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_02').asfloat, 0);
        end;
        prezzo := arrotonda(prezzo * trl.fieldbyname('valore_02').asfloat, decimali);
      end;
    end
    else if prezzo <= trl.fieldbyname('limite_03').asfloat then
    begin
      if trl.fieldbyname('valore_03').asfloat <> 0 then
      begin
        if superiore then
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_03').asfloat, 0, 2);
        end
        else
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_03').asfloat, 0);
        end;
        prezzo := arrotonda(prezzo * trl.fieldbyname('valore_03').asfloat, decimali);
      end;
    end
    else if prezzo <= trl.fieldbyname('limite_04').asfloat then
    begin
      if trl.fieldbyname('valore_04').asfloat <> 0 then
      begin
        if superiore then
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_04').asfloat, 0, 2);
        end
        else
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_04').asfloat, 0);
        end;
        prezzo := arrotonda(prezzo * trl.fieldbyname('valore_04').asfloat, decimali);
      end;
    end
    else if prezzo <= trl.fieldbyname('limite_05').asfloat then
    begin
      if trl.fieldbyname('valore_05').asfloat <> 0 then
      begin
        if superiore then
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_05').asfloat, 0, 2);
        end
        else
        begin
          prezzo := arrotonda(prezzo / trl.fieldbyname('valore_05').asfloat, 0);
        end;
        prezzo := arrotonda(prezzo * trl.fieldbyname('valore_05').asfloat, decimali);
      end;
    end
    else
    begin
      if superiore then
      begin
        prezzo := arrotonda(prezzo, decimali, 2);
      end
      else
      begin
        prezzo := arrotonda(prezzo, decimali);
      end;
    end;
  end;
  trl.free;

  result := prezzo;
end;

procedure TARC.chiamata_telefono(numero: string);
var
  comando, parametro, interno: string;
  username, password: string;
  i: Integer;
begin
  // campi utente
  if utn.fieldbyname('comando_voip').asstring <> '' then
  begin
    comando := trim(utn.fieldbyname('comando_voip').asstring);
    parametro := trim(utn.fieldbyname('parametro_voip').asstring);
  end
  else
  begin
    comando := trim(dit.fieldbyname('comando_voip').asstring);
    parametro := trim(dit.fieldbyname('parametro_voip').asstring);
  end;
  username := trim(utn.fieldbyname('user_id_voip').asstring);
  password := trim(utn.fieldbyname('user_password_voip').asstring);

  // separazione dell'interno telefonico
  i := pos(':', numero);
  if (i > 0) then
  begin
    interno := cancella_escluso(copy(numero, i + 1), ['0' .. '9', '#', '*']);
    if (interno = '') then
    begin
      messaggio(000, 'numero dell''interno mancante o errato');
    end;
    numero := copy(numero, 1, i - 1);
  end
  else
  begin
    i := pos(',', numero);
    if (i > 0) then
    begin
      interno := cancella_escluso(copy(numero, i + 1), ['0' .. '9', '#', '*']);
      if (interno = '') then
      begin
        messaggio(000, 'numero dell''interno mancante o errato');
      end;
      numero := copy(numero, 1, i - 1);
    end;
  end;

  numero := cancella_escluso(numero, ['0' .. '9', '+', '#', '*']);
  if (numero = '') then
  begin
    messaggio(000, 'numero di telefono mancante o errato');
    exit;
  end;

  // parametri in parametro_voip
  if (pos('%PRE_IT%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%PRE_IT%', '', [rfIgnoreCase]);
    // aggiunta prefisso internazionale italiano
    if (pos('+', numero) <> 1) and (pos('00', numero) <> 1) then
    begin
      numero := '+39' + numero;
    end;
  end;

  if (pos('%NUMERO%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%NUMERO%', numero, [rfIgnoreCase]);
  end
  else
  begin
    parametro := parametro + numero;
  end;

  if (pos('%INTERNO%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%INTERNO%', interno, [rfIgnoreCase]);
  end;

  if (pos('%USERNAME%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%USERNAME%', username, [rfIgnoreCase]);
  end;

  if (pos('%PASSWORD%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%PASSWORD%', password, [rfIgnoreCase]);
  end;

  if (pos('%PARAMETRO%', uppercase(parametro)) > 0) then
  begin
    parametro := stringreplace(parametro, '%PARAMETRO%', '', [rfIgnoreCase]);
  end;

  // parametri in comando_voip
  if (pos('%PRE_IT%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%PRE_IT%', '', [rfIgnoreCase]);
    // aggiunta prefisso internazionale italiano
    if (pos('+', numero) <> 1) and (pos('00', numero) <> 1) then
    begin
      numero := '+39' + numero;
    end;
  end;

  if (pos('%NUMERO%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%NUMERO%', numero, [rfIgnoreCase]);
  end;

  if (pos('%INTERNO%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%INTERNO%', interno, [rfIgnoreCase]);
  end;

  if (pos('%USERNAME%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%USERNAME%', username, [rfIgnoreCase]);
  end;

  if (pos('%PASSWORD%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%PASSWORD%', password, [rfIgnoreCase]);
  end;

  if (pos('%PARAMETRO%', uppercase(comando)) > 0) then
  begin
    comando := stringreplace(comando, '%PARAMETRO%', parametro, [rfIgnoreCase]);
    parametro := '';
  end;

  // esecuzione
  esegui_effettivo(comando, parametro);
end;

procedure TARC.chiamata_skype(skype_user: string);
begin
  if skype_user = '' then
  begin
    messaggio(000, 'manca il nome dell''utente skype');
  end
  else
  begin
    shellexecute(application.handle, pchar('open'), pchar('skype:' + skype_user + '?call'), nil, nil, SW_HIDE);
  end;
end;

procedure TARC.messaggio_skype(skype_user: string);
begin
  if skype_user = '' then
  begin
    messaggio(000, 'manca il nome dell''utente skype');
  end
  else
  begin
    shellexecute(application.handle, pchar('open'), pchar('skype:' + skype_user + '?chat'), nil, nil, SW_HIDE);
  end;
end;

function getValidTelephoneNumber(number: string): string;
var
  telephoneNumber: string;
  _number: string;
  i: Integer;
begin
  telephoneNumber := '';
  _number := trim(number);

  if _number = '' then
  begin
    telephoneNumber := '';
  end
  else
  begin
    if _number.StartsWith('0039') then
    begin
      _number := StringsReplace(_number, ['0039'], ['+39']);
    end;

    if not _number.StartsWith('+') then
    begin
      number := '+39' + number;
    end;

    if _number[1] = '+' then
    begin
      telephoneNumber := '+';
    end;
    for i := 2 to length(_number) do
    begin
      if _number[i].IsNumber then
      begin
        telephoneNumber := telephoneNumber + _number[i];
      end;
    end;
  end;

  result := telephoneNumber;
end;

procedure TARC.messaggio_whatsapp(cellulare: string);
var
  _cellulare_valido: string;
begin
  _cellulare_valido := getValidTelephoneNumber(cellulare);
  if _cellulare_valido = '' then
  begin
    messaggio(000, 'manca il cellulare');
  end
  else
  begin
    shellexecute(application.handle, pchar('open'), pchar('https://web.whatsapp.com/send?phone=' + _cellulare_valido), nil, nil, SW_SHOWNORMAL);
  end;
end;

function read_tabella(data_base: TMyConnection_go; nome_archivio, nome_codice: string; valore_codice: variant): Boolean;
begin
  result := read_tabella(data_base, nome_archivio, nome_codice, valore_codice, '*');
end;

function read_tabella(data_base: TMyConnection_go; nome_archivio, nome_codice: string; valore_codice: variant; nome_campi: string): Boolean;
var
  valore_codice_array: variant;
  indice, i: word;
  testo_sql, stringa, stringa_codice: string;
begin
  result := false;

  // array di varyant
  if vartype(valore_codice) = 8204 then
  begin
    valore_codice_array := valore_codice;
  end;

  indice := 0;
  stringa := nome_codice;
  while pos(';', stringa) > 0 do
  begin
    indice := indice + 1;
    stringa := copy(stringa, pos(';', stringa) + 1, length(stringa));
  end;

  testo_sql := 'select ' + nome_campi + ' from ' + nome_archivio + ' where';
  stringa := nome_codice;
  for i := 0 to indice do
  begin
    if pos(';', stringa) > 0 then
    begin
      stringa_codice := copy(stringa, 1, pos(';', stringa) - 1);
      stringa := copy(stringa, pos(';', stringa) + 1, length(stringa) + pos(';', stringa));
    end
    else
    begin
      stringa_codice := stringa;
      stringa := '';
    end;
    if i > 0 then
    begin
      testo_sql := testo_sql + ' and ';
    end;
    testo_sql := testo_sql + ' ' + stringa_codice;

    if vartype(valore_codice) = 8204 then
    begin
      testo_sql := testo_sql + ' = :' + stringa_codice;
    end
    else
    begin
      testo_sql := testo_sql + ' = :' + stringa_codice;
    end;
  end;

  if data_base = arc.arc then
  begin
    archivio_arc.sql.text := testo_sql;

    for i := 0 to indice do
    begin
      if vartype(valore_codice) = 8204 then
      begin
        archivio_arc.params[i].value := valore_codice_array[i];
      end
      else
      begin
        archivio_arc.params[i].value := valore_codice;
      end;
    end;

    archivio_arc.Close;
    archivio_arc.open;
    if not archivio_arc.isempty then
    begin
      result := true;
    end;
  end
  else if data_base = arc.arcdit then
  begin
    archivio.sql.text := testo_sql;

    for i := 0 to indice do
    begin
      if vartype(valore_codice) = 8204 then
      begin
        archivio.params[i].value := valore_codice_array[i];
      end
      else
      begin
        archivio.params[i].value := valore_codice;
      end;
    end;

    archivio.Close;
    archivio.open;
    if not archivio.isempty then
    begin
      result := true;
    end;
  end;
end;

function read_tabella(tabella_lettura: tmyquery_go; codice_archivio: variant): Boolean;
var
  i: word;
begin
  if tabella_lettura.params.count = 1 then
  begin
    if vartype(codice_archivio) = vardate then
    begin
      tabella_lettura.params[0].asdate := codice_archivio;
    end
    else
    begin
      tabella_lettura.params[0].value := codice_archivio;
    end;
  end
  else
  begin
    for i := 0 to tabella_lettura.params.count - 1 do
    begin
      if vartype(codice_archivio[i]) = vardate then
      begin
        tabella_lettura.params[i].asdate := codice_archivio[i];
      end
      else
      begin
        tabella_lettura.params[i].value := codice_archivio[i];
      end;
    end;
  end;
  tabella_lettura.Close;
  tabella_lettura.open;
  if not tabella_lettura.isempty then
  begin
    result := true;
  end
  else
  begin
    result := false;
  end;
end;

function read_tabella(tabella_lettura: tmyquery_go): Boolean;
begin
  tabella_lettura.Close;
  tabella_lettura.open;
  if not tabella_lettura.isempty then
  begin
    result := true;
  end
  else
  begin
    result := false;
  end;
end;

function read_tabella(data_base: TMyConnection_go; nome_tabella: string): Boolean;
begin
  result := false;
  if data_base = arc.arc then
  begin
    archivio_arc.Close;
    archivio_arc.sql.text := 'select * from ' + nome_tabella;
    archivio_arc.open;
    result := not archivio_arc.isempty;
  end
  else if data_base = arc.arcdit then
  begin
    archivio.Close;
    archivio.sql.text := 'select * from ' + nome_tabella;
    archivio.open;
    result := not archivio.isempty;
  end;
end;

function cambio(codice_tva: string; data_valuta: tdatetime): double;
var
  tvf, tva: tmyquery_go;
begin
  tvf := create_query(arc.arcdit, 'select cambio from tvf where tva_codice = :tva_codice and data = :data');

  result := 1;
  if read_tabella(tvf, vararrayof([codice_tva, data_valuta])) then
  begin
    result := tvf.fieldbyname('cambio').asfloat;
  end
  else
  begin
    tva := create_query(arc.arcdit, 'select cambio from tva where codice = :codice');
    if read_tabella(tva, codice_tva) then
    begin
      result := tva.fieldbyname('cambio').asfloat;
    end;
    tva.free;
  end;

  tvf.free;
end;

function create_query(data_base: TMyConnection_go; testo_sql: string): tmyquery_go;
begin
  result := tmyquery_go.create(nil);
  result.connection := data_base;
  result.sql.text := testo_sql;
end;

procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string);
var
  lista: tstringlist;
begin
  lista := tstringlist.create;
  esegui_visarc_effettivo(database, tabella, nome_lookup, codice_archivio, filtro_01, filtro_02, filtro_03, multiselezione, chiave_fissa, programma_gestione, lista);
  lista.free;
end;

procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; var lista_multiselezione: tstringlist);
begin
  esegui_visarc_effettivo(database, tabella, nome_lookup, codice_archivio, filtro_01, filtro_02, filtro_03, multiselezione, chiave_fissa, programma_gestione, lista_multiselezione);
end;

procedure esegui_visarc(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; obbligatorio: Boolean);
var
  lista: tstringlist;
begin
  lista := tstringlist.create;
  esegui_visarc_effettivo(database, tabella, nome_lookup, codice_archivio, filtro_01, filtro_02, filtro_03, multiselezione, chiave_fissa, programma_gestione, lista, obbligatorio);
  lista.free;
end;

procedure esegui_visarc_effettivo(database: TMyConnection_go; tabella: string; nome_lookup: string; var codice_archivio: variant; filtro_01, filtro_02: variant; filtro_03: variant; multiselezione, chiave_fissa, programma_gestione: string; var lista_multiselezione: tstringlist;
  obbligatorio: Boolean = false);
var
  vis: tvisinh;
  vis20: tvis20inh;
  query_personalizzata: Boolean;
begin
  query_personalizzata := true;
  if arc.dit.fieldbyname('nuovo_vis').asstring = 'si' then
  begin
    query_personalizzata := false;

    (*
      //  controllo per eseguire VIS se query personalizzata
      select
      case
      when (select query_personalizzata from vis where utn_codice = :utn_codice and codice = :codice) is not null then
      (select query_personalizzata from vis where utn_codice = :utn_codice and codice = :codice)
      else (select query_personalizzata from vis where utn_codice = 'OPEN' and codice = :codice)
      end query_personalizzata
      from ggg
 *)

    visarc_codice := codice_archivio;

    if (lowercase(copy(nome_lookup, 1, 3)) = 'cli') and (arc.utn.fieldbyname('tag_filtro').asstring = 'si') then
    begin
      nome_lookup := 'cliag';
      filtro_01 := arc.utn.fieldbyname('tag_codice').asstring;
    end;

    if (lowercase(copy(nome_lookup, 1, 3)) = 'nom') and (arc.utn.fieldbyname('tag_filtro').asstring = 'si') then
    begin
      nome_lookup := 'nomag';
      filtro_01 := arc.utn.fieldbyname('tag_codice').asstring;
    end;

    vis20 := tvis20inh.create(nil);
    try
      vis20.databasename := database;
      vis20.chiave_fissa := chiave_fissa;
      vis20.programma_gestione := programma_gestione;
      vis20.multiselezione := multiselezione;
      vis20.lista_multiselezione := lista_multiselezione;
      vis20.filtro_01 := filtro_01;
      vis20.filtro_02 := filtro_02;
      vis20.filtro_03 := filtro_03;
      vis20.nomelookup := nome_lookup;
      vis20.obbligatorio := obbligatorio;
      vis20.showmodal;
      codice_archivio := visarc_codice;
      lista_multiselezione := vis20.lista_multiselezione;
    finally
      vis20.free;
    end;
  end;

  if query_personalizzata then
  begin
    visarc_codice := codice_archivio;

    if (lowercase(copy(nome_lookup, 1, 3)) = 'cli') and (arc.utn.fieldbyname('tag_filtro').asstring = 'si') then
    begin
      nome_lookup := 'cliag';
      filtro_01 := arc.utn.fieldbyname('tag_codice').asstring;
    end;

    if (lowercase(copy(nome_lookup, 1, 3)) = 'nom') and (arc.utn.fieldbyname('tag_filtro').asstring = 'si') then
    begin
      nome_lookup := 'nomag';
      filtro_01 := arc.utn.fieldbyname('tag_codice').asstring;
    end;

    vis := tvisinh.create(nil);
    vis.databasename := database;
    vis.chiave_fissa := chiave_fissa;
    vis.programma_gestione := programma_gestione;
    vis.multiselezione := multiselezione;
    vis.lista_multiselezione := lista_multiselezione;
    vis.filtro_01 := filtro_01;
    vis.filtro_02 := filtro_02;
    vis.filtro_03 := filtro_03;
    vis.nomelookup := nome_lookup;
    vis.obbligatorio := obbligatorio;
    vis.showmodal;
    codice_archivio := visarc_codice;
    lista_multiselezione := vis.lista_multiselezione;
    vis.free;
  end;
end;

function decimali_importo(tva_codice: string): word;
var
  tva: tmyquery_go;
begin
  tva := create_query(arc.arcdit, 'select decimali_importo from tva where codice = :codice');

  result := 2;
  if read_tabella(tva, tva_codice) then
  begin
    result := tva.fieldbyname('decimali_importo').asinteger;
  end;

  tva.free;
end;

function decimali_quantita_art(art_codice: string; tipo_tum: string = ''): word;
var
  art, tum: tmyquery_go;
  tum_codice: string;
begin
  art := create_query(arc.arcdit, 'select tum_codice, tum_codice_vendite, tum_codice_acquisti, ' + 'tum_codice_dsb from art where codice = :codice');
  tum := create_query(arc.arcdit, 'select decimali from tum where codice = :codice');

  result := 4;
  if read_tabella(art, art_codice) then
  begin
    if tipo_tum = 'vendite' then
    begin
      tum_codice := art.fieldbyname('tum_codice_vendite').asstring;
    end
    else if tipo_tum = 'acquisti' then
    begin
      tum_codice := art.fieldbyname('tum_codice_acquisti').asstring;
    end
    else if tipo_tum = 'dsb' then
    begin
      tum_codice := art.fieldbyname('tum_codice_dsb').asstring;
    end
    else
    begin
      tum_codice := art.fieldbyname('tum_codice').asstring;
    end;
    if tum_codice = '' then
    begin
      tum_codice := art.fieldbyname('tum_codice').asstring;
    end;

    if read_tabella(tum, tum_codice) then
    begin
      result := tum.fieldbyname('decimali').asinteger;
    end;
  end;

  tum.free;
  art.free;
end;

function decimali_prezzo_nom(nom_codice: string): word;
var
  nom: tmyquery_go;
begin
  nom := create_query(arc.arcdit, 'select tva_codice from nom where codice = :codice');

  result := 6;
  if read_tabella(nom, nom_codice) then
  begin
    result := decimali_prezzo(nom.fieldbyname('tva_codice').asstring);
  end;

  nom.free;
end;

function decimali_prezzo(tva_codice: string): word;
var
  tva: tmyquery_go;
begin
  tva := create_query(arc.arcdit, 'select decimali_prezzo from tva where codice = :codice');

  result := 6;
  if read_tabella(tva, tva_codice) then
  begin
    result := tva.fieldbyname('decimali_prezzo').asinteger;
  end;

  tva.free;
end;

function decimali_prezzo_acq_nom(nom_codice: string): word;
var
  nom: tmyquery_go;
begin
  nom := create_query(arc.arcdit, 'select tva_codice from nom where codice = :codice');

  result := 6;
  if read_tabella(nom, nom_codice) then
  begin
    result := decimali_prezzo(nom.fieldbyname('tva_codice').asstring);
  end;

  nom.free;
end;

function decimali_prezzo_acq(tva_codice: string): word;
var
  tva: tmyquery_go;
begin
  tva := create_query(arc.arcdit, 'select decimali_prezzo_acq from tva where codice = :codice');

  result := 6;
  if read_tabella(tva, tva_codice) then
  begin
    result := tva.fieldbyname('decimali_prezzo_acq').asinteger;
  end;

  tva.free;
end;

function decimali_importo_nom(nom_codice: string): word;
var
  nom: tmyquery_go;
begin
  nom := create_query(arc.arcdit, 'select tva_codice from nom where codice = :codice');

  result := 2;
  if read_tabella(nom, nom_codice) then
  begin
    result := decimali_prezzo(nom.fieldbyname('tva_codice').asstring);
  end;

  nom.free;
end;

function cancella_escluso(const testo: string; caratteri: tsyscharset): string;
var
  i: Integer;
begin
  result := testo;
  for i := length(result) downto 1 do
  begin
    if not(result[i] in caratteri) then
      delete(result, i, 1);
  end;
end;

function EncodePwd(const valpwd: string): string;
var
  contro: string;
  i, p, v: Integer;
  n: array [1 .. 3] of Integer;
  valori: array [0 .. 61] of Integer;
begin
  for i := 0 to 10 do
  begin
    valori[i] := i + 48;
  end;
  for i := 0 to 25 do
  begin
    valori[i + 10] := i + 65;
  end;
  for i := 0 to 25 do
  begin
    valori[i + 36] := i + 97;
  end;
  contro := '';
  for i := 1 to length(valpwd) do
  begin
    v := Ord(byte(valpwd[i]));
    n[1] := 0;
    n[2] := 0;
    n[3] := 0;
    if v > 61 then
      n[1] := random(61)
    else
      n[1] := random(trunc(v / 3));
    p := v - n[1];
    if p > 61 then
      n[2] := random(61)
    else
      n[2] := random(trunc(p / 2));
    n[3] := v - (n[1] + n[2]);
    if n[3] > 61 then
    begin
      n[3] := 59;
      if n[1] > n[2] then
        n[1] := v - (n[2] + n[3])
      else
        n[2] := v - (n[1] + n[3]);
    end;
    contro := contro + Chr(valori[n[1]]) + Chr(valori[n[2]]) + Chr(valori[n[3]]);
  end;
  result := contro;
end;

function DecodePwd(const valpwd: string): string;
var
  contro, parter: string;
  numero, i, n, p, x: Integer;
  valori: array [0 .. 61] of Integer;
begin
  for i := 0 to 10 do
  begin
    valori[i] := i + 48;
  end;
  for i := 0 to 25 do
  begin
    valori[i + 10] := i + 65;
  end;
  for i := 0 to 25 do
  begin
    valori[i + 36] := i + 97;
  end;
  contro := '';
  for i := 1 to trunc(length(valpwd) / 3) do
  begin
    parter := MidStr(valpwd, ((3 * i) - 2), 3);
    numero := 0;
    for n := 1 to 3 do
    begin
      p := Ord(parter[n]);
      for x := 1 to 61 do
      begin
        if p = valori[x] then
          numero := numero + x;
      end;
    end;
    contro := contro + Chr(numero);
  end;
  result := contro;
end;

function md5print(parola: string): string;
var
  md5: imd5;
begin
  md5 := getmd5;
  md5.init;
  md5.update(tbytedynarray(rawbytestring(parola)), length(parola));
  result := lowercase(md5.asstring);
end;

procedure TARC.AddMainContentType(SL: TStrings; img_count: Integer; allegati: Boolean);
begin
  if allegati then
  begin
    SL.add('Content-Type: multipart/mixed;');
    SL.add(#9'boundary="' + boundary_main_and_attachments + '"');
  end
  else
  begin
    if img_count > 0 then
    begin
      SL.add('Content-Type: multipart/related;');
      SL.add(#9'type="multipart/alternative";');
      SL.add(#9'boundary="' + boundary_message_and_pictures + '"');
    end
    else
    begin
      SL.add('Content-Type: multipart/alternative;');
      SL.add(#9'boundary="' + boundary_text_and_html + '"');
    end;
  end;
end;

procedure TARC.escludi_tco_tna_iva_sospensione(tabella: tmyquery_go; nom_codice: string);
var
  tco, tna: tmyquery_go;
begin
  tco := tmyquery_go.create(nil);
  tna := tmyquery_go.create(nil);

  tco.connection := tabella.connection;
  tna.connection := tabella.connection;

  try
    tco.sql.text := 'select escluso_iva_cassa from tco where codice = ' + quotedstr(tabella.fieldbyname('tco_codice').asstring);
    tco.open;
  except
    tco.Close;
    tco.sql.text := 'select ''no'' escluso_iva_cassa';
    tco.open;
  end;

  tna.sql.text := 'select tna.escluso_iva_cassa from nom inner join tna on tna.codice = nom.tna_codice ' + 'where nom.codice = ' + quotedstr(nom_codice);
  tna.open;

  if (tco.fieldbyname('escluso_iva_cassa').asstring = 'si') or (tna.fieldbyname('escluso_iva_cassa').asstring = 'si') then
  begin
    tabella.fieldbyname('iva_sospensione').asstring := 'no';
  end;

  tco.free;
  tna.free;
end;

function etichetta_campo(griglia_devex: Pgriglia_devex; i: Integer; nome_campo_db, etichetta: string): Boolean;
var
  field: string;
begin
  result := true;
  field := lowercase(griglia_devex^.datacontroller.getitemfield(i).FieldName);
  if containsstr(field, nome_campo_db) then
  begin
    griglia_devex^.columns[i].caption := stringreplace(field, nome_campo_db, etichetta, [rfreplaceall, rfIgnoreCase]);
    griglia_devex^.columns[i].caption := stringreplace(griglia_devex^.columns[i].caption, '_', ' ', [rfreplaceall]);
    result := false;
  end;
end;

procedure rinomina_campi(griglia_devex: Pgriglia_devex; titoli_minuscoli: string = 'si');
var
  i: Integer;
  continua: Boolean;
begin
  for i := 0 to griglia_devex^.columncount - 1 do
  begin
    continua := true;

    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'oar', 'ordine d''acquisto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tum_codice', 'unità di misura');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tva_codice', 'valuta');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tma_codice', 'codice deposito');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tma_descrizione', 'deposito');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tmo_codice', 'causale movimento magazzino');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tmo_descrizione', 'movimento magazzino');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cfg_tipo', 'tipo progressivo contabile');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cfg_codice', 'codice progressivo contabile');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cfg_descrizione', 'progressivo contabile');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ese_codice', 'esercizio');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tsm_codice_art', 'codice sconto articolo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tsm_codice', 'codice sconto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'art_codice', 'codice articolo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'art_descrizione1', 'descrizione articolo (1a parte)');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'art_descrizione2', 'descrizione articolo (2a parte)');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'art_descrizione', 'articolo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'progressivo_reggen', 'progressivo registrazione piano dei conti');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tco_codice', 'codice causale contabile');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tpe_codice', 'ritenuta d''acconto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tco_descrizione', 'causale contabile');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'gen_codice', 'codice piano dei conti');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'gen_descrizione', 'voce piano dei conti');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'importo_saldo_tva', 'importo saldo in valuta');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tub_codice', 'codice ubicazione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cen_codice', 'codice centro di costo/ricavo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cen_descrizione', 'centro di costo/ricavo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'doc_progressivo_origine', 'progressivo documento d''origine');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tvc_codice', 'codice voce analitica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tvc_descrizione', 'voce contabilità analitica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cms_codice', 'codice commessa');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cms_descrizione', 'commessa');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tpo_descrizione', 'porto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cli_vet', 'cliente vettore');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cli_codice', 'codice cliente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cli_descrizione', 'cliente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cli_des', 'cliente destinazione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cli_citta', 'città cliente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'frn_codice', 'codice fornitore');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'frn_descrizione', 'fornitore');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'frn_citta', 'città fornitore');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cmd_descrizione', 'descrizione intervento commessa');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tba_codice', 'codice banca');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tba_descrizione', 'banca');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ind_codice', 'codice indirizzo spedizione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ind_descrizione', 'indirizzo spedizione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ind_via', 'indirizzo spedizione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ind_citta', 'città spedizione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'csl_codice', 'codice causale chiusura sollecito');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'csl_descrizione', 'causale chiusura sollecito');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tpa_codice', 'codice pagamento');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tpa_descrizione', 'pagamento');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tda_codice', 'codice causale documento');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tda_descrizione', 'causale documento d''acquisto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tdo_codice', 'codice causale documento');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tdo_descrizione', 'causale documento di vendita');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tln_codice', 'codice linea di produzione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tln_descrizione', 'linea di produzione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tnc_codice', 'codice non conformità');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tnc_descrizione', 'non conformità');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tac_codice', 'codice correzione non conformità');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tac_descrizione', 'correzione non conformità');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tsm_codice', 'codice sconto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tsm_descrizione', 'sconto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tgm_codice', 'codice gruppo merceologico');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tgm_descrizione', 'gruppo merceologico');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tag_codice', 'codice agente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tag_descrizione', 'agente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'lot_codice', 'lotto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'utn_codice', 'codice utente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'utn_descrizione', 'utente');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'cmt_descrizione', 'commessa');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'mtr_codice', 'codice matr. assistenza tecnica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'mtr_descrizione', 'matricola assistenza tecnica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tns_codice', 'codice tipo installazione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tns_descrizione', 'tipo installazione');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tcn_codice', 'codice contratto assistenza tecnica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tcn_descrizione', 'contratto assistenza tecnica');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tlv_codice', 'codice listino di vendita');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tlv_descrizione', 'listino di vendita');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'aut_codice', 'codice automezzo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'aut_descrizione', 'automezzo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tec_codice', 'codice tecnico assitenza');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tec_descrizione', 'tecnico assistenza');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tua_codice', 'codice stato apparecchiatura');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tua_descrizione', 'stato apparecchiatura');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ttc_codice', 'codice tipo contatto');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'ttc_descrizione', 'tipologia contatto cli/for');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'dit_codice', 'codice ditta');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'dit_descrizione', 'ditta');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'prg_codice', 'codice programma');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'prg_descrizione', 'programma');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'csp_codice', 'codice cespite');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'csp_descrizione1', 'descrizione cespite (1a parte)');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'csp_descrizione2', 'descrizione cespite (2a parte)');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tmc_codice', 'codice movimento cespite');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'tmc_descrizione', 'movimento cespite');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'nom_codice', 'codice nominativo');
    end;
    if continua then
    begin
      continua := etichetta_campo(griglia_devex, i, 'nom_descrizione', 'nominativo');
    end;

    if continua then
    begin
      // sostituisco eventuali '_' con blank e metto in lowercase
      if titoli_minuscoli = 'si' then
      begin
        griglia_devex^.columns[i].caption := lowercase(stringreplace(griglia_devex^.datacontroller.getitemfield(i).FieldName, '_', ' ', [rfreplaceall]));
      end
      else
      begin
        griglia_devex^.columns[i].caption := stringreplace(griglia_devex^.datacontroller.getitemfield(i).FieldName, '_', ' ', [rfreplaceall]);
      end;
    end;
  end;
end;

procedure assegna_parametri_passati;
var
  i: word;
begin
  if paramcount > 0 then
  begin
    parametro_globale := '';
    for i := 1 to paramcount do
    begin
      parametro_globale := parametro_globale + paramstr(i) + ' ';

      if parametro_utente = '' then
      begin
        parametro_utente := estrai_tag(paramstr(i), 'utente');
      end;
      if parametro_ditta = '' then
      begin
        parametro_ditta := estrai_tag(paramstr(i), 'ditta');
      end;
      if parametro_esercizio = '' then
      begin
        parametro_esercizio := estrai_tag(paramstr(i), 'esercizio');
      end;
      if parametro_salta_login = '' then
      begin
        parametro_salta_login := estrai_tag(paramstr(i), 'salta_login');
      end;

      if parametro_programma = '' then
      begin
        parametro_programma := estrai_tag(paramstr(i), 'programma');
      end;
      if parametro_programma <> '' then
      begin
        parametro_salta_login := 'si';
      end;

      if parametro_tvm_codice = '' then
      begin
        parametro_tvm_codice := estrai_tag(paramstr(i), 'tvm_codice');
      end;
      if parametro_tpresho_codice = '' then
      begin
        parametro_tpresho_codice := estrai_tag(paramstr(i), 'tpresho_codice');
      end;
      if parametro_schedulato = '' then
      begin
        parametro_schedulato := estrai_tag(paramstr(i), 'schedulato');
      end;
      if parametro_progressivo_gesven = '' then
      begin
        parametro_progressivo_gesven := estrai_tag(paramstr(i), 'progressivo_gesven');
      end;
      if parametro_codice_gesarc = '' then
      begin
        parametro_codice_gesarc := estrai_tag(paramstr(i), 'codice_gesarc');
      end;
      if parametro_tag_codice = '' then
      begin
        parametro_tag_codice := estrai_tag(paramstr(i), 'tag_codice');
      end;
      if parametro_vending_assistenza = '' then
      begin
        parametro_vending_assistenza := estrai_tag(paramstr(i), 'vending_assistenza');
      end;
      if parametro_vending_ordini = '' then
      begin
        parametro_vending_ordini := estrai_tag(paramstr(i), 'vending_ordini');
      end;
      if parametro_vending_segnalazioni = '' then
      begin
        parametro_vending_segnalazioni := estrai_tag(paramstr(i), 'vendind_segnalazioni');
      end;
      if parametro_multi = '' then
      begin
        parametro_multi := estrai_tag(paramstr(i), 'multi');
      end;
      if parametro_personalizzato = '' then
      begin
        parametro_personalizzato := estrai_tag(paramstr(i), 'personalizzato');
      end;
      if parametro_sessione = '' then
      begin
        parametro_sessione := estrai_tag(paramstr(i), 'sessione');
      end;
      if parametro_negozio = '' then
      begin
        parametro_negozio := estrai_tag(paramstr(i), 'negozio');
      end;
      if parametro_password = '' then
      begin
        parametro_password := estrai_tag(paramstr(i), 'password');
      end;
      if parametro_personalizzazioni = '' then
      begin
        parametro_personalizzazioni := estrai_tag(paramstr(i), 'personalizzazioni');
      end;
      if parametro_stampa_diretta_pdf = '' then
      begin
        parametro_stampa_diretta_pdf := estrai_tag(paramstr(i), 'stampa_diretta_pdf');
      end;
    end;
  end;

  if extractfilename(lowercase(application.exename)) = 'gestionale_bovini.exe' then
  begin
    parametro_salta_login := 'si';
    parametro_utente := 'BOVINI';
    parametro_ditta := 'AAAA';
    parametro_esercizio := 'AAAA';

    parametro_globale := parametro_utente + ' ' + parametro_ditta + ' ' + parametro_esercizio + ' ' + parametro_salta_login;
  end;
end;

function estrai_tag(campo, tag: string): string;
begin
  result := '';
  result := trim(campo);
  if pos('<', result) > 0 then
  begin
    result := copy(result, pos('<' + tag + '>', result) + length(tag) + 2);
    result := copy(result, 1, pos('</' + tag + '>', result) - 1);
  end
  else if pos('[', result) > 0 then
  begin
    result := copy(result, pos('[' + tag + ']', result) + length(tag) + 2);
    result := copy(result, 1, pos('[/' + tag + ']', result) - 1);
  end;
end;

procedure disabilita_campo(campo: TObject; tabstop: Boolean = true);
begin
  if (campo is trzedit_go) or (campo is trzdbedit_go) then
  begin
    setpropvalue(campo, 'color', clbtnface);
    setpropvalue(campo, 'readonly', true);
    if getpropvalue(campo, 'tabstop') = true then
    begin
      setpropvalue(campo, 'tabstop', false);
    end;
  end
  else if (campo is trznumericedit_go) or (campo is trzdbnumericedit_go) or (campo is trzdatetimeedit_go) or (campo is trzdbdatetimeedit_go) or (campo is trzcombobox) or (campo is trzdbcombobox) or (campo is trzcolorcombobox) or (campo is trzdbgrid_go) or (campo is trzmemo_go) or
    (campo is trzdbmemo_go) then
  begin
    setpropvalue(campo, 'color', clbtnface);
    setpropvalue(campo, 'readonly', true);
    setpropvalue(campo, 'tabstop', false);
  end
  else if (campo is trzrapidfirebutton) or (campo is trzcheckbox) or (campo is trzdbcheckbox) or (campo is trzbitbtn) or (campo is ttoolbutton) or (campo is trzgroupbox) or (campo is tgroupbox) or (campo is trzpagecontrol) or (campo is trzpanel) or (campo is timage) then
  begin
    setpropvalue(campo, 'enabled', false);
  end;
end;

function abilita_campo(campo: TObject; controllo_abilitato: Boolean = false): Boolean;
begin
  result := true;

  if (campo is trzedit_go) or (campo is trzdbedit_go) or (campo is trznumericedit_go) or (campo is trzdbnumericedit_go) or (campo is trzdatetimeedit_go) or (campo is trzdbdatetimeedit_go) or (campo is trzcombobox) or (campo is trzcolorcombobox) or (campo is trzdbcombobox) or (campo is trzmemo_go) or
    (campo is trzdbmemo_go) or (campo is trzdbgrid_go) then
  begin
    if controllo_abilitato then
    begin
      if (getpropvalue(campo, 'readonly') = true) or (getpropvalue(campo, 'enabled') = false) then
      begin
        result := false;
      end;
    end
    else
    begin
      setpropvalue(campo, 'color', clwindow);
      setpropvalue(campo, 'readonly', false);
      setpropvalue(campo, 'tabstop', true);
    end;
  end
  else if (campo is trzcheckbox) or (campo is trzdbcheckbox) or (campo is trzrapidfirebutton) or (campo is trzbitbtn) or (campo is ttoolbutton) or (campo is trzgroupbox) or (campo is tgroupbox) or (campo is trzpagecontrol) or (campo is trzpanel) or (campo is timage) then
  begin
    if controllo_abilitato then
    begin
      if (getpropvalue(campo, 'readonly') = true) then
      begin
        result := false;
      end;
    end
    else
    begin
      setpropvalue(campo, 'enabled', true);
    end;
  end;
end;

function TARC.controllo_dominio(dominio: string): Boolean;
begin
  result := false;
  server_controllo_dominio := '';
  ipwmx.resolve(dominio);
  if trim(server_controllo_dominio) = '' then
  begin
    result := true;
  end;
end;

procedure TARC.ipwMXResponse(Sender: TObject; RequestId: Integer; const Domain, MailServer: string; Precedence, TimeToLive, StatusCode: Integer; const Description: string; Authoritative: Boolean);
begin
  server_controllo_dominio := server_controllo_dominio + MailServer;
end;

procedure TARC.edit_note(var note_out: string; note: string; data_set: tmyquery_go; modifica: Boolean = true; font: word = 14);
var
  pr: tnote;
begin
  pr := tnote.create(nil);
  pr.note := note_out;
  pr.font := font;

  if not modifica then
  begin
    pr.v_note.readonly := true;
  end;
  pr.showmodal;

  if modifica then
  begin
    if assigned(data_set) then
    begin
      if tabella_edit(data_set) then
      begin
        data_set.fieldbyname(note).asstring := pr.note;
      end;
    end
    else
    begin
      note_out := pr.note;
    end;
  end;

  pr.free;
end;

procedure TARC.generazione_barcode(art: tmyquery_go; forzatura: Boolean = false);
var
  bar, bar_insert, query: tmyquery_go;

  prosegui: Boolean;
  stringa: string;
  progressivo, numero_parte_fissa: double;
begin
  if ((dit.fieldbyname('generazione_automatica_barcode').asstring = 'si') or forzatura) and (art.fieldbyname('codice_barre_peso').asstring = 'no') then
  begin
    bar := tmyquery_go.create(nil);
    bar.connection := arcdit;
    bar.sql.text := 'select codice_barre from bar where art_codice = :art_codice and codice_interno = ''si'' limit 1';
    read_tabella(bar, art.fieldbyname('codice').asstring);

    query := tmyquery_go.create(nil);
    query.connection := arcdit;

    bar_insert := tmyquery_go.create(nil);
    bar_insert.connection := arcdit;
    bar_insert.sql.text := 'insert into bar (art_codice, codice_barre, codice_interno) values (:art_codice, :codice_barre, :codice_interno)';

    prosegui := true;
    if (dit.fieldbyname('codice_barre').asstring = 'ean 13') or (dit.fieldbyname('codice_barre').asstring = 'ean 8') then
    begin
      bar.Close;
      bar.parambyname('art_codice').asstring := art.fieldbyname('codice').asstring;
      bar.open;
      if bar.eof then
      begin
        progressivo := 0;
        query.Close;
        query.sql.clear;
        query.sql.add('select max(codice_barre) codice_barre from bar where codice_interno = ''si''');
        query.sql.add('and codice_barre like ' + quotedstr(dit.fieldbyname('parte_fissa_codice_barre').asstring + '%'));
        query.open;
        // if query.fieldbyname('codice_barre').value = null then
        begin
          if query.fieldbyname('codice_barre').asstring = '' then
          begin
            progressivo := 0;
          end
          else
          begin
            if not numerico(query.fieldbyname('codice_barre').asstring) then
            begin
              messaggio(000, 'l''ultimo codice a barre interno presente in archivio non è numerico');
              prosegui := false;
            end
            else
            begin
              progressivo := strtofloat(query.fieldbyname('codice_barre').asstring);
            end;
          end;

          progressivo := trunc(progressivo / 10);

          numero_parte_fissa := strtofloat(dit.fieldbyname('parte_fissa_codice_barre').asstring);
          if dit.fieldbyname('codice_barre').asstring = 'ean 13' then
          begin
            case length(dit.fieldbyname('parte_fissa_codice_barre').asstring) of
              0:
                progressivo := progressivo;
              1:
                progressivo := trunc(progressivo - numero_parte_fissa * 100000000000);
              2:
                progressivo := trunc(progressivo - numero_parte_fissa * 10000000000);
              3:
                progressivo := trunc(progressivo - numero_parte_fissa * 1000000000);
              4:
                progressivo := trunc(progressivo - numero_parte_fissa * 100000000);
              5:
                progressivo := trunc(progressivo - numero_parte_fissa * 10000000);
              6:
                progressivo := trunc(progressivo - numero_parte_fissa * 1000000);
              7:
                progressivo := trunc(progressivo - numero_parte_fissa * 100000);
              8:
                progressivo := trunc(progressivo - numero_parte_fissa * 10000);
              9:
                progressivo := trunc(progressivo - numero_parte_fissa * 1000);
              10:
                progressivo := trunc(progressivo - numero_parte_fissa * 100);
            end;
          end
          else if dit.fieldbyname('codice_barre').asstring = 'ean 8' then
          begin
            case length(dit.fieldbyname('parte_fissa_codice_barre').asstring) of
              0:
                progressivo := progressivo;
              1:
                progressivo := trunc(progressivo - numero_parte_fissa * 1000000);
              2:
                progressivo := trunc(progressivo - numero_parte_fissa * 100000);
              3:
                progressivo := trunc(progressivo - numero_parte_fissa * 10000);
              4:
                progressivo := trunc(progressivo - numero_parte_fissa * 1000);
              5:
                progressivo := trunc(progressivo - numero_parte_fissa * 100);
              6:
                progressivo := trunc(progressivo - numero_parte_fissa * 10);
              7:
                progressivo := trunc(progressivo - numero_parte_fissa * 1);
            end;
          end;
        end;

        if prosegui then
        begin
          progressivo := progressivo + 1;

          if dit.fieldbyname('codice_barre').asstring = 'ean 8' then
          begin
            stringa := dit.fieldbyname('parte_fissa_codice_barre').asstring + setta_lunghezza(progressivo, 7 - length(dit.fieldbyname('parte_fissa_codice_barre').asstring), 0) + ' ';
            check_digit_codice_barre('ean 8', stringa);
          end
          else
          begin
            stringa := dit.fieldbyname('parte_fissa_codice_barre').asstring + setta_lunghezza(progressivo, 12 - length(dit.fieldbyname('parte_fissa_codice_barre').asstring), 0) + ' ';
            check_digit_codice_barre('ean 13', stringa);
          end;

          bar_insert.Close;

          bar_insert.parambyname('art_codice').asstring := art.fieldbyname('codice').asstring;
          bar_insert.parambyname('codice_barre').asstring := stringa;
          bar_insert.parambyname('codice_interno').asstring := 'si';

          bar_insert.execsql;
        end;
      end
      else
      begin
        messaggio(000, 'il codice a barre non è stato generato perchè esiste già');
      end;
    end
    else
    begin
      messaggio(200, 'il barcode non è stato creato perché il tipo di codice a barre' + #13 + 'deve essere ean 8 o ean 13');
    end;

    bar.free;
    query.free;
    bar_insert.free;
  end;
end;

function tabella_edit(tabella: tdataset): Boolean;
begin
  result := true;
  if (tabella.state <> dsedit) and (tabella.state <> dsinsert) then
  begin
    tabella.edit;
  end;
  if (tabella.state <> dsedit) and (tabella.state <> dsinsert) then
  begin
    result := false;
  end;
end;

procedure colore_control(contenitore: twincontrol; codice_abilitato: Boolean);
var
  i: Integer;
  colore_righe: tcolor;
  enabled_righe: Boolean;
begin
  if contenitore.enabled then
  begin
    if codice_abilitato then
    begin
      colore_righe := clwindow;
      enabled_righe := true;
    end
    else
    begin
      colore_righe := clbtnface;
      enabled_righe := false;
    end;

    for i := 0 to (contenitore.controlcount - 1) do
    begin
      if ((contenitore.Controls[i] is trzdbedit_go) and (not(contenitore.Controls[i] is trzdbeditdescrizione_go))) or (contenitore.Controls[i] is trzedit_go) or (contenitore.Controls[i] is trzdbnumericedit_go) then
      begin
        if getpropvalue(contenitore.Controls[i], 'readonly') = false then
        begin
          setpropvalue(contenitore.Controls[i], 'color', colore_righe);
          if (colore_righe = clwindow) and (contenitore.Controls[i].enabled = false) then
          begin
            setpropvalue(contenitore.Controls[i], 'color', clbtnface);
          end;
        end
      end
      else if (contenitore.Controls[i] is trznumericedit_go) or (contenitore.Controls[i] is trzdbdatetimeedit_go) or (contenitore.Controls[i] is trzdatetimeedit_go) or (contenitore.Controls[i] is trzdbmemo_go) or (contenitore.Controls[i] is trzmemo_go) or (contenitore.Controls[i] is trzdbgrid_go) or
        (contenitore.Controls[i] is trzdbcombobox_go) or (contenitore.Controls[i] is trzcombobox_go) then
      begin
        setpropvalue(contenitore.Controls[i], 'color', colore_righe);
        if (colore_righe = clwindow) and (contenitore.Controls[i].enabled = false) then
        begin
          setpropvalue(contenitore.Controls[i], 'color', clbtnface);
        end;
      end
      else if (contenitore.Controls[i] is trzdbcheckbox) or (contenitore.Controls[i] is trzcheckbox) or (contenitore.Controls[i] is trzbutton) or (contenitore.Controls[i] is trzbitbtn) or (contenitore.Controls[i] is trzrapidfirebutton) or (contenitore.Controls[i] is tupdown) then
      begin
        contenitore.Controls[i].enabled := enabled_righe;
      end
      else if (contenitore.Controls[i] is tgroupbox) or (contenitore.Controls[i] is trzgroupbox) or (contenitore.Controls[i] is trzpagecontrol) or (contenitore.Controls[i] is trztabsheet) or (contenitore.Controls[i] is trzpanel) then
      begin
        colore_control(twincontrol(contenitore.Controls[i]), codice_abilitato);
      end
    end;
  end;
end;

procedure TARC.calcola_peso_lordo(testata_documento: tmyquery_go; colli: Integer = 0);
var
  numero_colli: Integer;
begin
  numero_colli := 0;
  if colli <> 0 then
  begin
    numero_colli := colli;
  end
  else
  begin
    if testata_documento.fieldbyname('numero_colli').asinteger <> 0 then
    begin
      numero_colli := testata_documento.fieldbyname('numero_colli').asinteger;
    end
    else if testata_documento.fieldbyname('numero_confezioni').asinteger <> 0 then
    begin
      numero_colli := testata_documento.fieldbyname('numero_confezioni').asinteger;
    end;
  end;

  if (testata_documento.fieldbyname('peso_lordo').asfloat = 0) and (testata_documento.fieldbyname('peso_netto').asfloat <> 0) and (testata_documento.fieldbyname('tab_codice').asstring <> '') and (numero_colli <> 0) then
  begin
    read_tabella(arcdit, 'tab', 'codice', testata_documento.fieldbyname('tab_codice').asstring);
    if tabella_edit(testata_documento) then
    begin
      if archivio.fieldbyname('riferimento_tara').asstring = 'colli' then
      begin
        testata_documento.fieldbyname('peso_lordo').asfloat := testata_documento.fieldbyname('peso_netto').asfloat + (numero_colli * archivio.fieldbyname('tara').asfloat);
      end
      else
      begin
        testata_documento.fieldbyname('peso_lordo').asfloat := testata_documento.fieldbyname('peso_netto').asfloat + (numero_colli * archivio.fieldbyname('tara').asfloat);
      end;
    end;
  end;
end;

function TARC.tipo_variabile(campo: tfieldtype): string;
begin
  result := '';
  if campo in [ftstring, ftFixedChar, ftwidestring, ftFixedWideChar] then
  begin
    result := 'char';
  end
  else if campo in [ftmemo, ftWideMemo] then
  begin
    result := 'memo';
  end
  else if campo in [ftSmallInt, ftInteger, ftWord, ftAutoInc, ftLongWord, ftShortint] then
  begin
    result := 'int';
  end
  // else if campo in [ftFloat, ftCurrency, ftLargeint, ftExtended] then
  else if campo in [ftFloat, ftCurrency, ftLargeint, tfieldtype(45)] then
  begin
    result := 'num';
  end
  else if campo in [ftDate] then
  begin
    result := 'data';
  end
  else if campo in [ftDatetime] then
  begin
    result := 'data_ora';
  end;
end;

procedure TARC.cerca_valore_passato(valore_passato_ricerca: string; data_set: tdataset; campo_tabella: string; successivo: Boolean = false);
var
  trovato: Boolean;
begin
  data_set.disablecontrols;

  trovato := false;
  if successivo then
  begin
    data_set.next;
  end;
  while not data_set.eof do
  begin
    if (pos(lowercase(valore_passato_ricerca), lowercase(data_set.fieldbyname(campo_tabella).asstring)) > 0) then
    begin
      trovato := true;
      break;
    end
    else if (pos(lowercase(valore_passato_ricerca), multireplace(lowercase(data_set.fieldbyname(campo_tabella).asstring), ' |.|,|-|;|/|\', '')) > 0) then
    begin
      trovato := true;
      break;
    end;
    data_set.next;
  end;
  if not trovato then
  begin
    if messaggio(300, 'valore [' + valore_passato_ricerca + '] non trovato' + #13 + 'riesegui la ricerca dall''inizio') = 1 then
    begin
      data_set.first;
      cerca_valore_passato(valore_passato_ricerca, data_set, campo_tabella);
    end;
  end;

  data_set.enablecontrols;
end;

function TARC.multireplace(valore, old, new: string): string;
var
  str: string;
begin
  result := valore;
  while pos('|', old) > 0 do
  begin
    str := copy(old, 1, pos('|', old) - 1);
    result := stringreplace(result, str, new, [rfreplaceall]);
    old := trim(copy(old, pos('|', old) + 1, length(old)));
  end;
end;

procedure TARC.ricerca_griglia(griglia: trzdbgrid_go);
var
  pr: timpalf;
  indice: word;
begin
  indice := griglia.selectedindex;

  try
    pr := timpalf.create(nil);
    pr.v_form_caption := 'Ricerca ' + griglia.columns[indice].title.caption;
    pr.v_descrizione_caption := griglia.columns[indice].title.caption;

    if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'char' then
    begin
      pr.tipo_campo := 'alfa';
      if carattere_ricerca_griglia = '' then
      begin
        pr.valore_passato := '';
      end
      else
      begin
        pr.valore_passato := carattere_ricerca_griglia;
        pr.digitato_carattere := true;
        carattere_ricerca_griglia := '';
      end
    end
    else if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'memo' then
    begin
      pr.tipo_campo := 'alfa';
      if carattere_ricerca_griglia = '' then
      begin
        pr.valore_passato := '';
      end
      else
      begin
        pr.valore_passato := carattere_ricerca_griglia;
        pr.digitato_carattere := true;
        carattere_ricerca_griglia := '';
      end;
    end
    else if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'data' then
    begin
      pr.tipo_campo := 'data';
      pr.valore_passato := date;
    end
    else if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'data_ora' then
    begin
      pr.tipo_campo := 'data';
      pr.valore_passato := date;
    end
    else
    begin
      pr.tipo_campo := 'numero';
      if carattere_ricerca_griglia = '' then
      begin
        pr.valore_passato := 0;
      end
      else
      begin
        try
          pr.valore_passato := strtoint(carattere_ricerca_griglia);
        except
          pr.valore_passato := 0;
        end;
        pr.digitato_carattere := true;
        carattere_ricerca_griglia := '';
      end;
    end;

    pr.v_width_campo := 20;
    pr.showmodal;

    if pr.premuto_escape then
    begin
      valore_passato_ricerca := '';
    end
    else
    begin
      if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'char' then
      begin
        valore_passato_ricerca := pr.valore_passato;
      end
      else if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'memo' then
      begin
        valore_passato_ricerca := pr.valore_passato;
      end
      else if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'data' then
      begin
        try
          valore_passato_ricerca := datetostr(pr.valore_passato);
        except
          valore_passato_ricerca := '';
        end;
      end
      else
      begin
        try
          valore_passato_ricerca := floattostr(pr.valore_passato);
        except
          valore_passato_ricerca := '';
        end;
      end;
    end;
  finally
    freeandnil(pr);
  end;

  if valore_passato_ricerca <> '' then
  begin
    cerca_valore_passato(valore_passato_ricerca, griglia.datasource.dataset, griglia.columns[indice].FieldName);
  end;
end;

procedure TARC.filtro_griglia(griglia: trzdbgrid_go; var filtro_impostato: string);
var
  pr: timpalf;
  indice: word;
begin
  indice := griglia.selectedindex;

  try
    pr := timpalf.create(nil);
    pr.v_form_caption := 'Impostazione filtro: [' + filtro_impostato + ']';
    pr.v_descrizione_caption := griglia.columns[indice].title.caption;
    // if tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].fieldname).datatype) = 'char' then
    begin
      pr.tipo_campo := 'alfa';
      if carattere_ricerca_griglia = '' then
      begin
        pr.valore_passato := '';
      end
      else
      begin
        pr.valore_passato := carattere_ricerca_griglia;
        pr.digitato_carattere := true;
        carattere_ricerca_griglia := '';
      end;
      pr.v_width_campo := 400;
      pr.showmodal;

      if not pr.premuto_escape then
      begin
        if pr.valore_passato = '' then
        begin
          if filtro_impostato <> '' then
          begin
            griglia.datasource.dataset.filter := filtro_impostato;
          end
          else
          begin
            griglia.datasource.dataset.filter := '';
            griglia.datasource.dataset.filtered := false;
          end;
        end
        else
        begin
          if filtro_impostato <> '' then
          begin
            griglia.datasource.dataset.filter := '(' + filtro_impostato + ') and (lower(' + griglia.columns[indice].FieldName + ') like ' + quotedstr('%' + pr.valore_passato + '%') + ')';
          end
          else
          begin
            griglia.datasource.dataset.filter := 'lower(' + griglia.columns[indice].FieldName + ') like ' + quotedstr('%' + lowercase(pr.valore_passato) + '%');

            griglia.datasource.dataset.filtered := true;
          end;
        end;

        filtro_impostato := griglia.datasource.dataset.filter;
      end;
    end;

  finally
    freeandnil(pr);
  end;
end;

procedure TARC.totalizza_griglia(griglia: trzdbgrid_go);
var
  indice, i: word;
  totale: double;
begin
  indice := griglia.selectedindex;

  if (tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'num') or (tipo_variabile(griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).datatype) = 'int') then
  begin
    i := griglia.datasource.dataset.recno;
    griglia.datasource.dataset.disablecontrols;
    griglia.datasource.dataset.first;

    totale := 0;

    screen.Cursor := crhourglass;
    griglia.datasource.dataset.disablecontrols;
    while not griglia.datasource.dataset.eof do
    begin
      totale := totale + griglia.datasource.dataset.fieldbyname(griglia.columns[indice].FieldName).asfloat;

      griglia.datasource.dataset.next;
    end;
    griglia.datasource.dataset.enablecontrols;
    screen.Cursor := cursore;

    messaggio(100, 'totale ' + griglia.columns[indice].title.caption + ' = ' + Formatfloat(',0.000000;-,0.000000;#', totale));

    // griglia.datasource.dataset.first;
    griglia.datasource.dataset.recno := i;
    griglia.datasource.dataset.enablecontrols;
  end
  else
  begin
    messaggio(200, 'la colonna deve essere di tipo numerico');
  end;
end;

function TARC.controllo_data_utilizzo: Boolean;
var
  data: tdate;
begin
  result := false;
  data := strtodate('31/03/2022');

  if not programma_collegato_personalizzato then
  begin
    if date > data then
    begin
      result := true;
      messaggio(000, 'il programma non è utilizzabile oltre il ' + datetostr(data) + ' senza aggiornare la versione' + slinebreak + 'ricordiamo che il produttore del software che state utilizzando' + slinebreak + 'è indicato nella videata successiva');
      esegui_programma('ABOUTBOX', 'fisso', true);
    end
    else if date + 31 > data then
    begin
      messaggio(200, 'il ' + datetostr(data) + ' scade il termine di utilizzo del programma' + #13 + 'aggiornare la versione');
    end;
  end;
end;

function TARC.controllo_data_licenza: Boolean;
var
  data: tdate;
begin
  result := false;

  if (lowercase(codice_procedura) = 'go') and (versione_aggiornamento <> 0) then
  begin
    data := strtodate('31/03/2022');

    if date > data then
    begin
      messaggio(000, 'la data di utilizzo della Licenza d''uso della versione è scaduta' + slinebreak + slinebreak + 'aggiornare alla versione successiva');
      if date > data + 60 then
      begin
        messaggio(200, 'dal ' + datetostr(data + 60) + ' il programma non sarà più utilizzabile ' + slinebreak + 'aggiornare alla versione successiva');
        result := true;
      end;
    end
    else if date + 31 > data then
    begin
      messaggio(200, 'il ' + datetostr(data) + ' scade il termine di utilizzo della Licenza d''uso' + slinebreak + 'aggiornare alla versione successiva');
    end;
  end;
end;

function TARC.mag_esercizio(art_codice, tma_codice, ese_codice: string; dalla_data, alla_data: tdate): double;
begin
  query_mag_esercizio.Close;
  query_mag_esercizio.parambyname('art_codice').asstring := art_codice;
  query_mag_esercizio.parambyname('tma_codice').asstring := tma_codice;
  query_mag_esercizio.parambyname('ese_codice').asstring := ese_codice;
  query_mag_esercizio.parambyname('data_inizio').asdate := dalla_data;
  query_mag_esercizio.parambyname('data_bilancio').asdate := alla_data;
  query_mag_esercizio.open;
  result := query_mag_esercizio.fieldbyname('esistenza').asfloat;
end;

procedure filtra_eccezioni(const exceptintf: imeexception; var handled: Boolean);
begin
  if bugreport_nominale = 'si' then
  begin
    mesettings().bugreportfile := 'bugreport_' + utente + '.txt';
  end;

  if exceptintf.exceptclass = 'EAccessViolation' then
  begin
    handled := access_violation = 'no';
  end
end;

function tgridhelper.columnbyname(const aname: string): tcolumn;
var
  i: word;
begin
  result := nil;
  for i := 0 to columns.count - 1 do
  begin
    if (columns[i].field <> nil) and (columns[i].FieldName.tolower = aname.tolower) then
    begin
      result := columns[i];
      exit;
    end;
  end;
end;

function TARC.presenti_provvisori(ese_codice: string; dalla_data: tdate = 0; alla_data: tdate = 0): Boolean;
var
  pnt: tmyquery_go;
begin
  result := false;

  if dit.fieldbyname('movimenti_provvisori').asstring = 'si' then
  begin
    pnt := tmyquery_go.create(nil);
    pnt.connection := arcdit;

    pnt.sql.add('select id from pnt');
    if ese_codice <> '' then
    begin
      pnt.sql.add('where ese_codice = ' + quotedstr(ese_codice));
    end
    else
    begin
      pnt.sql.add('where data_registrazione between :dalla_data and :alla_data');
      pnt.parambyname('dalla_data').asdate := dalla_data;
      pnt.parambyname('alla_data').asdate := alla_data;
    end;
    pnt.sql.add('and movimento_provvisorio = ''si''');

    pnt.open;
    if not pnt.isempty then
    begin
      result := true;
    end;

    freeandnil(pnt);
  end;
end;

function TARC.numero_documento_alfa(tabella: tmyquery_go; campo_numero_documento, numero_documento_alfa: string): double;
var
  i: word;
  stringa: string;
begin
  inherited;

  result := 0;

  if numero_documento_alfa <> '' then
  begin
    stringa := '';

    (*
      for i := 1 to length(numero_documento_alfa) do
      begin
      if isnumeric(numero_documento_alfa[i]) then
      begin
      stringa := stringa + numero_documento_alfa[i];
      end
      else
      begin
      break;
      end;
      end;
      if stringa = '' then
      begin
      stringa := '0';
      end;

      if length(stringa) < 16 then
      begin
      result := strtofloat(copy(stringa, 1, 15));
      end
      else
      begin
      result := strtofloat(copy(stringa, length(stringa) - 14, 15));
      end;
 *)
    // assegna tutti i numerici
    for i := 1 to length(numero_documento_alfa) do
    begin
      if isnumeric(numero_documento_alfa[i]) then
      begin
        stringa := stringa + numero_documento_alfa[i];
      end;
    end;
    if stringa = '' then
    begin
      stringa := '0';
    end;
    result := strtofloat(trim(copy(stringa, 1, 15)));
    // assegna tutti i numerici

    if result = 0 then
    begin
      result := strtofloat(formatdatetime('yyyymmdd', date));
    end;

    if tabella <> nil then
    begin
      if tabella_edit(tabella) then
      begin
        tabella.fieldbyname(campo_numero_documento).asfloat := result;
      end;
    end;
  end;
end;

function TARC.serie_documento_alfa(tabella: tmyquery_go; campo_serie_documento, numero_documento_alfa: string): string;
var
  i, inizio: word;
  stringa: string;
begin
  inherited;

  result := '';
  if numero_documento_alfa <> '' then
  begin
    stringa := '';

    (*
      inizio := 0;
      for i := 1 to length(numero_documento_alfa) do
      begin
      if not isnumeric(numero_documento_alfa[i]) then
      begin
      inizio := i;
      break;
      end;
      end;

      for i := inizio to length(numero_documento_alfa) do
      begin
      stringa := stringa + numero_documento_alfa[i];
      end;
 *)

    // assegna tutti gli alfa
    for i := 1 to length(numero_documento_alfa) do
    begin
      if not isnumeric(numero_documento_alfa[i]) and (numero_documento_alfa[i] <> '/') and (numero_documento_alfa[i] <> '-') and (numero_documento_alfa[i] <> '_') then
      begin
        stringa := stringa + numero_documento_alfa[i];
      end;
    end;
    // assegna tutti gli alfa

    result := stringa;

    if tabella <> nil then
    begin
      if tabella_edit(tabella) then
      begin
        tabella.fieldbyname(campo_serie_documento).asstring := result;
      end;
    end;
  end;
end;

procedure apri_transazione(cursore_normale: string = 'no');
begin
  if not arc.arcdit.intransaction then
  begin
    if cursore_normale = 'no' then
    begin
      screen.Cursor := crhourglass;
    end;
    arc.arcdit.starttransaction;
  end;
end;

procedure commit_transazione(testo: string = 'transazione non eseguita');
begin
  if arc.arcdit.intransaction then
  begin
    try
      arc.arcdit.commit;
    except
      on e: exception do
      begin
        rollback_transazione(e.message);
      end;
    end;
  end;
end;

procedure rollback_transazione(testo: string = '');
begin
  screen.Cursor := cursore;

  if pos('deadlock', lowercase(testo)) = 0 then
  begin
    messaggio(000, testo);
  end;

  if arc.arcdit.intransaction then
  begin
    if pos('deadlock', lowercase(testo)) > 0 then
    begin
      messaggio(200, 'si è verificato un deadlock' + #13 + '(stesso record in aggiornamento da due utenti contemporaneamente)' + #13 + 'sul database durante la transazione' + #13 + #13 + 'RIESEGUIRE L''ELABORAZIONE INTERROTTA', true);
    end
    else
    begin
      messaggio(200, 'la transazione non è stata completata correttamente' + #13 + #13 + 'rieseguire l''elaborazione interrotta');
    end;

    try
      arc.arcdit.rollback;
    except
      on e: exception do
      begin
        messaggio(000, e.message);
      end;
    end;
  end;

  screen.Cursor := crhourglass;
end;

procedure chiudi_transazione;
begin
  screen.Cursor := cursore;

  // annulla eventuali transazioni aperte
  if arc.arcdit.intransaction then
  begin
    arc.arcdit.rollback;
  end;
end;

procedure TARC.assegna_peso_modulo(peso: Integer);
var
  i: word;
  conversioni: longint;
  esiste: Boolean;
  ggg: tmyquery_go;
begin
  ggg := tmyquery_go.create(nil);
  ggg.connection := arc;
  ggg.sql.text := 'update ggg set conversioni = ''0'' where conversioni not regexp ''[0][1][2][3][4][5][6][7][8][9]''';
  ggg.execsql;

  ggg.sql.clear;
  ggg.sql.text := 'select * from ggg where codice = ''G''';
  ggg.open;
  conversioni := strtoint(ggg.fieldbyname('conversioni').asstring);

  esiste := false;
  for i := 24 downto peso - 1 do
  begin
    if (i <> peso - 1) and (conversioni >= power(2, i)) then
    begin
      conversioni := conversioni - trunc(power(2, i));
    end;
    if (conversioni / peso >= 1) and (conversioni / peso < 2) then
    begin
      esiste := true;
      break;
    end;
  end;

  if not esiste then
  begin
    ggg.edit;
    conversioni := strtoint(ggg.fieldbyname('conversioni').asstring) + peso;
    ggg.fieldbyname('conversioni').asstring := inttostr(conversioni);
    ggg.post;
  end;

  freeandnil(ggg);
end;

procedure TARC.sconti_percentuale(componente: twincontrol);
var
  pr: tscontiperc;
  tsm_codice: string;
begin
  tsm_codice := trzcustomedit(componente).text;

  pr := tscontiperc.create(nil);
  try
    if pr.esegui_form then
    begin
      pr.tsm_codice := trzcustomedit(componente).text;
      pr.showmodal;
      tsm_codice := pr.tsm_codice;
    end;
  finally
    freeandnil(pr);
  end;

  if tsm_codice <> trzcustomedit(componente).text then
  begin
    if (componente is trzedit) then
    begin
      trzedit(componente).text := tsm_codice;
    end
    else
    begin
      if tabella_edit(trzdbedit(componente).datasource.dataset) then
      begin
        trzdbedit(componente).datasource.dataset.fieldbyname(trzdbedit(componente).datafield).asstring := tsm_codice;
      end;
    end;
  end;
end;

function TARC.crea_tsm(sconto_maggiorazione: string; percentuale_01, percentuale_02, percentuale_03, percentuale_04, percentuale_05, percentuale_06, percentuale_07, percentuale_08: double): string;
var
  tsm_ultimo, tsm: tmyquery_go;
  numero: Integer;
begin
  result := '';

  tsm := tmyquery_go.create(nil);
  tsm.connection := arcdit;
  tsm.sql.add('select * from tsm');
  tsm.sql.add('where sconto_maggiorazione = :sconto_maggiorazione');
  tsm.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 and percentuale_03 = :percentuale_03');
  tsm.sql.add('and percentuale_04 = :percentuale_04 and percentuale_05 = :percentuale_05 and percentuale_06 = :percentuale_06');
  tsm.sql.add('and percentuale_07 = :percentuale_07 and percentuale_08 = :percentuale_08');

  tsm.Close;
  tsm.parambyname('sconto_maggiorazione').asstring := sconto_maggiorazione;
  tsm.parambyname('percentuale_01').asfloat := percentuale_01;
  tsm.parambyname('percentuale_02').asfloat := percentuale_02;
  tsm.parambyname('percentuale_03').asfloat := percentuale_03;
  tsm.parambyname('percentuale_04').asfloat := percentuale_04;
  tsm.parambyname('percentuale_05').asfloat := percentuale_05;
  tsm.parambyname('percentuale_06').asfloat := percentuale_06;
  tsm.parambyname('percentuale_07').asfloat := percentuale_07;
  tsm.parambyname('percentuale_08').asfloat := percentuale_08;
  tsm.open;
  if not tsm.isempty then
  begin
    result := tsm.fieldbyname('codice').asstring;
  end
  else
  begin
    tsm_ultimo := tmyquery_go.create(nil);
    tsm_ultimo.connection := arcdit;
    tsm_ultimo.sql.add('select max(codice) codice from tsm');
    tsm_ultimo.sql.add('where mid(codice, 1, 1) regexp ''[0-9]''');
    tsm_ultimo.sql.add('and mid(codice, 2, 1) regexp ''[0-9]''');
    tsm_ultimo.sql.add('and mid(codice, 3, 1) regexp ''[0-9]''');
    tsm_ultimo.sql.add('and mid(codice, 4, 1) regexp ''[0-9]''');

    tsm_ultimo.open;
    if (tsm_ultimo.isempty) or (tsm_ultimo.fieldbyname('codice').value = null) then
    begin
      numero := 1;
    end
    else
    begin
      numero := strtoint(tsm_ultimo.fieldbyname('codice').asstring) + 1;
    end;
    tsm.append;
    tsm.fieldbyname('codice').asstring := setta_lunghezza(numero, 4, 0);
    tsm.fieldbyname('sconto_maggiorazione').asstring := sconto_maggiorazione;
    tsm.fieldbyname('descrizione').asstring := floattostr(percentuale_01);
    if percentuale_02 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_02);
    end;
    if percentuale_03 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_03);
    end;
    if percentuale_04 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_04);
    end;
    if percentuale_05 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_05);
    end;
    if percentuale_06 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_06);
    end;
    if percentuale_07 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_07);
    end;
    if percentuale_08 <> 0 then
    begin
      tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + '+' + floattostr(percentuale_08);
    end;
    tsm.fieldbyname('percentuale_01').asfloat := percentuale_01;
    tsm.fieldbyname('percentuale_02').asfloat := percentuale_02;
    tsm.fieldbyname('percentuale_03').asfloat := percentuale_03;
    tsm.fieldbyname('percentuale_04').asfloat := percentuale_04;
    tsm.fieldbyname('percentuale_05').asfloat := percentuale_05;
    tsm.fieldbyname('percentuale_06').asfloat := percentuale_06;
    tsm.fieldbyname('percentuale_07').asfloat := percentuale_07;
    tsm.fieldbyname('percentuale_08').asfloat := percentuale_08;

    tsm.fieldbyname('percentuale_totale').asfloat := 100;

    if sconto_maggiorazione = 'sconto' then
    begin
      if tsm.fieldbyname('percentuale_08').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) * (100 - tsm.fieldbyname('percentuale_04').asfloat) *
          (100 - tsm.fieldbyname('percentuale_05').asfloat) * (100 - tsm.fieldbyname('percentuale_06').asfloat) * (100 - tsm.fieldbyname('percentuale_07').asfloat) * (100 - tsm.fieldbyname('percentuale_08').asfloat) / 100000000000000;
      end
      else if tsm.fieldbyname('percentuale_07').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) * (100 - tsm.fieldbyname('percentuale_04').asfloat) *
          (100 - tsm.fieldbyname('percentuale_05').asfloat) * (100 - tsm.fieldbyname('percentuale_06').asfloat) * (100 - tsm.fieldbyname('percentuale_07').asfloat) / 1000000000000;
      end
      else if tsm.fieldbyname('percentuale_06').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) * (100 - tsm.fieldbyname('percentuale_04').asfloat) *
          (100 - tsm.fieldbyname('percentuale_05').asfloat) * (100 - tsm.fieldbyname('percentuale_06').asfloat) / 10000000000;
      end
      else if tsm.fieldbyname('percentuale_05').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) * (100 - tsm.fieldbyname('percentuale_04').asfloat) *
          (100 - tsm.fieldbyname('percentuale_05').asfloat) / 100000000;
      end
      else if tsm.fieldbyname('percentuale_04').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) * (100 - tsm.fieldbyname('percentuale_04').asfloat) / 1000000;
      end
      else if tsm.fieldbyname('percentuale_03').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) * (100 - tsm.fieldbyname('percentuale_03').asfloat) / 10000;
      end
      else if tsm.fieldbyname('percentuale_02').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) * (100 - tsm.fieldbyname('percentuale_02').asfloat) / 100;
      end
      else if tsm.fieldbyname('percentuale_01').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 - tsm.fieldbyname('percentuale_01').asfloat) / 1;
      end;
    end
    else
    begin
      if tsm.fieldbyname('percentuale_08').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) * (100 + tsm.fieldbyname('percentuale_04').asfloat) *
          (100 + tsm.fieldbyname('percentuale_05').asfloat) * (100 + tsm.fieldbyname('percentuale_06').asfloat) * (100 + tsm.fieldbyname('percentuale_07').asfloat) * (100 + tsm.fieldbyname('percentuale_08').asfloat) / 100000000000000;
      end
      else if tsm.fieldbyname('percentuale_07').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) * (100 + tsm.fieldbyname('percentuale_04').asfloat) *
          (100 + tsm.fieldbyname('percentuale_05').asfloat) * (100 + tsm.fieldbyname('percentuale_06').asfloat) * (100 + tsm.fieldbyname('percentuale_07').asfloat) / 1000000000000;
      end
      else if tsm.fieldbyname('percentuale_06').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) * (100 + tsm.fieldbyname('percentuale_04').asfloat) *
          (100 + tsm.fieldbyname('percentuale_05').asfloat) * (100 + tsm.fieldbyname('percentuale_06').asfloat) / 10000000000;
      end
      else if tsm.fieldbyname('percentuale_05').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) * (100 + tsm.fieldbyname('percentuale_04').asfloat) *
          (100 + tsm.fieldbyname('percentuale_05').asfloat) / 100000000;
      end
      else if tsm.fieldbyname('percentuale_04').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) * (100 + tsm.fieldbyname('percentuale_04').asfloat) / 1000000;
      end
      else if tsm.fieldbyname('percentuale_03').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) * (100 + tsm.fieldbyname('percentuale_03').asfloat) / 10000;
      end
      else if tsm.fieldbyname('percentuale_02').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) * (100 + tsm.fieldbyname('percentuale_02').asfloat) / 100;
      end
      else if tsm.fieldbyname('percentuale_01').asfloat <> 0 then
      begin
        tsm.fieldbyname('percentuale_totale').asfloat := (100 + tsm.fieldbyname('percentuale_01').asfloat) / 1;
      end;
    end;

    tsm.post;

    result := tsm.fieldbyname('codice').asstring;
    freeandnil(tsm_ultimo);
  end;

  freeandnil(tsm);
end;

function decodifica_html(atesto_html: string): string;
var
  contenuto_testo: Boolean;
  i: Integer;
begin
  result := '';
  contenuto_testo := false;

  for i := 1 to length(atesto_html) do
  begin
    if atesto_html[i] = '<' then
    begin
      contenuto_testo := false;
    end;

    if contenuto_testo then
    begin
      result := result + atesto_html[i];
    end;

    if atesto_html[i] = '>' then
    begin
      contenuto_testo := true;
    end;
  end;

  result := stringreplace(result, '&quot;', '"', [rfreplaceall]);
  result := stringreplace(result, '&apos;', '''', [rfreplaceall]);
  result := stringreplace(result, '&gt;', '>', [rfreplaceall]);
  result := stringreplace(result, '&lt;', '<', [rfreplaceall]);
  result := stringreplace(result, '&amp;', '&', [rfreplaceall]);
  result := stringreplace(result, '&nbsp;', ' ', [rfreplaceall]);
  result := stringreplace(result, '&#232;', 'è', [rfreplaceall]);
  result := stringreplace(result, '&#233;', 'é', [rfreplaceall]);
  result := stringreplace(result, '&#224;', 'à', [rfreplaceall]);
  result := stringreplace(result, '&#8217;', '''', [rfreplaceall]);
  result := stringreplace(result, '&#242;', 'ò', [rfreplaceall]);
  result := stringreplace(result, '&#236;', 'ì', [rfreplaceall]);
  result := stringreplace(result, '&#249;', 'ù', [rfreplaceall]);
  result := stringreplace(result, '&#8220;', '“', [rfreplaceall]);
  result := stringreplace(result, '&#8221;', '”', [rfreplaceall]);
end;

function exe_in_esecuzione(nome_file_exe: string): Boolean;
var
  prosegui: Boolean;
  fsnapshothandle: thandle;
  processo_32bit: tprocessentry32;
  pathname_eseguibile: string;
  myhandle: thandle;
  mypid: DWORD;
begin
  fsnapshothandle := createtoolhelp32snapshot(th32cs_snapprocess, 0);
  processo_32bit.dwsize := SizeOf(processo_32bit);
  prosegui := process32first(fsnapshothandle, processo_32bit);
  result := false;
  while Integer(prosegui) <> 0 do
  begin
    if uppercase(extractfilename(processo_32bit.szexefile)) = uppercase(extractfilename(nome_file_exe)) then
    begin
      mypid := processo_32bit.th32processid;
      myhandle := openprocess(process_query_information or process_vm_read, false, mypid);
      if myhandle <> 0 then
        try
          SetLength(pathname_eseguibile, max_path);
          if getmodulefilenameex(myhandle, 0, pchar(pathname_eseguibile), max_path) > 0 then
          begin
            SetLength(pathname_eseguibile, strlen(pchar(pathname_eseguibile)));
            if uppercase(extractfilename(pathname_eseguibile)) = uppercase(nome_file_exe) then
            begin
              result := true;
            end;
          end
          else
          begin
            pathname_eseguibile := '';
          end;
        finally
          CloseHandle(myhandle);
        end;
      if result then
      begin
        break;
      end;
    end;
    prosegui := process32next(fsnapshothandle, processo_32bit);
  end;
  CloseHandle(fsnapshothandle);
end;

procedure TARC.aggiorna_database;
var
  aggiornare: Boolean;
  aggiornamento_in_esecuzione, versione_aggiornamento_effettuato: word;

  dit, q, ggg: tmyquery_go;
  s: tmyscript_go;

  avviso: tavviso;
  connessione: TMyConnection_go;
begin
  ggg := tmyquery_go.create(nil);
  ggg.connection := arc;
  ggg.sql.text := 'select * from ggg where codice = ''G''';
  ggg.open;

  try
    versione_aggiornamento_effettuato := ggg.fieldbyname('versione_aggiornamento').asinteger;
  except
    versione_aggiornamento_effettuato := 0;
  end;

  aggiornare := true;
  if versione_aggiornamento_effettuato >= versione_aggiornamento then
  begin
    aggiornare := false;
  end;

  if aggiornare then
  begin
    ggg.edit;
    ggg.fieldbyname('utn_codice_manutenzione').asstring := utente;
    ggg.post;

    avviso := tavviso.create(nil);
    avviso.v_messaggio.caption := 'aggiornamento database numero ' + inttostr(versione_aggiornamento);
    avviso.update;
    avviso.show;

    dit := tmyquery_go.create(nil);
    dit.connection := arc;
    dit.sql.text := 'select codice from dit';
    dit.open;

    q := tmyquery_go.create(nil);

    s := tmyscript_go.create(nil);
    s.connection := arcdit;
    s.delimiter := '#';

    connessione := TMyConnection_go.create(self);
    connessione.server := TMyConnection_go(ggg.connection).server;
    connessione.port := TMyConnection_go(ggg.connection).port;
    connessione.username := utente;
    connessione.password := password_database;

    try
      for aggiornamento_in_esecuzione := 1 to versione_aggiornamento do
      begin
        /// /////////////////////////////////////////////////////////////////////
        // aggiornamenti arc
        /// /////////////////////////////////////////////////////////////////////
        (*
          avviso.v_messaggio.caption := 'aggiornamento database [arc] numero ' + inttostr(aggiornamento_in_esecuzione);
          avviso.update;
          application.processmessages;

          if (aggiornamento_in_esecuzione = 2) and (aggiornamento_in_esecuzione > versione_aggiornamento_effettuato) then
          begin
          //  10.12.02
          q.close;
          q.connection := arc;

          q.close;
          q.sql.clear;
          q.sql.add('ALTER TABLE dit add partita_iva_controllante VARCHAR(20) NULL DEFAULT '''';');
          try
          q.execsql;
          except
          end;
          end;
 *)

        /// /////////////////////////////////////////////////////////////////////
        // aggiornamenti arc_ordinamento
        /// /////////////////////////////////////////////////////////////////////
        (*
          avviso.v_messaggio.caption := 'aggiornamento database [arc_ordinamento] numero ' + inttostr(aggiornamento_in_esecuzione);
          avviso.update;
          application.processmessages;

          //  10.12.02
          if (aggiornamento_in_esecuzione = 2) and (aggiornamento_in_esecuzione > versione_aggiornamento_effettuato) then
          begin
          q.close;
          q.connection := arcsor;

          q.close;
          q.sql.clear;
          q.sql.add('ALTER TABLE stasta add importo_anno DECIMAL(18,6) NULL DEFAULT ''0'';');
          q.sql.add('ALTER TABLE stasta add importo_anno_raffronto DECIMAL(18,6) NULL DEFAULT ''0'';');
          q.sql.add('ALTER TABLE stastv add importo_anno DECIMAL(18,6) NULL DEFAULT ''0'';');
          q.sql.add('ALTER TABLE stastv add importo_anno_raffronto DECIMAL(18,6) NULL DEFAULT ''0'';');
          try
          q.execsql;
          except
          end;
          end;
 *)

        /// /////////////////////////////////////////////////////////////////////
        // aggiornamenti arc_ditta
        /// /////////////////////////////////////////////////////////////////////
        dit.Close;
        dit.open;
        while not dit.eof do
        begin
          avviso.v_messaggio.caption := 'aggiornamento database [arc] numero ' + inttostr(aggiornamento_in_esecuzione);
          avviso.update;
          application.processmessages;

          connessione.connected := false;
          connessione.database := 'arc_' + lowercase(dit.fieldbyname('codice').asstring);
          try
            connessione.connected := true;

            // 11.01.02
            if (aggiornamento_in_esecuzione = 2) and (aggiornamento_in_esecuzione > versione_aggiornamento_effettuato) then
            begin
              q.Close;
              q.connection := connessione;

              q.Close;
              q.sql.clear;
              q.sql.add('ALTER TABLE bvr add data_inizio_lavorazione DATE NULL;');
              q.sql.add('ALTER TABLE cvr add data_inizio_lavorazione DATE NULL;');
              q.sql.add('ALTER TABLE dvr add data_inizio_lavorazione DATE NULL;');
              q.sql.add('ALTER TABLE fvr add data_inizio_lavorazione DATE NULL;');
              q.sql.add('ALTER TABLE ovr add data_inizio_lavorazione DATE NULL;');
              q.sql.add('ALTER TABLE pvr add data_inizio_lavorazione DATE NULL;');
              try
                q.execsql;
              except
              end;

              q.Close;
              q.sql.clear;
              q.sql.add('ALTER TABLE opt add data_inizio_lavorazione DATE NULL;');
              try
                q.execsql;
              except
              end;
            end;
          except
          end;

          dit.next;
        end;
      end;
    finally
      ggg.edit;
      ggg.fieldbyname('utn_codice_manutenzione').asstring := '';
      ggg.fieldbyname('versione_aggiornamento').asinteger := versione_aggiornamento;
      ggg.post;

      freeandnil(connessione);
      freeandnil(avviso);
      freeandnil(q);
      freeandnil(s);
      freeandnil(dit);
    end;

    messaggio(200, 'essendo stato aggiornato il database' + slinebreak + 'il programma viene terminato' + slinebreak + 'per consentire l''attivazione delle modifiche');
    application.terminate;
  end;
  freeandnil(ggg);
end;

initialization

registerexceptionhandler(filtra_eccezioni, stdontsync, epquickfiltering);

end.
