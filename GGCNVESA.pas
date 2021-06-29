unit GGCNVESA;

interface

uses
  Windows, Messages, SysUtils, DateUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, GGELABORA, Grids, dbgrids, RzDBGrid, ADODB, DB, MyAccess, query_go, adstable,
  adsdata, adsfunc, Menus, StdCtrls, Buttons, ComCtrls, RzTabs, ExtCtrls,
  ToolWin, Mask, DBTables, MemDS, VirtualTable,

  RzButton, rzLabel, RzPanel, RzDBEdit, RzListVw, RzTreeVw, RzDBChk,
  RzRadChk, RzSplit, RzCmboBx, RzPrgres,
  RzSpnEdt, RzShellDialogs, RzDBCmbo, raizeedit_go, RzEdit, DBAccess;

type

  TCNVESA = class(TELABORA)
    v_griglia: TRzDBGrid_go;
    v_tabella_01: TRzlabel;
    v_tabella: TRzlabel;
    tabella_01: tmytable;
    tabella_02: tmytable;
    cfg: tmytable;
    tabella_01_ds: tmydatasource;
    GroupBox1: TGroupBox;
    Label8: TRzlabel;
    Label7: TRzlabel;
    Label9: TRzlabel;
    v_sottoconti: TRzcheckbox;
    v_clienti: TRzcheckbox;
    v_fornitori: TRzcheckbox;
    v_articoli: TRzcheckbox;
    v_ind_inf: TRzcheckbox;
    v_lsv: TRzcheckbox;
    v_pnt: TRzcheckbox;
    v_scadenze: TRzcheckbox;
    v_mov: TRzcheckbox;
    v_tabelle: TRzcheckbox;
    v_ordini: TRzcheckbox;
    v_codifica_clienti: TRzcheckbox;
    v_codifica_fornitori: TRzcheckbox;
    v_ricavi: TRzEdit_go;
    v_acquisti: TRzEdit_go;
    v_art_codice: TRzEdit_go;
    v_codice_aggiuntivi: TRzcheckbox;
    v_provvigioni: TRzcheckbox;
    tsm: tmytable;
    ADOEsatto: TADOConnection;
    tabella_esa_01: TADOQuery;
    tabella_clifor: TADOQuery;
    tabella_esa_02: TADOQuery;
    tabella_clienti_forn: TADOQuery;
    tpa: tmytable;
    tco: tmytable;
    tsa: tmytable;
    cpa: tmytable;
    cpv: tmytable;
    tca: tmytable;
    query_02: tmyquery_go;
    tabella_03: tmytable;
    tabella_esa_03: TADOQuery;
    query_03: tmyquery_go;
    dit: tmyquery_go;
    procedure v_confermaClick(Sender: TObject);
    procedure tabella_01BeforePost(DataSet: TDataSet);
    procedure tabella_02BeforePost(DataSet: TDataSet);
    procedure tabella_03BeforePost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  protected
    { Protected declarations }
    test_alfa: string;
    test_data: tdate;
    test_datas: string;
    test_numero: double;
    test_cod_causale: string;
    test_numero_stringa: string;
    codice_sor_01, codice_sor_02, codice_sor_03, codice_sor_04: string;
    primo_livello, secondo_livello, terzo_livello, quarto_livello, quinto_livello: double;
    file_archivio: textfile;
    record_tabella, valore: string;
    stringa, codice_articolo: string;
    esiste_articolo: boolean;
    riga, riga_iva: word;
    cli_for, codice_clifor: string;
    inizio_tabella: word;
    tsm_codice: word;

    function Converti_data(data: string): TDateTime;

    procedure calcola_importo;

    procedure converti_tcm;
    procedure converti_tsa;
    procedure converti_tba;
    procedure converti_tva;
    procedure converti_tiv;
    procedure converti_tna;
    procedure converti_tma;
    procedure converti_tsp;
    procedure converti_tzo;
    procedure converti_tlv;
    procedure converti_tag;
    procedure converti_tgm;
    procedure converti_tum;
    procedure converti_tdo;
    procedure converti_tcc;
    procedure converti_tca;
    procedure converti_tcf;
    procedure converti_tpo;
    procedure converti_tab;
    procedure converti_tpa;

    procedure converti_tmo;
    procedure converti_tco;

    procedure converti_sottoconti;
    procedure converti_clienti;
    procedure converti_fornitori;
    procedure converti_articoli;
    procedure converti_bar;
    procedure converti_ind_inf;
    procedure converti_lsv;
    procedure converti_pnt;
    procedure crea_pnt(progressivo: integer);
    procedure crea_pnr(progressivo: integer);
    procedure crea_pni(progressivo: integer);
    procedure converti_par;
    procedure converti_mov;
    procedure crea_mmt(progressivo: integer);
    procedure crea_mmr;

    procedure converti_ordini_clienti;
    procedure crea_ovt;
    procedure crea_ovr(serie_documento, numero_documento: string);
    procedure converti_ordini_fornitori;
    procedure crea_oat;
    procedure crea_oar(serie_documento, numero_documento: string);

    procedure assegna_codice_cli(codice_adhoc: string; var cli_for: string);
    procedure assegna_codice_frn(codice_adhoc: string; var cli_for: string);

    procedure crea_tsm(sconto1, sconto2: double);
    procedure converti_provvigioni;
    procedure converti_fatture_vendita(tipo_documento: string);
    procedure cancella_tabella(tabella: string);
  public
    { Public declarations }
    procedure controllo_campi; override;
  end;

var
  CNVESA: TCNVESA;

implementation

{$r *.dfm}


uses DMARC;

procedure TCNVESA.v_confermaClick(Sender: TObject);
begin
  dit.close;
  dit.parambyname('codice').asstring := ditta;
  dit.open;

  tsm.open;
  if tsm.eof then
  begin
    tsm_codice := 0;
  end
  else
  begin
    // tsm.last;

    while not tsm.Eof do
    begin
      if tsm.fieldbyname('codice').asstring <= '9999' then
      begin
        tsm_codice := strtoint(tsm.fieldbyname('codice').asstring);
      end;
      tsm.Next;
    end;

  end;

  if v_tabelle.checked then
  begin
    converti_tva;
    converti_tiv;
    converti_tna;
    converti_tma;
    converti_tag;
    converti_tlv;

    converti_tcm;
    converti_tgm;
    converti_tna;
    converti_tum;
    converti_tdo; // tipo documenti vendita
    converti_tmo;
    converti_tco;
    converti_tpa;
    converti_tsp;
    converti_tpo;
    converti_tab;
    converti_tsa; // codice statistico  articoli
    converti_tcc;
    converti_tca;
    converti_tcf;
    (*
      conver    ti_tzo;
      converti_tmo; CAUSALIDIMAGAZZINO
      converti_tco;
    *)
  end;

  if (v_clienti.checked) and (v_fornitori.checked) then
  begin
    query.Close;
    query.sql.clear;
    query.sql.add('delete from nom');
    query.ExecSQL;
    if query.Connection.InTransaction then
    begin
      query.Connection.Commit
    end;
  end;

  if v_clienti.checked then
  begin
    converti_clienti;
  end;

  if v_fornitori.checked then
  begin
    converti_fornitori;
  end;

  if v_sottoconti.checked then
  begin
    converti_sottoconti;
  end;

  if v_articoli.checked then
  begin
    converti_articoli;
    converti_bar;
  end;

  if v_codice_aggiuntivi.checked then
  begin
    converti_bar;
  end;

  if v_scadenze.checked then
  begin
    converti_par;
  end;

  if v_pnt.checked then
  begin
    converti_pnt;
  end;

  if v_lsv.checked then
  begin
    converti_lsv;
  end;

  if v_mov.checked then
  begin
    converti_mov;
  end;

  if v_ind_inf.checked then
  begin
    converti_ind_inf;
  end;

  (*

    if v_lsv.checked then
    begin
    converti_lsv;
    end;

    if v_pnt.checked then
    begin
    converti_pnt;
    end;

    if v_mov.checked then
    begin
    converti_mov;
    end;
  *)

  if v_ordini.checked then
  begin
    converti_ordini_clienti;

    converti_ordini_fornitori;
  end;

  if v_provvigioni.checked then
  begin
    converti_provvigioni;
  end;

  dit.close;

  close;
end;

procedure TCNVESA.converti_clienti;
var
  i, j: word;
  data_nascita: TDateTime;
begin
  v_tabella.caption := 'clienti';
  application.processmessages;

  read_tabella(arc.arc, 'dit', 'codice', ditta, '*');

  query_02.Close;
  query_02.sql.clear;
  query_02.sql.add('select * from esa_tag');
  query_02.sql.add('where esa_codice =:esa_codice');
  query_02.parambyname('esa_codice').asstring := '0000';
  try
    query_02.open;
  except
    converti_tag;
  end;

  cancella_tabella('cli');
  tabella_01.close;
  tabella_01.tablename := 'cli';
  tabella_01.open;

  cancella_tabella('nom');
  tabella_02.tablename := 'nom';
  tabella_02.open;

  query.sql.add('delete from cfg ');
  query.sql.add('where');
  query.sql.add('cfg_tipo =' + quotedstr('C'));
  cfg.tablename := 'cfg';
  cfg.open;

  tabella_01_ds.DataSet := tabella_clifor;

  tabella_clifor.Close;
  tabella_clifor.Parameters.ParamByName('ind_Clifor').Value := 'C';
  tabella_clifor.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * from BANCHECLIENTI');
    tabella_esa_02.Sql.Add('where');
    tabella_esa_02.Sql.Add('Parte_Fissa=:PF');
    tabella_esa_02.Sql.Add('order by Codice_Banca');
    tabella_esa_02.Parameters.ParamByName('PF').Value := 'BN';
    tabella_esa_02.open;
  except
    messaggio(000, 'manca la tabella delle banche (BANCHECLIENTI)');
    close;
    abort;
  end;

  // --------------------------------------------------------
  // TABELLA ANAGRAFICHECOMUNI
  // --------------------------------------------------------
  try

    tabella_esa_03.Close;
    tabella_esa_03.Sql.Clear;
    tabella_esa_03.Sql.Add('SELECT * from ANAGRAFICHECOMUNI');
    tabella_esa_03.Sql.Add('where codice_anagrafica=:codice_anagrafica');
    tabella_esa_03.Sql.Add('order by Codice_Anagrafica');
  except
    messaggio(000, 'manca l''anagrafica comune (ANAGRAFICHECOMUNI)');
    close;
    abort;
  end;

  while not tabella_clifor.eof do
  begin
    Application.processMessages;
    // ------------------------------------
    // CERCO CODICE IN TABELLA CLIENTI-FORNITORI
    // ------------------------------------
    tabella_esa_03.close;
    tabella_esa_03.parameters.parambyname('Codice_Anagrafica').Value := tabella_clifor.FieldByName('Codice_Cli_For').AsString;
    tabella_esa_03.open;
    if not tabella_esa_03.eof then
    begin
      // nominativi

      if tabella_clifor.fieldbyname('Codice_Cli_For').asstring <> '' then
      begin
        tabella_02.append;

        assegna_codice_cli(tabella_esa_03.fieldbyname('Codice_Anagrafica').asstring, cli_for);
        tabella_02.fieldbyname('codice').asstring := cli_for;

        if length(trim(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring)) > 30 then
        begin
          for i := 30 downto 1 do
          begin
            if tabella_esa_03.fieldbyname('Ragione_Sociale').asstring[i] = ' ' then
            begin
              j := i;
              break;
            end;
          end;
        end
        else
        begin
          j := length(trim(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring));
        end;
        tabella_02.fieldbyname('descrizione1').asstring := trim(copy(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring, 1, j));

        if tabella_02.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_02.fieldbyname('descrizione1').asstring := '.';
        end;
        tabella_02.fieldbyname('descrizione2').asstring :=
          trim(copy(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring, j + 1, length(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring) - j));

        tabella_02.fieldbyname('via').asstring := trim(tabella_esa_03.fieldbyname('Indirizzo').asstring);
        tabella_02.fieldbyname('cap').asstring := trim(tabella_esa_03.fieldbyname('CAP').asstring);
        tabella_02.fieldbyname('citta').asstring := trim(tabella_esa_03.fieldbyname('Localita').asstring);
        tabella_02.fieldbyname('provincia').asstring := trim(tabella_esa_03.fieldbyname('Provincia').asstring);
        tabella_02.fieldbyname('tna_codice').asstring := trim(tabella_esa_03.fieldbyname('Nazione').asstring);
        if tabella_02.fieldbyname('tna_codice').asstring = '' then
        begin
          tabella_02.fieldbyname('tna_codice').asstring := 'IT';
        end;
        tabella_02.fieldbyname('partita_iva').asstring := tabella_esa_03.fieldbyname('Partita_Iva').asstring;
        tabella_02.fieldbyname('codice_fiscale').asstring := tabella_esa_03.fieldbyname('Codice_Fiscale').asstring;
        tabella_02.fieldbyname('telefono').asstring := trim(tabella_esa_03.fieldbyname('Telefono').asstring);
        tabella_02.fieldbyname('fax').asstring := trim(tabella_esa_03.fieldbyname('Fax').asstring);
        tabella_02.fieldbyname('cellulare').asstring := trim(tabella_esa_03.fieldbyname('Num_Cellulare').asstring);
        if tabella_esa_03.fieldbyname('Flag_Persona_Fisica').asstring = '1' then
        begin
          tabella_02.fieldbyname('codice_alternativo').asstring := trim(tabella_esa_03.fieldbyname('Ragione_Sociale').asString);

          if trim(tabella_esa_03.fieldbyname('Cognome').asString) <> '' then
          begin
            tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_esa_03.fieldbyname('Cognome').asString);
            tabella_02.fieldbyname('descrizione2').asstring := trim(tabella_esa_03.fieldbyname('Nome').asString);
          end;

          tabella_02.fieldbyname('persona_fisica').asstring := 'si';
          if tabella_esa_03.fieldbyname('Flag_Sesso').asstring = '0' then
          begin
            tabella_02.fieldbyname('sesso').asstring := 'femminile';
          end
          else
          begin
            tabella_02.fieldbyname('sesso').asstring := 'maschile';
          end;

          data_nascita := Converti_data(tabella_esa_03.fieldbyname('Data_Nascita').asString);
          if data_nascita <> StrToDate('01/01/1900') then
            tabella_02.fieldbyname('data_nascita').asdatetime := Converti_data(tabella_esa_03.fieldbyname('Data_Nascita').asString);

          // tabella_02.fieldbyname('data_nascita').asdatetime := tabella_esa_03.fieldbyname('Data_Nascita').asdatetime;
          tabella_02.fieldbyname('citta_nascita').asstring := tabella_esa_03.fieldbyname('Luogo_di_Nascita').asstring;
          tabella_02.fieldbyname('provincia_nascita').asstring := tabella_esa_03.fieldbyname('Provincia_Nascita').asstring;
        end
        else
        begin
          tabella_02.fieldbyname('persona_fisica').asstring := 'no';
        end;

        if trim(tabella_clifor.fieldbyname('Codice_Valuta').asString) <> '' then
        begin
          tabella_02.fieldbyname('tva_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Valuta').asString);
        end
        else
        begin
          tabella_02.fieldbyname('tva_codice').asstring := divisa_di_conto;
        end;

        tabella_02.fieldbyname('web').asstring := trim(tabella_esa_03.fieldbyname('Sito_Internet').asstring);
        tabella_02.fieldbyname('e_mail_amministrazione').asstring := trim(tabella_esa_03.fieldbyname('EMAIL').asstring);

        tabella_02.fieldbyname('via_legale').asstring := tabella_02.fieldbyname('via').asstring;
        tabella_02.fieldbyname('cap_legale').asstring := tabella_02.fieldbyname('cap').asstring;
        tabella_02.fieldbyname('citta_legale').asstring := tabella_02.fieldbyname('citta').asstring;
        tabella_02.fieldbyname('provincia_legale').asstring := tabella_02.fieldbyname('provincia').asstring;
        tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_02.fieldbyname('tna_codice').asstring;

        tabella_02.post;

        // ---------------------------------------------------------------------------
        // clienti
        // ---------------------------------------------------------------------------
        tabella_01.append;

        tabella_01.fieldbyname('codice').asstring := tabella_02.fieldbyname('codice').asstring;
        tabella_01.fieldbyname('descrizione1').asstring := trim(tabella_02.fieldbyname('descrizione1').asstring);
        tabella_01.fieldbyname('descrizione2').asstring := trim(tabella_02.fieldbyname('descrizione2').asstring);
        tabella_01.fieldbyname('via').asstring := tabella_02.fieldbyname('via').asstring;
        tabella_01.fieldbyname('citta').asstring := tabella_02.fieldbyname('citta').asstring;
        tabella_01.fieldbyname('partita_iva').asstring := tabella_02.fieldbyname('partita_iva').asstring;
        tabella_01.fieldbyname('codice_fiscale').asstring := tabella_02.fieldbyname('codice_fiscale').asstring;
        tabella_01.fieldbyname('tna_codice').asstring := tabella_02.fieldbyname('tna_codice').asstring;

        // ------------------------------------
        // default ditta
        // ------------------------------------
        tabella_01.fieldbyname('gen_codice').asstring := dit.fieldbyname('gen_codice_cli').asstring;
        tabella_01.fieldbyname('tba_codice').asstring := dit.fieldbyname('tba_codice_cli').asstring;
        tabella_01.fieldbyname('tpa_codice').asstring := dit.fieldbyname('tpa_codice_cli').asstring;
        tabella_01.fieldbyname('tcc_codice').asstring := dit.fieldbyname('tcc_codice_cli').asstring;
        tabella_01.fieldbyname('tlv_codice').asstring := dit.fieldbyname('tlv_codice_cli').asstring;
        tabella_01.fieldbyname('ts1_codice').asstring := dit.fieldbyname('ts1_codice_cli').asstring;
        tabella_01.fieldbyname('tzo_codice').asstring := dit.fieldbyname('tzo_codice_cli').asstring;
        tabella_01.fieldbyname('tzo_codice_assistenza').asstring := dit.fieldbyname('tzo_codice_cli').asstring;
        tabella_01.fieldbyname('tsc_codice').asstring := dit.fieldbyname('tsc_codice_cli').asstring;
        tabella_01.fieldbyname('tsp_codice').asstring := dit.fieldbyname('tsp_codice_cli').asstring;
        tabella_01.fieldbyname('tpo_codice').asstring := dit.fieldbyname('tpo_codice_cli').asstring;
        tabella_01.fieldbyname('tag_codice').asstring := dit.fieldbyname('tag_codice_cli').asstring;
        tabella_01.fieldbyname('tp1_codice').asstring := dit.fieldbyname('tp1_codice_cli').asstring;
        tabella_01.fieldbyname('tpf_codice').asstring := dit.fieldbyname('tpf_codice_cli').asstring;
        tabella_01.fieldbyname('tst_codice').asstring := dit.fieldbyname('tst_codice_cli').asstring;
        tabella_01.fieldbyname('addebito_spese_fattura').asstring := dit.fieldbyname('addebito_spese_fattura_clienti').asstring;
        tabella_01.fieldbyname('tar_codice').asstring := dit.fieldbyname('tar_codice_cli').asstring;
        tabella_01.fieldbyname('tcg_codice').asstring := dit.fieldbyname('tcg_codice_cli').asstring;

        tabella_01.fieldbyname('contatto').asstring := tabella_esa_03.fieldbyname('contatto').asstring;
        if tabella_clifor.fieldbyname('Sottoconto_Apparten').asstring <> '' then
        begin
          tabella_01.fieldbyname('gen_codice').asstring := tabella_clifor.fieldbyname('Sottoconto_Apparten').asstring;
        end;

        // if trim(tabella_clifor.fieldbyname('Flag_Gest_Partite').asstring) = '1' then
        // begin
        tabella_01.fieldbyname('partitario').asstring := 'si';
        // end
        // else
        // begin
        // tabella_01.fieldbyname('partitario').asstring := 'no';
        // end;

        if trim(tabella_clifor.fieldbyname('Codice_Pagamento').asString) <> '' then
        begin
          tabella_01.fieldbyname('tpa_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Pagamento').asString);
        end;

        if trim(tabella_clifor.fieldbyname('Vettore').asString) <> '' then
        begin
          tabella_01.fieldbyname('tsp_codice').asstring := trim(tabella_clifor.fieldbyname('Vettore').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('Porto').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tpo_codice').asstring := trim(tabella_clifor.fieldbyname('Porto').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('Codice_Nazione_Zona').asString) <> '' then
        begin
          tabella_01.fieldbyname('tzo_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Nazione_Zona').asString);
        end;

        if trim(tabella_clifor.fieldbyname('Agente').asString) <> '' then
        begin
          query_02.close;
          query_02.parambyname('esa_codice').asstring := trim(tabella_clifor.fieldbyname('Agente').asstring);
          query_02.sql.savetofile('c:\temp\esa_tag.sql');
          query_02.open;
          if not query_02.eof then
          begin
            tabella_01.fieldbyname('tag_codice').asstring := query_02.fieldbyname('Codice').asstring;
          end
          else
          begin
            tabella_01.fieldbyname('tag_codice').asstring := '0000';
          end;
        end;

        if trim(tabella_clifor.fieldbyname('Codice_Iva_Esente').asString) <> '' then
        begin
          tabella_01.fieldbyname('tiv_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Iva_Esente').asString);
        end;

        if tabella_esa_02.locate('Codice_Banca', trim(tabella_clifor.fieldbyname('Codice_Banca').asstring), []) then
        begin
          tabella_01.fieldbyname('codice_abi').asstring := trim(tabella_clifor.fieldbyname('Codice_Banca').asstring);
          tabella_01.fieldbyname('codice_cab').asstring := trim(tabella_clifor.fieldbyname('Codice_Agenzia').asstring);
        end;
        if tabella_01.fieldbyname('codice_abi').asstring <> '' then
        begin
          if length(tabella_01.fieldbyname('codice_abi').asstring) = 4 then
          begin
            tabella_01.fieldbyname('codice_abi').asstring := '0' + tabella_01.fieldbyname('codice_abi').asstring;
          end;
        end;
        tabella_01.fieldbyname('mese_01').asInteger := 0;

        if trim(tabella_clifor.fieldbyname('Mese_1_Escluso_Pagam').asString) <> '' then
          tabella_01.fieldbyname('mese_01').asInteger := tabella_clifor.fieldbyname('Mese_1_Escluso_Pagam').asInteger;

        if tabella_01.fieldbyname('mese_01').asinteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_01').asinteger := tabella_clifor.fieldbyname('Giorno_Succ_x_Scaden').asinteger;
        end;
        tabella_01.fieldbyname('mese_02').asinteger := 0;

        if trim(tabella_clifor.fieldbyname('Mese_2_Escluso_Pagam').asString) <> '' then
          tabella_01.fieldbyname('mese_02').asinteger := tabella_clifor.fieldbyname('Mese_2_Escluso_Pagam').asinteger;

        if tabella_01.fieldbyname('mese_02').asinteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_02').asinteger := tabella_clifor.fieldbyname('Giorno_Succ_x_Scaden').asinteger;
        end;

        if (tabella_01.fieldbyname('mese_01').asinteger <> 0) or (tabella_01.fieldbyname('mese_02').asinteger <> 0) then
        begin
          tabella_01.fieldbyname('mesi_esclusi').asstring := 'si';
        end;

        if trim(tabella_clifor.fieldbyname('Numero_Listino').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tlv_codice').asstring := trim(tabella_clifor.fieldbyname('Numero_Listino').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring) = '0' then
        begin
          tabella_01.fieldbyname('riepilogo_fattura').asstring := 'nessuno';
        end
        else if tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring = '1' then
        begin
          tabella_01.fieldbyname('riepilogo_fattura').asstring := 'globale';
        end;

        if (tabella_clifor.fieldbyname('Ind_Bolli_in_Fattura').asstring = '1') or
          (tabella_clifor.fieldbyname('Indic_Spese_Incasso').asstring = '1') then
        begin
          tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'si';
        end
        else
        begin
          tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'no';
        end;

        if (tabella_clifor.fieldbyname('Stampa Prezzo Bolla').asstring = '1') then
        begin
          tabella_01.fieldbyname('valori_in_bolla').asstring := 'si';
        end
        else
        begin
          tabella_01.fieldbyname('valori_in_bolla').asstring := 'no';
        end;

        tabella_01.fieldbyname('fido').asFloat := tabella_clifor.fieldbyname('fido').asFloat;

        tabella_01.fieldbyname('conto_corrente').asstring := trim(tabella_clifor.fieldbyname('NumContoC').asstring);
        if tabella_01.fieldbyname('conto_corrente').asstring <> '' then
        begin
          if length(tabella_01.fieldbyname('conto_corrente').asstring) <> 12 then
          begin
            for i := 1 to (12 - length(tabella_01.fieldbyname('conto_corrente').asstring)) do
            begin
              tabella_01.fieldbyname('conto_corrente').asstring := '0' + tabella_01.fieldbyname('conto_corrente').asstring;
            end;
          end;
        end;
        // tabella_01.fieldbyname('note').asstring := trim(tabella_clifor.fieldbyname('CL__NOTE').asstring;
        (*
          if trim(tabella_clifor.fieldbyname('CLSTACOD').asstring = 'S' then
          begin
          tabella_01.fieldbyname('stampa_codice_articolo_cliente').asstring := 'si';
          end
          else
          begin
          tabella_01.fieldbyname('stampa_codice_articolo_cliente').asstring := 'no';
          end;
        *)
        tabella_01.fieldbyname('stampa_codice_articolo_cliente').asstring := 'no';

        tabella_01.post;

        // file cfg

        if cfg.locate('cfg_tipo;cfg_codice', vararrayof(['C', tabella_01.fieldbyname('codice').asstring]), []) then
        begin
          cfg.edit;
          cfg.fieldbyname('descrizione1').asstring := trim(tabella_01.fieldbyname('descrizione1').asstring) + ' ' +
            trim(tabella_01.fieldbyname('descrizione2').asstring);
          cfg.fieldbyname('descrizione2').asstring := trim(tabella_01.fieldbyname('citta').asstring);
          cfg.fieldbyname('utente').asstring := utente;
          cfg.fieldbyname('data_ora').asdatetime := now;
          cfg.post;
        end;

      end;

    end; // if locate anagrafiche comuni

    tabella_clifor.next;
  end;

  query_02.close;
  tabella_01.close;
  tabella_02.close;
  tabella_clifor.close;
  tabella_esa_02.close;
  tabella_esa_03.close;
  cfg.close;
end;

procedure TCNVESA.converti_fornitori;
var
  i, j: word;
  data_nascita: TDateTime;
begin
  v_tabella_01.caption := 'fornitori';
  application.processmessages;

  read_tabella(arc.arc, 'dit', 'codice', ditta, '*');

  tabella_01.close;
  tabella_01.tablename := 'frn';
  tabella_01.open;

  cancella_tabella('frn');

  tabella_02.tablename := 'nom';
  tabella_02.open;

  query.sql.clear;
  query.sql.add('delete from cfg ');
  query.sql.add('where');
  query.sql.add('cfg_tipo =' + quotedstr('F'));

  cfg.tablename := 'cfg';
  cfg.open;

  tabella_clifor.Close;
  tabella_clifor.Parameters.ParamByName('ind_Clifor').Value := 'F';
  tabella_clifor.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * from BANCHECLIENTI');
    tabella_esa_02.Sql.Add('where');
    tabella_esa_02.Sql.Add('Parte_Fissa=:PF');
    tabella_esa_02.Sql.Add('order by Codice_Banca');
    tabella_esa_02.Parameters.ParamByName('PF').Value := 'BN';
    tabella_esa_02.open;

  except
    messaggio(000, 'manca la tabella delle banche (BANCHECLIENTI)');
    close;
    abort;
  end;

  // --------------------------------------------------------
  // TABELLA ANAGRAFICHECOMUNI
  // --------------------------------------------------------
  try
    tabella_esa_03.Close;
    tabella_esa_03.Sql.Clear;
    tabella_esa_03.Sql.Add('SELECT * from ANAGRAFICHECOMUNI');
    tabella_esa_03.Sql.Add('where Codice_Anagrafica=:codice_anagrafica');
    tabella_esa_03.Sql.Add('order by Codice_Anagrafica');
  except
    messaggio(000, 'manca l''anagrafica comune (ANAGRAFICHECOMUNI)');
    close;
    abort;
  end;

  while not tabella_clifor.eof do
  begin
    Application.ProcessMessages;

    // ------------------------------------
    // CERCO CODICE IN TABELLA CLIENTI-FORNITORI
    // ------------------------------------
    tabella_esa_03.close;
    tabella_esa_03.parameters.parambyname('codice_anagrafica').value := tabella_clifor.FieldByName('Codice_Cli_For').AsString;
    tabella_esa_03.open;
    if not tabella_esa_03.eof then
    begin
      assegna_codice_frn(tabella_esa_03.fieldbyname('Codice_Anagrafica').asstring, cli_for);

      // nominativi
      if not tabella_02.Locate('codice', cli_for, []) then
      begin

        if tabella_clifor.fieldbyname('Codice_Cli_For').asstring <> '' then
        begin
          tabella_02.append;

          tabella_02.fieldbyname('codice').asstring := cli_for;
          ;

          if length(trim(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring)) > 30 then
          begin
            for i := 30 downto 1 do
            begin
              if tabella_esa_03.fieldbyname('Ragione_Sociale').asstring[i] = ' ' then
              begin
                j := i;
                break;
              end;
            end;
          end
          else
          begin
            j := length(trim(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring));
          end;
          tabella_02.fieldbyname('descrizione1').asstring := trim(copy(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring, 1, j));

          if tabella_02.fieldbyname('descrizione1').asstring = '' then
          begin
            tabella_02.fieldbyname('descrizione1').asstring := '.';
          end;
          tabella_02.fieldbyname('descrizione2').asstring :=
            trim(copy(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring, j + 1, length(tabella_esa_03.fieldbyname('Ragione_Sociale').asstring) - j));

          tabella_02.fieldbyname('via').asstring := tabella_esa_03.fieldbyname('Indirizzo').asstring;
          tabella_02.fieldbyname('cap').asstring := tabella_esa_03.fieldbyname('CAP').asstring;
          tabella_02.fieldbyname('citta').asstring := tabella_esa_03.fieldbyname('Localita').asstring;
          tabella_02.fieldbyname('provincia').asstring := tabella_esa_03.fieldbyname('Provincia').asstring;

          tabella_02.fieldbyname('tna_codice').asstring := tabella_esa_03.fieldbyname('Nazione').asstring;
          if tabella_02.fieldbyname('tna_codice').asstring = '' then
          begin
            tabella_02.fieldbyname('tna_codice').asstring := 'IT';
          end;
          tabella_02.fieldbyname('partita_iva').asstring := tabella_esa_03.fieldbyname('Partita_Iva').asstring;
          tabella_02.fieldbyname('codice_fiscale').asstring := tabella_esa_03.fieldbyname('Codice_Fiscale').asstring;
          tabella_02.fieldbyname('telefono').asstring := tabella_esa_03.fieldbyname('Telefono').asstring;
          tabella_02.fieldbyname('fax').asstring := tabella_esa_03.fieldbyname('Fax').asstring;
          tabella_02.fieldbyname('cellulare').asstring := tabella_esa_03.fieldbyname('Num_Cellulare').asstring;
          if tabella_esa_03.fieldbyname('Flag_Persona_Fisica').asstring = '1' then
          begin
            tabella_02.fieldbyname('codice_alternativo').asstring := tabella_esa_03.fieldbyname('Ragione_Sociale').asString;

            tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_esa_03.fieldbyname('Cognome').asString);
            tabella_02.fieldbyname('descrizione2').asstring := trim(tabella_esa_03.fieldbyname('Nome').asString);

            tabella_02.fieldbyname('persona_fisica').asstring := 'si';
            if tabella_esa_03.fieldbyname('Flag_Sesso').asstring = '0' then
            begin
              tabella_02.fieldbyname('sesso').asstring := 'femminile';
            end
            else
            begin
              tabella_02.fieldbyname('sesso').asstring := 'maschile';
            end;

            data_nascita := Converti_data(tabella_esa_03.fieldbyname('Data_Nascita').asString);
            if data_nascita <> StrToDate('01/01/1900') then
              tabella_02.fieldbyname('data_nascita').asdatetime := Converti_data(tabella_esa_03.fieldbyname('Data_Nascita').asString);

            tabella_02.fieldbyname('citta_nascita').asstring := tabella_esa_03.fieldbyname('Luogo_di_Nascita').asstring;
            tabella_02.fieldbyname('provincia_nascita').asstring := tabella_esa_03.fieldbyname('Provincia_Nascita').asstring;
          end
          else
          begin
            tabella_02.fieldbyname('persona_fisica').asstring := 'no';
          end;

          if trim(tabella_clifor.fieldbyname('Codice_Valuta').asString) <> '' then
          begin
            tabella_02.fieldbyname('tva_codice').asstring := tabella_clifor.fieldbyname('Codice_Valuta').asString;
          end
          else
          begin
            tabella_02.fieldbyname('tva_codice').asstring := divisa_di_conto;
          end;

          tabella_02.fieldbyname('web').asstring := tabella_esa_03.fieldbyname('Sito_Internet').asstring;
          tabella_02.fieldbyname('e_mail_amministrazione').asstring := tabella_esa_03.fieldbyname('EMAIL').asstring;

          tabella_02.fieldbyname('via_legale').asstring := tabella_02.fieldbyname('via').asstring;
          tabella_02.fieldbyname('cap_legale').asstring := tabella_02.fieldbyname('cap').asstring;
          tabella_02.fieldbyname('citta_legale').asstring := tabella_02.fieldbyname('citta').asstring;
          tabella_02.fieldbyname('provincia_legale').asstring := tabella_02.fieldbyname('provincia').asstring;
          tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_02.fieldbyname('tna_codice').asstring;

          tabella_02.post;
        end; // if
      end;
      // ---------------------------------------------------------------------------
      // fornitori
      // ---------------------------------------------------------------------------
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := tabella_02.fieldbyname('codice').asstring;
      tabella_01.fieldbyname('descrizione1').asstring := trim(tabella_02.fieldbyname('descrizione1').asstring);
      tabella_01.fieldbyname('descrizione2').asstring := trim(tabella_02.fieldbyname('descrizione2').asstring);
      tabella_01.fieldbyname('via').asstring := tabella_02.fieldbyname('via').asstring;
      tabella_01.fieldbyname('citta').asstring := tabella_02.fieldbyname('citta').asstring;
      tabella_01.fieldbyname('partita_iva').asstring := tabella_02.fieldbyname('partita_iva').asstring;
      tabella_01.fieldbyname('codice_fiscale').asstring := tabella_02.fieldbyname('codice_fiscale').asstring;
      tabella_01.fieldbyname('tna_codice').asstring := tabella_02.fieldbyname('tna_codice').asstring;

      // ------------------------------------
      // default ditta
      // ------------------------------------
      tabella_01.fieldbyname('gen_codice').asstring := dit.fieldbyname('gen_codice_frn').asstring;
      tabella_01.fieldbyname('tba_codice').asstring := dit.fieldbyname('tba_codice_frn').asstring;
      tabella_01.fieldbyname('tpa_codice').asstring := dit.fieldbyname('tpa_codice_frn').asstring;
      tabella_01.fieldbyname('tcf_codice').asstring := dit.fieldbyname('tcf_codice_frn').asstring;
      tabella_01.fieldbyname('tla_codice').asstring := dit.fieldbyname('tla_codice_frn').asstring;
      tabella_01.fieldbyname('ts2_codice').asstring := dit.fieldbyname('ts2_codice_frn').asstring;
      tabella_01.fieldbyname('tzo_codice').asstring := dit.fieldbyname('tzo_codice_frn').asstring;
      tabella_01.fieldbyname('tsc_codice').asstring := dit.fieldbyname('tsc_codice_frn').asstring;
      tabella_01.fieldbyname('tsp_codice').asstring := dit.fieldbyname('tsp_codice_frn').asstring;
      tabella_01.fieldbyname('tpo_codice').asstring := dit.fieldbyname('tpo_codice_frn').asstring;

      tabella_01.fieldbyname('contatto').asstring := tabella_esa_03.fieldbyname('contatto').asstring;
      if tabella_clifor.fieldbyname('Sottoconto_Apparten').asstring <> '' then
      begin
        tabella_01.fieldbyname('gen_codice').asstring := tabella_clifor.fieldbyname('Sottoconto_Apparten').asstring;
      end;

      // if tabella_clifor.fieldbyname('Flag_Gest_Partite').asstring = '1' then
      // begin
      tabella_01.fieldbyname('partitario').asstring := 'si';
      // end
      // else
      // begin
      // tabella_01.fieldbyname('partitario').asstring := 'no';
      // end;

      if trim(tabella_clifor.fieldbyname('Codice_Pagamento').asString) <> '' then
      begin
        tabella_01.fieldbyname('tpa_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Pagamento').asString);
      end;

      if trim(tabella_clifor.fieldbyname('Vettore').asString) <> '' then
      begin
        tabella_01.fieldbyname('tsp_codice').asstring := trim(tabella_clifor.fieldbyname('Vettore').asString);
      end;

      if trim(tabella_clifor.fieldbyname('Porto').asstring) <> '' then
      begin
        tabella_01.fieldbyname('tpo_codice').asstring := trim(tabella_clifor.fieldbyname('Porto').asstring);
      end;

      if trim(tabella_clifor.fieldbyname('Codice_Nazione_Zona').asString) <> '' then
      begin
        tabella_01.fieldbyname('tzo_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Nazione_Zona').asString);
      end;

      if trim(tabella_clifor.fieldbyname('Codice_Iva_Esente').asString) <> '' then
      begin
        tabella_01.fieldbyname('tiv_codice').asstring := trim(tabella_clifor.fieldbyname('Codice_Iva_Esente').asString);
      end;

      if tabella_esa_02.locate('Codice_Banca', trim(tabella_clifor.fieldbyname('Codice_Banca').asstring), []) then
      begin
        tabella_01.fieldbyname('codice_abi').asstring := trim(tabella_clifor.fieldbyname('Codice_Banca').asstring);
        tabella_01.fieldbyname('codice_cab').asstring := trim(tabella_clifor.fieldbyname('Codice_Agenzia').asstring);
      end;
      if tabella_01.fieldbyname('codice_abi').asstring <> '' then
      begin
        if length(tabella_01.fieldbyname('codice_abi').asstring) = 4 then
        begin
          tabella_01.fieldbyname('codice_abi').asstring := '0' + tabella_01.fieldbyname('codice_abi').asstring;
        end;
      end;
      tabella_01.fieldbyname('mese_01').asInteger := 0;

      if TRIM(tabella_clifor.fieldbyname('Mese_1_Escluso_Pagam').asString) <> '' then
        tabella_01.fieldbyname('mese_01').asInteger := tabella_clifor.fieldbyname('Mese_1_Escluso_Pagam').asInteger;

      if tabella_01.fieldbyname('mese_01').asinteger <> 0 then
      begin
        if trim(tabella_clifor.fieldbyname('Giorno_Succ_x_Scaden').asstring) <> '' then
          tabella_01.fieldbyname('giorno_01').asinteger := tabella_clifor.fieldbyname('Giorno_Succ_x_Scaden').asinteger;
      end;
      tabella_01.fieldbyname('mese_02').asinteger := 0;

      if TRIM(tabella_clifor.fieldbyname('Mese_2_Escluso_Pagam').asString) <> '' then
        tabella_01.fieldbyname('mese_02').asinteger := tabella_clifor.fieldbyname('Mese_2_Escluso_Pagam').asinteger;

      if tabella_01.fieldbyname('mese_02').asinteger <> 0 then
      begin
        tabella_01.fieldbyname('giorno_02').asinteger := tabella_clifor.fieldbyname('Giorno_Succ_x_Scaden').asinteger;
      end;

      if (tabella_01.fieldbyname('mese_01').asinteger <> 0) or (tabella_01.fieldbyname('mese_02').asinteger <> 0) then
      begin
        tabella_01.fieldbyname('mesi_esclusi').asstring := 'si';
      end;

      if trim(tabella_clifor.fieldbyname('Numero_Listino').asstring) <> '' then
      begin
        tabella_01.fieldbyname('tla_codice').asstring := trim(tabella_clifor.fieldbyname('Numero_Listino').asstring);
      end;
      (*
        if tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring = '0' then
        begin
        tabella_01.fieldbyname('riepilogo_fattura').asstring := 'nessuno';
        end
        else if tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring = '1' then
        begin
        tabella_01.fieldbyname('riepilogo_fattura').asstring := 'globale';
        end;

        if (tabella_clifor.fieldbyname('Ind_Bolli_in_Fattura').asstring = '1') or
        (tabella_clifor.fieldbyname('Indic_Spese_Incasso').asstring = '1') then
        begin
        tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'si';
        end
        else
        begin
        tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'no';
        end;

        if (tabella_clifor.fieldbyname('Stampa Prezzo Bolla').asstring = '1') then
        begin
        tabella_01.fieldbyname('valori_in_bolla').asstring := 'si';
        end
        else
        begin
        tabella_01.fieldbyname('valori_in_bolla').asstring := 'no';
        end;

        tabella_01.fieldbyname('fido').asFloat := tabella_clifor.fieldbyname('fido').asFloat;
      *)
      tabella_01.fieldbyname('conto_corrente').asstring := trim(tabella_clifor.fieldbyname('NumContoC').asstring);
      if tabella_01.fieldbyname('conto_corrente').asstring <> '' then
      begin
        if length(tabella_01.fieldbyname('conto_corrente').asstring) <> 12 then
        begin
          for i := 1 to (12 - length(tabella_01.fieldbyname('conto_corrente').asstring)) do
          begin
            tabella_01.fieldbyname('conto_corrente').asstring := '0' + tabella_01.fieldbyname('conto_corrente').asstring;
          end;
        end;
      end;

      tabella_01.post;

      // file cfg
      if cfg.locate('cfg_tipo;cfg_codice', vararrayof(['F', tabella_01.fieldbyname('codice').asstring]), []) then
      begin
        cfg.edit;
        cfg.fieldbyname('descrizione1').asstring := trim(tabella_01.fieldbyname('descrizione1').asstring) + ' ' +
          trim(tabella_01.fieldbyname('descrizione2').asstring);
        cfg.fieldbyname('descrizione2').asstring := trim(tabella_01.fieldbyname('citta').asstring);
        cfg.fieldbyname('utente').asstring := utente;
        cfg.fieldbyname('data_ora').asdatetime := now;

        cfg.post;
      end;

    end; // if locate anagrafiche comuni

    tabella_clifor.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_clifor.close;
  tabella_esa_02.close;
  tabella_esa_03.close;
  cfg.close;

end;

procedure TCNVESA.converti_sottoconti;
begin
  v_tabella_01.caption := 'piano dei conti';
  application.processmessages;

  tabella_01.close;
  tabella_01.tablename := 'gen';
  tabella_01.open;

  cancella_tabella('gen');

  tabella_02.close;
  tabella_02.tablename := 'tpc';
  tabella_02.open;

  cancella_tabella('tpc');

  cfg.open;
  while not cfg.eof do
  begin
    if cfg.fieldbyname('cfg_tipo').asstring = 'G' then
    begin
      cfg.delete;
    end
    else
    begin
      cfg.next;
    end;
  end;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM PIANODEICONTI');
  tabella_esa_01.Sql.Add('WHERE CHIAVE =:Chiave');
  tabella_esa_01.Sql.Add('ORDER BY GRUPPO, CONTO, SOTTOCONTO');
  tabella_esa_01.Parameters.ParamByname('Chiave').Value := 'P';
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    if (trim(tabella_esa_01.fieldbyname('Sottoconto').asstring) <> '')
      and (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) <> '') then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring :=
        trim(tabella_esa_01.fieldbyname('Gruppo').asstring) +
        trim(tabella_esa_01.fieldbyname('Conto').asstring) +
        trim(tabella_esa_01.fieldbyname('Sottoconto').asstring);

      tabella_01.fieldbyname('descrizione1').asstring := trim(copy(tabella_esa_01.fieldbyname('descrizione').asstring, 1, 30));
      // tabella_01.fieldbyname('descrizione2').asstring := copy(tabella_esa_01.fieldbyname('pcdespia').asstring, 31, 10);
      tabella_01.fieldbyname('tpc_codice_01').asstring := trim(tabella_esa_01.fieldbyname('Gruppo').asstring);
      tabella_01.fieldbyname('tpc_codice_02').asstring := trim(tabella_esa_01.fieldbyname('Conto').asstring);

      tabella_01.post;

    end
    else
    begin
      tabella_02.append;

      tabella_02.fieldbyname('codice_01').asstring := trim(tabella_esa_01.fieldbyname('Gruppo').asstring);

      if trim(tabella_esa_01.fieldbyname('Conto').asstring) <> '' then
      begin
        tabella_02.fieldbyname('codice_02').asstring := trim(tabella_esa_01.fieldbyname('Conto').asstring);
        tabella_02.fieldbyname('tipo').asstring := '';
      end
      else
      begin
        if (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) = '01') or
          (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) = '02') or
          (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) = '03') or
          (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) = '06') or
          (trim(tabella_esa_01.fieldbyname('Gruppo').asstring) = '07') then
        begin
          tabella_02.fieldbyname('tipo').asstring := 'patrimoniale';
        end
        else
        begin
          tabella_02.fieldbyname('tipo').asstring := 'economico';
        end;
      end;
      tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('descrizione').asstring);

      tabella_02.post;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_articoli;
var
  gruppo,
    sottogru: string;
  tcc_codice,
    tcf_codice,
    codice_stat: string;
  tca_codice: Word;
begin
  v_tabella_01.caption := 'articoli';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_02;

  query_03.Close;
  query_03.SQL.Add('SELECT * FROM esa_tcm');
  query_03.SQL.Add('where esa_gruppo=:esa_gruppo');

  tabella_01.close;
  tabella_01.tablename := 'art';
  tabella_01.open;
  cancella_tabella('art');

  cancella_tabella('cpv');
  cpv.close;
  cpv.tablename := 'cpv';
  cpv.open;

  cancella_tabella('cpa');
  cpa.close;
  cpa.tablename := 'cpa';
  cpa.open;

  cancella_tabella('tca');
  tca.close;
  tca.tablename := 'tca';
  tca.open;

  tsa.close;
  tsa.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * from ANAGRAFICAARTICOLI');
    tabella_esa_02.Sql.Add('order by Codice_Articolo');
    tabella_esa_02.open;

  except
    messaggio(000, 'manca la tabella delle banche (BANCHECLIENTI)');
    close;
    abort;
  end;

  query_02.sql.clear;
  query_02.sql.add('select * from cpv inner join cpa on cpv.tca_codice = cpa.taq_codice');
  query_02.sql.add('where cpv.gen_codice = :cpv_gen_codice and cpa.gen_codice = :cpa_gen_codice');

  tcc_codice := dit.fieldbyname('tcc_codice_cli').asstring;
  tcf_codice := dit.fieldbyname('tcf_codice_frn').asstring;
  tca_codice := 0;

  while not tabella_esa_02.eof do
  begin

    if tabella_esa_02.fieldbyname('Codice_Articolo').asstring <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('tum_codice').asstring := dit.fieldbyname('tum_codice_art').asstring;
      tabella_01.fieldbyname('tub_codice').asstring := dit.fieldbyname('tub_codice_art').asstring;
      tabella_01.fieldbyname('tiv_codice_vendite').asstring := dit.fieldbyname('tiv_codice_vendite_art').asstring;
      tabella_01.fieldbyname('tiv_codice_acquisti').asstring := dit.fieldbyname('tiv_codice_acquisti_art').asstring;
      tabella_01.fieldbyname('tca_codice').asstring := dit.fieldbyname('tca_codice_art').asstring;
      tabella_01.fieldbyname('taq_codice').asstring := dit.fieldbyname('taq_codice_art').asstring;
      tabella_01.fieldbyname('tni_codice').asstring := dit.fieldbyname('tni_codice_art').asstring;
      tabella_01.fieldbyname('tcm_codice').asstring := dit.fieldbyname('tcm_codice_art').asstring;
      tabella_01.fieldbyname('tgm_codice').asstring := dit.fieldbyname('tgm_codice_art').asstring;
      tabella_01.fieldbyname('tin_codice').asstring := dit.fieldbyname('tin_codice_art').asstring;
      tabella_01.fieldbyname('ts3_codice').asstring := dit.fieldbyname('ts3_codice_art').asstring;
      tabella_01.fieldbyname('tp2_codice').asstring := dit.fieldbyname('tp2_codice_art').asstring;
      tabella_01.fieldbyname('tsa_codice').asstring := dit.fieldbyname('tsa_codice_art').asstring;
      tabella_01.fieldbyname('taa_codice').asstring := dit.fieldbyname('taa_codice_art').asstring;

      tabella_01.fieldbyname('codice').asstring := tabella_esa_02.fieldbyname('Codice_Articolo').asstring;
      tabella_01.fieldbyname('descrizione1').asstring := trim(tabella_esa_02.fieldbyname('DescrizPrimariaArtic').asstring);
      if tabella_01.fieldbyname('descrizione1').asstring = '' then
      begin
        tabella_01.fieldbyname('descrizione1').asstring := '.';
      end;
      tabella_01.fieldbyname('descrizione2').asstring := trim(tabella_esa_02.fieldbyname('DescrSecond_articolo').asstring);
      tabella_01.fieldbyname('tum_codice').asstring := tabella_esa_02.fieldbyname('Unita_Misura_Princip').asstring;

      gruppo := trim(tabella_esa_02.fieldbyname('Gruppo_Merceologico').asstring);
      sottogru := trim(tabella_esa_02.fieldbyname('SottoGruppoMerceolog').asstring);
      codice_stat := trim(tabella_esa_02.fieldbyname('codice_statistico').AsString);

      if codice_stat <> '' then
      begin

        (*
          if not tsa.Locate('codice', codice_stat, []) then
          begin
          tsa.Append;
          tsa.FieldByName('codice').asstring := codice_stat;
          tsa.FieldByName('descrizione').asstring := codice_stat;
          tsa.post
          end;
        *)

        tabella_01.FieldByName('codice_alternativo').asstring := codice_stat;

      end;

      if (gruppo <> '') then
      begin
        query_03.close;
        query_03.parambyname('esa_gruppo').asstring := gruppo;
        query_03.open;
        if not query_03.eof then
          tabella_01.fieldbyname('tcm_codice').asstring := query_03.fieldbyname('codice').asstring;
      end;

      if (gruppo + sottogru <> '') then
      begin

        tabella_01.fieldbyname('tgm_codice').asstring := tabella_01.fieldbyname('tcm_codice').asstring + sottogru;
      end;
      tabella_01.fieldbyname('tiv_codice_vendite').asstring := copy(tabella_esa_02.fieldbyname('Codice_Iva').asstring, 2, 2);
      tabella_01.fieldbyname('tiv_codice_acquisti').asstring := copy(tabella_esa_02.fieldbyname('Codice_Iva').asstring, 2, 2);

      tabella_01.fieldbyname('tiv_codice_vendite').asstring := copy(tabella_esa_02.fieldbyname('Codice_Iva').asstring, 2, 2);
      tabella_01.fieldbyname('tiv_codice_acquisti').asstring := copy(tabella_esa_02.fieldbyname('Codice_Iva').asstring, 2, 2);

      query_02.params[0].asstring := trim(tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring);
      if trim(tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring) = '' then
      begin
        query_02.params[0].asstring := trim(v_ricavi.text);
      end;
      query_02.params[1].asstring := tabella_esa_02.fieldbyname('Contropartita_Costo').asstring;
      if trim(tabella_esa_02.fieldbyname('Contropartita_Costo').asstring) = '' then
      begin
        query_02.params[1].asstring := trim(v_acquisti.text);
      end;
      query_02.close;
      query_02.open;
      if query_02.eof then
      begin
        tca_codice := tca_codice + 1;

        tca.append;

        tca.fieldbyname('codice').asstring := setta_lunghezza(tca_codice, 4, 0);
        tca.fieldbyname('descrizione').asstring := 'vendite: ' + trim
          (tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring) +
          '  /  acquisti: ' + tabella_esa_02.fieldbyname('Contropartita_Costo').asstring;
        tca.fieldbyname('utente').asstring := utente;
        tca.fieldbyname('data_ora').asdatetime := now;

        tca.post;

        tabella_01.fieldbyname('tca_codice').asstring := tca.fieldbyname('codice').asstring;

        cpv.append;

        cpv.fieldbyname('tca_codice').asstring := setta_lunghezza(tca_codice, 4, 0);
        cpv.fieldbyname('tcc_codice').asstring := tcc_codice;
        cpv.fieldbyname('gen_codice').asstring := trim(tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring);
        cpv.fieldbyname('gen_codice_omaggi').asstring := cpv.fieldbyname('gen_codice').asstring;
        cpv.fieldbyname('gen_codice_sconti').asstring := cpv.fieldbyname('gen_codice').asstring;
        cpv.fieldbyname('utente').asstring := utente;
        cpv.fieldbyname('data_ora').asdatetime := now;

        cpv.post;

        cpa.append;

        cpa.fieldbyname('taq_codice').asstring := setta_lunghezza(tca_codice, 4, 0);
        cpa.fieldbyname('tcf_codice').asstring := tcf_codice;
        cpa.fieldbyname('gen_codice').asstring := trim(tabella_esa_02.fieldbyname('Contropartita_Costo').AsString);
        cpa.fieldbyname('gen_codice_omaggi').asstring := cpa.fieldbyname('gen_codice').asstring;
        cpa.fieldbyname('gen_codice_sconti').asstring := cpa.fieldbyname('gen_codice').asstring;
        cpa.fieldbyname('utente').asstring := utente;
        cpa.fieldbyname('data_ora').asdatetime := now;

        cpa.post;
      end
      else
      begin
        tabella_01.fieldbyname('tca_codice').asstring := query_02.fieldbyname('tca_codice').asstring;
      end;
      tabella_01.post;

      (*
        if trim(tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring) <> '' then
        begin

        if not tabella_02.locate('tca_codice;tcc_codice',
        vararrayof([tabella_01.fieldByName('tca_codice').AsString, '0']), []) then
        begin
        tabella_02.append;
        tabella_02.fieldvalues['tca_codice'] := tabella_01.fieldByName('tca_codice').AsString;
        tabella_02.fieldvalues['tcc_codice'] := '0';
        tabella_02.fieldvalues['gen_codice'] := trim(tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring);
        tabella_02.fieldvalues['gen_codice_omaggi'] := tabella_02.fieldvalues['gen_codice'];
        tabella_02.fieldvalues['gen_codice_sconti'] := tabella_02.fieldvalues['gen_codice'];
        tabella_02.post;

        end;

        end;

        if trim(tabella_esa_02.fieldbyname('Contropartita_Costo').asstring) <> '' then
        begin
        if not tabella_dettaglio.locate(
        'tcf_codice;tca_codice',
        vararrayof(['0', tabella_01.fieldByName('tca_codice').AsString]), []) then
        begin
        tabella_dettaglio.append;
        tabella_dettaglio.fieldvalues['tcf_codice'] := '0';
        tabella_dettaglio.fieldvalues['tca_codice'] := tabella_01.fieldByName('tca_codice').AsString;
        tabella_dettaglio.fieldvalues['gen_codice'] := trim(tabella_esa_02.fieldbyname('Contropartita_Costo').asstring);
        tabella_dettaglio.fieldvalues['gen_codice_omaggi'] := tabella_dettaglio.fieldvalues['gen_codice'];
        tabella_dettaglio.fieldvalues['gen_codice_sconti'] := tabella_dettaglio.fieldvalues['gen_codice'];
        tabella_dettaglio.post;

        end;

        end;

      *)
    end; // if

    tabella_esa_02.next;
  end; // while

  query.close;
  query_02.close;
  cpa.close;
  cpv.close;
  tabella_01.close;
  tabella_esa_02.close;
  tsa.close;
end;

procedure TCNVESA.converti_bar;
begin
  v_tabella_01.caption := 'barcode articoli';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_02;

  cancella_tabella('bar');
  tabella_01.close;
  tabella_01.tablename := 'bar';
  tabella_01.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * FROM CODICIAGGIUNTIVI');
    tabella_esa_02.Sql.Add('order by Codice_Articolo');
    tabella_esa_02.open;

  except
    messaggio(000, 'manca la tabella barcode articoli (CODICIAGGIUNTIVI)');
    close;
    abort;
  end;

  while not tabella_esa_02.eof do
  begin

    if not tabella_01.locate('art_codice;codice_barre', Vararrayof([
      trim(tabella_esa_02.fieldbyname('codice_articolo').asstring), trim(tabella_esa_02.fieldbyname('CodArticoloAggiuntiv').asstring)]), []) then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('art_codice').asstring := trim(tabella_esa_02.fieldbyname('codice_articolo').asstring);
      tabella_01.fieldbyname('codice_barre').asstring := trim(tabella_esa_02.fieldbyname('CodArticoloAggiuntiv').asstring);

      tabella_01.post;
    end;

    tabella_esa_02.next;
  end;

  tabella_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_tmo;
begin
  v_tabella_01.caption := 'causali magazzino';
  application.processmessages;

  tabella_01_ds.Dataset := tabella_02;

  cancella_tabella('tmo');
  tabella_01.close;
  tabella_01.tablename := 'tmo';
  tabella_01.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * from CAUSALIDIMAGAZZINO');
    tabella_esa_02.Sql.Add('where');
    tabella_esa_02.Sql.Add('Parte_Fissa=:PF');
    tabella_esa_02.Sql.Add('order by Causale_di_Magazzino');

    tabella_esa_02.Parameters.ParamByName('PF').Value := 'CM';
    tabella_esa_02.open;
  except
    messaggio(000, 'manca la tabella causali di magazzino (CAUSALIDIMAGAZZINO)');
    close;
    abort;
  end;

  while not tabella_esa_02.eof do
  begin
    if tabella_esa_02.fieldbyname('Causale_di_Magazzino').asstring <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_02.fieldbyname('Causale_di_Magazzino').asstring);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_02.fieldbyname('Descrizione_Causale').asstring);
      if tabella_01.fieldbyname('descrizione').asstring = '' then
      begin
        tabella_01.fieldbyname('descrizione').asstring := '.';
      end;

      if (tabella_esa_02.fieldbyname('Causale_collegata').asstring <> '000') and
        (tabella_esa_02.fieldbyname('Causale_collegata').asstring <> '') then
        tabella_01.fieldbyname('tmo_codice_collegato').asstring := tabella_esa_02.fieldbyname('Causale_collegata').asstring;
      (*
        if tabella_esa_02.fieldbyname('cmdatcar').asstring = '=' then
        begin
        tabella_01.fieldbyname('ultimi_valori').asstring := 'carico';
        end;

        if tabella_esa_02.fieldbyname('cmdatsca').asstring = '=' then
        begin
        tabella_01.fieldbyname('ultimi_valori').asstring := 'scarico';
        end;

        if tabella_esa_02.fieldbyname('cmqtacar').asstring = '+' then
        begin
        tabella_01.fieldbyname('esistenza').asstring := 'incrementa';
        end
        else if tabella_esa_02.fieldbyname('cmqtacar').asstring = '-' then
        begin
        tabella_01.fieldbyname('esistenza').asstring := 'decrementa';
        end
        else if tabella_esa_02.fieldbyname('cmqtasca').asstring = '+' then
        begin
        tabella_01.fieldbyname('esistenza').asstring := 'decrementa';
        end
        else if tabella_esa_02.fieldbyname('cmqtasca').asstring = '-' then
        begin
        tabella_01.fieldbyname('esistenza').asstring := 'incrementa';
        end
        else
        begin
        tabella_01.fieldbyname('esistenza').asstring := 'ignora';
        end;

        if (tabella_esa_02.fieldbyname('cmflaval').asstring = 'S')
        and (tabella_esa_02.fieldbyname('cmqtacar').asstring = '+')
        and (tabella_esa_02.fieldbyname('cmvalcar').asstring = '+') then
        begin
        tabella_01.fieldbyname('valorizzazione').asstring := 'incrementa';
        end
        else if (tabella_esa_02.fieldbyname('cmflaval').asstring = 'S')
        and (tabella_esa_02.fieldbyname('cmqtacar').asstring = '-')
        and (tabella_esa_02.fieldbyname('cmvalcar').asstring = '-') then
        begin
        tabella_01.fieldbyname('valorizzazione').asstring := 'decrementa';
        end;
      *)
      tabella_01.post;
    end;

    tabella_esa_02.next;
  end;

  tabella_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_tpa;
var
  i, riga: integer;
  nome_campo: string;

begin
  v_tabella_01.caption := 'codici pagamento';
  application.processmessages;

  tabella_01_ds.Dataset := tabella_02;

  cancella_tabella('tpa');
  tabella_01.close;
  tabella_01.tablename := 'tpa';
  tabella_01.open;

  try
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT * from PAGAMENTITESTATA');
    tabella_esa_01.Sql.Add('order by codice_pagamento');
    tabella_esa_01.open;
  except
    messaggio(000, 'manca la tabella patamenti testata (PAGAMENTITESTATA)');
    close;
    abort;
  end;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * from PAGAMENTIRIGA');
    tabella_esa_02.Sql.Add('order by 2,3');
    tabella_esa_02.open;
  except
    messaggio(000, 'manca la tabella causali di magazzino (PAGAMANETIRIGA)');
    close;
    abort;
  end;

  while not tabella_esa_01.eof do
  begin
    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('Codice_Pagamento').asString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('DescrizionePagamento').asstring);
    if tabella_01.fieldbyname('descrizione').asstring = '' then
    begin
      tabella_01.fieldbyname('descrizione').asstring := '.';
    end;

    tabella_01.fieldbyname('numero_rate').asinteger := tabella_esa_01.fieldbyname('N_Rate').asinteger;
    tabella_01.fieldbyname('tipo_rate').asstring := 'variabili';
    tabella_01.fieldbyname('giorni_prima_rata_fisse').asinteger := 0;
    tabella_01.fieldbyname('giorni_rate_fisse').asinteger := 0;

    if tabella_esa_02.locate('Codice_Pagamento', trim(tabella_esa_01.fieldbyname('Codice_Pagamento').asString), []) then
    begin
      riga := 0;
      while not tabella_esa_02.Eof and
        (trim(tabella_esa_02.fieldbyname('Codice_Pagamento').asString) = trim(tabella_esa_01.fieldbyname('Codice_Pagamento').asString)) do
      begin
        riga := riga + 1;

        if (riga <= tabella_01.fieldbyname('numero_rate').asinteger) and (riga < 9) then
        begin
          nome_campo := 'tts_codice_variabili_0' + IntToStr(riga);

          if (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '3') or
            (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '4') or
            (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '7') or
            (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '8') then
            tabella_01.fieldbyname(nome_campo).asString := 'rimessa diretta'
          else if (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '1') then
            tabella_01.fieldbyname(nome_campo).asstring := 'R.I.B.A.'
          else if (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '6') then
            tabella_01.fieldbyname(nome_campo).asstring := 'bonifico bancario'
          else if (trim(tabella_esa_02.fieldbyname('Tipo_Rata').asString) = '9') then
            tabella_01.fieldbyname(nome_campo).asstring := 'R.I.D.';

          // fine mese
          nome_campo := 'fine_mese_variabili_0' + IntToStr(riga);

          if (trim(tabella_esa_02.fieldbyname('Tipo_Scadenza').asString)) = '99' then
            tabella_01.fieldbyname(nome_campo).asstring := 'si';

          // perc da pag
          nome_campo := 'percentuale_variabili_0' + IntToStr(riga);
          tabella_01.fieldbyname(nome_campo).asfloat := tabella_esa_02.fieldbyname('Perc_imponibile').asFloat;

          // rata
          nome_campo := 'giorni_variabili_0' + IntToStr(riga);
          tabella_01.fieldbyname(nome_campo).asfloat := tabella_esa_02.fieldbyname('Giorni_Scadenza').asFloat;

        end; // if

        tabella_esa_02.next;
      end; // while

      riga := riga + 1;
      if riga < 8 then
      begin

        for i := riga to 8 do
        begin
          nome_campo := 'tts_codice_variabili_0' + IntToStr(i);
          tabella_01.fieldbyname(nome_campo).asString := ''
        end; // for
      end;

    end; // if
    (*
      if pos('DDFM', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('fine_mese_fisse').asstring := 'si';
      end
      else
      begin
      tabella_01.fieldbyname('fine_mese_fisse').asstring := 'no';
      end;

      if pos('RD', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('tipo_scadenza_fisse').asstring := 'rimessa diretta';
      end
      else if pos('R.D', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('tipo_scadenza_fisse').asstring := 'rimessa diretta';
      end
      else if pos('RB', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('tipo_scadenza_fisse').asstring := 'R.I.B.A.';
      end
      else if pos('R.B', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('tipo_scadenza_fisse').asstring := 'R.I.B.A.';
      end
      else if pos('BB', tabella_esa_02.fieldbyname('DescrizionePagamento').asstring) > 0 then
      begin
      tabella_01.fieldbyname('tipo_scadenza_fisse').asstring := 'bonifico bancario';
      end;
    *)
    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_ind_inf;
begin
  v_tabella_01.caption := 'indirizzi spedizione';
  application.processmessages;

  cancella_tabella('ind');
  tabella_01.close;
  tabella_01.tablename := 'ind';
  tabella_01.open;

  cancella_tabella('inf');
  tabella_02.close;
  tabella_02.tablename := 'inf';
  tabella_02.open;

  try
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT * from SEDEAMMINISTRATIVA');
    tabella_esa_01.Sql.Add('WHERE');
    tabella_esa_01.Sql.Add('Ind_Sede_Amm_Diversa =' + quotedstr('D'));
    tabella_esa_01.Sql.Add('order by 2');
    tabella_esa_01.open;
  except
    raise;
    messaggio(000, 'manca la tabella INDIRIZZI DI SPEDIZIONE (SEDEAMMINISTRATIVA)');
    close;
    abort;
  end;

  tabella_clienti_forn.close;
  tabella_clienti_forn.prepared := True;

  while not tabella_esa_01.eof do
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('codice_anagra_Clifor').asstring, cli_for);

    tabella_clienti_forn.close;
    tabella_clienti_forn.parameters.parambyname('codice_nom').value := tabella_esa_01.fieldbyname('codice_anagra_Clifor').asstring;
    tabella_clienti_forn.open;

    if tabella_clienti_forn.fieldbyname('Ind_clienteFornitore').asstring = 'C' then
    begin

      if read_tabella(arc.arcdit, 'cli', 'codice', cli_for) then
      begin
        tabella_01.append;

        tabella_01.fieldbyname('cli_codice').asstring := cli_for;
        tabella_01.fieldbyname('indirizzo').asstring := trim(tabella_esa_01.fieldbyname('codice_sede').asstring);
        tabella_01.fieldbyname('descrizione1').asstring := trim(tabella_esa_01.fieldbyname('descrizione_sede').asstring);
        if tabella_01.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_01.fieldbyname('descrizione1').asstring := trim(archivio.fieldbyname('descrizione1').asstring);
          tabella_01.fieldbyname('descrizione2').asstring := trim(archivio.fieldbyname('descrizione2').asstring);
        end;
        tabella_01.fieldbyname('via').asstring := trim(tabella_esa_01.fieldbyname('indirizzo_sede').asstring);
        tabella_01.fieldbyname('cap').asstring := trim(tabella_esa_01.fieldbyname('cap').asstring);
        tabella_01.fieldbyname('citta').asstring := trim(tabella_esa_01.fieldbyname('localita').asstring);
        tabella_01.fieldbyname('provincia').asstring := trim(tabella_esa_01.fieldbyname('provincia').asstring);
        tabella_01.fieldbyname('telefono').asstring := trim(tabella_esa_01.fieldbyname('numero_telefono_sede').asstring);

        tabella_01.fieldbyname('tna_codice').asstring := archivio.fieldbyname('tna_codice').asstring;
        tabella_01.fieldbyname('tzo_codice').asstring := archivio.fieldbyname('tzo_codice').asstring;
        tabella_01.fieldbyname('tsp_codice').asstring := archivio.fieldbyname('tsp_codice').asstring;
        tabella_01.fieldbyname('tpo_codice').asstring := archivio.fieldbyname('tpo_codice').asstring;
        tabella_01.fieldbyname('tag_codice').asstring := archivio.fieldbyname('tag_codice').asstring;
        tabella_01.fieldbyname('tst_codice').asstring := archivio.fieldbyname('tst_codice').asstring;
        tabella_01.fieldbyname('tzo_codice_assistenza').asstring := archivio.fieldbyname('tzo_codice').asstring;

        tabella_01.post;
      end;
    end
    else if tabella_clienti_forn.fieldbyname('Ind_clienteFornitore').asstring = 'F' then
    begin

      assegna_codice_frn(tabella_esa_01.fieldbyname('Codice_anagra_Clifor').asstring, cli_for);
      if read_tabella(arc.arcdit, 'frn', 'codice', cli_for) then
      begin
        tabella_02.append;

        tabella_02.fieldbyname('frn_codice').asstring := cli_for;
        tabella_02.fieldbyname('indirizzo').asstring := trim(tabella_esa_01.fieldbyname('codice_sede').asstring);
        tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_esa_01.fieldbyname('descrizione_sede').asstring);
        if tabella_02.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_02.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
          tabella_02.fieldbyname('descrizione2').asstring := archivio.fieldbyname('descrizione2').asstring;
        end;
        tabella_02.fieldbyname('via').asstring := trim(tabella_esa_01.fieldbyname('indirizzo_sede').asstring);
        tabella_02.fieldbyname('cap').asstring := trim(tabella_esa_01.fieldbyname('cap').asstring);
        tabella_02.fieldbyname('citta').asstring := trim(tabella_esa_01.fieldbyname('localita').asstring);
        tabella_02.fieldbyname('provincia').asstring := trim(tabella_esa_01.fieldbyname('provincia').asstring);
        tabella_02.fieldbyname('telefono').asstring := trim(tabella_esa_01.fieldbyname('numero_telefono_sede').asstring);

        tabella_02.fieldbyname('tna_codice').asstring := archivio.fieldbyname('tna_codice').asstring;
        tabella_02.fieldbyname('tsp_codice').asstring := archivio.fieldbyname('tsp_codice').asstring;
        tabella_02.fieldbyname('tpo_codice').asstring := archivio.fieldbyname('tpo_codice').asstring;

        tabella_02.post;
      end;
    end;
    tabella_esa_01.next;
  end;

  tabella_clienti_forn.close;
  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;

end;

procedure TCNVESA.converti_lsv;
var
  articolo, listino: string;
begin
  v_tabella_01.caption := 'listini di vendita';
  application.processmessages;

  cancella_tabella('lsv');
  tabella_01.close;
  tabella_01.tablename := 'lsv';
  tabella_01.open;

  cancella_tabella('tlv');
  tabella_02.close;
  tabella_02.tablename := 'tlv';
  tabella_02.open;

  tabella_01_ds.dataset := tabella_01;
  try
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT Listino, Articolo, Prezzo, Sconto_1, Sconto_2 FROM PREZZILISTINO');
    tabella_esa_01.Sql.Add('order by listino, Articolo');
    tabella_esa_01.open;
  except

    messaggio(000, 'manca la tabella dei listini (PREZZILISTINO)');

    close;
    abort;
  end;

  try
    tabella_esa_02.Close;
    tabella_esa_02.Sql.Clear;
    tabella_esa_02.Sql.Add('SELECT * FROM ANAGRAFICALISTINI');
    tabella_esa_02.Sql.Add('order by Codice_Listino');
    tabella_esa_02.open;
  except

    messaggio(000, 'manca la tabella anagrafic listini');

    close;
    abort;
  end;

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin

    articolo := trim(tabella_esa_01.fieldbyname('articolo').asstring);
    listino := trim(tabella_esa_01.fieldbyname('listino').asstring);

    while
      not(tabella_esa_01.eof) and
      (articolo = trim(tabella_esa_01.fieldbyname('articolo').asstring)) and
      (listino = trim(tabella_esa_01.fieldbyname('listino').asstring)) do
    begin

      if read_tabella(arc.arcdit, 'art', 'codice', tabella_esa_01.fieldbyname('articolo').asstring) then
      begin
        if not tabella_01.locate('art_codice;tlv_codice',
          VarArrayOf([
          trim(tabella_esa_01.fieldbyname('Articolo').asstring),
          trim(tabella_esa_01.fieldbyname('listino').asstring)])
          , []) then
        begin
          tabella_01.append;

          tabella_01.fieldbyname('art_codice').asstring := tabella_esa_01.fieldbyname('Articolo').asstring;
          tabella_01.fieldbyname('tlv_codice').asstring := tabella_esa_01.fieldbyname('listino').asstring;
          tabella_01.fieldbyname('data_inizio').asdatetime := StrToDate('01/01/1900');
          tabella_01.fieldbyname('data_fine').asdatetime := StrToDate('31/12/9999');
          tabella_01.fieldbyname('prezzo').asfloat := tabella_esa_01.fieldbyname('prezzo').asfloat;

          if not((tabella_esa_01.fieldbyname('Sconto_1').asfloat = 0) and (tabella_esa_01.fieldbyname('Sconto_2').asfloat = 0)) then
          begin
            query.params[0].asfloat := tabella_esa_01.fieldbyname('Sconto_1').asfloat;
            query.params[1].asfloat := tabella_esa_01.fieldbyname('Sconto_2').asfloat;
            query.close;
            query.open;
            if not query.eof then
            begin
              tabella_01.fieldbyname('tsm_codice').asstring := query.fieldbyname('codice').asstring;
            end
            else
            begin
              crea_tsm(
                tabella_esa_01.fieldbyname('Sconto_1').asfloat,
                tabella_esa_01.fieldbyname('Sconto_2').asfloat);
              tabella_01.fieldbyname('tsm_codice').asstring := setta_lunghezza(tsm_codice, 4, 0);
            end;
          end;

          tabella_01.post;
        end; // if

      end;

      tabella_esa_01.next;
    end; // while

    while
      not(tabella_esa_02.eof) do
    begin
      tabella_02.append;

      tabella_02.fieldbyname('codice').asstring := trim(tabella_esa_02.fieldbyname('Codice_listino').asstring);
      tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_02.fieldbyname('Descrizione_listino').asstring);
      tabella_02.fieldbyname('tva_codice').asstring := divisa_di_conto;

      tabella_02.post;

      tabella_esa_02.next;
    end;

  end;

  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_pnt;
var
  i: integer;
  progressivo: integer;
begin
  v_tabella_01.caption := 'primanota';
  application.processmessages;

  // arc.attivazione_trigger(arc.arcdit, false, false);

  query.sql.clear;
  query.sql.add('delete from cfg ');
  query.sql.add('where ');
  query.sql.add('cfg_tipo =' + quotedstr('G'));
  query.ExecSQL;

  query.sql.clear;
  query.sql.add('delete from cfgese ');
  query.sql.add('where ');
  query.sql.add('cfg_tipo =' + quotedstr('G'));
  query.ExecSQL;

  cancella_tabella('pni');
  tabella_03.close;
  tabella_03.tablename := 'pni';
  tabella_03.open;

  cancella_tabella('pnr');
  tabella_02.close;
  tabella_02.tablename := 'pnr';
  tabella_02.open;

  cancella_tabella('pnt');
  tabella_01.close;
  tabella_01.tablename := 'pnt';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * from primanota');
  // tabella_esa_01.Sql.Add('where data_registrazione >=''20122131'' ');
  tabella_esa_01.Sql.Add('order by data_registrazione,numero_registrazione');
  tabella_esa_01.open;

  test_datas := '';
  test_numero_stringa := '';
  progressivo := 0;

  while not tabella_esa_01.eof do
  begin
    Application.ProcessMessages;
    v_tabella_01.caption := 'primanota ' + tabella_esa_01.fieldbyname('data_registrazione').asstring;

    if ((test_datas <> tabella_esa_01.fieldbyname('data_registrazione').asstring) or
      (test_numero_stringa <> tabella_esa_01.fieldbyname('numero_registrazione').asstring)) then
    begin
      codice_clifor := '';
      progressivo := progressivo + 1;
      crea_pnt(progressivo);
      test_datas := tabella_esa_01.fieldbyname('data_registrazione').asstring;
      test_numero_stringa := tabella_esa_01.fieldbyname('numero_registrazione').asstring;
    end;

    crea_pnr(progressivo);

    if (tabella_01.fieldbyname('cfg_tipo').asstring = 'C') or (tabella_01.fieldbyname('cfg_tipo').asstring = 'F') and
      (tabella_01.fieldbyname('cfg_codice').asstring = '') then
    begin
      try
        IF codice_clifor = '' then
        begin
          codice_clifor := '99999999';
        end;
        tabella_01.edit;
        tabella_01.fieldbyname('cfg_codice').asstring := codice_clifor;
        tabella_01.post;
      except
      end;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_03.close;
  tabella_esa_01.close;

  // arc.attivazione_trigger(arc.arcdit, false, TRUE);
end;

procedure TCNVESA.crea_pnt(progressivo: integer);
begin

  tabella_01.append;

  //
  // dtacom, codca2, numreg, fcfiva,
  //

  tabella_01.fieldbyname('data_registrazione').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_registrazione').asstring);
  tabella_01.fieldbyname('progressivo').asinteger := progressivo;
  tabella_01.fieldbyname('tco_codice').asstring := tabella_esa_01.fieldbyname('codice_causale').asstring;

  // ------------------------------
  read_tabella(arc.arcdit, 'tco', 'codice', tabella_01.fieldbyname('tco_codice').asstring);

  tabella_01.fieldbyname('tco_descrizione').asstring := archivio.fieldbyname('descrizione').asstring;
  tabella_01.fieldbyname('documento_iva').asstring := archivio.fieldbyname('movimento_iva').asstring;
  tabella_01.fieldbyname('tipo_documento_iva').asstring := archivio.fieldbyname('tipo_registro_iva').asstring;
  tabella_01.fieldbyname('descrizione').asstring := tabella_esa_01.fieldbyname('numero_registrazione').asstring;
  // ------------------------------

  tabella_01.fieldbyname('numero_documento').asinteger := tabella_esa_01.fieldbyname('Numero_documento').asinteger;
  if tabella_esa_01.fieldbyname('data_documento').asstring <> '' then
  begin
    tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_documento').asstring);
  end
  else
  begin
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_01.fieldbyname('data_registrazione').asdatetime;
  end;

  if tabella_01.fieldbyname('data_documento').asdatetime = strtodate('01/01/1900') then
  begin
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_01.fieldbyname('data_registrazione').asdatetime;
  end;

  if tabella_esa_01.fieldbyname('data_competenza').asstring <> '' then
  begin
    tabella_01.fieldbyname('data_competenza_iva').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_competenza').asstring);
  end
  else
  begin
    tabella_01.fieldbyname('data_competenza_iva').asdatetime := tabella_01.fieldbyname('data_registrazione').asdatetime;
  end;

  tabella_01.fieldbyname('protocollo').asinteger := tabella_esa_01.fieldbyname('numero_di_protocollo').asinteger;
  if tabella_01.fieldbyname('documento_iva').asstring = 'si' then
  begin
    if archivio.fieldbyname('tipo_registro_iva').asstring = 'vendite' then
    begin
      tabella_01.fieldbyname('cfg_tipo').asstring := 'C';
      assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);
      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
      if archivio.fieldbyname('serie_numerazione').asstring <> '0' then
      begin
        tabella_01.fieldbyname('serie_documento').asstring := archivio.fieldbyname('serie_numerazione').asstring;
      end;
    end
    else if archivio.fieldbyname('tipo_registro_iva').asstring = 'acquisti' then
    begin
      tabella_01.fieldbyname('cfg_tipo').asstring := 'F';
      assegna_codice_frn(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);
      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
      tabella_01.fieldbyname('serie_documento').asstring := trim(tabella_esa_01.fieldbyname('Serie_Num_documento').asstring);

      if archivio.fieldbyname('serie_numerazione').asstring <> '0' then
      begin
        tabella_01.fieldbyname('serie_protocollo').asstring := archivio.fieldbyname('serie_numerazione').asstring;
      end;
    end
    else if archivio.fieldbyname('tipo_registro_iva').asstring = 'corrispettivi' then
    begin
      tabella_01.fieldbyname('cfg_tipo').asstring := 'G';
      tabella_01.fieldbyname('cfg_codice').asstring := tabella_esa_01.fieldbyname('Conto_Prima_Nota').asstring;
    end;
  end;
  tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;
  tabella_01.fieldbyname('ese_codice').asstring := inttostr(yearof(tabella_01.fieldbyname('data_competenza_iva').asdatetime));
  tabella_01.fieldbyname('cambio').asfloat := 1;
  try

    tabella_01.post;
  except
  end;
  riga := 0;
  riga_iva := 0;
end;

procedure TCNVESA.crea_pnr(progressivo: integer);
begin

  tabella_02.append;

  tabella_02.fieldbyname('progressivo').asfloat := progressivo;
  riga := riga + 1;

  tabella_02.fieldbyname('riga').asinteger := riga;

  // ----------------------------------------------------------
  // il cfg_ese viene inserito nel trigger
  // ----------------------------------------------------------

  if tabella_esa_01.fieldbyname('Ind_ClienteFornitore').asstring = 'C' then
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);

    tabella_02.fieldbyname('cfg_tipo').asstring := 'C';
    tabella_02.fieldbyname('cfg_codice').asstring := cli_for;

    if codice_clifor = '' then
    begin
      codice_clifor := cli_for;
    end;

  end
  else if tabella_esa_01.fieldbyname('Ind_ClienteFornitore').asstring = 'F' then
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);
    tabella_02.fieldbyname('cfg_tipo').asstring := 'F';
    tabella_02.fieldbyname('cfg_codice').asstring := cli_for;
    if codice_clifor = '' then
    begin
      codice_clifor := cli_for;
    end;
  end
  else
  begin
    tabella_02.fieldbyname('cfg_tipo').asstring := 'G';
    tabella_02.fieldbyname('cfg_codice').asstring := tabella_esa_01.fieldbyname('Conto_Prima_Nota').asstring;
  end;

  if tabella_esa_01.fieldbyname('Dare_Avere').asstring = 'D' then
  begin
    tabella_02.fieldbyname('importo_dare').asfloat := tabella_esa_01.fieldbyname('Importo_in_Valuta').asfloat;
    tabella_02.fieldbyname('importo_dare_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Movimento').asfloat;
    tabella_02.fieldbyname('importo_avere').asfloat := 0;
    tabella_02.fieldbyname('importo_avere_euro').asfloat := 0;
  end
  else
  begin
    tabella_02.fieldbyname('importo_dare').asfloat := 0;
    tabella_02.fieldbyname('importo_dare_euro').asfloat := 0;
    tabella_02.fieldbyname('importo_avere').asfloat := tabella_esa_01.fieldbyname('Importo_in_Valuta').asfloat;
    tabella_02.fieldbyname('importo_avere_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Movimento').asfloat;

  end;

  tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('Descrizione_Supplem').asstring);
  try

    tabella_02.post;
  except
  end;

  if (tabella_esa_01.fieldbyname('TipoRiga_Iva_Normale').asstring = 'I') and (tabella_esa_01.fieldbyname('Codice_IVA').asinteger <> 0) then
  begin
    crea_pni(progressivo);
  end;

end;

procedure TCNVESA.crea_pni;
begin

  tabella_03.append;

  tabella_03.fieldbyname('progressivo').asfloat := tabella_01.fieldbyname('progressivo').asfloat;
  riga_iva := riga_iva + 1;
  tabella_03.fieldbyname('riga').asinteger := riga_iva;

  tabella_03.fieldbyname('tiv_codice').asstring := trim(tabella_esa_01.fieldbyname('Codice_IVA').asstring);
  tabella_03.fieldbyname('importo_imponibile').asfloat := tabella_esa_01.fieldbyname('Imponibile_Valuta').asfloat;
  tabella_03.fieldbyname('importo_imponibile_euro').asfloat := tabella_esa_01.fieldbyname('Imponibile').asfloat;

  tabella_03.fieldbyname('importo_iva').asfloat := tabella_esa_01.fieldbyname('Importo_in_Valuta').asfloat;
  tabella_03.fieldbyname('importo_iva_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Movimento').asfloat;

  // Name := 'importo_indetraibile';
  // Name := 'importo_indetraibile_euro';
  read_tabella(arc.arcdit, 'tiv', 'codice', tabella_03.fieldbyname('tiv_codice').asstring);

  tabella_03.fieldbyname('importo_indetraibile').asfloat := arrotonda
    (tabella_03.fieldbyname('importo_iva').asfloat * archivio.fieldbyname('indetraibile').asfloat / 100);

  tabella_03.fieldbyname('importo_indetraibile_euro').asfloat := arrotonda
    (tabella_03.fieldbyname('importo_iva').asfloat * archivio.fieldbyname('indetraibile').asfloat / 100);

  try
    tabella_03.post;
  except
  end;

end;

procedure TCNVESA.converti_par;
var
  tipo_cliente,
    codice_clifor: string;
  data_doc: TDateTime;
  numero_doc: integer;
  serie_doc: string;
  progressivo, nr_riga: integer;

  totale_pagare, totale_pagare_euro: double;
begin
  v_conferma.enabled := false;

  v_tabella_01.caption := 'scadenze';
  application.processmessages;

  cancella_tabella('pat');
  tabella_01.close;
  tabella_01.tablename := 'pat';
  tabella_01.open;

  cancella_tabella('pas');
  tabella_02.close;
  tabella_02.tablename := 'pas';
  tabella_02.open;

  tpa.close;
  tpa.open;

  tabella_01_ds.Dataset := tabella_01;
  try
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT * from SCADENZE');
    tabella_esa_01.Sql.Add('order by ind_clientefornitore,Cod_clientefornitore,data_documento,Numero_documento, Serie_documento, Numero_rata');
    tabella_esa_01.open;
  except
    messaggio(000, 'manca la tabella delle scadenze (SCADENZE)');
    close;
    abort;
  end;

  progressivo := 0;
  while not tabella_esa_01.eof do
  begin
    Application.ProcessMessages;

    v_tabella_01.caption := tabella_esa_01.fieldbyname('data_documento').asstring;

    tipo_cliente := tabella_esa_01.fieldbyname('ind_clientefornitore').asstring;
    codice_clifor := tabella_esa_01.fieldbyname('cod_clientefornitore').asstring;
    Data_doc := converti_data(tabella_esa_01.fieldbyname('data_documento').asstring);
    Numero_doc := tabella_esa_01.fieldbyname('numero_documento').asinteger;
    Serie_doc := trim(tabella_esa_01.fieldbyname('serie_documento').asString);

    if (tabella_esa_01.fieldbyname('Flag_Scad_Unificata').asstring = '2') then
    begin

      nr_riga := 0;
      while not(tabella_esa_01.eof) and
        (tipo_cliente = tabella_esa_01.fieldbyname('ind_clientefornitore').asstring) and
        (codice_clifor = tabella_esa_01.fieldbyname('cod_clientefornitore').asstring) and
        (Converti_data(tabella_esa_01.fieldbyname('data_documento').asString) = data_doc) and
        (tabella_esa_01.fieldbyname('numero_documento').asInteger = numero_doc) and
        (trim(tabella_esa_01.fieldbyname('serie_documento').asString) = serie_doc) do
      begin

        tabella_esa_01.Next;

      end;

    end
    else
    begin

      progressivo := progressivo + 1;
      tabella_01.append;
      tabella_01.fieldbyname('progressivo').asinteger := progressivo;

      tabella_01.fieldbyname('cfg_tipo').asstring := tabella_esa_01.fieldbyname('Ind_ClienteFornitore').asstring;
      if (tabella_esa_01.fieldbyname('Ind_ClienteFornitore').asstring = 'F') then
      begin
        assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);
        tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
        tabella_01.fieldbyname('importo_dovuto').asfloat := tabella_esa_01.fieldbyname('Importo_Pagare_Valut').asfloat * -1;
        tabella_01.fieldbyname('importo_dovuto_euro').asfloat := tabella_esa_01.fieldbyname('Importo_da_Pagare').asfloat * -1;
        tabella_01.fieldbyname('importo_pagato').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato_Valut').asfloat * -1;
        tabella_01.fieldbyname('importo_pagato_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato').asfloat * -1;
      end
      else
      begin
        assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_ClienteFornitore').asstring, cli_for);
        tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
        tabella_01.fieldbyname('importo_dovuto').asfloat := tabella_esa_01.fieldbyname('Importo_Pagare_Valut').asfloat;
        tabella_01.fieldbyname('importo_dovuto_euro').asfloat := tabella_esa_01.fieldbyname('Importo_da_Pagare').asfloat;
        // tabella_01.fieldbyname('importo_pagato').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato_Valut').asfloat;
        // tabella_01.fieldbyname('importo_pagato_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato').asfloat;
      end;

      tabella_01.fieldbyname('data_registrazione').asdatetime := converti_data(tabella_esa_01.fieldbyname('Data_Documento').asstring);
      tabella_01.fieldbyname('numero_documento').asinteger := tabella_esa_01.fieldbyname('numero_documento').asinteger;
      tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('Data_Documento').asstring);
      tabella_01.fieldbyname('serie_documento').asstring := trim(tabella_esa_01.fieldbyname('Serie_Documento').asstring);
      tabella_01.fieldbyname('data_scadenza').asdatetime := converti_data(tabella_esa_01.fieldbyname('Data_scadenza').AsString);
      tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;
      tabella_01.fieldbyname('cambio').asfloat := 1;

      if (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '3') or
        (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '4') or
        (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '7') or
        (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '8') then
        tabella_01.fieldbyname('tts_codice').asString := 'rimessa diretta'
      else if (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '1') then
        tabella_01.fieldbyname('tts_codice').asstring := 'R.I.B.A.'
      else if (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '6') then
        tabella_01.fieldbyname('tts_codice').asstring := 'bonifico bancario'
      else if (trim(tabella_esa_01.fieldbyname('Tipo_Pagamento').asString) = '9') then
        tabella_01.fieldbyname('tts_codice').asstring := 'R.I.D.';

      if not tpa.locate('tts_codice_variabili_01', tabella_01.fieldbyname('tts_codice').asstring, []) then
      begin
        tpa.First;
      end;

      tabella_01.fieldbyname('tpa_codice').asstring := tpa.fieldByNAme('codice').Asstring;

      // tabella_01.fieldbyname('descrizione').asstring := tabella_esa_01.fieldbyname('scdescri').asstring;

      tabella_01.post;

      totale_pagare := 0;
      totale_pagare_euro := 0;

      nr_riga := 1;

      if (tabella_esa_01.fieldbyname('importo_pagato').asfloat <> 0) then
      begin
        tabella_02.append;

        tabella_02.fieldbyname('progressivo').asfloat := progressivo;
        tabella_02.fieldbyname('riga').asinteger := nr_riga;
        tabella_02.fieldbyname('data_registrazione').asdatetime := data_doc;

        if (tabella_esa_01.fieldbyname('Ind_ClienteFornitore').asstring = 'F') then
        begin

          tabella_02.fieldbyname('importo_pagato').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato_Valut').asfloat * -1;
          tabella_02.fieldbyname('importo_pagato_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato').asfloat * -1;

        end
        else
        begin
          tabella_02.fieldbyname('importo_pagato').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato_Valut').asfloat;
          tabella_02.fieldbyname('importo_pagato_euro').asfloat := tabella_esa_01.fieldbyname('Importo_Pagato').asfloat;

        end;

        tabella_02.post;
      end;

      tabella_esa_01.Next;

    end; // if

  end; // while

  tpa.close;
  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;
  v_conferma.enabled := true;

end;

procedure TCNVESA.tabella_01BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_01.fieldbyname('utente').asstring := utente;
  tabella_01.fieldbyname('data_ora').asdatetime := now;

end;

procedure TCNVESA.tabella_02BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_02.fieldbyname('utente').asstring := utente;
  tabella_02.fieldbyname('data_ora').asdatetime := now;

end;

procedure TCNVESA.tabella_03BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_03.fieldbyname('utente').asstring := utente;
  tabella_03.fieldbyname('data_ora').asdatetime := now;

end;

procedure TCNVESA.controllo_campi;
begin
  if v_clienti.checked then
  begin
    v_fornitori.checked := true;
  end;
  if v_fornitori.checked then
  begin
    v_clienti.checked := true;
  end;
end;

procedure TCNVESA.converti_tva;
begin
  v_tabella_01.caption := 'tabella valute';
  application.processmessages;

  cancella_tabella('tva');
  tabella_01.close;
  tabella_01.tablename := 'tva';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM VALUTE');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('PARTE_FISSA=:PF');
  tabella_esa_01.Parameters.ParamByName('PF').Value := 'VL';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice_Valuta').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_valuta').AsString);
    tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.FieldByName('Cambio').AsFloat;

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tiv;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella iva';
  application.processmessages;

  cancella_tabella('tiv');
  tabella_01.close;
  tabella_01.tablename := 'tiv';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM TABELLAIVA');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('Parte_Fissa=:PF');
  tabella_esa_01.Parameters.ParamByName('PF').Value := 'CI';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice_Iva').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('DescrizioneCodiceIva').AsString);
    tabella_01.fieldbyname('percentuale').asfloat := tabella_esa_01.FieldByName('Percentuale_IVA').AsFloat;
    tabella_01.fieldbyname('indetraibile').asfloat := tabella_esa_01.FieldByName('Percent_Indetraibili').AsFloat;

    if tabella_esa_01.FieldByName('Codice_IVA_Ventilare').AsString <> '' then
    begin
      tabella_01.fieldbyname('tiv_codice_ventilazione').asString := tabella_esa_01.FieldByName('Codice_IVA_Ventilare').AsString;
    end;

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tna;
begin
  v_tabella_01.caption := 'tabella nazioni';
  application.processmessages;

  cancella_tabella('tna');
  tabella_01.close;
  tabella_01.tablename := 'tna';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM NAZIONI');
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione').AsString);
    tabella_01.fieldbyname('codice_iso').asstring := tabella_esa_01.FieldByName('Codice_ISO').AsString;
    ;
    tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;

end;

procedure TCNVESA.converti_tag;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella agenti';
  application.processmessages;

  try
    query.close;
    query.sql.clear;
    query.sql.add('create table esa_tag (');
    query.sql.add(' codice varchar(4),');
    query.sql.add(' descrizione varchar(40),');
    query.sql.add(' esa_codice varchar(6) )');
    query.execsql;

    query.Connection.commit;
  except
  end;

  query_02.sql.clear;
  query_02.sql.add('select * from esa_tag');
  query_02.sql.add('order by esa_codice');
  query_02.ReadOnly := false;
  query_02.open;

  cancella_tabella('tag');
  cancella_tabella('esa_tag');

  tabella_01.close;
  tabella_01.tablename := 'tag';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM AGENTI');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('chiave=:chiave');
  tabella_esa_01.Parameters.ParamByName('chiave').Value := 'AG';
  tabella_esa_01.Open;

  tabella_esa_02.Close;
  tabella_esa_02.Sql.Clear;
  tabella_esa_02.Sql.Add('SELECT * FROM ANAGRAFICHECOMUNI');
  tabella_esa_02.Sql.Add('ORDER BY Codice_Anagrafica');
  tabella_esa_02.Open;

  tabella_01.append;
  tabella_01.fieldbyname('codice').asstring := '0000';
  tabella_01.fieldbyname('descrizione').asstring := 'Agente standard';
  tabella_01.post;

  i := 0;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;
    i := i + 1;

    if tabella_esa_02.Locate('Codice_anagrafica', tabella_esa_01.FieldByName('Codice_Anagrafica').AsString, []) then
    begin
      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := setta_lunghezza(i, 4, 0);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_02.FieldByName('Ragione_sociale').AsString);
      tabella_01.post;

      query_02.append;
      query_02.fieldbyname('esa_codice').asstring := tabella_esa_01.FieldByName('Codice_Anagrafica').AsString;
      query_02.fieldbyname('codice').asstring := setta_lunghezza(i, 4, 0);
      query_02.fieldbyname('descrizione').asstring := trim(tabella_esa_02.FieldByName('Ragione_sociale').AsString);
      query_02.post;
    end;

    tabella_esa_01.Next;
  end;

  query.close;
  query_02.close;
  tabella_01.close;
  tabella_esa_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_tcm;
var
  i: word;
  codice, campi: string;
begin
  v_tabella_01.caption := 'tabella categorie merceologiche';
  application.processmessages;

  campi :=
    'codice,C,4,0;' +
    'descrizione,C,30,0;' +
    'esa_gruppo,C,15,0;' +
    'esa_sottogru,C,15,0;';

  try
    query.sql.clear;
    query.sql.Add('create table esa_tcm (');
    query.sql.Add('codice varchar(4), ');
    query.sql.Add(' descrizione varchar(40), ');
    query.sql.Add(' esa_gruppo varchar(15), ');
    query.sql.Add(' esa_sottogru varchar(15)) ');
    query.execsql;
  except
  end;

  query_02.sql.Clear;
  query_02.sql.add('select * from esa_tcm');
  query_02.open;

  cancella_tabella('tcm');
  tabella_01.close;
  tabella_01.tablename := 'tcm';
  tabella_01.open;

  cancella_tabella('tgm');
  tabella_02.close;
  tabella_02.tablename := 'tgm';
  tabella_02.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM GRUPPIMERCEOLOGICI');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('PARTE_FISSA=:PF');
  tabella_esa_01.Sql.Add('ORDER BY Gruppo_Merceologico,SottogruppoMerceolog');
  tabella_esa_01.Parameters.ParamByName('PF').Value := 'GM';
  tabella_esa_01.Open;

  i := 0;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    codice := tabella_esa_01.FieldByName('Gruppo_merceologico').AsString;

    if (codice <> tabella_esa_01.FieldByName('Gruppo_merceologico').AsString) or
      ((codice = tabella_esa_01.FieldByName('Gruppo_merceologico').AsString) and
      (trim(tabella_esa_01.FieldByName('SottogruppoMerceolog').AsString) = '')) then
    begin
      i := i + 1;

      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := setta_lunghezza(i, 2, 0);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_gruppo').AsString);

      tabella_01.post;

      query_02.append;
      query_02.fieldbyname('codice').asstring := trim(tabella_01.fieldbyname('codice').asstring);
      query_02.fieldbyname('descrizione').asstring := trim(tabella_01.fieldbyname('descrizione').asstring);
      query_02.fieldbyname('esa_gruppo').asstring := trim(tabella_esa_01.FieldByName('Gruppo_Merceologico').AsString);
      query_02.fieldbyname('esa_sottogru').asstring := '';
      query_02.post;

      // if i > 9 then
      // showmessage('categorie > 10 !');
    end
    else
    begin

      tabella_02.append;
      tabella_02.fieldbyname('codice').asstring := trim(tabella_01.fieldByName('codice').AsString + trim(tabella_esa_01.FieldByName('SottogruppoMerceolog').AsString));
      tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_gruppo').AsString);
      tabella_02.post;

      query_02.append;
      query_02.fieldbyname('codice').asstring := trim(tabella_01.fieldbyname('codice').asstring);
      query_02.fieldbyname('descrizione').asstring := trim(tabella_01.fieldbyname('descrizione').asstring);
      query_02.fieldbyname('esa_gruppo').asstring := trim(tabella_esa_01.FieldByName('Gruppo_Merceologico').AsString);
      query_02.fieldbyname('esa_sottogru').asstring := trim(tabella_esa_01.FieldByName('SottogruppoMerceolog').AsString);
      query_02.post;

    end;

    tabella_esa_01.Next;
  end;

  query.close;
  query_02.close;
  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;

end;

procedure TCNVESA.converti_tgm;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella gruppi merceologici';
  application.processmessages;

  cancella_tabella('tgm');
  tabella_01.close;
  tabella_01.tablename := 'tgm';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;
  tabella_01.close;
end;

procedure TCNVESA.converti_tma;
begin
  v_tabella_01.caption := 'tabella depositi';
  application.processmessages;

  cancella_tabella('tma');
  tabella_01.close;
  tabella_01.tablename := 'tma';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM ANAGRAFICADEPOSITI');
  tabella_esa_01.Sql.Add('where');
  tabella_esa_01.Sql.Add('chiave=:chiave');
  tabella_esa_01.Parameters.ParamByName('chiave').Value := 'DP';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('cod_deposito').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_Deposito').AsString);
    if tabella_esa_01.FieldByName('Flag_proprieta_merce').AsString = '1' then
    begin
      tabella_01.fieldbyname('proprieta').asstring := 'si';
    end
    else
    begin
      tabella_01.fieldbyname('proprieta').asstring := 'no';
    end;

    tabella_01.fieldbyname('descrizione1').asstring := '';
    tabella_01.fieldbyname('descrizione2').asstring := '';
    tabella_01.fieldbyname('via').asstring := tabella_esa_01.FieldByName('Indirizzo_Deposito').asstring;
    tabella_01.fieldbyname('cap').asstring := tabella_esa_01.FieldByName('CAP_Deposito').asstring;
    tabella_01.fieldbyname('citta').asstring := tabella_esa_01.FieldByName('Localita_Deposito').asstring;
    tabella_01.fieldbyname('provincia').asstring := tabella_esa_01.FieldByName('Provincia_Deposito').asstring;

    tabella_01.fieldbyname('tna_codice').asstring := dit.fieldbyname('codice').asstring;

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tba;
begin
  v_tabella_01.caption := 'tabella banche';
  application.processmessages;

  cancella_tabella('tba');

  tabella_01.close;
  tabella_01.tablename := 'tba';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM BANCHEAZIENDA1');
  tabella_esa_01.Sql.Add('where');
  tabella_esa_01.Sql.Add('chiave=:chiave');
  tabella_esa_01.Parameters.ParamByName('chiave').Value := 'BA1';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('cod_deposito').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_Banca_Aziend').AsString);

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tum;
begin
  v_tabella_01.caption := 'tabella unita misura';
  application.processmessages;

  cancella_tabella('tum');
  tabella_01.close;
  tabella_01.tablename := 'tum';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT distinct unita_misura_Princip FROM ANAGRAFICAARTICOLI');
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    if not tabella_01.locate('codice', trim(tabella_esa_01.FieldByName('Unita_misura_Princip').AsString), []) then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Unita_misura_Princip').AsString);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Unita_misura_Princip').AsString);
      tabella_01.fieldbyname('decimali').asInteger := 0;

      tabella_01.post;
    end;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tlv;
begin
  v_tabella_01.caption := 'tabella listini';
  application.processmessages;

  cancella_tabella('tlv');
  tabella_01.close;
  tabella_01.tablename := 'tlv';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM ANAGRAFICALISTINI');
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice_listino').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione_Listino').AsString);
    tabella_01.fieldbyname('tva_codice').asstring := trim(tabella_esa_01.FieldByName('Codice_Valuta').AsString);
    if tabella_esa_01.FieldByName('Ind_prezzo_compr_iva').AsString = '0' then
      tabella_01.fieldbyname('iva_inclusa').asstring := 'no'
    else
      tabella_01.fieldbyname('iva_inclusa').asstring := 'si';

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tco;
begin
  v_tabella_01.caption := 'causali contabili';
  application.processmessages;

  cancella_tabella('tco');
  tabella_01.close;
  tabella_01.tablename := 'tco';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM CAUSALICONTABILI');
  tabella_esa_01.Sql.Add('where');
  tabella_esa_01.Sql.Add('CHIAVE =:CHIAVE');
  tabella_esa_01.Parameters.ParamByName('CHIAVE').Value := 'C';

  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    if tabella_esa_01.fieldbyname('Codice_Causale').asstring <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('Codice_Causale').asstring);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('Descrizione_Causale').asstring);
      if tabella_01.fieldbyname('descrizione').asstring = '' then
      begin
        tabella_01.fieldbyname('descrizione').asstring := '.';
      end;

      if tabella_esa_01.fieldbyname('Tipo_registro_Iva').asstring = '1' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'vendite';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';
      end
      else if tabella_esa_01.fieldbyname('Tipo_registro_Iva').asstring = '2' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'acquisti';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';
      end
      else if tabella_esa_01.fieldbyname('Tipo_registro_Iva').asstring = '3' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'corrispettivi';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';

      end;

      if trim(tabella_esa_01.fieldbyname('Ind_gestione_partite').asstring) = '0' then
      begin
        tabella_01.fieldbyname('gestione_partite').asstring := 'no';
      end
      else
      begin
        tabella_01.fieldbyname('gestione_partite').asstring := 'si';
      end;

      if tabella_esa_01.fieldbyname('Codice_causale').asstring = 'BA' then
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'apertura bilancio';
      end
      else if tabella_esa_01.fieldbyname('Codice_causale').asstring = 'BC' then
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'chiusura bilancio';
      end
      else
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'normale';
      end;

      if trim(tabella_esa_01.fieldbyname('Flag_insoluto').asstring) = '1' then
      begin
        tabella_01.fieldbyname('insoluto').asstring := 'si';
      end
      else
      begin
        tabella_01.fieldbyname('insoluto').asstring := 'no';
      end;

      if (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite') or
        (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'acquisti') then
      begin
        if (trim(tabella_esa_01.fieldbyname('Tipo_causale').asstring) <> '1') then
        begin
          tabella_01.fieldbyname('segno_registro_iva').asstring := 'incrementa';
        end
        else
        begin
          tabella_01.fieldbyname('segno_registro_iva').asstring := 'decrementa';
        end;
      end
      else if tabella_01.fieldbyname('tipo_registro_iva').asstring <> '' then
      begin
        tabella_01.fieldbyname('segno_registro_iva').asstring := 'incrementa';
      end;

      if (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite') or
        (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'acquisti') then
      begin
        tabella_01.fieldbyname('serie_numerazione').asstring :=
          trim(tabella_esa_01.fieldbyname('nRegistroivaproposto').asString) +
          trim(tabella_esa_01.fieldbyname('SerieNumerazProposta').asString);
      end;

      // ------------------------
      // da controllare
      // ------------------------
      (*
        if tabella_01.fieldbyname('tipo_registro_iva').asstring <> '' then
        begin
        tabella_01.fieldbyname('gen_codice_iva').asstring := tabella_esa_01.fieldbyname('sottoconto_Dare_1').asstring;
        end;
      *)
      if tabella_01.fieldbyname('tipo_registro_iva').asstring = 'acquisti' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'F';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'A';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_codice_02').asstring := tabella_01.fieldbyname('gen_codice_iva').asstring;
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'D';
      end;
      if tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'C';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'D';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_codice_02').asstring := tabella_01.fieldbyname('gen_codice_iva').asstring;
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'A';
      end;
      if tabella_01.fieldbyname('tipo_registro_iva').asstring = 'corrispettivi' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'G';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'D';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_codice_02').asstring := tabella_01.fieldbyname('gen_codice_iva').asstring;
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'A';
      end;

      if tabella_01.fieldbyname('corrispettivi_da_ventilare').asstring = 'si' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'G';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'D';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'A';
      end;

      tabella_01.post;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tsp;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella tipo spedizione';
  application.processmessages;

  cancella_tabella('tsp');
  tabella_01.close;
  tabella_01.tablename := 'tsp';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM DATIDISPEDIZIONE');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('parte_fissa=:pf AND ');
  tabella_esa_01.Sql.Add('tipo_record=:tr ');
  tabella_esa_01.Parameters.ParamByName('pf').Value := 'SP';
  tabella_esa_01.Parameters.ParamByName('tr').Value := '2';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice').AsString);

    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione').AsString);

    if tabella_esa_01.FieldByName('Tipo_Mezzo_Trasporto').AsString = '0' then
      tabella_01.fieldbyname('tipo').asstring := 'mittente'
    else if tabella_esa_01.FieldByName('Tipo_Mezzo_Trasporto').AsString = '1' then
      tabella_01.fieldbyname('tipo').asstring := 'destinatario'
    else if tabella_esa_01.FieldByName('Tipo_Mezzo_Trasporto').AsString = '2' then
      tabella_01.fieldbyname('tipo').asstring := 'vettore';

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;

end;

procedure TCNVESA.converti_tzo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella zone';
  application.processmessages;

  (*
    tabella_01.close;
    tabella_01.tablename := 'tzo';
    tabella_01.exclusive := true;
    tabella_01.open;
    tabella_01.adszaptable;
    tabella_01.AdsPackTable;

    assignfile(file_archivio, tabella_esa_01.databasename + '\tabelle.txt');
    reset(file_archivio);
    while not eof(file_archivio) do
    begin
    readln(file_archivio, record_tabella);

    v_tabella_01.caption := copy(record_tabella, 1, inizio_tabella - 1);
    application.processmessages;

    if (copy(record_tabella, 1, 5) = 'ZONE_') and (copy(record_tabella, 6, 1) <> '\') then
    begin
    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := '';
    for i := 6 to 8 do
    begin
    if copy(record_tabella, i, 1) = ' ' then
    begin
    tabella_01.fieldbyname('codice').asstring := tabella_01.fieldbyname('codice').asstring + '0';
    end
    else
    begin
    tabella_01.fieldbyname('codice').asstring := tabella_01.fieldbyname('codice').asstring + copy(record_tabella, i, 1);
    end;
    end;

    tabella_01.fieldbyname('descrizione').asstring := copy(record_tabella, inizio_tabella, 30);

    tabella_01.post;
    end;
    end;

    closefile(file_archivio);
    tabella_01.close;
  *)
end;

procedure TCNVESA.converti_tdo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella documenti vendita';
  application.processmessages;

  (*
    tabella_01.close;
    tabella_01.tablename := 'tdo';
    tabella_01.exclusive := true;
    tabella_01.open;
    tabella_01.adszaptable;
    tabella_01.AdsPackTable;

    tco.Open;

    //--------------------------------------------------------------------
    // ordini clienti
    //--------------------------------------------------------------------
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT DISTINCT Causale_di_magazzino FROM TESTATEORDINI');
    tabella_esa_01.Sql.Add('where');
    tabella_esa_01.Sql.Add('tipo_Documento_1=:td');
    tabella_esa_01.Parameters.ParamByName('td').Value := 'I';
    tabella_esa_01.Open;

    while not tabella_esa_01.eof do
    begin
    application.processmessages;

    if not tabella_01.locate('codice', trim(tabella_esa_01.fieldByName('Causale_di_magazzino').AsString), []) then
    begin

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldByName('Causale_di_magazzino').AsString);

    if tco.locate('codice', tabella_01.fieldbyname('codice').asstring, []) then
    tabella_01.fieldbyname('descrizione').asstring := tco.Fieldbyname('descrizione').AsString
    else
    tabella_01.fieldbyname('descrizione').asstring := tabella_01.Fieldbyname('codice').AsString;

    tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
    tabella_01.fieldbyname('tma_codice').asstring := '000';

    //      tabella_01.fieldbyname('tco_codice').asstring := copy(record_tabella, inizio_tabella + 33, 3);

    tabella_01.post;
    end; // if

    tabella_esa_01.Next;
    end;

    //--------------------------------------------------------------------
    // bolle clienti
    //--------------------------------------------------------------------

    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT DISTINCT CodCausale_magazzino FROM BOLLE');
    tabella_esa_01.Open;

    while not tabella_esa_01.eof do
    begin
    application.processmessages;

    if not tabella_01.locate('codice', trim(tabella_esa_01.fieldByName('CodCausale_magazzino').AsString), []) then
    begin

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldByName('CodCausale_magazzino').AsString);

    if tco.locate('codice', tabella_01.fieldbyname('codice').asstring, []) then
    tabella_01.fieldbyname('descrizione').asstring := tco.Fieldbyname('descrizione').AsString
    else
    tabella_01.fieldbyname('descrizione').asstring := tabella_01.Fieldbyname('codice').AsString;

    tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
    tabella_01.fieldbyname('tma_codice').asstring := '000';

    //      tabella_01.fieldbyname('tco_codice').asstring := copy(record_tabella, inizio_tabella + 33, 3);

    tabella_01.post;
    end; // if
    tabella_esa_01.Next;

    end;

    tco.Close;
    tabella_esa_01.Close;
    tabella_01.close;
  *)
end;

procedure TCNVESA.converti_tcc;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella categorie contabili clienti';
  application.processmessages;

  cancella_tabella('tcc');
  tabella_01.close;
  tabella_01.tablename := 'tcc';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;
  tabella_01.close;
end;

procedure TCNVESA.converti_tca;
var
  i: word;
begin

  v_tabella_01.caption := 'tabella categorie contabili articoli';
  application.processmessages;

  cancella_tabella('tca');
  tabella_01.close;
  tabella_01.tablename := 'tca';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;

  tabella_01.close;
end;

procedure TCNVESA.converti_tcf;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella categorie contabili fornitori';
  application.processmessages;

  cancella_tabella('tcf');
  tabella_01.close;
  tabella_01.tablename := 'tcf';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;
  tabella_01.close;
end;

procedure TCNVESA.converti_tpo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella porti';
  application.processmessages;

  cancella_tabella('tpo');
  tabella_01.close;
  tabella_01.tablename := 'tpo';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM DATIDISPEDIZIONE');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('parte_fissa=:pf AND ');
  tabella_esa_01.Sql.Add('tipo_record=:tr ');
  tabella_esa_01.Parameters.ParamByName('pf').Value := 'SP';
  tabella_esa_01.Parameters.ParamByName('tr').Value := '3';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice').AsString);

    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione').AsString);

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;

end;

procedure TCNVESA.converti_tab;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella aspetto beni';
  application.processmessages;

  cancella_tabella('tab');
  tabella_01.close;
  tabella_01.tablename := 'tab';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM DATIDISPEDIZIONE');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('parte_fissa=:pf AND ');
  tabella_esa_01.Sql.Add('tipo_record=:tr ');
  tabella_esa_01.Parameters.ParamByName('pf').Value := 'SP';
  tabella_esa_01.Parameters.ParamByName('tr').Value := '0';
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice').AsString);

    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('Descrizione').AsString);

    tabella_01.post;

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

procedure TCNVESA.converti_tsa;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella codice statistico';
  application.processmessages;

  cancella_tabella('tsa');
  tabella_01.close;
  tabella_01.tablename := 'tsa';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT DISTINCT Codice_Statistico FROM ANAGRAFICAARTICOLI');
  tabella_esa_01.Open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    if trim(tabella_esa_01.FieldByName('Codice_Statistico').AsString) <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('Codice_Statistico').AsString);
      tabella_01.fieldbyname('descrizione').asstring := trim(trim(tabella_esa_01.FieldByName('Codice_Statistico').AsString));

      tabella_01.post;
    end; // if

    tabella_esa_01.Next;
  end;

  tabella_01.close;
  tabella_esa_01.close;
end;

// ******************************************************************************

procedure TCNVESA.converti_mov;
var
  i, progressivo: integer;
begin
  v_tabella_01.caption := 'movimenti magazzino';
  application.processmessages;

  // arc.attivazione_trigger(arc.arcdit, false, false);

  cancella_tabella('magese');

  cancella_tabella('mag');

  cancella_tabella('mmt');
  tabella_01.close;
  tabella_01.tablename := 'mmt';
  tabella_01.open;

  cancella_tabella('mmr');
  tabella_02.close;
  tabella_02.tablename := 'mmr';
  tabella_02.open;

  tabella_01_ds.Dataset := tabella_01;
  try
    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT * FROM MOVIMENTIMAGAZZINO');
    tabella_esa_01.Sql.Add('order by Progressivo, numero_riga');
    tabella_esa_01.open;
  except

    messaggio(000, 'manca la tabella del magazzino (MOVIMENTIMAGAZZINO)');

    close;
    abort;
  end;

  test_data := 0;
  test_numero_stringa := '';
  test_cod_causale := '';
  progressivo := 0;

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    v_tabella_01.caption := 'movimenti magazzino ' + tabella_esa_01.fieldbyname('data_registrazione').asstring;

    if (pos('-D', tabella_esa_01.fieldbyname('Codice_Articolo').asstring) = 0) and
      (pos('-M', tabella_esa_01.fieldbyname('Codice_Articolo').asstring) = 0) and
      (trim(tabella_esa_01.fieldbyname('cod_causale').asstring) <> '040') and
      (trim(tabella_esa_01.fieldbyname('cod_causale').asstring) <> '041') then
    begin
      codice_articolo := trim(tabella_esa_01.fieldbyname('Codice_Articolo').asstring);

      if (test_data <> converti_data(trim(tabella_esa_01.fieldbyname('data_registrazione').asstring))) or
        (test_numero_stringa <> tabella_esa_01.fieldbyname('Numero_Documento').asstring) or
        (test_cod_causale <> trim(tabella_esa_01.fieldbyname('cod_causale').asstring)) then
      begin
        progressivo := progressivo + 1;
        crea_mmt(progressivo);
        test_data := converti_data(trim(tabella_esa_01.fieldbyname('data_registrazione').asstring));
        try
          test_numero_stringa := trim(tabella_esa_01.fieldbyname('Numero_Documento').asstring);
        except
          test_numero_stringa := '';
        end;
        test_cod_causale := tabella_esa_01.fieldbyname('cod_causale').asstring;
      end;

      crea_mmr;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_esa_01.close;

  // arc.attivazione_trigger(arc.arcdit, false, true);
end;

procedure TCNVESA.crea_mmt(progressivo: integer);
begin

  tabella_01.append;
  tabella_01.fieldbyname('progressivo').asinteger := progressivo;
  tabella_01.fieldbyname('data_registrazione').asdatetime := converti_data(trim(tabella_esa_01.fieldbyname('data_registrazione').asstring));
  tabella_01.fieldbyname('tmo_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_causale').asstring);
  tabella_01.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('Codice_Deposito').asstring;
  try
    tabella_01.fieldbyname('numero_documento').asinteger := tabella_esa_01.fieldbyname('numero_documento').asinteger;
  except
    tabella_01.fieldbyname('numero_documento').asinteger := 0;
  end;

  try
    tabella_01.fieldbyname('data_documento').asdatetime := converti_data(trim(tabella_esa_01.fieldbyname('data_documento').asstring));
  except
    tabella_01.fieldbyname('data_documento').asdatetime := converti_data(trim(tabella_esa_01.fieldbyname('data_registrazione').asstring));
  end;

  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('Serie_Numeraz_Docum').asstring;

  if trim(tabella_esa_01.fieldbyname('Indicatore_Cli_For').asstring) = 'C' then
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('Codice_Cli_For_Contr').asstring, cli_for);
    tabella_01.fieldbyname('cfg_tipo').asstring := 'C';
    tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
  end

  else if trim(tabella_esa_01.fieldbyname('Indicatore_Cli_For').asstring) = 'F' then
  begin
    assegna_codice_frn(tabella_esa_01.fieldbyname('Codice_Cli_For_Contr').asstring, cli_for);
    tabella_01.fieldbyname('cfg_tipo').asstring := 'F';
    tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
  end;

  tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;
  tabella_01.fieldbyname('ese_codice').asstring := Copy(tabella_esa_01.fieldbyname('data_registrazione').asstring, 1, 4);
  tabella_01.fieldbyname('cambio').asfloat := tabella_esa_01.fieldbyname('cambio').asfloat;

  tabella_01.post;

  riga := 0;

end;

procedure TCNVESA.crea_mmr;
begin

  if read_tabella(arc.arcdit, 'art', 'codice', codice_articolo) then
  begin
    tabella_02.append;

    tabella_02.fieldbyname('progressivo').asfloat := tabella_01.fieldbyname('progressivo').asfloat;
    riga := riga + 1;
    tabella_02.fieldbyname('riga').asinteger := riga;
    tabella_02.fieldbyname('art_codice').asstring := codice_articolo;

    tabella_02.fieldbyname('quantita').asfloat := tabella_esa_01.fieldbyname('Qta_movimento').asfloat;
    tabella_02.fieldbyname('prezzo').asfloat := tabella_esa_01.fieldbyname('prz_unitario_valuta').asfloat;
    tabella_02.fieldbyname('importo').asfloat := tabella_esa_01.fieldbyname('Val_Movimento_Valuta').asfloat;
    tabella_02.fieldbyname('importo_euro').asfloat := tabella_esa_01.fieldbyname('Valore_movimento').asfloat;
    tabella_02.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('Codice_Deposito').asstring;

    read_tabella(arc.arcdit, 'tmo', 'codice', tabella_01.fieldbyname('tmo_codice').asstring);
    if archivio.fieldbyname('esistenza').asstring = 'incrementa' then
    begin
      tabella_02.fieldbyname('quantita_entrate').AsFloat := tabella_02.fieldbyname('quantita').asfloat;
      tabella_02.fieldbyname('quantita_uscite').AsFloat := 0;
    end
    else if archivio.fieldbyname('esistenza').asstring = 'decrementa' then
    begin
      tabella_02.fieldbyname('quantita_entrate').AsFloat := 0;
      tabella_02.fieldbyname('quantita_uscite').AsFloat := tabella_02.fieldbyname('quantita').asfloat;
    end
    else if archivio.fieldbyname('esistenza').asstring = 'ignora' then
    begin
      tabella_02.fieldbyname('quantita_entrate').AsFloat := 0;
      tabella_02.fieldbyname('quantita_uscite').AsFloat := 0;
    end;

    if not((tabella_esa_01.fieldbyname('Sconto_1_su_articolo').asfloat = 0) and (tabella_esa_01.fieldbyname('Sconto_2_su_articolo').asfloat = 0)) then
    begin
      query.params[0].asfloat := tabella_esa_01.fieldbyname('Sconto_1_su_articolo').asfloat;
      query.params[1].asfloat := tabella_esa_01.fieldbyname('Sconto_2_su_articolo').asfloat;

      query.close;
      query.open;
      if not query.eof then
      begin
        tabella_02.fieldbyname('tsm_codice').asstring := query.fieldbyname('codice').asstring;
      end
      else
      begin
        crea_tsm(
          tabella_esa_01.fieldbyname('Sconto_1_su_articolo').asfloat,
          tabella_esa_01.fieldbyname('Sconto_2_su_articolo').asfloat);
        tabella_02.fieldbyname('tsm_codice').asstring := setta_lunghezza(tsm_codice, 4, 0);
      end;
    end;

    tabella_02.post;

  end;

end;

procedure TCNVESA.converti_ordini_clienti;
var
  i: integer;
  test_serie, test_numero_doc: string;

begin
  v_tabella_01.caption := 'ordini clienti';
  application.processmessages;

  query.SQL.clear;
  query.sql.add('select * from esa_tag');
  query.sql.add('where');
  query.sql.add('esa_codice=:esa_codice');
  try
    query.Open;
  except
    converti_tag;
  end;

  cancella_tabella('ovt');
  tabella_01.close;
  tabella_01.tablename := 'ovt';
  tabella_01.open;

  cancella_tabella('ovr');
  tabella_02.close;
  tabella_02.tablename := 'ovr';
  tabella_02.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM TESTATEORDINI');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('TIPO_RECORD_1=:TP AND TIPO_documento_1=:TD AND Stato_Ordine<>:SO');
  tabella_esa_01.Sql.Add('ORDER BY Serie_documento,Numero_documento');
  tabella_esa_01.Parameters.ParamByName('TP').Value := 'T';
  tabella_esa_01.Parameters.ParamByName('TD').Value := 'I';
  tabella_esa_01.Parameters.ParamByName('SO').Value := 'E';

  tabella_esa_01.Open;

  tabella_esa_02.Close;
  tabella_esa_02.Sql.Clear;
  tabella_esa_02.Sql.Add('SELECT * FROM RIGHEORDINI');
  tabella_esa_02.Sql.Add('WHERE');
  tabella_esa_02.Sql.Add('TIPO_RECORD=:TP AND TIPO_DOCUMENTO=:TD AND Stato_riga_ordine<>:SO');
  tabella_esa_02.Sql.Add('ORDER BY Serie_documento_1,Numero_documento_1,Numero_riga');

  tabella_esa_02.Parameters.ParamByName('TP').Value := 'R';
  tabella_esa_02.Parameters.ParamByName('TD').Value := 'I';
  tabella_esa_02.Parameters.ParamByName('SO').Value := 'E';
  tabella_esa_02.Open;

  test_data := 0;
  test_alfa := '';
  test_numero := 0;

  test_serie := '';
  test_numero_doc := '';

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin
    crea_ovt;

    test_serie := tabella_esa_01.FieldByName('Serie_documento').AsString;
    test_numero_doc := tabella_esa_01.FieldByName('Numero_documento').AsString;

    if tabella_esa_02.Locate('Serie_documento_1;Numero_documento_1', VarArrayof([test_serie, test_numero_doc]), []) then
    begin
      crea_ovr(test_serie, test_numero_doc);
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_03.close;
  tabella_esa_01.close;
  tabella_esa_02.close;

end;

procedure TCNVESA.crea_ovt;
begin

  tabella_01.append;

  tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
  tabella_01.fieldbyname('tdo_codice').asstring := 'ORDV';
  tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_ordine').asstring);
  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('serie_documento').asstring;
  tabella_01.fieldbyname('numero_documento').asinteger := StrToInt(trim(tabella_esa_01.fieldbyname('numero_documento').asstring));

  assegna_codice_cli(tabella_esa_01.fieldbyname('Cod_clifor_Contropar').asstring, cli_for);

  tabella_01.fieldbyname('cli_codice').asstring := cli_for;
  tabella_01.fieldbyname('tma_codice').asstring := '000'; // tabella_esa_01.fieldbyname('ortesmag').asstring;
  if trim(tabella_esa_01.fieldbyname('Codice_Agente').asstring) = '' then
    tabella_01.fieldbyname('tag_codice').asstring := '0000'
  else
  begin
    query.close;
    query.parambyname('esa_codice').asstring := trim(tabella_esa_01.fieldbyname('Codice_Agente').asstring);
    query.open;
    if not query.eof then
    begin
      tabella_01.fieldbyname('tag_codice').asstring := query.fieldbyname('Codice').asstring;
    end
    else
    begin
      tabella_01.fieldbyname('tag_codice').asstring := '0000';
    end;
  end;

  tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_01.fieldbyname('Codice_Pagamento').asstring;
  tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_01.fieldbyname('Codice_Banca').asstring;
  tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_01.fieldbyname('Codice_Agenzia').asstring;

  tabella_01.fieldbyname('tlv_codice').asstring := tabella_esa_01.fieldbyname('num_listino').asstring;
  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('codice_valuta').asstring;
  tabella_01.fieldbyname('cambio').asfloat := tabella_esa_01.fieldbyname('cambio').asfloat;
  // tabella_01.fieldbyname('tpo_codice').asstring := tabella_esa_01.fieldbyname('orcodpor').asstring;

  tabella_01.fieldbyname('nostro_riferimento').asstring := tabella_esa_01.fieldbyname('Riferimento').asstring;

  (*
    if tabella_esa_01.fieldbyname('ORCODDES').asinteger <> 0 then
    begin
    if read_tabella(arc.arcdit,'ind', 'cli_codice', tabella_01.fieldbyname('cli_codice').asstring,
    'indirizzo', tabella_esa_01.fieldbyname('ORCODDES').asinteger) then
    begin
    tabella_01.fieldbyname('indirizzo').asstring := inttostr(tabella_esa_01.fieldbyname('orcoddes').asinteger);

    tabella_01.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
    tabella_01.fieldbyname('descrizione2').asstring := archivio.fieldbyname('descrizione2').asstring;
    tabella_01.fieldbyname('via').asstring := archivio.fieldbyname('via').asstring;
    tabella_01.fieldbyname('cap').asstring := archivio.fieldbyname('cap').asstring;
    tabella_01.fieldbyname('citta').asstring := archivio.fieldbyname('citta').asstring;
    tabella_01.fieldbyname('provincia').asstring := archivio.fieldbyname('provincia').asstring;
    tabella_01.fieldbyname('tna_codice').asstring := archivio.fieldbyname('tna_codice').asstring;
    end;
    end;
  *)
  read_tabella(arc.arcdit, 'tlv', 'codice', tabella_01.fieldbyname('tlv_codice').asstring);
  tabella_01.fieldbyname('listino_con_iva').asstring := archivio.fieldbyname('iva_inclusa').asstring;

  read_tabella(arc.arcdit, 'cli', 'codice', tabella_01.fieldbyname('cli_codice').asstring);
  tabella_01.fieldbyname('tsp_codice').asstring := '01'; // archivio.fieldbyname('tsp_codice').asstring;
  tabella_01.fieldbyname('tst_codice').asstring := archivio.fieldbyname('tst_codice').asstring;
  tabella_01.fieldbyname('addebito_spese_fattura').asstring := archivio.fieldbyname('addebito_spese_fattura').asstring;

  tabella_01.fieldbyname('tpo_codice').asstring := '01';
  read_tabella(arc.arcdit, 'tpo', 'codice', tabella_01.fieldbyname('tpo_codice').asstring);
  tabella_01.fieldbyname('addebito_spese_trasporto').asstring := archivio.fieldbyname('addebito').asstring;

  if tabella_esa_01.FieldByName('stato_ordine').AsString = 'E' then
    tabella_01.fieldbyname('situazione').asstring := 'evaso'
  else if tabella_esa_01.FieldByName('stato_ordine').AsString = 'P' then
    tabella_01.fieldbyname('situazione').asstring := 'evaso parziale'
  else
    tabella_01.fieldbyname('situazione').asstring := 'inserito';

  tabella_01.post;

  riga := 0;
end;

procedure TCNVESA.crea_ovr(serie_documento, numero_documento: string);
var
  codice_articolo: string;
begin

  while
    not(tabella_esa_02.eof) and
    (tabella_esa_02.FieldByName('Serie_documento_1').AsString = serie_documento) and
    (tabella_esa_02.FieldByName('Numero_documento_1').AsString = numero_documento) do
  begin

    codice_Articolo := trim(tabella_esa_02.fieldbyname('Codice_Articolo').asstring);
    if (codice_articolo <> '-M') and
      (codice_articolo <> '-D') then
    begin
      esiste_articolo := true;
      if (not read_tabella(arc.arcdit, 'art', 'codice', codice_articolo)) then
      begin
        esiste_articolo := false;
      end;
    end
    else
    begin
      esiste_articolo := true;
    end;

    if esiste_articolo then
    begin
      tabella_02.append;

      tabella_02.fieldbyname('progressivo').asfloat := tabella_01.fieldbyname('progressivo').asfloat;

      tabella_02.fieldbyname('riga').asinteger := tabella_esa_02.fieldbyname('Numero_riga').asinteger;

      if (codice_articolo = '-M') or
        (codice_articolo = '-D') then
      begin
        tabella_02.fieldbyname('art_codice').asstring := '';
      end
      else
      begin
        tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
      end;

      tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_esa_02.fieldbyname('Descrizione_Articolo').asstring);
      tabella_02.fieldbyname('descrizione2').asstring := trim(tabella_esa_02.fieldbyname('Descrizione2Articolo').asstring);
      // tabella_02.fieldbyname('note').asstring := tabella_esa_02.fieldbyname('OR__NOTE').asstring;
      tabella_02.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring;
      tabella_02.fieldbyname('tiv_codice').asstring := tabella_esa_02.fieldbyname('Codice_Iva').asstring;

      tabella_02.fieldbyname('quantita').asfloat := arrotonda(tabella_esa_02.fieldbyname('qta_movimento').asfloat);

      tabella_02.fieldbyname('prezzo').asfloat := tabella_esa_02.fieldbyname('Prezzo_Unitario').asfloat;
      (*
        if not ((tabella_esa_02.fieldbyname('orscont1').asfloat = 0) and (tabella_esa_02.fieldbyname('orscont2').asfloat = 0)) then
        begin
        query.params[0].asfloat := tabella_esa_02.fieldbyname('orscont1').asfloat;
        query.params[1].asfloat := tabella_esa_02.fieldbyname('orscont2').asfloat;
        query.close;
        query.open;
        if not query.eof then
        begin
        tabella_02.fieldbyname('tsm_codice').asstring := query.fieldbyname('codice').asstring;
        end
        else
        begin
        messaggio(000, 'manca la tabella sconto ' + floattostr(tabella_esa_02.fieldbyname('orscont1').asfloat) + ' + ' +
        floattostr(tabella_esa_02.fieldbyname('orscont2').asfloat));
        end;
        end;
      *)
      calcola_importo;
      tabella_02.fieldbyname('tma_codice').asstring := tabella_01.fieldbyname('tma_codice').asstring;

      tabella_02.post;
    end;

    tabella_esa_02.Next;
  end; // while

end;

procedure TCNVESA.converti_ordini_fornitori;
var
  i: integer;
  test_serie,
    test_numero_doc: string;
begin
  v_tabella_01.caption := 'ordini fornitori';
  application.processmessages;

  cancella_tabella('oat');
  tabella_01.close;
  tabella_01.tablename := 'oat';
  tabella_01.open;

  cancella_tabella('oat');
  tabella_02.close;
  tabella_02.tablename := 'oar';
  tabella_02.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM TESTATEORDINI');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('TIPO_RECORD_1=:TP AND TIPO_DOCUMENTO_1=:TD AND Stato_Ordine<>:SO');
  tabella_esa_01.Sql.Add('ORDER BY Serie_documento,Numero_documento');
  tabella_esa_01.Parameters.ParamByName('TP').Value := 'T';
  tabella_esa_01.Parameters.ParamByName('TD').Value := 'O';
  tabella_esa_01.Parameters.ParamByName('SO').Value := 'E';
  tabella_esa_01.Open;

  tabella_esa_02.Close;
  tabella_esa_02.Sql.Clear;
  tabella_esa_02.Sql.Add('SELECT * FROM RIGHEORDINI');
  tabella_esa_02.Sql.Add('WHERE');
  tabella_esa_02.Sql.Add('TIPO_RECORD=:TP AND TIPO_DOCUMENTO=:TD AND Stato_riga_Ordine<>:SO');
  tabella_esa_02.Sql.Add('ORDER BY Serie_documento_1,Numero_documento_1,Numero_riga');

  tabella_esa_02.Parameters.ParamByName('TP').Value := 'R';
  tabella_esa_02.Parameters.ParamByName('TD').Value := 'O';
  tabella_esa_02.Parameters.ParamByName('SO').Value := 'E';
  tabella_esa_02.Open;

  test_data := 0;
  test_alfa := '';
  test_numero := 0;

  test_serie := '';
  test_numero_doc := '';

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin
    crea_oat;

    test_serie := tabella_esa_01.FieldByName('Serie_documento').AsString;
    test_numero_doc := tabella_esa_01.FieldByName('Numero_documento').AsString;

    if tabella_esa_02.Locate('Serie_documento_1;Numero_documento_1', VarArrayof([test_serie, test_numero_doc]), []) then
    begin
      crea_oar(test_serie, test_numero_doc);
    end;

    tabella_esa_01.next;
  end;

  tabella_01.close;
  tabella_02.close;
  tabella_03.close;
  tabella_esa_01.close;
  tabella_esa_02.close;

  (*

    tabella_esa_01.exclusive := true;
    tabella_esa_01.tablename := 'ord_ini';
    tabella_esa_01.indexname := 'tag1';
    tabella_esa_01.filtered := true;
    tabella_esa_01.filter := 'ORTIPORD = ''OR'' and ORDATORD > ''31/12/2001''';
    tabella_esa_01.open;
    tabella_esa_01.AdsPackTable;

    tabella_esa_02.filtered := false;
    tabella_esa_02.tablename := 'ban_che';
    try
    tabella_esa_02.open;
    except
    messaggio(000, 'manca la tabella delle banche (BAN_CHE)');
    close;
    abort;
    end;

    test_data := 0;
    test_alfa := '';
    test_numero := 0;

    query.sql.clear;
    query.sql.add('select first 1 codice from tsm where sconto_maggiorazione = ''sconto''');
    query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02');

    while not tabella_esa_01.eof do
    begin
    if (test_data <> tabella_esa_01.fieldbyname('ordatord').asdatetime) or
    (test_alfa <> tabella_esa_01.fieldbyname('oralford').asstring) or
    (test_numero <> tabella_esa_01.fieldbyname('ornumord').asinteger) then
    begin
    crea_oat;
    test_data := tabella_esa_01.fieldbyname('ordatord').asdatetime;
    test_alfa := tabella_esa_01.fieldbyname('oralford').asstring;
    test_numero := tabella_esa_01.fieldbyname('ornumord').asinteger;
    end;

    crea_oar;

    tabella_esa_01.next;
    end;

    tabella_01.close;
    tabella_02.close;
    tabella_03.close;
    tabella_esa_01.close;
    tabella_esa_02.close;
  *)
end;

procedure TCNVESA.crea_oat;
begin
  tabella_01.append;

  tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
  tabella_01.fieldbyname('tda_codice').asstring := 'ORDA';
  tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_ordine').asstring);
  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('serie_documento').asstring;
  tabella_01.fieldbyname('numero_documento').asinteger := StrToInt(trim(tabella_esa_01.fieldbyname('numero_documento').asstring));

  assegna_codice_frn(tabella_esa_01.fieldbyname('Cod_clifor_Contropar').asstring, cli_for);

  tabella_01.fieldbyname('frn_codice').asstring := cli_for;

  tabella_01.fieldbyname('tma_codice').asstring := '000'; // tabella_esa_01.fieldbyname('ortesmag').asstring;
  // tabella_01.fieldbyname('tag_codice').asstring := tabella_esa_01.fieldbyname('Codice_Agente').asstring;
  tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_01.fieldbyname('Codice_Pagamento').asstring;
  // tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_01.fieldbyname('Codice_Banca').asstring;
  // tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_01.fieldbyname('Codice_Agenzia').asstring;

  tabella_01.fieldbyname('tla_codice').asstring := tabella_esa_01.fieldbyname('num_listino').asstring;
  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('codice_valuta').asstring;
  tabella_01.fieldbyname('cambio').asfloat := tabella_esa_01.fieldbyname('cambio').asfloat;
  tabella_01.fieldbyname('tpo_codice').asstring := '01';
  // tabella_01.fieldbyname('tpo_codice').asstring := tabella_esa_01.fieldbyname('orcodpor').asstring;

  tabella_01.fieldbyname('riferimento').asstring := tabella_esa_01.fieldbyname('Riferimento').asstring;

  (*
    if tabella_esa_01.fieldbyname('ORCODDES').asinteger <> 0 then
    begin
    if read_tabella(arc.arcdit,'ind', 'cli_codice', tabella_01.fieldbyname('cli_codice').asstring,
    'indirizzo', tabella_esa_01.fieldbyname('ORCODDES').asinteger) then
    begin
    tabella_01.fieldbyname('indirizzo').asstring := inttostr(tabella_esa_01.fieldbyname('orcoddes').asinteger);

    tabella_01.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
    tabella_01.fieldbyname('descrizione2').asstring := archivio.fieldbyname('descrizione2').asstring;
    tabella_01.fieldbyname('via').asstring := archivio.fieldbyname('via').asstring;
    tabella_01.fieldbyname('cap').asstring := archivio.fieldbyname('cap').asstring;
    tabella_01.fieldbyname('citta').asstring := archivio.fieldbyname('citta').asstring;
    tabella_01.fieldbyname('provincia').asstring := archivio.fieldbyname('provincia').asstring;
    tabella_01.fieldbyname('tna_codice').asstring := archivio.fieldbyname('tna_codice').asstring;
    end;
    end;
  *)

  read_tabella(arc.arcdit, 'tla', 'codice', tabella_01.fieldbyname('tla_codice').asstring);
  tabella_01.fieldbyname('listino_con_iva').asstring := archivio.fieldbyname('iva_inclusa').asstring;

  read_tabella(arc.arcdit, 'frn', 'codice', tabella_01.fieldbyname('frn_codice').asstring);
  tabella_01.fieldbyname('tsp_codice').asstring := '01';
  archivio.fieldbyname('tsp_codice').asstring;
  if tabella_01.fieldbyname('tla_codice').asstring = '' then
  begin
    tabella_01.fieldbyname('tla_codice').asstring := archivio.fieldbyname('tla_codice').asstring;
  end;

  tabella_01.post;

  riga := 0;

end;

procedure TCNVESA.crea_oar(serie_documento, numero_documento: string);
var
  codice_articolo: string;
begin

  while
    not(tabella_esa_02.eof) and
    (tabella_esa_02.FieldByName('Serie_documento_1').AsString = serie_documento) and
    (tabella_esa_02.FieldByName('Numero_documento_1').AsString = numero_documento) do
  begin

    codice_Articolo := trim(tabella_esa_02.fieldbyname('Codice_Articolo').asstring);

    if (codice_articolo <> '-M') and
      (codice_articolo <> '-D') then
    begin
      if (not read_tabella(arc.arcdit, 'art', 'codice', codice_articolo)) then
      begin
        esiste_articolo := false;
      end;
    end
    else
    begin
      esiste_articolo := true;
    end;

    if esiste_articolo then
    begin
      tabella_02.append;

      tabella_02.fieldbyname('progressivo').asfloat := tabella_01.fieldbyname('progressivo').asfloat;

      tabella_02.fieldbyname('riga').asinteger := tabella_esa_02.fieldbyname('Numero_riga').asinteger;

      if (codice_articolo = '-M') or
        (codice_articolo = '-D') then
      begin
        tabella_02.fieldbyname('art_codice').asstring := '';
      end
      else
      begin
        tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
      end;

      tabella_02.fieldbyname('descrizione1').asstring := tabella_esa_02.fieldbyname('Descrizione_Articolo').asstring;
      tabella_02.fieldbyname('descrizione2').asstring := tabella_esa_02.fieldbyname('Descrizione2Articolo').asstring;
      // tabella_02.fieldbyname('note').asstring := tabella_esa_02.fieldbyname('OR__NOTE').asstring;
      tabella_02.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring;
      tabella_02.fieldbyname('tiv_codice').asstring := tabella_esa_02.fieldbyname('Codice_Iva').asstring;

      tabella_02.fieldbyname('quantita').asfloat := arrotonda(tabella_esa_02.fieldbyname('qta_movimento').asfloat);

      tabella_02.fieldbyname('prezzo').asfloat := tabella_esa_02.fieldbyname('Prezzo_Unitario').asfloat;
      (*
        if not ((tabella_esa_02.fieldbyname('orscont1').asfloat = 0) and (tabella_esa_02.fieldbyname('orscont2').asfloat = 0)) then
        begin
        query.params[0].asfloat := tabella_esa_02.fieldbyname('orscont1').asfloat;
        query.params[1].asfloat := tabella_esa_02.fieldbyname('orscont2').asfloat;
        query.close;
        query.open;
        if not query.eof then
        begin
        tabella_02.fieldbyname('tsm_codice').asstring := query.fieldbyname('codice').asstring;
        end
        else
        begin
        messaggio(000, 'manca la tabella sconto ' + floattostr(tabella_esa_02.fieldbyname('orscont1').asfloat) + ' + ' +
        floattostr(tabella_esa_02.fieldbyname('orscont2').asfloat));
        end;
        end;
      *)
      calcola_importo;
      tabella_02.fieldbyname('tma_codice').asstring := tabella_01.fieldbyname('tma_codice').asstring;

      tabella_02.post;
    end;

    tabella_esa_02.Next;
  end; // while

end;

procedure TCNVESA.calcola_importo;
var
  tiv_codice: string;
  imponibile: double;
begin
  if not((tabella_02.fieldbyname('quantita').asfloat = 0) or (tabella_02.fieldbyname('prezzo').asfloat = 0)) then
  begin
    tabella_02.fieldbyname('importo').asfloat :=
      arrotonda(tabella_02.fieldbyname('quantita').asfloat * tabella_02.fieldbyname('prezzo').asfloat *
      sconto(tabella_02.fieldbyname('tsm_codice').asstring) / 100);
  end;
  tabella_02.fieldbyname('importo_euro').asfloat := arrotonda
    (tabella_02.fieldbyname('importo').asfloat / tabella_01.fieldbyname('cambio').asfloat);

  tiv_codice := tabella_02.fieldbyname('tiv_codice').asstring;
  if read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring) then
  begin
    if tabella_01.fieldbyname('listino_con_iva').Asstring = 'no' then
    begin
      tabella_02.fieldbyname('importo_iva').AsFloat :=
        arrotonda(tabella_02.fieldbyname('importo').asfloat * archivio.fieldbyname('percentuale').asfloat / 100);
      read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring);
      tabella_02.fieldbyname('importo_iva_euro').AsFloat :=
        arrotonda(tabella_02.fieldbyname('importo_iva').asfloat / tabella_01.fieldbyname('cambio').asfloat);
    end
    else
    begin
      imponibile := arrotonda(tabella_02.fieldbyname('importo').asfloat / (1 + archivio.fieldbyname('percentuale').asfloat / 100));
      tabella_02.fieldbyname('importo_iva').asfloat :=
        arrotonda(tabella_02.fieldbyname('importo').asfloat - imponibile);
      read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring);
      tabella_02.fieldbyname('importo_iva_euro').AsFloat :=
        arrotonda(tabella_02.fieldbyname('importo_iva').asfloat / tabella_01.fieldbyname('cambio').asfloat);
    end;
  end;
end;

procedure TCNVESA.assegna_codice_cli(codice_adhoc: string; var cli_for: string);
begin
  codice_adhoc := trim(codice_adhoc);
  if v_codifica_clienti.checked then
  begin
    if codice_nom_numerico = 'si' then
    begin
      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);
    end
    else

    begin
      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 6, 0); // codice_adhoc;
    end;
  end
  else
  begin
    // cli_for := 'C' + codice_adhoc;
    if codice_nom_numerico = 'si' then
    begin

      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);
    end
    else
    begin
      cli_for := trim(codice_adhoc);
    end;

  end;
end;

procedure TCNVESA.assegna_codice_frn(codice_adhoc: string; var cli_for: string);
begin
  codice_adhoc := trim(codice_adhoc);
  if v_codifica_fornitori.checked then
  begin
    if codice_nom_numerico = 'si' then
    begin

      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);

    end
    else
    begin
      cli_for := trim(codice_adhoc);
      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 6, 0); // codice_adhoc;

      // cli_for := codice_adhoc;
    end;
  end
  else
  begin
    // cli_for := 'F' + codice_adhoc;
    if codice_nom_numerico = 'si' then
    begin

      if trim(codice_adhoc) <> '' then
        cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);
    end
    else
    begin
      cli_for := trim(codice_adhoc);
    end;

  end;
end;

procedure TCNVESA.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  adoEsatto.Connected := False;
end;

procedure TCNVESA.FormShow(Sender: TObject);
begin
  inherited;
  adoEsatto.Connected := false;
  // adoEsatto.ConnectionString :=' FILE NAME='+cartella_installazione + 'ESATTO.udl';
  // ShowMessage(adoEsatto.ConnectionString);
  adoEsatto.Connected := true;
end;

function TCNVESA.Converti_data(data: string): TDateTime;
begin
  try
    if Length(data) = 8 then
      Result := StrToDate(Copy(data, 7, 2) + DateSeparator + Copy(data, 5, 2) + Dateseparator + Copy(data, 1, 4))
    else
      Result := StrToDate('01/01/1900');
  except
    if Copy(data, 5, 2) <> '2' then
    begin
      Result := StrToDate('30' + DateSeparator + Copy(data, 5, 2) + Dateseparator + Copy(data, 1, 4))
    end
    else
    begin
      Result := StrToDate('28' + DateSeparator + Copy(data, 5, 2) + Dateseparator + Copy(data, 1, 4))
    end;

  end;
end;

procedure TCNVESA.crea_tsm(sconto1, sconto2: double);
begin
  tsm.append;

  tsm_codice := tsm_codice + 1;
  tsm.fieldbyname('codice').asstring := setta_lunghezza(tsm_codice, 4, 0);
  tsm.fieldbyname('descrizione').asstring := floattostr(sconto1);
  if sconto2 <> 0 then
  begin
    tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring +
      ' + ' + floattostr(sconto2);
  end;
  tsm.fieldbyname('sconto_maggiorazione').asstring := 'sconto';
  tsm.fieldbyname('percentuale_01').asfloat := sconto1;
  tsm.fieldbyname('percentuale_02').asfloat := sconto2;

  tsm.fieldbyname('percentuale_totale').asfloat := 100;
  if sconto2 <> 0 then
  begin
    tsm.fieldbyname('percentuale_totale').asfloat := (100 - sconto1) *
      (100 - sconto2) / 100;
  end
  else
  begin
    if sconto1 <> 0 then
    begin
      tsm.fieldbyname('percentuale_totale').asfloat := (100 - sconto1) / 1;
    end;
  end;

  tsm.post;
end;

procedure TCNVESA.converti_provvigioni;
begin
  v_tabella_01.caption := 'provvigioni agenti';
  application.processmessages;

  query_02.Close;
  query_02.sql.add('select * from esa_tag');
  query_02.sql.add('where esa_codice =:esa_codice');
  query_02.parambyname('esa_codice').asstring := '0000';
  try
    query_02.open;
  except
    converti_tag;
  end;

  cancella_tabella('pro');
  tabella_01.close;
  tabella_01.tablename := 'pro';
  tabella_01.open;

  tabella_01_ds.Dataset := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.Sql.Clear;
  tabella_esa_01.Sql.Add('SELECT * FROM TESTATEPROVVIGIONI');
  tabella_esa_01.Sql.Add('WHERE');
  tabella_esa_01.Sql.Add('PARTE_FISSA=:PF ');
  tabella_esa_01.Sql.Add('ORDER BY Serie_documento,Numero_documento');
  tabella_esa_01.Parameters.ParamByName('PF').Value := 'T';

  tabella_esa_01.Open;

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin

    tabella_01.append;

    tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_documento').asstring);
    tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('serie_documento').asstring;
    tabella_01.fieldbyname('numero_documento').asinteger := StrToInt(trim(tabella_esa_01.fieldbyname('numero_documento').asstring));
    tabella_01.fieldbyname('data_documento').asdatetime := converti_data(tabella_esa_01.fieldbyname('data_documento').asstring);
    tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;

    assegna_codice_cli(tabella_esa_01.fieldbyname('Codice_cliente').asstring, cli_for);

    tabella_01.fieldbyname('cli_codice').asstring := cli_for;
    if trim(tabella_esa_01.fieldbyname('Codice_Agente').asstring) = '' then
      tabella_01.fieldbyname('tag_codice').asstring := '0000'
    else
    begin
      query_02.close;
      query_02.parambyname('esa_codice').asstring := trim(tabella_esa_01.fieldbyname('Codice_Agente').asstring);
      query_02.open;
      if not query.Eof then
      begin
        tabella_01.fieldbyname('tag_codice').asstring := query_02.fieldbyname('Codice').asstring;
      end
      else
      begin
        tabella_01.fieldbyname('tag_codice').asstring := '0000';
      end;
    end;
    tabella_01.fieldbyname('cambio').asfloat := 1;

    if tabella_esa_01.fieldbyname('tipo_documento').asstring = 'F' then
    begin
      tabella_01.fieldbyname('importo_imponibile').asfloat := tabella_esa_01.fieldbyname('totale_imponibile').asfloat;
      tabella_01.fieldbyname('importo_provvigioni').asfloat := tabella_esa_01.fieldbyname('totale_provvigioni').asfloat;
    end
    else
    begin
      tabella_01.fieldbyname('importo_imponibile').asfloat := abs(tabella_esa_01.fieldbyname('totale_imponibile').asfloat) * -1;
      tabella_01.fieldbyname('importo_provvigioni').asfloat := abs(tabella_esa_01.fieldbyname('totale_provvigioni').asfloat) * -1;
    end;

    tabella_01.post;

    tabella_esa_01.next;
  end;

  query_02.Close;
  tabella_01.close;
  tabella_02.close;
  tabella_03.close;
  tabella_esa_01.close;
  tabella_esa_02.close;
end;

procedure TCNVESA.converti_fatture_vendita(tipo_documento: string);
var
  flag_fatturati, flag_tipo_documento, flag_codice_co, flag_sconto, flag_agente: boolean;

  indice, contatore, i, pos: integer;
  data_consegna: tdatetime;
  sconto_cli, sconto_art, note: string;
  imponibile: double;
  codice_esercizio, situazione, tma_codice: string;
  trovato_testata: boolean;
  comodo_stringa: string;
begin
end;

procedure TCNVESA.cancella_tabella(tabella: string);
begin

  query.sql.clear;
  query.sql.Add('delete from ' + tabella);
  query.execsql;

end;

initialization

registerclass(tcnvesa);

end.
