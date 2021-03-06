unit GGCONVEREADY;

interface

uses
  Windows, Messages, SysUtils, DateUtils, Variants, Classes, Graphics, Controls, Forms,
  Vcl.Dialogs, GGELABORA, Grids, dbgrids, RzDBGrid, ADODB, DB, MyAccess, query_go,
  Menus, StdCtrls, Buttons, ComCtrls, RzTabs, ExtCtrls,
  ToolWin, Mask, MemDS, VirtualTable,

  RzButton, rzLabel, RzPanel, RzDBEdit, RzListVw, RzTreeVw, RzDBChk,
  RzRadChk, RzSplit, RzCmboBx, RzPrgres,
  RzSpnEdt, RzShellDialogs, RzDBCmbo, raizeedit_go, RzEdit, DBAccess, zzlibrerie, ZZTOTVEN;

type

  TCONVEREADY = class(TELABORA)
    v_griglia: TRzDBGrid_go;
    v_tabella_01: TRzlabel;
    v_tabella: TRzlabel;
    tabella_01: tmytable;
    tabella_02: tmytable;
    cfg: tmytable;
    tabella_01_ds: tmydatasource;
    GroupBox1: TGroupBox;
    v_sottoconti: TRzcheckbox;
    v_clienti: TRzcheckbox;
    v_fornitori: TRzcheckbox;
    v_articoli: TRzcheckbox;
    v_sedi_amministrative: TRzcheckbox;
    v_lsv: TRzcheckbox;
    v_pnt: TRzcheckbox;
    v_scadenze: TRzcheckbox;
    v_mov: TRzcheckbox;
    v_tabelle: TRzcheckbox;
    v_ordini: TRzcheckbox;
    v_codifica_clienti: TRzcheckbox;
    v_codifica_fornitori: TRzcheckbox;
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
    DBGrid1: TDBGrid;
    v_tma_codice: TRzEdit;
    Label1: TLabel;
    taq: tmytable;
    v_pni: TRzcheckbox;
    v_fatture_vendita: TRzcheckbox;
    v_tdo: TRzcheckbox;
    v_nml: TRzcheckbox;
    procedure v_confermaClick(Sender: TObject);
    procedure tabella_02BeforePost(DataSet: TDataSet);
    procedure tabella_03BeforePost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tabella_01BeforePost(DataSet: TDataSet);
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
    procedure converti_contatti;
    procedure converti_fornitori;
    procedure converti_articoli;
    procedure crea_cpv;
    procedure crea_cpa;
    procedure converti_bar;
    procedure converti_sedi_amministrative;
    procedure converti_listino(nome_tabella, tipo_listino: string);
    procedure converti_pnt;
    procedure converti_pni;
    procedure crea_pnt(progressivo: integer);
    procedure crea_pnr(progressivo: integer);
    procedure crea_pni(progressivo: integer);
    procedure converti_par;
    procedure converti_mov;
    procedure crea_mmt(progressivo: integer);
    procedure crea_mmr;

    procedure converti_ordini_clienti;
    procedure crea_ovt;
    procedure crea_ovr(progressivo: integer);
    procedure converti_ordini_fornitori;
    procedure crea_oat;
    procedure crea_oar(progressivo: integer);
    procedure converti_fatture_vendita;
    procedure crea_fvt;
    procedure crea_fvi;
    procedure crea_fvr(progressivo: integer);

    procedure assegna_codice_cli(codice_adhoc: string; var cli_for: string);
    procedure assegna_codice_frn(codice_adhoc: string; var cli_for: string);

    procedure crea_tsm(sconto1, sconto2: double);
    procedure converti_provvigioni;
    procedure cancella_tabella(tabella: string);

  public
    { Public declarations }
    procedure controllo_campi; override;
  end;

var
  CONVEREADY: TCONVEREADY;

implementation

{$R *.dfm}

uses DMARC;

procedure TCONVEREADY.v_confermaClick(Sender: TObject);
begin
  v_conferma.enabled := false;

  (*
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
 *)
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

    converti_tzo;

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
    converti_contatti;
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
    // converti_bar;
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

  if v_pni.checked then
  begin
    converti_pni;
  end;

  if v_lsv.checked then
  begin
    converti_listino('lsv', 'C');
    converti_listino('lsa', 'F');
  end;

  if v_mov.checked then
  begin
    converti_mov;
  end;

  if v_sedi_amministrative.checked then
  begin
    converti_sedi_amministrative;
  end;

  if v_ordini.checked then
  begin
    converti_ordini_clienti;
    // converti_ordini_fornitori;
  end;

  if v_fatture_vendita.checked then
  begin
    converti_fatture_vendita;
  end;

  if v_provvigioni.checked then
  begin
    // converti_provvigioni;
  end;
  if v_tdo.checked then
  begin
    converti_tdo;
  end;

  if v_nml.checked then
  begin
    converti_contatti;
  end;

  v_conferma.enabled := true;
  Close;
end;

procedure TCONVEREADY.converti_clienti;
var
  i, j: word;
  progressivo: integer;
  data_nascita: TDateTime;
  stringa: string;
  iban: tiban;
begin
  progressivo := 0;
  iban := tiban.create;

  v_tabella.caption := 'clienti';
  application.processmessages;

  read_tabella(arc.arc, 'dit', 'codice', ditta, '*');

  cancella_tabella('cli');
  tabella_01.Close;
  tabella_01.tablename := 'cli';
  tabella_01.open;

  cancella_tabella('nom');
  tabella_02.tablename := 'nom';
  tabella_02.open;

  cancella_tabella('cfg');
  cfg.tablename := 'cfg';
  cfg.open;

  tabella_01_ds.DataSet := tabella_clifor;

  tabella_clifor.Close;
  tabella_clifor.Parameters.ParamByName('ind_clifor').Value := 'C';
  tabella_clifor.Parameters.ParamByName('ind_cliforF').Value := 'C';
  tabella_clifor.open;
  while not tabella_clifor.eof do
  begin
    application.processmessages;
    // ------------------------------------
    // CERCO CODICE IN TABELLA CLIENTI-FORNITORI
    // ------------------------------------

    if tabella_clifor.fieldbyname('cod_anagra').asstring <> '' then
    begin
      if (codice_nom_numerico = 'si') and (v_codifica_clienti.checked) then
      begin
        progressivo := progressivo + 1;
        cli_for := setta_lunghezza(progressivo, 8, 0);
      end
      else
      begin
        assegna_codice_cli(tabella_clifor.fieldbyname('cod_anagra').asstring, cli_for);
      end;

      if length(cli_for) > 0 then
      begin
        tabella_02.append;

        tabella_02.fieldbyname('codice').asstring := cli_for;

        if length(trim(tabella_clifor.fieldbyname('des_ragsoc').asstring)) > 30 then
        begin
          for i := 30 downto 1 do
          begin
            if tabella_clifor.fieldbyname('des_ragsoc').asstring[i] = ' ' then
            begin
              j := i;
              break;
            end;
          end;
        end
        else
        begin
          j := length(trim(tabella_clifor.fieldbyname('des_ragsoc').asstring));
        end;
        tabella_02.fieldbyname('descrizione1').asstring := trim(copy(tabella_clifor.fieldbyname('des_ragsoc').asstring, 1, j));

        if tabella_02.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_02.fieldbyname('descrizione1').asstring := '.';
        end;
        tabella_02.fieldbyname('descrizione2').asstring := trim(copy(tabella_clifor.fieldbyname('des_ragsoc').asstring, j + 1, length(tabella_clifor.fieldbyname('des_ragsoc').asstring) - j));

        tabella_02.fieldbyname('via').asstring := trim(tabella_clifor.fieldbyname('des_indir_res').asstring);
        tabella_02.fieldbyname('cap').asstring := trim(tabella_clifor.fieldbyname('cod_cap_res').asstring);
        tabella_02.fieldbyname('citta').asstring := trim(tabella_clifor.fieldbyname('des_localita_res').asstring);
        tabella_02.fieldbyname('provincia').asstring := trim(tabella_clifor.fieldbyname('sig_prov_res').asstring);
        tabella_02.fieldbyname('tna_codice').asstring := trim(tabella_clifor.fieldbyname('cod_naz_res').asstring);
        if tabella_02.fieldbyname('tna_codice').asstring = '' then
        begin
          tabella_02.fieldbyname('tna_codice').asstring := arc.dit.fieldbyname('tna_codice').asstring;
        end;

        tabella_02.fieldbyname('partita_iva').asstring := tabella_clifor.fieldbyname('cod_piva').asstring;
        tabella_02.fieldbyname('codice_fiscale').asstring := tabella_clifor.fieldbyname('cod_cfisc').asstring;
        tabella_02.fieldbyname('telefono').asstring := trim(tabella_clifor.fieldbyname('sig_tel_res').asstring);
        tabella_02.fieldbyname('telefono_01').asstring := trim(tabella_clifor.fieldbyname('sig_tel1_res').asstring);
        tabella_02.fieldbyname('fax').asstring := trim(tabella_clifor.fieldbyname('sig_fax_res').asstring);
        tabella_02.fieldbyname('cellulare').asstring := trim(tabella_clifor.fieldbyname('sig_fax1_res').asstring);
        if (tabella_clifor.fieldbyname('flg_pers_fis').asstring = '1') and (tabella_02.fieldbyname('partita_iva').asstring = '') then
        begin
          tabella_02.fieldbyname('codice_alternativo').asstring := trim(tabella_clifor.fieldbyname('des_ragsoc').asstring);

          if trim(tabella_clifor.fieldbyname('des_cogn').asstring) <> '' then
          begin
            tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_clifor.fieldbyname('des_cogn').asstring);
            tabella_02.fieldbyname('descrizione2').asstring := trim(tabella_clifor.fieldbyname('des_nome').asstring);
          end;

          tabella_02.fieldbyname('persona_fisica').asstring := 'si';
          if tabella_clifor.fieldbyname('ind_sesso').asstring = '0' then
          begin
            tabella_02.fieldbyname('sesso').asstring := 'femminile';
          end
          else
          begin
            tabella_02.fieldbyname('sesso').asstring := 'maschile';
          end;

          if tabella_clifor.fieldbyname('dat_nasc').asvariant <> null then
          begin
            tabella_02.fieldbyname('data_nascita').asdatetime := tabella_clifor.fieldbyname('dat_nasc').asdatetime;
          end;

          tabella_02.fieldbyname('citta_nascita').asstring := tabella_clifor.fieldbyname('des_localita_nasc').asstring;
          tabella_02.fieldbyname('provincia_nascita').asstring := tabella_clifor.fieldbyname('sig_prov_nasc').asstring;
        end
        else
        begin
          tabella_02.fieldbyname('persona_fisica').asstring := 'no';
        end;

        if trim(tabella_clifor.fieldbyname('cod_valuta').asstring) <> '' then
        begin
          tabella_02.fieldbyname('tva_codice').asstring := trim(tabella_clifor.fieldbyname('cod_valuta').asstring);
        end
        else
        begin
          tabella_02.fieldbyname('tva_codice').asstring := divisa_di_conto;
        end;

        // tabella_02.fieldbyname('web').asstring := trim(tabella_clifor.fieldbyname('Sito_Internet').asstring);
        tabella_02.fieldbyname('e_mail_amministrazione').asstring := trim(tabella_clifor.fieldbyname('des_email1').asstring);

        tabella_02.fieldbyname('via_legale').asstring := tabella_clifor.fieldbyname('des_indir_fis').asstring;
        tabella_02.fieldbyname('cap_legale').asstring := tabella_clifor.fieldbyname('cod_cap_fis').asstring;
        tabella_02.fieldbyname('citta_legale').asstring := tabella_clifor.fieldbyname('des_localita_fis').asstring;
        tabella_02.fieldbyname('provincia_legale').asstring := tabella_clifor.fieldbyname('sig_prov_fis').asstring;
        tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_clifor.fieldbyname('cod_naz_fis').asstring;

        if tabella_02.fieldbyname('tna_codice_legale').asstring = '' then
        begin
          tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_02.fieldbyname('tna_codice').asstring;
        end;

        tabella_02.fieldbyname('e_mail_pec_fe').asstring := tabella_clifor.fieldbyname('des_email_pec_fe').asstring;

        // ---------------------------------------------------------------------------
        // clienti
        // ---------------------------------------------------------------------------
        tabella_01.append;

        tabella_01.fieldbyname('codice').asstring := tabella_02.fieldbyname('codice').asstring;
        if (codice_nom_numerico = 'si') then
        begin
          tabella_01.fieldbyname('codice_alternativo').asstring := tabella_clifor.fieldbyname('cod_anagra').asstring;
        end;

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
        tabella_01.fieldbyname('gen_codice').asstring := arc.dit.fieldbyname('gen_codice_cli').asstring;
        tabella_01.fieldbyname('tba_codice').asstring := arc.dit.fieldbyname('tba_codice_cli').asstring;
        tabella_01.fieldbyname('tpa_codice').asstring := arc.dit.fieldbyname('tpa_codice_cli').asstring;
        tabella_01.fieldbyname('tcc_codice').asstring := arc.dit.fieldbyname('tcc_codice_cli').asstring;
        tabella_01.fieldbyname('tlv_codice').asstring := arc.dit.fieldbyname('tlv_codice_cli').asstring;
        tabella_01.fieldbyname('ts1_codice').asstring := arc.dit.fieldbyname('ts1_codice_cli').asstring;
        tabella_01.fieldbyname('tzo_codice').asstring := arc.dit.fieldbyname('tzo_codice_cli').asstring;
        tabella_01.fieldbyname('tzo_codice_assistenza').asstring := arc.dit.fieldbyname('tzo_codice_cli').asstring;
        tabella_01.fieldbyname('tsc_codice').asstring := arc.dit.fieldbyname('tsc_codice_cli').asstring;
        tabella_01.fieldbyname('tsp_codice').asstring := arc.dit.fieldbyname('tsp_codice_cli').asstring;
        tabella_01.fieldbyname('tpo_codice').asstring := arc.dit.fieldbyname('tpo_codice_cli').asstring;
        tabella_01.fieldbyname('tag_codice').asstring := arc.dit.fieldbyname('tag_codice_cli').asstring;
        tabella_01.fieldbyname('tp1_codice').asstring := arc.dit.fieldbyname('tp1_codice_cli').asstring;
        tabella_01.fieldbyname('tpf_codice').asstring := arc.dit.fieldbyname('tpf_codice_cli').asstring;
        tabella_01.fieldbyname('tst_codice').asstring := arc.dit.fieldbyname('tst_codice_cli').asstring;
        tabella_01.fieldbyname('addebito_spese_fattura').asstring := arc.dit.fieldbyname('addebito_spese_fattura_clienti').asstring;
        tabella_01.fieldbyname('tar_codice').asstring := arc.dit.fieldbyname('tar_codice_cli').asstring;
        tabella_01.fieldbyname('tcg_codice').asstring := arc.dit.fieldbyname('tcg_codice_cli').asstring;

        tabella_01.fieldbyname('contatto').asstring := tabella_clifor.fieldbyname('des_contatto').asstring;

        if tabella_clifor.fieldbyname('cod_piacont').asstring <> '' then
        begin
          tabella_01.fieldbyname('gen_codice').asstring := tabella_clifor.fieldbyname('cod_piacont').asstring;
        end;

        if tabella_clifor.fieldbyname('cod_contropartita').asstring <> '' then
        begin
          tabella_01.fieldbyname('cfg_tipo_01').asstring := 'G';
          tabella_01.fieldbyname('cfg_codice_01').asstring := tabella_clifor.fieldbyname('cod_contropartita').asstring;
        end;

        if trim(tabella_clifor.fieldbyname('flg_gest_part').asstring) = '1' then
        begin
          tabella_01.fieldbyname('partitario').asstring := 'si';
        end
        else
        begin
          tabella_01.fieldbyname('partitario').asstring := 'no';
        end;

        if trim(tabella_clifor.fieldbyname('cod_pag').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tpa_codice').asstring := trim(tabella_clifor.fieldbyname('cod_pag').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_vettore').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tsp_codice').asstring := trim(tabella_clifor.fieldbyname('cod_vettore').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_porto').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tpo_codice').asstring := trim(tabella_clifor.fieldbyname('cod_porto').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_zona').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tzo_codice').asstring := trim(tabella_clifor.fieldbyname('cod_zona').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_agente').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tag_codice').asstring := copy(tabella_clifor.fieldbyname('cod_agente').asstring, 3, 4);
        end;

        if trim(tabella_clifor.fieldbyname('cod_iva').asstring) <> '21' then
        begin
          tabella_01.fieldbyname('tiv_codice').asstring := trim(tabella_clifor.fieldbyname('cod_iva').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_banca1').asstring) <> '' then
        begin
          tabella_01.fieldbyname('codice_abi').asstring := trim(tabella_clifor.fieldbyname('cod_banca1').asstring);
          tabella_01.fieldbyname('codice_cab').asstring := trim(tabella_clifor.fieldbyname('cod_agenzia1').asstring);
        end;
        if tabella_01.fieldbyname('codice_abi').asstring <> '' then
        begin
          if length(tabella_01.fieldbyname('codice_abi').asstring) = 4 then
          begin
            tabella_01.fieldbyname('codice_abi').asstring := '0' + tabella_01.fieldbyname('codice_abi').asstring;
          end;
        end;
        tabella_01.fieldbyname('mese_01').asInteger := 0;

        if trim(tabella_clifor.fieldbyname('num_mese_escl1').asstring) <> '' then
          tabella_01.fieldbyname('mese_01').asInteger := tabella_clifor.fieldbyname('num_mese_escl1').asInteger;

        if tabella_01.fieldbyname('mese_01').asInteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_01').asInteger := tabella_clifor.fieldbyname('num_gio_mese_suc').asInteger;
        end;
        tabella_01.fieldbyname('mese_02').asInteger := 0;

        if trim(tabella_clifor.fieldbyname('num_mese_escl2').asstring) <> '' then
          tabella_01.fieldbyname('mese_02').asInteger := tabella_clifor.fieldbyname('num_mese_escl2').asInteger;

        if tabella_01.fieldbyname('mese_02').asInteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_02').asInteger := tabella_clifor.fieldbyname('num_gio_mese_suc2').asInteger;
        end;

        if (tabella_01.fieldbyname('mese_01').asInteger <> 0) or (tabella_01.fieldbyname('mese_02').asInteger <> 0) then
        begin
          tabella_01.fieldbyname('mesi_esclusi').asstring := 'si';
        end;

        if trim(tabella_clifor.fieldbyname('cod_listino').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tlv_codice').asstring := trim(tabella_clifor.fieldbyname('cod_listino').asstring);
        end;
        (*
          if trim(tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring) = '0' then
          begin
          tabella_01.fieldbyname('riepilogo_fattura').asstring := 'nessuno';
          end
          else if tabella_clifor.fieldbyname('Flag_Tipo_Fatturaz').asstring = '1' then
          begin
          tabella_01.fieldbyname('riepilogo_fattura').asstring := 'globale';
          end;
 *)
        if (tabella_clifor.fieldbyname('flg_addebito_bollo').asstring = '1') then
        begin
        end;

        if (tabella_clifor.fieldbyname('flg_addebito_spese').asstring = '1') then
        begin
          tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'si';
        end
        else
        begin
          tabella_01.fieldbyname('addebito_spese_fattura').asstring := 'no';
        end;

        (*
          if (tabella_clifor.fieldbyname('Stampa Prezzo Bolla').asstring = '1') then
          begin
          tabella_01.fieldbyname('valori_in_bolla').asstring := 'si';
          end
          else
          begin
          tabella_01.fieldbyname('valori_in_bolla').asstring := 'no';
          end;
 *)
        try
          tabella_01.fieldbyname('fido').asFloat := tabella_clifor.fieldbyname('val_fido').asFloat;
        except
          tabella_01.fieldbyname('fido').asFloat := 0;
        end;

        if trim(tabella_clifor.fieldbyname('sig_bban1').asstring) > '' then
        begin
          tabella_01.fieldbyname('cin').asstring := copy(trim(tabella_clifor.fieldbyname('sig_bban1').asstring), 1, 1);
          tabella_01.fieldbyname('codice_abi').asstring := copy(trim(tabella_clifor.fieldbyname('sig_bban1').asstring), 2, 5);
          tabella_01.fieldbyname('codice_cab').asstring := copy(trim(tabella_clifor.fieldbyname('sig_bban1').asstring), 7, 5);
          tabella_01.fieldbyname('conto_corrente').asstring := copy(trim(tabella_clifor.fieldbyname('sig_bban1').asstring), 12, 12);

          if length(tabella_01.fieldbyname('conto_corrente').asstring) < 12 then
          begin
            for i := 1 to (12 - length(tabella_01.fieldbyname('conto_corrente').asstring)) do
            begin
              tabella_01.fieldbyname('conto_corrente').asstring := '0' + tabella_01.fieldbyname('conto_corrente').asstring;
            end;
          end;

          stringa := iban.calcola_iban('IT', tabella_01.fieldbyname('codice_abi').asstring, tabella_01.fieldbyname('codice_cab').asstring, tabella_01.fieldbyname('conto_corrente').asstring);

          tabella_01.fieldbyname('cin').asstring := copy(stringa, 5, 1);
          tabella_01.fieldbyname('iban').asstring := copy(stringa, 1, 4) + ' ' + copy(stringa, 5, 4) + ' ' + copy(stringa, 9, 4) + ' ' + copy(stringa, 13, 4) + ' ' + copy(stringa, 17, 4) + ' ' + copy(stringa, 21, 4) + ' ' + copy(stringa, 25, 3);

        end;

        tabella_01.fieldbyname('stampa_codice_articolo_cliente').asstring := 'no';
        tabella_01.fieldbyname('codice_ufficio_pa').asstring := tabella_clifor.fieldbyname('cod_idest').asstring;
        tabella_01.post;

        // file cfg

        if cfg.locate('cfg_tipo;cfg_codice', vararrayof(['C', tabella_01.fieldbyname('codice').asstring]), []) then
        begin
          cfg.edit;
          cfg.fieldbyname('descrizione1').asstring := trim(tabella_01.fieldbyname('descrizione1').asstring) + ' ' + trim(tabella_01.fieldbyname('descrizione2').asstring);
          cfg.fieldbyname('descrizione2').asstring := trim(tabella_01.fieldbyname('citta').asstring);
          cfg.fieldbyname('utente').asstring := utente;
          cfg.fieldbyname('data_ora').asdatetime := now;
          cfg.post;
        end;
      end; // if!
    end;

    tabella_clifor.next;
  end;

  query_02.Close;
  tabella_01.Close;
  tabella_02.Close;
  tabella_clifor.Close;
  cfg.Close;
end;

procedure TCONVEREADY.converti_contatti;
var
  i, j: word;
  progressivo: integer;
  data_nascita: TDateTime;
  stringa: string;
  iban: tiban;
begin
  progressivo := 0;
  iban := tiban.create;

  v_tabella.caption := 'contatti clienti';
  application.processmessages;

  read_tabella(arc.arc, 'dit', 'codice', ditta, '*');

  cancella_tabella('nml');
  tabella_02.Close;
  tabella_02.tablename := 'nml';
  tabella_02.open;

  tabella_01_ds.DataSet := tabella_esa_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT c.* ');
  tabella_esa_01.sql.add('from CA_CONTATTI c');
  tabella_esa_01.sql.add('where ');
  tabella_esa_01.sql.add('(c.des_cogn is not null ) or ');
  tabella_esa_01.sql.add('(c.des_nome is not null )  ');
  tabella_esa_01.sql.add('order by c.cod_anagra');
  tabella_esa_01.open;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;
    // ------------------------------------
    // CERCO CODICE IN TABELLA CLIENTI-FORNITORI
    // ------------------------------------

    if tabella_esa_01.fieldbyname('cod_anagra').asstring <> '' then
    begin
      if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_anagra').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;

        if archivio.eof then
        begin
          read_tabella(arc.arcdit, 'frn', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_anagra').asstring);
          cli_for := archivio.fieldbyname('codice').asstring;
        end;
      end
      else
      begin
        assegna_codice_cli(tabella_esa_01.fieldbyname('cod_anagra').asstring, cli_for);
      end;

      tabella_02.append;
      tabella_02.fieldbyname('nom_codice').asstring := cli_for;
      tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_cogn').asstring + ' ' + tabella_esa_01.fieldbyname('des_nome').asstring);
      if tabella_02.fieldbyname('descrizione').asstring = '' then
      begin
        tabella_02.fieldbyname('descrizione').asstring := tabella_02.fieldbyname('id').asstring;
      end;
      tabella_02.fieldbyname('mansione').asstring := trim(tabella_esa_01.fieldbyname('cod_reparto_crm').asstring);
      tabella_02.fieldbyname('telefono').asstring := trim(tabella_esa_01.fieldbyname('sig_tel').asstring);
      tabella_02.fieldbyname('telefono_01').asstring := trim(tabella_esa_01.fieldbyname('sig_tel1').asstring);
      tabella_02.fieldbyname('cellulare').asstring := trim(tabella_esa_01.fieldbyname('sig_tel2').asstring);
      tabella_02.fieldbyname('fax').asstring := trim(tabella_esa_01.fieldbyname('sig_fax1').asstring);
      tabella_02.fieldbyname('email').asstring := trim(tabella_esa_01.fieldbyname('des_email1').asstring);
      tabella_02.fieldbyname('note').asstring := trim(tabella_esa_01.fieldbyname('des_note').asstring);
      tabella_02.post;
    end;

    tabella_esa_01.next;
  end;

  tabella_esa_01.Close;
  tabella_02.Close;
end;

procedure TCONVEREADY.converti_fornitori;
var
  progressivo: integer;
  i, j: word;
  data_nascita: TDateTime;
begin
  v_tabella_01.caption := 'fornitori';
  application.processmessages;

  read_tabella(arc.arc, 'dit', 'codice', ditta, '*');

  tabella_01.Close;
  tabella_01.tablename := 'frn';
  tabella_01.open;

  cancella_tabella('frn');

  tabella_02.tablename := 'nom';
  tabella_02.open;
  if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
  begin
    tabella_02.last;
    progressivo := tabella_02.fieldbyname('codice').asInteger;
  end;

  cfg.tablename := 'cfg';
  cfg.open;

  tabella_01_ds.DataSet := tabella_clifor;

  tabella_clifor.Close;
  tabella_clifor.Parameters.ParamByName('ind_clifor').Value := 'F';
  tabella_clifor.Parameters.ParamByName('ind_cliforF').Value := 'F';
  tabella_clifor.open;

  // --------------------------------------------------------
  // TABELLA ANAGRAFICHECOMUNI
  // --------------------------------------------------------
  try
    tabella_esa_03.Close;
    tabella_esa_03.sql.clear;
    tabella_esa_03.sql.add('SELECT * from CA_ANAGRAFICHE');
    tabella_esa_03.sql.add('where cod_anagra=:cod_anagra');
  except
    messaggio(000, 'manca l''anagrafica comune (CA_ANAGRAFICHE');
    Close;
    abort;
  end;

  while not tabella_clifor.eof do
  begin
    application.processmessages;

    // ------------------------------------
    // CERCO CODICE IN TABELLA CLIENTI-FORNITORI
    // ------------------------------------
    tabella_esa_03.Close;
    tabella_esa_03.Parameters.ParamByName('cod_anagra').Value := tabella_clifor.fieldbyname('cod_anagra').asstring;
    tabella_esa_03.open;
    if not tabella_esa_03.eof then
    begin
      if (codice_nom_numerico = 'si') and (v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_clifor.fieldbyname('cod_anagra').asstring);
        if archivio.eof then
        begin
          progressivo := progressivo + 1;
          cli_for := setta_lunghezza(progressivo, 8, 0);
        end
        else
        begin
          cli_for := archivio.fieldbyname('codice').asstring;
        end;
      end
      else
      begin
        assegna_codice_frn(tabella_esa_03.fieldbyname('cod_anagra').asstring, cli_for);
      end;

      // nominativi
      if not tabella_02.locate('codice', cli_for, []) then
      begin

        if tabella_clifor.fieldbyname('cod_clifor').asstring <> '' then
        begin
          tabella_02.append;

          tabella_02.fieldbyname('codice').asstring := cli_for;

          if length(trim(tabella_esa_03.fieldbyname('des_ragsoc').asstring)) > 30 then
          begin
            for i := 30 downto 1 do
            begin
              if tabella_esa_03.fieldbyname('des_ragsoc').asstring[i] = ' ' then
              begin
                j := i;
                break;
              end;
            end;
          end
          else
          begin
            j := length(trim(tabella_esa_03.fieldbyname('des_ragsoc').asstring));
          end;
          tabella_02.fieldbyname('descrizione1').asstring := trim(copy(tabella_esa_03.fieldbyname('des_ragsoc').asstring, 1, j));

          if tabella_02.fieldbyname('descrizione1').asstring = '' then
          begin
            tabella_02.fieldbyname('descrizione1').asstring := '.';
          end;
          tabella_02.fieldbyname('descrizione2').asstring := trim(copy(tabella_esa_03.fieldbyname('des_ragsoc').asstring, j + 1, length(tabella_esa_03.fieldbyname('des_ragsoc').asstring) - j));

          tabella_02.fieldbyname('via').asstring := trim(tabella_clifor.fieldbyname('des_indir_res').asstring);
          tabella_02.fieldbyname('cap').asstring := trim(tabella_clifor.fieldbyname('cod_cap_res').asstring);
          tabella_02.fieldbyname('citta').asstring := trim(tabella_clifor.fieldbyname('des_localita_res').asstring);
          tabella_02.fieldbyname('provincia').asstring := trim(tabella_clifor.fieldbyname('sig_prov_res').asstring);
          tabella_02.fieldbyname('tna_codice').asstring := trim(tabella_clifor.fieldbyname('cod_naz_res').asstring);
          if tabella_02.fieldbyname('tna_codice').asstring = '' then
          begin
            tabella_02.fieldbyname('tna_codice').asstring := 'IT';
          end;
          tabella_02.fieldbyname('partita_iva').asstring := tabella_clifor.fieldbyname('cod_piva').asstring;
          tabella_02.fieldbyname('codice_fiscale').asstring := tabella_clifor.fieldbyname('cod_cfisc').asstring;
          tabella_02.fieldbyname('telefono').asstring := trim(tabella_clifor.fieldbyname('sig_tel1_res').asstring);
          tabella_02.fieldbyname('fax').asstring := trim(tabella_clifor.fieldbyname('sig_fax_res').asstring);
          tabella_02.fieldbyname('cellulare').asstring := trim(tabella_clifor.fieldbyname('sig_tel2_res').asstring);
          if tabella_clifor.fieldbyname('flg_pers_fis').asstring = '1' then
          begin
            tabella_02.fieldbyname('codice_alternativo').asstring := trim(tabella_clifor.fieldbyname('des_ragsoc').asstring);

            if trim(tabella_clifor.fieldbyname('des_cogn').asstring) <> '' then
            begin
              tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_clifor.fieldbyname('des_cogn').asstring);
              tabella_02.fieldbyname('descrizione2').asstring := trim(tabella_clifor.fieldbyname('des_nome').asstring);
            end;

            tabella_02.fieldbyname('persona_fisica').asstring := 'si';
            if tabella_clifor.fieldbyname('ind_sesso').asstring = '0' then
            begin
              tabella_02.fieldbyname('sesso').asstring := 'femminile';
            end
            else
            begin
              tabella_02.fieldbyname('sesso').asstring := 'maschile';
            end;

            if tabella_clifor.fieldbyname('dat_nasc').asvariant <> null then
            begin
              tabella_02.fieldbyname('data_nascita').asdatetime := tabella_clifor.fieldbyname('dat_nasc').asdatetime;
            end;

            tabella_02.fieldbyname('citta_nascita').asstring := tabella_clifor.fieldbyname('des_localita_nasc').asstring;
            tabella_02.fieldbyname('provincia_nascita').asstring := tabella_clifor.fieldbyname('sig_prov_nasc').asstring;
          end
          else
          begin
            tabella_02.fieldbyname('persona_fisica').asstring := 'no';
          end;

          if trim(tabella_clifor.fieldbyname('cod_valuta').asstring) <> '' then
          begin
            tabella_02.fieldbyname('tva_codice').asstring := trim(tabella_clifor.fieldbyname('cod_valuta').asstring);
          end
          else
          begin
            tabella_02.fieldbyname('tva_codice').asstring := divisa_di_conto;
          end;

          if codice_nom_numerico = 'si' then
          begin
            tabella_02.fieldbyname('codice_alternativo').asstring := trim(tabella_clifor.fieldbyname('cod_anagra').asstring);
          end;
          // tabella_02.fieldbyname('web').asstring := trim(tabella_clifor.fieldbyname('Sito_Internet').asstring);
          tabella_02.fieldbyname('e_mail_amministrazione').asstring := trim(tabella_clifor.fieldbyname('des_email1').asstring);

          tabella_02.fieldbyname('via_legale').asstring := tabella_clifor.fieldbyname('des_indir_fis').asstring;
          tabella_02.fieldbyname('cap_legale').asstring := tabella_clifor.fieldbyname('cod_cap_fis').asstring;
          tabella_02.fieldbyname('citta_legale').asstring := tabella_clifor.fieldbyname('des_localita_fis').asstring;
          tabella_02.fieldbyname('provincia_legale').asstring := tabella_clifor.fieldbyname('sig_prov_fis').asstring;
          tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_clifor.fieldbyname('cod_naz_fis').asstring;

          if tabella_02.fieldbyname('tna_codice_legale').asstring = '' then
          begin
            tabella_02.fieldbyname('tna_codice_legale').asstring := tabella_02.fieldbyname('tna_codice').asstring;
          end;

          tabella_02.post;
        end;
      end;

      // ---------------------------------------------------------------------------
      // fornitori
      // ---------------------------------------------------------------------------
      try
        tabella_01.append;

        tabella_01.fieldbyname('codice').asstring := cli_for;
        if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
        begin
          tabella_01.fieldbyname('codice_alternativo').asstring := tabella_clifor.fieldbyname('cod_anagra').asstring;
        end;

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
        tabella_01.fieldbyname('gen_codice').asstring := arc.dit.fieldbyname('gen_codice_frn').asstring;
        tabella_01.fieldbyname('tba_codice').asstring := arc.dit.fieldbyname('tba_codice_frn').asstring;
        tabella_01.fieldbyname('tpa_codice').asstring := arc.dit.fieldbyname('tpa_codice_frn').asstring;
        tabella_01.fieldbyname('tcf_codice').asstring := arc.dit.fieldbyname('tcf_codice_frn').asstring;
        tabella_01.fieldbyname('tla_codice').asstring := arc.dit.fieldbyname('tla_codice_frn').asstring;
        tabella_01.fieldbyname('ts2_codice').asstring := arc.dit.fieldbyname('ts2_codice_frn').asstring;
        tabella_01.fieldbyname('tzo_codice').asstring := arc.dit.fieldbyname('tzo_codice_frn').asstring;
        tabella_01.fieldbyname('tsc_codice').asstring := arc.dit.fieldbyname('tsc_codice_frn').asstring;
        tabella_01.fieldbyname('tsp_codice').asstring := arc.dit.fieldbyname('tsp_codice_frn').asstring;
        tabella_01.fieldbyname('tpo_codice').asstring := arc.dit.fieldbyname('tpo_codice_frn').asstring;
        tabella_01.fieldbyname('tcf_codice').asstring := arc.dit.fieldbyname('tcf_codice_frn').asstring;

        tabella_01.fieldbyname('contatto').asstring := tabella_clifor.fieldbyname('des_contatto').asstring;

        if tabella_clifor.fieldbyname('cod_piacont').asstring <> '' then
        begin
          tabella_01.fieldbyname('gen_codice').asstring := tabella_clifor.fieldbyname('cod_piacont').asstring;
        end;

        if tabella_clifor.fieldbyname('cod_contropartita').asstring <> '' then
        begin
          tabella_01.fieldbyname('cfg_tipo_01').asstring := 'G';
          tabella_01.fieldbyname('cfg_codice_01').asstring := tabella_clifor.fieldbyname('cod_contropartita').asstring;
        end;

        if trim(tabella_clifor.fieldbyname('flg_gest_part').asstring) = '1' then
        begin
          tabella_01.fieldbyname('partitario').asstring := 'si';
        end
        else
        begin
          tabella_01.fieldbyname('partitario').asstring := 'no';
        end;

        if trim(tabella_clifor.fieldbyname('cod_pag').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tpa_codice').asstring := trim(tabella_clifor.fieldbyname('cod_pag').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_vettore').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tsp_codice').asstring := trim(tabella_clifor.fieldbyname('cod_vettore').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_porto').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tpo_codice').asstring := trim(tabella_clifor.fieldbyname('cod_porto').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_zona').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tzo_codice').asstring := trim(tabella_clifor.fieldbyname('cod_zona').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_iva').asstring) <> '21' then
        begin
          tabella_01.fieldbyname('tiv_codice').asstring := trim(tabella_clifor.fieldbyname('cod_iva').asstring);
        end;

        if trim(tabella_clifor.fieldbyname('cod_banca1').asstring) <> '' then
        begin
          tabella_01.fieldbyname('codice_abi').asstring := trim(tabella_clifor.fieldbyname('cod_banca1').asstring);
          tabella_01.fieldbyname('codice_cab').asstring := trim(tabella_clifor.fieldbyname('cod_agenzia1').asstring);
        end;
        if tabella_01.fieldbyname('codice_abi').asstring <> '' then
        begin
          if length(tabella_01.fieldbyname('codice_abi').asstring) = 4 then
          begin
            tabella_01.fieldbyname('codice_abi').asstring := '0' + tabella_01.fieldbyname('codice_abi').asstring;
          end;
        end;
        tabella_01.fieldbyname('mese_01').asInteger := 0;

        if trim(tabella_clifor.fieldbyname('num_mese_escl1').asstring) <> '' then
          tabella_01.fieldbyname('mese_01').asInteger := tabella_clifor.fieldbyname('num_mese_escl1').asInteger;

        if tabella_01.fieldbyname('mese_01').asInteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_01').asInteger := tabella_clifor.fieldbyname('num_gio_mese_suc').asInteger;
        end;
        tabella_01.fieldbyname('mese_02').asInteger := 0;

        if trim(tabella_clifor.fieldbyname('num_mese_escl2').asstring) <> '' then
          tabella_01.fieldbyname('mese_02').asInteger := tabella_clifor.fieldbyname('num_mese_escl2').asInteger;

        if tabella_01.fieldbyname('mese_02').asInteger <> 0 then
        begin
          tabella_01.fieldbyname('giorno_02').asInteger := tabella_clifor.fieldbyname('num_gio_mese_suc2').asInteger;
        end;

        if (tabella_01.fieldbyname('mese_01').asInteger <> 0) or (tabella_01.fieldbyname('mese_02').asInteger <> 0) then
        begin
          tabella_01.fieldbyname('mesi_esclusi').asstring := 'si';
        end;

        if trim(tabella_clifor.fieldbyname('cod_listino').asstring) <> '' then
        begin
          tabella_01.fieldbyname('tla_codice').asstring := trim(tabella_clifor.fieldbyname('cod_listino').asstring);
        end;

        tabella_01.fieldbyname('conto_corrente').asstring := trim(tabella_clifor.fieldbyname('sig_bban1').asstring);
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
      except
        on E: exception do
        begin
          showmessage(cli_for + '[errore :' + E.message + ']');
        end;
      end;
      (*
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
 *)

      tabella_clifor.next;
    end;
  end;
  tabella_01.Close;
  tabella_02.Close;
  tabella_clifor.Close;
  tabella_esa_02.Close;
  tabella_esa_03.Close;
  cfg.Close;

end;

procedure TCONVEREADY.converti_sottoconti;
begin
  v_tabella_01.caption := 'piano dei conti';
  application.processmessages;

  tabella_01.Close;
  tabella_01.tablename := 'gen';
  tabella_01.open;

  cancella_tabella('gen');

  tabella_02.Close;
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
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CO_PIACONT');
  tabella_esa_01.sql.add('ORDER BY cod_piacont');
  tabella_esa_01.open;

  tabella_01_ds.DataSet := tabella_esa_01;
  while not tabella_esa_01.eof do
  begin
    if tabella_esa_01.fieldbyname('num_livello').asInteger = 3 then
    begin
      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_piacont').asstring);
      tabella_01.fieldbyname('descrizione1').asstring := trim(copy(tabella_esa_01.fieldbyname('des_piacont').asstring, 1, 30));
      tabella_01.fieldbyname('tpc_codice_01').asstring := copy(trim(tabella_esa_01.fieldbyname('cod_piacont_padre').asstring), 1, 2);
      tabella_01.fieldbyname('tpc_codice_02').asstring := copy(trim(tabella_esa_01.fieldbyname('cod_piacont_padre').asstring), 3, 2);
      tabella_01.post;

    end
    else
    begin
      tabella_02.append;

      tabella_02.fieldbyname('codice_01').asstring := copy(trim(tabella_esa_01.fieldbyname('cod_piacont').asstring), 1, 2);

      if tabella_esa_01.fieldbyname('num_livello').asInteger = 2 then
      begin
        tabella_02.fieldbyname('codice_02').asstring := copy(trim(tabella_esa_01.fieldbyname('cod_piacont').asstring), 3, 2);
        tabella_02.fieldbyname('tipo').asstring := '';
      end
      else
      begin
        if (tabella_02.fieldbyname('codice_01').asstring = '03') or (tabella_02.fieldbyname('codice_01').asstring = '04') then
        begin
          tabella_02.fieldbyname('tipo').asstring := 'economico';
        end
        else
        begin
          tabella_02.fieldbyname('tipo').asstring := 'patrimoniale';
        end;
      end;
      tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_piacont').asstring);

      tabella_02.post;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_articoli;
var
  descrizione1, descrizione2, gruppo, sottogru: string;
  tcc_codice, tcf_codice, codice_stat: string;
  tca_codice: word;
  progressivo: integer;
begin
  v_tabella_01.caption := 'articoli';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_02;

  tabella_01.Close;
  tabella_01.tablename := 'art';
  tabella_01.open;

  tabella_02.Close;
  tabella_02.sql.clear;
  tabella_02.sql.add('select codice from tgm');
  tabella_02.sql.add('where');
  tabella_02.sql.add('cod_grumerc=:cod_grumerc and ');
  tabella_02.sql.add('cod_sgrumerc=:cod_sgrumerc ');

  cancella_tabella('art');
  cancella_tabella('cpv');
  cancella_tabella('cpa');

  tca.Close;
  tca.tablename := 'tca';
  tca.open;

  taq.Close;
  taq.tablename := 'taq';
  taq.open;

  tsa.Close;
  tsa.open;

  cpv.Close;
  cpv.tablename := 'cpv';
  cpv.open;

  cpa.Close;
  cpa.tablename := 'cpa';
  cpa.open;

  crea_cpv;
  crea_cpa;

  try
    tabella_esa_02.Close;
    tabella_esa_02.sql.clear;
    tabella_esa_02.sql.add('SELECT * from MG_ARTBASE');
    tabella_esa_02.sql.add('order by cod_art');
    tabella_esa_02.open;

  except
    messaggio(000, 'manca la tabella delle articoli (MG_ARTBASE)');
    Close;
    abort;
  end;

  query_02.sql.clear;
  query_02.sql.add('select * from cpv inner join cpa on cpv.tca_codice = cpa.taq_codice');
  query_02.sql.add('where cpv.gen_codice = :cpv_gen_codice and cpa.gen_codice = :cpa_gen_codice');

  tcc_codice := arc.dit.fieldbyname('tcc_codice_cli').asstring;
  tcf_codice := arc.dit.fieldbyname('tcf_codice_frn').asstring;
  tca_codice := 0;

  tabella_01_ds.DataSet := tabella_esa_02;
  while not tabella_esa_02.eof do
  begin

    if tabella_esa_02.fieldbyname('cod_art').asstring <> '' then
    begin
      tabella_01.append;
      if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
      begin
        progressivo := progressivo + 1;
        tabella_01.fieldbyname('codice').asstring := setta_lunghezza(inttostr(progressivo), 9, true, '0');
        tabella_01.fieldbyname('codice_alternativo').asstring := tabella_esa_02.fieldbyname('cod_art').asstring;
      end
      else
      begin
        tabella_01.fieldbyname('codice').asstring := tabella_esa_02.fieldbyname('cod_art').asstring;
      end;

      tabella_01.fieldbyname('tum_codice').asstring := arc.dit.fieldbyname('tum_codice_art').asstring;
      tabella_01.fieldbyname('tmr_codice').asstring := arc.dit.fieldbyname('tmr_codice_art').asstring;
      tabella_01.fieldbyname('tub_codice').asstring := arc.dit.fieldbyname('tub_codice_art').asstring;
      tabella_01.fieldbyname('tiv_codice_vendite').asstring := arc.dit.fieldbyname('tiv_codice_vendite_art').asstring;
      tabella_01.fieldbyname('tiv_codice_acquisti').asstring := arc.dit.fieldbyname('tiv_codice_acquisti_art').asstring;
      tabella_01.fieldbyname('tca_codice').asstring := arc.dit.fieldbyname('tca_codice_art').asstring;
      tabella_01.fieldbyname('taq_codice').asstring := arc.dit.fieldbyname('taq_codice_art').asstring;
      tabella_01.fieldbyname('tni_codice').asstring := arc.dit.fieldbyname('tni_codice_art').asstring;
      tabella_01.fieldbyname('tcm_codice').asstring := arc.dit.fieldbyname('tcm_codice_art').asstring;
      tabella_01.fieldbyname('tgm_codice').asstring := arc.dit.fieldbyname('tgm_codice_art').asstring;
      tabella_01.fieldbyname('tin_codice').asstring := arc.dit.fieldbyname('tin_codice_art').asstring;
      tabella_01.fieldbyname('ts3_codice').asstring := arc.dit.fieldbyname('ts3_codice_art').asstring;
      tabella_01.fieldbyname('tp2_codice').asstring := arc.dit.fieldbyname('tp2_codice_art').asstring;
      tabella_01.fieldbyname('tsa_codice').asstring := arc.dit.fieldbyname('tsa_codice_art').asstring;
      tabella_01.fieldbyname('taa_codice').asstring := arc.dit.fieldbyname('taa_codice_art').asstring;

      descrizione1 := '';
      descrizione2 := '';
      arc.spezza_descrizione(trim(tabella_esa_02.fieldbyname('des_articolo').asstring), descrizione1, descrizione2, 40);

      tabella_01.fieldbyname('descrizione1').asstring := descrizione1;
      tabella_01.fieldbyname('descrizione2').asstring := descrizione2;
      tabella_01.fieldbyname('tum_codice').asstring := tabella_esa_02.fieldbyname('cod_um').asstring;

      if tabella_esa_02.fieldbyname('cod_linea').asstring <> '' then
      begin
        tabella_01.fieldbyname('tmr_codice').asstring := tabella_esa_02.fieldbyname('cod_linea').asstring;
      end;

      gruppo := trim(tabella_esa_02.fieldbyname('cod_grumerc').asstring);
      sottogru := trim(tabella_esa_02.fieldbyname('cod_sgrumerc').asstring);
      codice_stat := trim(tabella_esa_02.fieldbyname('cod_gruppostat').asstring);

      if codice_stat <> '' then
      begin
        tabella_01.fieldbyname('tsa_codice').asstring := codice_stat;
      end;

      if (gruppo <> '') then
      begin
        tabella_01.fieldbyname('tcm_codice').asstring := gruppo;
      end;

      if (sottogru <> '') then
      begin
        tabella_02.Close;
        tabella_02.ParamByName('cod_grumerc').asstring := gruppo;
        tabella_02.ParamByName('cod_sgrumerc').asstring := sottogru;
        tabella_02.open;

        tabella_01.fieldbyname('tgm_codice').asstring := tabella_02.fieldbyname('codice').asstring;
      end;

      if (tabella_esa_02.fieldbyname('cod_fornitore').asstring <> '') then
      begin
        tabella_01.fieldbyname('frn_codice').asstring := setta_lunghezza(tabella_esa_02.fieldbyname('cod_fornitore').asstring, 8, true, '0');
      end;

      tabella_01.fieldbyname('tiv_codice_vendite').asstring := tabella_esa_02.fieldbyname('cod_iva_ven').asstring;
      tabella_01.fieldbyname('tiv_codice_acquisti').asstring := tabella_esa_02.fieldbyname('cod_iva_acq').asstring;

      if trim(tabella_esa_02.fieldbyname('cod_cpart_ric').asstring) <> '' then
      begin
        if cpv.locate('gen_codice', tabella_esa_02.fieldbyname('cod_cpart_ric').asstring, []) then
        begin
          tabella_01.fieldbyname('tca_codice').asstring := cpv.fieldbyname('tca_codice').asstring;
        end;
      end;

      if trim(tabella_esa_02.fieldbyname('cod_cpart_cos').asstring) <> '' then
      begin
        if cpa.locate('gen_codice', tabella_esa_02.fieldbyname('cod_cpart_cos').asstring, []) then
        begin
          tabella_01.fieldbyname('taq_codice').asstring := cpa.fieldbyname('taq_codice').asstring;
        end;
      end;

      tabella_01.post;
    end; // if

    tabella_esa_02.next;
  end; // while

  query.Close;
  query_02.Close;
  cpa.Close;
  cpv.Close;
  tabella_01.Close;
  tabella_esa_02.Close;
  tsa.Close;
end;

procedure TCONVEREADY.converti_bar;
begin
  v_tabella_01.caption := 'barcode articoli';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_esa_02;

  tabella_02.Close;
  tabella_02.tablename := 'art';
  tabella_02.open;

  cancella_tabella('bar');
  tabella_01.Close;
  tabella_01.tablename := 'bar';
  tabella_01.open;

  try
    tabella_esa_02.Close;
    tabella_esa_02.sql.clear;
    tabella_esa_02.sql.add('SELECT * FROM MG_ARTICOLIALIAS');
    tabella_esa_02.sql.add('order by cod_art, prg_art, prg_alias');
    tabella_esa_02.open;

  except
    messaggio(000, 'manca la tabella barcode articoli (MG_ARTICOLIALIAS)');
    Close;
    abort;
  end;

  while not tabella_esa_02.eof do
  begin

    if not tabella_01.locate('art_codice;codice_barre', vararrayof([trim(tabella_esa_02.fieldbyname('cod_art').asstring), trim(tabella_esa_02.fieldbyname('cod_agg').asstring)]), []) then
    begin
      tabella_01.append;

      if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
      begin
        tabella_02.locate('codice_alternativo', trim(tabella_esa_02.fieldbyname('codice_articolo').asstring), []);
        tabella_01.fieldbyname('art_codice').asstring := tabella_02.fieldbyname('codice').asstring;
      end
      else
      begin
        tabella_01.fieldbyname('art_codice').asstring := trim(tabella_esa_02.fieldbyname('cod_art').asstring);
      end;

      tabella_01.fieldbyname('codice_barre').asstring := trim(tabella_esa_02.fieldbyname('cod_agg').asstring);

      tabella_01.post;
    end;

    tabella_esa_02.next;
  end;

  tabella_01.Close;
  tabella_esa_02.Close;
end;

procedure TCONVEREADY.converti_tmo;
begin
  v_tabella_01.caption := 'causali magazzino';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_02;

  cancella_tabella('tmo');
  tabella_01.Close;
  tabella_01.tablename := 'tmo';
  tabella_01.open;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT * from CA_CAUMAG');
  tabella_esa_02.sql.add('order by cod_caumag');
  tabella_esa_02.open;

  tabella_01_ds.DataSet := tabella_esa_02;
  while not tabella_esa_02.eof do
  begin
    if tabella_esa_02.fieldbyname('cod_caumag').asstring <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_02.fieldbyname('cod_caumag').asstring);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_02.fieldbyname('des_caumag').asstring);
      if tabella_01.fieldbyname('descrizione').asstring = '' then
      begin
        tabella_01.fieldbyname('descrizione').asstring := '.';
      end;

      tabella_01.fieldbyname('tipo_movimento').asstring := 'normale';
      if tabella_esa_02.fieldbyname('ind_qta_giaciniz').asstring = '+' then
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'apertura inventario';
      end;

      if tabella_esa_02.fieldbyname('ind_qta_esistenza').asstring = '+' then
      begin
        tabella_01.fieldbyname('esistenza').asstring := 'incrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_qta_esistenza').asstring = '-' then
      begin
        tabella_01.fieldbyname('esistenza').asstring := 'decrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_qta_car_rett').asstring = '+' then
      begin
        tabella_01.fieldbyname('esistenza').asstring := 'incrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_qta_car_rett').asstring = '-' then
      begin
        tabella_01.fieldbyname('esistenza').asstring := 'decrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_val_giaciniz').asstring = '+' then
      begin
        tabella_01.fieldbyname('valorizzazione').asstring := 'incrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_val_carichifor').asstring = '+' then
      begin
        tabella_01.fieldbyname('valorizzazione').asstring := 'incrementa';
      end;

      if tabella_esa_02.fieldbyname('ind_val_carichifor').asstring = '-' then
      begin
        tabella_01.fieldbyname('valorizzazione').asstring := 'decrementa';
      end;

      if (tabella_esa_02.fieldbyname('cod_caumag_r').asstring <> '000') and (tabella_esa_02.fieldbyname('cod_caumag_r').asstring <> '') then
        tabella_01.fieldbyname('tmo_codice_collegato').asstring := tabella_esa_02.fieldbyname('cod_caumag_r').asstring;

      tabella_01.post;
    end;

    tabella_esa_02.next;
  end;

  tabella_01.Close;
  tabella_esa_02.Close;
end;

procedure TCONVEREADY.converti_tpa;
var
  i, riga: integer;
  nome_campo: string;

begin
  v_tabella_01.caption := 'codici pagamento';
  application.processmessages;

  tabella_01_ds.DataSet := tabella_02;

  cancella_tabella('tpa');
  tabella_01.Close;
  tabella_01.tablename := 'tpa';
  tabella_01.open;

  try
    tabella_esa_01.Close;
    tabella_esa_01.sql.clear;
    tabella_esa_01.sql.add('SELECT * from CA_PAGAMT');
    tabella_esa_01.sql.add('order by cod_pag');
    tabella_esa_01.open;
  except
    messaggio(000, 'manca la tabella patamenti testata (CA_PAGMT)');
    Close;
    abort;
  end;

  try
    tabella_esa_02.Close;
    tabella_esa_02.sql.clear;
    tabella_esa_02.sql.add('SELECT * from CA_PAGAMR');
    tabella_esa_02.sql.add('where');
    tabella_esa_02.sql.add('cod_pag=:cod_pag');
    tabella_esa_02.sql.add('order by 1,2,3');
    tabella_esa_02.open;
  except
    messaggio(000, 'manca la tabella causali di magazzino (CA_PAGAMR)');
    Close;
    abort;
  end;

  while not tabella_esa_01.eof do
  begin
    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_pag').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_pag').asstring);
    if tabella_01.fieldbyname('descrizione').asstring = '' then
    begin
      tabella_01.fieldbyname('descrizione').asstring := '.';
    end;

    tabella_esa_02.Close;
    tabella_esa_02.Parameters.ParamByName('cod_pag').Value := trim(tabella_esa_01.fieldbyname('cod_pag').asstring);
    tabella_esa_02.open;

    tabella_01.fieldbyname('numero_rate').asInteger := tabella_esa_02.recordcount;
    tabella_01.fieldbyname('tipo_rate').asstring := 'costanti';
    tabella_01.fieldbyname('giorni_prima_rata_fisse').asInteger := 0;
    tabella_01.fieldbyname('giorni_rate_fisse').asInteger := 0;

    nome_campo := 'tts_codice_fisse';

    if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RD') then
      tabella_01.fieldbyname(nome_campo).asstring := 'rimessa diretta'
    else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RB') then
      tabella_01.fieldbyname(nome_campo).asstring := 'R.I.B.A.'
    else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'BO') then
      tabella_01.fieldbyname(nome_campo).asstring := 'bonifico bancario'
    else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RID') then
      tabella_01.fieldbyname(nome_campo).asstring := 'R.I.D.';

    riga := 0;
    while not tabella_esa_02.eof do
    begin
      riga := riga + 1;

      if (riga <= tabella_01.fieldbyname('numero_rate').asInteger) and (riga < 9) then
      begin
        nome_campo := 'tts_codice_variabili_0' + inttostr(riga);

        if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RD') then
          tabella_01.fieldbyname(nome_campo).asstring := 'rimessa diretta'
        else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RB') then
          tabella_01.fieldbyname(nome_campo).asstring := 'R.I.B.A.'
        else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'BO') then
          tabella_01.fieldbyname(nome_campo).asstring := 'bonifico bancario'
        else if (trim(tabella_esa_02.fieldbyname('cod_tipo_pag').asstring) = 'RID') then
          tabella_01.fieldbyname(nome_campo).asstring := 'R.I.D.';

        // fine mese
        nome_campo := 'fine_mese_variabili_0' + inttostr(riga);

        if (trim(tabella_esa_02.fieldbyname('ind_data_scad').asstring)) = '4' then
          tabella_01.fieldbyname(nome_campo).asstring := 'si';

        // perc da pag
        nome_campo := 'percentuale_variabili_0' + inttostr(riga);
        tabella_01.fieldbyname(nome_campo).asFloat := tabella_esa_02.fieldbyname('prc_totpag_rata').asFloat;

        // rata
        nome_campo := 'giorni_variabili_0' + inttostr(riga);
        tabella_01.fieldbyname(nome_campo).asFloat := tabella_esa_02.fieldbyname('num_giorni').asFloat;

      end; // if

      tabella_esa_02.next;
    end; // while

    riga := riga + 1;
    if riga < 8 then
    begin

      for i := riga to 8 do
      begin
        nome_campo := 'tts_codice_variabili_0' + inttostr(i);
        tabella_01.fieldbyname(nome_campo).asstring := ''
      end; // for
    end;

    tabella_01.post;
    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;
end;

procedure TCONVEREADY.converti_sedi_amministrative;
var
  campo_codice: string;
begin
  v_tabella_01.caption := 'indirizzi spedizione';
  application.processmessages;

  cancella_tabella('ind');
  tabella_01.Close;
  tabella_01.tablename := 'ind';
  tabella_01.open;

  cancella_tabella('inf');
  tabella_02.Close;
  tabella_02.tablename := 'inf';
  tabella_02.open;

  try
    tabella_esa_01.Close;
    tabella_esa_01.sql.clear;
    tabella_esa_01.sql.add('SELECT * from CA_SEDI');
    tabella_esa_01.sql.add('WHERE');
    tabella_esa_01.sql.add('ind_sede =' + quotedstr('D'));
    tabella_esa_01.sql.add('order by 1,3');
    tabella_esa_01.open;
  except
    raise;
    messaggio(000, 'manca la tabella INDIRIZZI DI SPEDIZIONE (SEDEAMMINISTRATIVA)');
    Close;
    abort;
  end;

  tabella_01_ds.DataSet := tabella_esa_01;

  tabella_clienti_forn.Close;
  tabella_clienti_forn.prepared := true;

  while not tabella_esa_01.eof do
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('cod_anagra_sede').asstring, cli_for);

    tabella_clienti_forn.Close;
    tabella_clienti_forn.Parameters.ParamByName('codice_nom').Value := tabella_esa_01.fieldbyname('cod_anagra_sede').asstring;
    tabella_clienti_forn.open;

    if tabella_clienti_forn.fieldbyname('ind_clifor').asstring = 'C' then
    begin
      if (arc.dit.fieldbyname('codice_nom_numerico').asstring = 'si') and (v_codifica_clienti.checked) then
      begin
        campo_codice := 'codice';
      end
      else if (arc.dit.fieldbyname('codice_nom_numerico').asstring = 'si') and not(v_codifica_clienti.checked) then
      begin
        campo_codice := 'codice_alternativo';
        read_tabella(arc.arcdit, 'cli', campo_codice, tabella_esa_01.fieldbyname('cod_anagra_sede').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;
      end
      else
      begin
        campo_codice := 'codice';
      end;

      if read_tabella(arc.arcdit, 'cli', 'codice', cli_for) then
      begin
        tabella_01.append;

        tabella_01.fieldbyname('cli_codice').asstring := cli_for;
        tabella_01.fieldbyname('indirizzo').asstring := trim(tabella_esa_01.fieldbyname('cdn_sede').asstring);
        tabella_01.fieldbyname('descrizione1').asstring := trim(tabella_esa_01.fieldbyname('des_sede').asstring);
        if tabella_01.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_01.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
          tabella_01.fieldbyname('descrizione2').asstring := archivio.fieldbyname('descrizione2').asstring;
        end;
        tabella_01.fieldbyname('via').asstring := trim(tabella_esa_01.fieldbyname('des_indir').asstring);
        tabella_01.fieldbyname('cap').asstring := trim(tabella_esa_01.fieldbyname('cod_cap').asstring);
        tabella_01.fieldbyname('citta').asstring := trim(tabella_esa_01.fieldbyname('des_localita').asstring);
        tabella_01.fieldbyname('provincia').asstring := trim(tabella_esa_01.fieldbyname('sig_prov').asstring);
        tabella_01.fieldbyname('telefono').asstring := trim(tabella_esa_01.fieldbyname('sig_tel').asstring);

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
    else if tabella_clienti_forn.fieldbyname('ind_clifor').asstring = 'F' then
    begin

      assegna_codice_frn(tabella_esa_01.fieldbyname('cod_anagra_sede').asstring, cli_for);
      if read_tabella(arc.arcdit, 'frn', 'codice', cli_for) then
      begin
        tabella_02.append;

        tabella_02.fieldbyname('frn_codice').asstring := cli_for;
        tabella_02.fieldbyname('indirizzo').asstring := trim(tabella_esa_01.fieldbyname('cdn_sede').asstring);
        tabella_02.fieldbyname('descrizione1').asstring := trim(tabella_esa_01.fieldbyname('des_sede').asstring);
        if tabella_02.fieldbyname('descrizione1').asstring = '' then
        begin
          tabella_02.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
          tabella_02.fieldbyname('descrizione2').asstring := archivio.fieldbyname('descrizione2').asstring;
        end;
        tabella_02.fieldbyname('via').asstring := trim(tabella_esa_01.fieldbyname('des_indir').asstring);
        tabella_02.fieldbyname('cap').asstring := trim(tabella_esa_01.fieldbyname('cod_cap').asstring);
        tabella_02.fieldbyname('citta').asstring := trim(tabella_esa_01.fieldbyname('des_localita').asstring);
        tabella_02.fieldbyname('provincia').asstring := trim(tabella_esa_01.fieldbyname('sig_prov').asstring);
        tabella_02.fieldbyname('telefono').asstring := trim(tabella_esa_01.fieldbyname('sig_tel').asstring);

        tabella_02.fieldbyname('tna_codice').asstring := archivio.fieldbyname('tna_codice').asstring;
        tabella_02.fieldbyname('tsp_codice').asstring := archivio.fieldbyname('tsp_codice').asstring;
        tabella_02.fieldbyname('tpo_codice').asstring := archivio.fieldbyname('tpo_codice').asstring;

        tabella_02.post;
      end;
    end;
    tabella_esa_01.next;
  end;

  tabella_clienti_forn.Close;
  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;

end;

procedure TCONVEREADY.converti_listino(nome_tabella, tipo_listino: string);
var
  articolo, listino, campo_listino: string;
  prg_listino: integer;
begin
  v_tabella_01.caption := 'listini ' + nome_tabella;
  application.processmessages;

  if tipo_listino = 'C' then
  begin
    campo_listino := 'tlv_codice';
  end
  else
  begin
    campo_listino := 'tla_codice';
  end;

  query_02.Close;
  query_02.sql.clear;
  query_02.sql.add('select * from tsm');
  query_02.sql.add('where');
  query_02.sql.add('sconto_maggiorazione=''sconto'' and ');
  query_02.sql.add('percentuale_01=:sconto_01 and');
  query_02.sql.add('percentuale_02=:sconto_02 ');

  cancella_tabella(nome_tabella);

  tsm.Close;
  tsm.open;

  tabella_02.Close;
  tabella_02.tablename := 'art';
  tabella_02.open;

  tabella_01.Close;
  tabella_01.tablename := nome_tabella;
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_esa_01;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT distinct cod_listino FROM CO_CLIFOR');
  tabella_esa_02.sql.add('where');
  tabella_esa_02.sql.add('ind_clifor=:tipo_listino');
  tabella_esa_02.sql.add('order by 1 ');
  tabella_esa_02.Parameters.ParamByName('tipo_listino').Value := tipo_listino;
  tabella_esa_02.open;

  while not tabella_esa_02.eof do
  begin

    tabella_esa_01.Close;
    tabella_esa_01.sql.clear;
    tabella_esa_01.sql.add('SELECT * FROM PZ_PREZZIBASE');
    tabella_esa_01.sql.add('where');
    tabella_esa_01.sql.add('cod_listino=:cod_listino and ');
    tabella_esa_01.sql.add('cod_dep=:cod_dep and ');
    tabella_esa_01.sql.add('dat_obsoleto is null ');
    tabella_esa_01.sql.add('order by cod_listino, prg_listino, cod_art ');
    tabella_esa_01.Parameters.ParamByName('cod_listino').Value := tabella_esa_02.fieldbyname('cod_listino').asstring;
    tabella_esa_01.Parameters.ParamByName('cod_dep').Value := v_tma_codice.text;
    tabella_esa_01.open;

    while not tabella_esa_01.eof do
    begin

      articolo := trim(tabella_esa_01.fieldbyname('cod_art').asstring);
      prg_listino := tabella_esa_01.fieldbyname('prg_listino').asInteger;
      listino := trim(tabella_esa_01.fieldbyname('cod_listino').asstring);

      while not(tabella_esa_01.eof) and (articolo = trim(tabella_esa_01.fieldbyname('cod_art').asstring)) and (listino = trim(tabella_esa_01.fieldbyname('cod_listino').asstring)) and (prg_listino = tabella_esa_01.fieldbyname('prg_listino').asInteger) do
      begin

        if read_tabella(arc.arcdit, 'art', 'codice', tabella_esa_01.fieldbyname('cod_art').asstring) then
        begin
          if not tabella_01.locate('art_codice;' + campo_listino, vararrayof([trim(tabella_esa_01.fieldbyname('cod_art').asstring), trim(tabella_esa_01.fieldbyname('cod_listino').asstring)]), []) then
          begin
            tabella_01.append;

            if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
            begin
              tabella_02.locate('codice_alternativo', tabella_esa_01.fieldbyname('cod_art').asstring, []);
              tabella_01.fieldbyname('art_codice').asstring := tabella_02.fieldbyname('codice').asstring;
            end
            else
            begin
              tabella_01.fieldbyname('art_codice').asstring := tabella_esa_01.fieldbyname('cod_art').asstring;
            end;

            tabella_01.fieldbyname(campo_listino).asstring := tabella_esa_01.fieldbyname('cod_listino').asstring;
            tabella_01.fieldbyname('data_inizio').asdatetime := StrToDate('01/01/' + esercizio);
            tabella_01.fieldbyname('data_fine').asdatetime := StrToDate('31/12/2999');
            tabella_01.fieldbyname('prezzo').asFloat := tabella_esa_01.fieldbyname('prz_listino').asFloat;

            if not((tabella_esa_01.fieldbyname('prc_Sconto1').asFloat = 0) or not(tabella_esa_01.fieldbyname('prc_Sconto2').asFloat = 0)) then
            begin
              query_02.Close;
              query_02.ParamByName('sconto_01').asFloat := tabella_esa_01.fieldbyname('prc_sconto1').asFloat;
              query_02.ParamByName('sconto_02').asFloat := tabella_esa_01.fieldbyname('prc_sconto2').asFloat;
              query_02.open;
              if not query_02.eof then
              begin
                tabella_01.fieldbyname('tsm_codice').asstring := query_02.fieldbyname('codice').asstring;
              end
              else
              begin
                crea_tsm(tabella_esa_01.fieldbyname('prc_sconto1').asFloat, tabella_esa_01.fieldbyname('prc_sconto2').asFloat);
                tabella_01.fieldbyname('tsm_codice').asstring := setta_lunghezza(tsm_codice, 4, 0);
              end;
            end;

            tabella_01.post;
          end; // if

        end;

        tabella_esa_01.next;
      end; // while

    end; // while

    tabella_esa_02.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;
  tsm.Close;

end;

procedure TCONVEREADY.converti_pnt;
var
  i: integer;
  progressivo: integer;
begin
  v_tabella_01.caption := 'primanota';
  application.processmessages;

  // arc.attivazione_trigger(arc.arcdit, false, false);

  try
    query.sql.clear;
    query.sql.add('delete from cfgese ');
    query.ExecSQL;
  except
  end;

  try
    query.sql.clear;
    query.sql.add('delete from cfg ');
    query.ExecSQL;
  except
  end;

  cancella_tabella('pni');
  tabella_03.Close;
  tabella_03.tablename := 'pni';
  tabella_03.open;

  cancella_tabella('pnr');
  tabella_02.Close;
  tabella_02.tablename := 'pnr';
  tabella_02.open;

  cancella_tabella('pnt');
  tabella_01.Close;
  tabella_01.tablename := 'pnt';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_esa_02;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT * from CG_PRINOTAR_CONT');
  tabella_esa_02.sql.add('order by prg_prinota, prg_prinota_riga');
  tabella_esa_02.open;

  tabella_esa_03.Close;
  tabella_esa_03.sql.clear;
  tabella_esa_03.sql.add('SELECT * from CG_PRINOTAR_IVA');
  tabella_esa_03.sql.add('where');
  tabella_esa_03.sql.add('prg_prinota=:prg_prinota');
  tabella_esa_03.sql.add('order by prg_prinota_riga');

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * from CG_PRINOTAT');
  tabella_esa_01.sql.add('where');
  tabella_esa_01.sql.add('prg_prinota=:prg_prinota');

  test_datas := '';
  test_numero_stringa := '';
  progressivo := 0;

  while not tabella_esa_02.eof do
  begin
    application.processmessages;

    tabella_esa_01.Close;
    tabella_esa_01.Parameters.ParamByName('prg_prinota').Value := tabella_esa_02.fieldbyname('prg_prinota').asInteger;
    tabella_esa_01.open;

    if tabella_esa_02.fieldbyname('prg_prinota').asInteger <> progressivo then
    begin
      codice_clifor := '';
      progressivo := tabella_esa_02.fieldbyname('prg_prinota').asInteger;
      crea_pnt(progressivo);

      test_datas := tabella_esa_01.fieldbyname('dat_registrazione').asstring;
      test_numero_stringa := tabella_esa_01.fieldbyname('num_registrazione').asstring;
    end;

    crea_pnr(progressivo);

    if (tabella_01.fieldbyname('cfg_tipo').asstring = 'C') or (tabella_01.fieldbyname('cfg_tipo').asstring = 'F') and (tabella_01.fieldbyname('cfg_codice').asstring = '') then
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

    tabella_esa_02.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_03.Close;
  tabella_esa_01.Close;

  // arc.attivazione_trigger(arc.arcdit, false, TRUE);
end;

procedure TCONVEREADY.converti_pni;
var
  i: integer;
  progressivo: integer;
begin
  v_tabella_01.caption := 'pri solo iva';
  application.processmessages;

  // arc.attivazione_trigger(arc.arcdit, false, false);
  tabella_03.Close;
  tabella_03.tablename := 'pni';
  tabella_03.open;

  tabella_01_ds.DataSet := tabella_esa_01;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;

  tabella_esa_03.Close;
  tabella_esa_03.sql.clear;
  tabella_esa_03.sql.add('SELECT * from CG_PRINOTAR_IVA');
  tabella_esa_03.sql.add('where');
  tabella_esa_03.sql.add('prg_prinota=:prg_prinota');
  tabella_esa_03.sql.add('order by prg_prinota_riga');

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * from CG_PRINOTAT');
  tabella_esa_01.sql.add('order by prg_prinota');
  tabella_esa_01.open;

  test_datas := '';
  test_numero_stringa := '';
  progressivo := 0;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    crea_pni(tabella_esa_01.fieldbyname('prg_prinota').asInteger);

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_03.Close;
  tabella_esa_01.Close;

  // arc.attivazione_trigger(arc.arcdit, false, TRUE);
end;

procedure TCONVEREADY.crea_pnt(progressivo: integer);
var
  cod_caucon: string;
  numero_documento_alfa: string;
begin
  cod_caucon := tabella_esa_01.fieldbyname('cod_caucon').asstring;

  if (tabella_esa_01.fieldbyname('cod_attivita').asstring = '01') and (tabella_esa_01.fieldbyname('num_reg_iva').asstring = '2') then
  begin
    cod_caucon := cod_caucon + tabella_esa_01.fieldbyname('num_reg_iva').asstring;
  end;

  if (tabella_esa_01.fieldbyname('cod_attivita').asstring = '02') and (tabella_esa_01.fieldbyname('sig_serie_doc').asstring = '') then
  begin
    cod_caucon := cod_caucon + tabella_esa_01.fieldbyname('cod_attivita').asstring;
  end;

  if (tabella_esa_01.fieldbyname('cod_attivita').asstring = '02') and (tabella_esa_01.fieldbyname('sig_serie_doc').asstring = 'BIS') then
  begin
    cod_caucon := cod_caucon + '2B';
  end;

  tabella_01.append;
  tabella_01.fieldbyname('data_registrazione').asdatetime := tabella_esa_01.fieldbyname('dat_registrazione').asdatetime;
  tabella_01.fieldbyname('progressivo').asInteger := tabella_esa_01.fieldbyname('prg_prinota').asInteger;
  tabella_01.fieldbyname('tco_codice').asstring := tabella_esa_01.fieldbyname('cod_caucon').asstring;

  numero_documento_alfa := tabella_esa_01.fieldbyname('des_num_doc').asstring;
  tabella_01.fieldbyname('numero_documento').asFloat := arc.numero_documento_alfa(nil, 'numero_documento', numero_documento_alfa);
  tabella_01.fieldbyname('serie_documento').asstring := arc.serie_documento_alfa(nil, 'serie_documento', numero_documento_alfa);

// ------------------------------
  read_tabella(arc.arcdit, 'tco', 'codice', tabella_01.fieldbyname('tco_codice').asstring);

  tabella_01.fieldbyname('tco_descrizione').asstring := archivio.fieldbyname('descrizione').asstring;
  tabella_01.fieldbyname('documento_iva').asstring := archivio.fieldbyname('movimento_iva').asstring;
  tabella_01.fieldbyname('tipo_documento_iva').asstring := archivio.fieldbyname('tipo_registro_iva').asstring;
  tabella_01.fieldbyname('descrizione').asstring := tabella_esa_01.fieldbyname('num_registrazione').asstring;
  // ------------------------------

  try
    tabella_01.fieldbyname('numero_documento').asInteger := tabella_esa_01.fieldbyname('des_num_doc').asInteger;
  except
    tabella_01.fieldbyname('descrizione').asstring := tabella_01.fieldbyname('descrizione').asstring + ' - nr doc:' + tabella_esa_01.fieldbyname('des_num_doc').asstring;
  end;

  if tabella_esa_01.fieldbyname('dat_doc').asdatetime <> 0 then
  begin
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
  end
  else
  begin
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_01.fieldbyname('data_registrazione').asdatetime;
  end;

  if tabella_esa_01.fieldbyname('dat_comp_iva').asdatetime <> 0 then
  begin
    tabella_01.fieldbyname('data_competenza_iva').asdatetime := tabella_esa_01.fieldbyname('dat_comp_iva').asdatetime;
  end
  else
  begin
    tabella_01.fieldbyname('data_competenza_iva').asdatetime := tabella_01.fieldbyname('data_registrazione').asdatetime;
  end;

  tabella_01.fieldbyname('protocollo').asInteger := tabella_esa_01.fieldbyname('num_protocollo').asInteger;
  if tabella_01.fieldbyname('documento_iva').asstring = 'si' then
  begin
    if archivio.fieldbyname('tipo_registro_iva').asstring = 'vendite' then
    begin
      tabella_01.fieldbyname('cfg_tipo').asstring := 'C';
      if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;
      end
      else
      begin
        assegna_codice_cli(tabella_esa_01.fieldbyname('cod_clifor').asstring, cli_for);
      end;

      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
      if tabella_esa_01.fieldbyname('sig_serie_doc').asstring <> '' then
      begin
        tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('sig_serie_doc').asstring;
      end;

      (*
        if tabella_esa_01.fieldbyname('num_reg_iva').asstring <> '' then
        begin
        tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('num_reg_iva').asstring;
        end;
 *)

    end
    else if archivio.fieldbyname('tipo_registro_iva').asstring = 'acquisti' then
    begin
      tabella_01.fieldbyname('cfg_tipo').asstring := 'F';
      if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'frn', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;
      end
      else
      begin
        assegna_codice_frn(tabella_esa_01.fieldbyname('cod_clifor').asstring, cli_for);
      end;

      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
      if tabella_esa_01.fieldbyname('sig_serie_doc').asstring <> '' then
      begin
        tabella_01.fieldbyname('serie_documento').asstring := trim(tabella_esa_01.fieldbyname('sig_serie_doc').asstring);
      end;

      (*
        if tabella_esa_01.fieldbyname('num_reg_iva').asstring <> '' then
        begin
        tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('num_reg_iva').asstring;

        if tabella_esa_01.fieldbyname('num_reg_iva').asstring > '1' then
        begin
        tabella_01.fieldbyname('tco_codice').asstring := tabella_01.fieldbyname('tco_codice').asstring + tabella_esa_01.fieldbyname('num_reg_iva').asstring;
        end;

        end;
 *)
    end
    else if archivio.fieldbyname('tipo_registro_iva').asstring = 'corrispettivi' then
    begin
      // tabella_01.fieldbyname('cfg_tipo').asstring := 'G';
      // tabella_01.fieldbyname('cfg_codice').asstring := tabella_esa_01.fieldbyname('Conto_Prima_Nota').asstring;
    end;
  end
  else
  begin

  end;

  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
  tabella_01.fieldbyname('ese_codice').asstring := tabella_esa_01.fieldbyname('num_esercizio').asstring;
  tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_valuta').asFloat;

  tabella_01.fieldbyname('descrizione').asstring := tabella_01.fieldbyname('descrizione').asstring + ' cod.attivita ' + tabella_esa_01.fieldbyname('cod_attivita').asstring + ' serie doc' + tabella_esa_01.fieldbyname('sig_serie_doc').asstring;
  try

    tabella_01.post;

    if (tabella_esa_01.fieldbyname('cod_reg_iva').asstring <> '') then
    begin
      crea_pni(progressivo);
    end;

  except
  end;

  riga := 0;
  riga_iva := 0;

end;

procedure TCONVEREADY.crea_pnr(progressivo: integer);
begin
  tabella_02.append;

  tabella_02.fieldbyname('progressivo').asFloat := tabella_esa_01.fieldbyname('prg_prinota').asInteger;
  riga := riga + 1;

  tabella_02.fieldbyname('riga').asInteger := tabella_esa_02.fieldbyname('num_riga').asInteger;

  // ----------------------------------------------------------
  // il cfg_ese viene inserito nel trigger
  // ----------------------------------------------------------

  if tabella_esa_02.fieldbyname('ind_clifor').asstring = 'C' then
  begin
    if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
    begin
      read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_02.fieldbyname('cod_clifor').asstring);
      cli_for := archivio.fieldbyname('codice').asstring;
    end
    else
    begin
      assegna_codice_cli(tabella_esa_02.fieldbyname('cod_clifor').asstring, cli_for);
    end;

    tabella_02.fieldbyname('cfg_tipo').asstring := 'C';
    tabella_02.fieldbyname('cfg_codice').asstring := cli_for;
    tabella_02.fieldbyname('partite').asstring := 'A';

    if codice_clifor = '' then
    begin
      codice_clifor := cli_for;
    end;

  end
  else if tabella_esa_02.fieldbyname('ind_clifor').asstring = 'F' then
  begin
    if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
    begin
      read_tabella(arc.arcdit, 'frn', 'codice_alternativo', tabella_esa_02.fieldbyname('cod_clifor').asstring);
      cli_for := archivio.fieldbyname('codice').asstring;
    end
    else
    begin
      assegna_codice_frn(tabella_esa_02.fieldbyname('cod_clifor').asstring, cli_for);
    end;

    tabella_02.fieldbyname('cfg_tipo').asstring := 'F';
    tabella_02.fieldbyname('cfg_codice').asstring := cli_for;
    tabella_02.fieldbyname('partite').asstring := 'A';

    if codice_clifor = '' then
    begin
      codice_clifor := cli_for;
    end;
  end
  else
  begin
    tabella_02.fieldbyname('cfg_tipo').asstring := 'G';
    tabella_02.fieldbyname('cfg_codice').asstring := tabella_esa_02.fieldbyname('cod_piacont').asstring;
  end;

  if tabella_esa_02.fieldbyname('ind_da_av').asstring = 'D' then
  begin
    tabella_02.fieldbyname('importo_dare').asFloat := tabella_esa_02.fieldbyname('val_importo_val').asFloat;
    tabella_02.fieldbyname('importo_dare_euro').asFloat := tabella_esa_02.fieldbyname('val_importo').asFloat;
    tabella_02.fieldbyname('importo_avere').asFloat := 0;
    tabella_02.fieldbyname('importo_avere_euro').asFloat := 0;
  end
  else
  begin
    tabella_02.fieldbyname('importo_dare').asFloat := 0;
    tabella_02.fieldbyname('importo_dare_euro').asFloat := 0;
    tabella_02.fieldbyname('importo_avere').asFloat := tabella_esa_02.fieldbyname('val_importo_val').asFloat;
    tabella_02.fieldbyname('importo_avere_euro').asFloat := tabella_esa_02.fieldbyname('val_importo').asFloat;
  end;

  tabella_02.fieldbyname('descrizione').asstring := trim(tabella_esa_02.fieldbyname('des_aggiuntiva').asstring);

  tabella_02.post;

end;

procedure TCONVEREADY.crea_pni;
begin
  tabella_esa_03.Close;
  tabella_esa_03.Parameters.ParamByName('prg_prinota').Value := tabella_esa_01.fieldbyname('prg_prinota').asInteger;
  tabella_esa_03.open;

  while not tabella_esa_03.eof do
  begin

    if not tabella_03.locate('progressivo;riga', vararrayof([tabella_esa_03.fieldbyname('prg_prinota').asInteger, tabella_esa_03.fieldbyname('prg_prinota_riga').asInteger]), []) then
    begin
      tabella_03.append;

      tabella_03.fieldbyname('progressivo').asFloat := tabella_esa_03.fieldbyname('prg_prinota').asInteger;
      tabella_03.fieldbyname('riga').asInteger := tabella_esa_03.fieldbyname('prg_prinota_riga').asInteger;

      tabella_03.fieldbyname('tiv_codice').asstring := tabella_esa_03.fieldbyname('cod_iva').asstring;
      tabella_03.fieldbyname('importo_imponibile').asFloat := tabella_esa_03.fieldbyname('val_imponibile_val').asFloat;
      tabella_03.fieldbyname('importo_imponibile_euro').asFloat := tabella_esa_03.fieldbyname('val_imponibile').asFloat;

      tabella_03.fieldbyname('importo_iva').asFloat := tabella_esa_03.fieldbyname('val_iva_val').asFloat;
      tabella_03.fieldbyname('importo_iva_euro').asFloat := tabella_esa_03.fieldbyname('val_iva').asFloat;

      read_tabella(arc.arcdit, 'tiv', 'codice', tabella_03.fieldbyname('tiv_codice').asstring);
      tabella_03.fieldbyname('importo_indetraibile').asFloat := arrotonda(tabella_03.fieldbyname('importo_iva').asFloat * archivio.fieldbyname('indetraibile').asFloat / 100);

      tabella_03.fieldbyname('importo_indetraibile_euro').asFloat := arrotonda(tabella_03.fieldbyname('importo_iva').asFloat * archivio.fieldbyname('indetraibile').asFloat / 100);

      try
        tabella_03.post;
      except
      end;

    end; // if

    tabella_esa_03.next;
  end; // while

end;

procedure TCONVEREADY.converti_par;
var
  tipo_cliente, codice_clifor: string;
  data_doc: TDateTime;
  numero_doc: integer;
  serie_doc: string;
  descrizione: string;
  nr_col1, nr_col2, progressivo, nr_riga: integer;

  totale_pagare, totale_pagare_euro: double;
begin
  v_conferma.enabled := false;

  v_tabella_01.caption := 'scadenze';
  application.processmessages;

  cancella_tabella('pat');
  tabella_01.Close;
  tabella_01.tablename := 'pat';
  tabella_01.open;

  cancella_tabella('pas');
  tabella_02.Close;
  tabella_02.tablename := 'pas';
  tabella_02.open;

  tpa.Close;
  tpa.open;

  tabella_01_ds.DataSet := tabella_esa_01;

  tabella_esa_03.Close;
  tabella_esa_03.sql.clear;
  tabella_esa_03.sql.add('SELECT * FROM CO_PAGAMENTI T');
  tabella_esa_03.sql.add('where');
  tabella_esa_03.sql.add('T.prg_scadenza=:prg_scadenza');
  tabella_esa_03.sql.add('order by T.prg_pagamento');

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT T.*, C.cod_pag from CG_PRINOTAT T');
  tabella_esa_02.sql.add('INNER JOIN CO_CLIFOR C ON C.ind_clifor=T.ind_clifor and C.cod_clifor=T.cod_clifor');
  tabella_esa_02.sql.add('where');
  tabella_esa_02.sql.add('T.ind_clifor=:ind_clifor and ');
  tabella_esa_02.sql.add('T.cod_clifor=:cod_clifor and ');
  tabella_esa_02.sql.add('T.dat_doc=:dat_doc  and ');
  tabella_esa_02.sql.add('T.des_num_doc=:des_num_doc ');

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT');
  tabella_esa_01.sql.add('CO.cod_caucon, CO.num_reg_iva, CP.prg_prinota, ');
  tabella_esa_01.sql.add('T.dat_doc,');
  tabella_esa_01.sql.add('T.des_num_doc,');
  tabella_esa_01.sql.add('T.sig_serie_doc, T.cod_valuta, T.cmb_valuta, C.cod_pag, ');
  tabella_esa_01.sql.add('p.*, s.*');
  tabella_esa_01.sql.add('from CO_SCADENZE s');
  tabella_esa_01.sql.add('inner join CO_PARTITE p ON p.prg_partita=s.prg_partita');
  tabella_esa_01.sql.add('inner join CO_DOCUMENTI CO on CO.prg_documento=s.prg_documento');
  tabella_esa_01.sql.add('inner join CO_DOCUMENTI_PN CP on CP.prg_documento=s.prg_documento');
  tabella_esa_01.sql.add('left join CG_PRINOTAT T on T.prg_prinota=CP.prg_prinota');
  tabella_esa_01.sql.add('inner join CO_CLIFOR C ON C.ind_clifor=p.ind_clifor and C.cod_clifor=p.cod_clifor');
  tabella_esa_01.sql.add('where ');
  tabella_esa_01.sql.add('CO.num_reg_iva is not null');
  tabella_esa_01.sql.add('order by s.prg_partita, s.num_riga');
  tabella_esa_01.open;

  progressivo := 0;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tipo_cliente := tabella_esa_01.fieldbyname('ind_clifor').asstring;
    codice_clifor := tabella_esa_01.fieldbyname('cod_clifor').asstring;

    descrizione := '';

    try
      nr_col1 := pos('Num: ', tabella_esa_01.fieldbyname('des_partita').asstring);
      nr_col2 := pos('del', tabella_esa_01.fieldbyname('des_partita').asstring);

      if nr_col1 > 0 then
      begin

        numero_doc := strtoint(copy(tabella_esa_01.fieldbyname('des_partita').asstring, nr_col1 + 5, nr_col2 - 1 - nr_col1 - 5));
      end;

    except
      on E: exception do
      begin
        numero_doc := 0;
        descrizione := tabella_esa_01.fieldbyname('des_partita').asstring;
      end;
    end;

    try
      nr_col2 := pos('del', tabella_esa_01.fieldbyname('des_partita').asstring);

      if nr_col2 > 0 then
      begin
        data_doc := StrToDate(copy(tabella_esa_01.fieldbyname('des_partita').asstring, nr_col2 + 4, 10));
      end;

    except
      on E: exception do
      begin
        data_doc := 0;
      end;
    end;

    (*
      tabella_esa_02.Close;
      tabella_esa_02.parameters.parambyname('ind_clifor').value := tipo_cliente;
      tabella_esa_02.parameters.parambyname('cod_clifor').value := codice_clifor;
      tabella_esa_02.parameters.parambyname('dat_doc').value := data_doc;
      tabella_esa_02.parameters.parambyname('des_num_doc').value := numero_doc;
      tabella_esa_02.open;
 *)
    serie_doc := trim(tabella_esa_01.fieldbyname('sig_serie_doc').asstring);

    progressivo := progressivo + 1;
    tabella_01.append;
    tabella_01.fieldbyname('progressivo').asInteger := tabella_esa_01.fieldbyname('prg_scadenza').asInteger;
    tabella_01.fieldbyname('descrizione').asstring := descrizione;
    tabella_01.fieldbyname('pnr_progressivo').asInteger := tabella_esa_01.fieldbyname('prg_partita').asInteger;

    tabella_01.fieldbyname('cfg_tipo').asstring := tipo_cliente;
    if (tipo_cliente = 'C') then
    begin
      if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;
      end
      else
      begin
        assegna_codice_cli(tabella_esa_01.fieldbyname('cod_clifor').asstring, cli_for);
      end;

      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
    end
    else
    begin
      if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
      begin
        read_tabella(arc.arcdit, 'frn', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor').asstring);
        cli_for := archivio.fieldbyname('codice').asstring;
      end
      else
      begin
        assegna_codice_frn(tabella_esa_01.fieldbyname('cod_clifor').asstring, cli_for);
      end;

      tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
    end;

    if (tabella_esa_01.fieldbyname('val_da_pagare').asFloat < 0) and (tabella_esa_01.fieldbyname('cod_caucon').asstring <> 'NC') then
    begin
      tabella_01.fieldbyname('importo_dovuto').asFloat := tabella_esa_01.fieldbyname('val_da_pagare_doc').asFloat * -1;
      tabella_01.fieldbyname('importo_dovuto_euro').asFloat := tabella_esa_01.fieldbyname('val_da_pagare').asFloat * -1;
    end
    else
    begin
      tabella_01.fieldbyname('importo_dovuto').asFloat := tabella_esa_01.fieldbyname('val_da_pagare_doc').asFloat;
      tabella_01.fieldbyname('importo_dovuto_euro').asFloat := tabella_esa_01.fieldbyname('val_da_pagare').asFloat;
    end;

    tabella_01.fieldbyname('importo_pagato').asFloat := 0;
    tabella_01.fieldbyname('importo_pagato_euro').asFloat := 0;

    tabella_01.fieldbyname('data_registrazione').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
    try
      tabella_01.fieldbyname('numero_documento').asInteger := tabella_esa_01.fieldbyname('des_num_doc').asInteger;
    except
    end;
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
    tabella_01.fieldbyname('serie_documento').asstring := trim(tabella_esa_01.fieldbyname('sig_serie_doc').asstring);
    tabella_01.fieldbyname('data_scadenza').asdatetime := tabella_esa_01.fieldbyname('dat_scadenza').asdatetime;
    tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
    tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_valuta').asFloat;

    tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_01.fieldbyname('cod_pag').asstring;

    read_tabella(arc.arcdit, 'tpa', 'codice', tabella_esa_01.fieldbyname('cod_pag').asstring, 'tts_codice_fisse');
    tabella_01.fieldbyname('tts_codice').asstring := archivio.fieldbyname('tts_codice_fisse').asstring;

    if tabella_esa_01.fieldbyname('cod_agenzia').asstring <> '' then
    begin
      tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_01.fieldbyname('cod_banca').asstring;
      tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_01.fieldbyname('cod_agenzia').asstring;
    end;

    tabella_01.post;

    totale_pagare := 0;
    totale_pagare_euro := 0;

    nr_riga := 0;

    if (tabella_esa_01.fieldbyname('val_pagato').asFloat <> 0) then
    begin
      tabella_esa_03.Close;
      tabella_esa_03.Parameters.ParamByName('prg_scadenza').Value := tabella_esa_01.fieldbyname('prg_scadenza').asInteger;
      tabella_esa_03.open;

      while not tabella_esa_03.eof do
      begin
        nr_riga := nr_riga + 1;

        tabella_02.append;

        tabella_02.fieldbyname('progressivo').asFloat := tabella_esa_01.fieldbyname('prg_scadenza').asInteger;
        tabella_02.fieldbyname('riga').asInteger := nr_riga;
        tabella_02.fieldbyname('data_registrazione').asdatetime := data_doc;

        if (tabella_esa_03.fieldbyname('val_pagato').asFloat < 0) and (tabella_esa_01.fieldbyname('cod_caucon').asstring <> 'NC') then
        begin
          tabella_02.fieldbyname('importo_pagato').asFloat := tabella_esa_03.fieldbyname('val_pagato_doc').asFloat * -1;
          tabella_02.fieldbyname('importo_pagato_euro').asFloat := tabella_esa_03.fieldbyname('val_pagato').asFloat * -1;
        end
        else if (tabella_esa_03.fieldbyname('val_pagato').asFloat > 0) and (tabella_esa_01.fieldbyname('cod_caucon').asstring = 'NC') then
        begin
          tabella_02.fieldbyname('importo_pagato').asFloat := tabella_esa_03.fieldbyname('val_pagato_doc').asFloat * -1;
          tabella_02.fieldbyname('importo_pagato_euro').asFloat := tabella_esa_03.fieldbyname('val_pagato').asFloat * -1;
        end

        else
        begin
          tabella_02.fieldbyname('importo_pagato').asFloat := tabella_esa_03.fieldbyname('val_pagato_doc').asFloat;
          tabella_02.fieldbyname('importo_pagato_euro').asFloat := tabella_esa_03.fieldbyname('val_pagato').asFloat;
        end;

        tabella_02.post;
        tabella_esa_03.next;
      end;
      tabella_esa_03.Close;

    end;

    tabella_esa_01.next;

  end; // while

  tpa.Close;
  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;
  v_conferma.enabled := true;

end;

procedure TCONVEREADY.tabella_01BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_01.fieldbyname('utente').asstring := utente;
  tabella_01.fieldbyname('data_ora').asdatetime := now;
end;

procedure TCONVEREADY.tabella_02BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_02.fieldbyname('utente').asstring := utente;
  tabella_02.fieldbyname('data_ora').asdatetime := now;

end;

procedure TCONVEREADY.tabella_03BeforePost(DataSet: TDataSet);
begin
  inherited;
  tabella_03.fieldbyname('utente').asstring := utente;
  tabella_03.fieldbyname('data_ora').asdatetime := now;

end;

procedure TCONVEREADY.controllo_campi;
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

procedure TCONVEREADY.converti_tva;
begin
  v_tabella_01.caption := 'tabella valute';
  application.processmessages;

  cancella_tabella('tva');
  tabella_01.Close;
  tabella_01.tablename := 'tva';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_VALUTE');
  tabella_esa_01.sql.add('order by cod_valuta');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_valuta').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_valuta').asstring);
    tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_cambio_paniere').asFloat;
    tabella_01.fieldbyname('codice_iso').asstring := tabella_esa_01.fieldbyname('cod_iso').asstring;
    if tabella_01.fieldbyname('codice_iso').asstring = '' then
    begin
      tabella_01.fieldbyname('codice_iso').asstring := tabella_esa_01.fieldbyname('sig_valuta').asstring;
    end;

    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tiv;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella iva';
  application.processmessages;

  cancella_tabella('tiv');
  tabella_01.Close;
  tabella_01.tablename := 'tiv';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_IVA');
  tabella_esa_01.sql.add('ORDER BY cod_iva');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_iva').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_iva').asstring);
    tabella_01.fieldbyname('percentuale').asFloat := tabella_esa_01.fieldbyname('prc_iva').asFloat;
    tabella_01.fieldbyname('indetraibile').asFloat := tabella_esa_01.fieldbyname('prc_indetraibile').asFloat;

    if tabella_esa_01.fieldbyname('cod_iva_vent').asstring <> '' then
    begin
      tabella_01.fieldbyname('tiv_codice_ventilazione').asstring := tabella_esa_01.fieldbyname('cod_iva_vent').asstring;
    end;

    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tna;
begin
  v_tabella_01.caption := 'tabella nazioni';
  application.processmessages;

  cancella_tabella('tna');
  tabella_01.Close;
  tabella_01.tablename := 'tna';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_NAZIONI');
  tabella_esa_01.sql.add('order by cod_naz');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_naz').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_naz').asstring);
    tabella_01.fieldbyname('codice_iso').asstring := tabella_esa_01.fieldbyname('sig_iso').asstring;
    tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
    if tabella_01.fieldbyname('tva_codice').asstring = '' then
    begin
      tabella_01.fieldbyname('tva_codice').asstring := divisa_di_conto;
    end;

    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;

end;

procedure TCONVEREADY.converti_tag;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella agenti';
  application.processmessages;

  (*
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
 *)
  cancella_tabella('tag');
  // cancella_tabella('esa_tag');

  tabella_01.Close;
  tabella_01.tablename := 'tag';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT C.*, A.* FROM CO_AGENTI C');
  tabella_esa_01.sql.add('INNER JOIN CA_ANAGRAFICHE A on A.cod_anagra=C.cod_agente');
  tabella_esa_01.sql.add('order by cod_agente');
  tabella_esa_01.open;

  tabella_01.append;
  tabella_01.fieldbyname('codice').asstring := '0000';
  tabella_01.fieldbyname('descrizione').asstring := 'Agente standard';
  tabella_01.post;

  i := 0;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;
    i := i + 1;

    if tabella_esa_01.locate('Codice_anagrafica', copy(tabella_esa_01.fieldbyname('cod_anagra').asstring, 3, 4), []) then
    begin
      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := copy(tabella_esa_01.fieldbyname('cod_anagra').asstring, 3, 4);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('Ragione_sociale').asstring);
      tabella_01.post;
    end;

    tabella_esa_01.next;
  end;

  query.Close;
  query_02.Close;
  tabella_01.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;
end;

procedure TCONVEREADY.converti_tcm;
var
  i: word;
  codice, campi: string;
begin
  v_tabella_01.caption := 'tabella categorie merceologiche';
  application.processmessages;

  cancella_tabella('tcm');
  tabella_01.Close;
  tabella_01.tablename := 'tcm';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM MG_GRUMERC');
  tabella_esa_01.sql.add('ORDER BY cod_grumerc');
  tabella_esa_01.open;

  i := 0;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    codice := tabella_esa_01.fieldbyname('cod_grumerc').asstring;

    if codice <> '' then
    begin
      i := i + 1;

      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := codice;
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_grumerc').asstring);

      tabella_01.post;

    end;
    tabella_esa_01.next;
  end;

  query.Close;
  query_02.Close;
  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;

end;

procedure TCONVEREADY.converti_tgm;
var
  i: word;
  codice: string;
begin
  v_tabella_01.caption := 'tabella gruppi merceologici';
  application.processmessages;

  cancella_tabella('tgm');

  try
    tabella_01.Close;
    tabella_01.sql.clear;
    tabella_01.sql.add('alter table tgm ');
    tabella_01.sql.add('add cod_grumerc varchar(3) default '''' ');
    tabella_01.ExecSQL;
  except
  end;

  try
    tabella_01.Close;
    tabella_01.tablename := '';
    tabella_01.sql.clear;
    tabella_01.sql.add(' alter table tgm ');
    tabella_01.sql.add(' add  cod_sgrumerc varchar(3) default ''''  ');
    tabella_01.ExecSQL;

  except
  end;

  tabella_01.Close;
  tabella_01.sql.clear;
  tabella_01.sql.add('select * from tgm');
  tabella_01.sql.add('where');
  tabella_01.sql.add('codice=:codice ');

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('select * from MG_SGRUMERC');
  tabella_esa_01.sql.add('order by cod_grumerc, cod_sgrumerc');
  tabella_esa_01.open;

  i := 0;
  tabella_01_ds.DataSet := tabella_esa_01;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    i := i + 1;

    codice := setta_lunghezza(inttostr(i), 4, true, '0');

    tabella_01.Close;
    tabella_01.ParamByName('codice').asstring := codice;
    tabella_01.open;
    if tabella_01.eof then
    begin
      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := codice;
      tabella_01.fieldbyname('cod_grumerc').asstring := tabella_esa_01.fieldbyname('cod_grumerc').asstring;
      tabella_01.fieldbyname('cod_sgrumerc').asstring := tabella_esa_01.fieldbyname('cod_sgrumerc').asstring;
      tabella_01.fieldbyname('descrizione').asstring := tabella_esa_01.fieldbyname('des_sgrumerc').asstring;
      tabella_01.post;
    end;

    tabella_esa_01.next;
  end;

  tabella_01.Close;

end;

procedure TCONVEREADY.converti_tma;
begin
  v_tabella_01.caption := 'tabella depositi';
  application.processmessages;

  cancella_tabella('tma');
  tabella_01.Close;
  tabella_01.tablename := 'tma';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM MG_DEPOSITI');
  tabella_esa_01.sql.add('ORDER BY cod_dep');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_dep').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_dep').asstring);
    if tabella_esa_01.fieldbyname('flg_prop_merce').asstring = '1' then
    begin
      tabella_01.fieldbyname('proprieta').asstring := 'si';
    end
    else
    begin
      tabella_01.fieldbyname('proprieta').asstring := 'no';
    end;

    tabella_01.fieldbyname('descrizione1').asstring := '';
    tabella_01.fieldbyname('descrizione2').asstring := '';
    tabella_01.fieldbyname('via').asstring := tabella_esa_01.fieldbyname('des_indir').asstring;
    tabella_01.fieldbyname('cap').asstring := tabella_esa_01.fieldbyname('cod_cap').asstring;
    tabella_01.fieldbyname('citta').asstring := tabella_esa_01.fieldbyname('des_loc').asstring;
    tabella_01.fieldbyname('provincia').asstring := tabella_esa_01.fieldbyname('sig_prov').asstring;

    tabella_01.fieldbyname('tna_codice').asstring := tabella_esa_01.fieldbyname('cod_naz').asstring;

    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tba;
begin
  v_tabella_01.caption := 'tabella banche';
  application.processmessages;

  (*
    cancella_tabella('tba');

    tabella_01.close;
    tabella_01.tablename := 'tba';
    tabella_01.open;

    tabella_01_ds.Dataset := tabella_01;

    tabella_esa_01.Close;
    tabella_esa_01.Sql.Clear;
    tabella_esa_01.Sql.Add('SELECT * FROM CA_BANCHE');
    tabella_esa_01.Sql.Add('order by cod_banca');
    tabella_esa_01.Open;

    while not tabella_esa_01.eof do
    begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.FieldByName('cod_banca').AsString);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.FieldByName('des_banca').AsString);

    tabella_01.post;

    tabella_esa_01.Next;
    end;
 *)
  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tum;
begin
  v_tabella_01.caption := 'tabella unita misura';
  application.processmessages;

  cancella_tabella('tum');
  tabella_01.Close;
  tabella_01.tablename := 'tum';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_UM');
  tabella_esa_01.sql.add('order by cod_um');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;
    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_um').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_um').asstring);
    tabella_01.fieldbyname('decimali').asInteger := tabella_esa_01.fieldbyname('num_dec_qta').asInteger;
    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tlv;
begin
  v_tabella_01.caption := 'tabella listini';
  application.processmessages;

  cancella_tabella('tlv');
  tabella_01.Close;
  tabella_01.tablename := 'tlv';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM PZ_LISTINI');
  tabella_esa_01.sql.add('order by cod_listino');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;

    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_listino').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_listino').asstring);
    tabella_01.fieldbyname('tva_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_valuta').asstring);
    if tabella_esa_01.fieldbyname('flg_iva_compresa').asstring = '0' then
      tabella_01.fieldbyname('iva_inclusa').asstring := 'no'
    else
      tabella_01.fieldbyname('iva_inclusa').asstring := 'si';

    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tco;
begin
  v_tabella_01.caption := 'causali contabili';
  application.processmessages;

  cancella_tabella('tco');
  tabella_01.Close;
  tabella_01.tablename := 'tco';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_CAUCON');
  tabella_esa_01.sql.add('where');
  tabella_esa_01.sql.add('cod_piacont_mod=' + quotedstr('STD'));
  tabella_esa_01.sql.add('order by cod_caucon');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    if tabella_esa_01.fieldbyname('cod_caucon').asstring <> '' then
    begin
      tabella_01.append;

      tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_caucon').asstring);
      tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_caucon').asstring);
      if tabella_01.fieldbyname('descrizione').asstring = '' then
      begin
        tabella_01.fieldbyname('descrizione').asstring := '.';
      end;

      if tabella_esa_01.fieldbyname('cod_reg_iva').asstring = 'V' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'vendite';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';
      end
      else if tabella_esa_01.fieldbyname('cod_reg_iva').asstring = 'A' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'acquisti';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';
      end
      else if tabella_esa_01.fieldbyname('cod_reg_iva').asstring = 'C' then
      begin
        tabella_01.fieldbyname('tipo_registro_iva').asstring := 'corrispettivi';
        tabella_01.fieldbyname('movimento_iva').asstring := 'si';

      end;

      if trim(tabella_esa_01.fieldbyname('flg_partscad').asstring) = '0' then
      begin
        tabella_01.fieldbyname('gestione_partite').asstring := 'no';
      end
      else
      begin
        tabella_01.fieldbyname('gestione_partite').asstring := 'si';
      end;

      if tabella_esa_01.fieldbyname('cod_caucon').asstring = '51' then
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'apertura bilancio';
      end
      else if tabella_esa_01.fieldbyname('cod_caucon').asstring = '50' then
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'chiusura bilancio';
      end
      else
      begin
        tabella_01.fieldbyname('tipo_movimento').asstring := 'normale';
      end;
      (*
        if trim(tabella_esa_01.fieldbyname('Flag_insoluto').asstring) = '1' then
        begin
        tabella_01.fieldbyname('insoluto').asstring := 'si';
        end
        else
        begin
        tabella_01.fieldbyname('insoluto').asstring := 'no';
        end;
 *)
      if (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite') or (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'acquisti') then
      begin
        if (trim(tabella_esa_01.fieldbyname('ind_tipo_oper_iva').asstring) <> '1') then
        begin
          tabella_01.fieldbyname('segno_registro_iva').asstring := 'incrementa';
        end
        else
        begin
          tabella_01.fieldbyname('segno_registro_iva').asstring := 'decrementa';
        end;
      end;

      if (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite') or (tabella_01.fieldbyname('tipo_registro_iva').asstring = 'acquisti') then
      begin
        // tabella_01.fieldbyname('serie_numerazione').asstring :=
        // trim(tabella_esa_01.fieldbyname('nRegistroivaproposto').asString) +
        // trim(tabella_esa_01.fieldbyname('SerieNumerazProposta').asString);
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
        tabella_01.fieldbyname('cfg_codice_02').asstring := tabella_esa_01.fieldbyname('cod_piacont_iva').asstring;
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'D';
      end;
      if tabella_01.fieldbyname('tipo_registro_iva').asstring = 'vendite' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'C';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'D';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_codice_02').asstring := tabella_esa_01.fieldbyname('cod_piacont_iva').asstring;
        tabella_01.fieldbyname('cfg_dare_avere_02').asstring := 'A';
      end;
      if tabella_01.fieldbyname('tipo_registro_iva').asstring = 'corrispettivi' then
      begin
        tabella_01.fieldbyname('cfg_tipo_01').asstring := 'G';
        tabella_01.fieldbyname('cfg_dare_avere_01').asstring := 'D';

        tabella_01.fieldbyname('cfg_tipo_02').asstring := 'G';
        tabella_01.fieldbyname('cfg_codice_02').asstring := '';
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

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tsp;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella tipo spedizione';
  application.processmessages;
  (*
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
 *)
end;

procedure TCONVEREADY.converti_tzo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella zone';
  application.processmessages;

  cancella_tabella('tzo');
  tabella_01.Close;
  tabella_01.tablename := 'tzo';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CO_ZONE');
  tabella_esa_01.sql.add('ORDER BY cod_zona');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    tabella_01.append;
    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_zona').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_zona').asstring);
    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tdo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella documenti vendita';
  application.processmessages;

  cancella_tabella('tdo');
  tabella_01.Close;
  tabella_01.tablename := 'tdo';
  tabella_01.open;

  tco.open;

  // --------------------------------------------------------------------
  // ordini clienti
  // --------------------------------------------------------------------
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * from MG_TIPODOC');
  tabella_esa_01.sql.add('where');
  tabella_esa_01.sql.add('cod_gruppo_tipodoc in (''BLV'', ''FAV'' ) ');
  tabella_esa_01.sql.add('order by 1');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;
    if not tabella_01.locate('codice', tabella_esa_01.fieldbyname('cod_caumag').asstring, []) then
    begin
      tabella_01.append;
      tabella_01.fieldbyname('codice').asstring := tabella_esa_01.fieldbyname('cod_caumag').asstring;
      tabella_01.fieldbyname('descrizione').asstring := tabella_esa_01.fieldbyname('des_tipodoc').asstring;
      if tabella_esa_01.fieldbyname('cod_caucon').asstring = 'FE' then
      begin
        if tabella_esa_01.fieldbyname('cod_caumag').asstring = '023' then
        begin
          tabella_01.fieldbyname('tipo_documento').asstring := 'fattura differita';
        end
        else
        begin
          tabella_01.fieldbyname('tipo_documento').asstring := 'fattura immediata';
        end;
      end
      else if tabella_esa_01.fieldbyname('cod_caucon').asstring = 'NC' then
      begin
        tabella_01.fieldbyname('tipo_documento').asstring := 'nota credito';
      end
      else
      begin
        tabella_01.fieldbyname('tipo_documento').asstring := 'ddt';
        tabella_01.fieldbyname('tma_codice').asstring := '000';
      end;

      tabella_01.fieldbyname('tco_codice').asstring := tabella_esa_01.fieldbyname('cod_caucon').asstring;

      try
        tabella_01.post;
      except
      end;
    end;

    tabella_esa_01.next;
  end;

  tabella_esa_01.Close;
  tabella_01.Close;

end;

procedure TCONVEREADY.converti_tcc;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella categorie contabili clienti';
  application.processmessages;

  cancella_tabella('tcc');
  tabella_01.Close;
  tabella_01.tablename := 'tcc';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;
  tabella_01.Close;
end;

procedure TCONVEREADY.converti_tca;
var
  i: word;
begin

  v_tabella_01.caption := 'tabella categorie contabili articoli';
  application.processmessages;

  cancella_tabella('tca');
  tabella_01.Close;
  tabella_01.tablename := 'tca';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;

  tabella_01.Close;
end;

procedure TCONVEREADY.converti_tcf;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella categorie contabili fornitori';
  application.processmessages;

  cancella_tabella('tcf');
  tabella_01.Close;
  tabella_01.tablename := 'tcf';
  tabella_01.open;

  tabella_01.append;
  tabella_01.fieldvalues['codice'] := '0';
  tabella_01.fieldvalues['descrizione'] := '.';
  tabella_01.post;
  tabella_01.Close;
end;

procedure TCONVEREADY.converti_tpo;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella porti';
  application.processmessages;

  cancella_tabella('tpo');
  tabella_01.Close;
  tabella_01.tablename := 'tpo';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_PORTO');
  tabella_esa_01.sql.add('order by cod_porto');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;
    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_porto').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_porto').asstring);
    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;

end;

procedure TCONVEREADY.converti_tab;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella aspetto beni';
  application.processmessages;

  cancella_tabella('tab');
  tabella_01.Close;
  tabella_01.tablename := 'tab';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;

  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM CA_ASPBENI');
  tabella_esa_01.sql.add('order by cod_asp_beni');
  tabella_esa_01.open;

  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    tabella_01.append;
    tabella_01.fieldbyname('codice').asstring := trim(tabella_esa_01.fieldbyname('cod_asp_beni').asstring);
    tabella_01.fieldbyname('descrizione').asstring := trim(tabella_esa_01.fieldbyname('des_asp_beni').asstring);
    tabella_01.post;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_esa_01.Close;
end;

procedure TCONVEREADY.converti_tsa;
var
  i: word;
begin
  v_tabella_01.caption := 'tabella codice statistico';
  application.processmessages;

  (*
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
 *)
  tabella_01.Close;
  tabella_esa_01.Close;
end;

// ******************************************************************************

procedure TCONVEREADY.converti_mov;
var
  i, progressivo: integer;
begin
  v_tabella_01.caption := 'movimenti magazzino';
  application.processmessages;

  cancella_tabella('magese');
  cancella_tabella('mag');

  cancella_tabella('mmt');
  tabella_01.Close;
  tabella_01.tablename := 'mmt';
  tabella_01.open;

  cancella_tabella('mmr');
  tabella_02.Close;
  tabella_02.tablename := 'mmr';
  tabella_02.open;

  tabella_01_ds.DataSet := tabella_esa_01;
  try
    tabella_esa_01.Close;
    tabella_esa_01.sql.clear;
    tabella_esa_01.sql.add('SELECT ');
    tabella_esa_01.sql.add('T.prg_movimento,');
    tabella_esa_01.sql.add('T.num_esercizio,');
    tabella_esa_01.sql.add('T.dat_movimento,');
    tabella_esa_01.sql.add('T.dat_doc,');
    tabella_esa_01.sql.add('T.des_num_doc,');
    tabella_esa_01.sql.add('T.sig_serie_doc,');
    tabella_esa_01.sql.add('T.sig_serie_prot,');
    tabella_esa_01.sql.add('T.ind_clifor_c,');
    tabella_esa_01.sql.add('T.cod_clifor_c,');
    tabella_esa_01.sql.add('T.cod_dep_num,');
    tabella_esa_01.sql.add('T.cod_caumag,');
    tabella_esa_01.sql.add('T.cod_dep_mov as cod_dep_mov_t,');
    tabella_esa_01.sql.add('T.cmb_valuta,');
    tabella_esa_01.sql.add('T.cod_valuta,');
    tabella_esa_01.sql.add('R.*');
    tabella_esa_01.sql.add('FROM MG_MOVMAGT T');
    tabella_esa_01.sql.add('inner join MG_MOVMAGR R on R.prg_movimento=T.prg_movimento');
    tabella_esa_01.sql.add('WHERE');
    tabella_esa_01.sql.add('T.cod_gruppo_tipodoc=''MMM'' ');
    tabella_esa_01.sql.add('order by R.prg_movimento,R.prg_movimento_riga');
    tabella_esa_01.open;
  except

    messaggio(000, 'manca la tabella del magazzino (MG_MOVMAGT)');

    Close;
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

    v_tabella_01.caption := 'movimenti magazzino ' + tabella_esa_01.fieldbyname('dat_movimento').asstring;

    codice_articolo := trim(tabella_esa_01.fieldbyname('cod_art').asstring);

    if (test_data <> tabella_esa_01.fieldbyname('dat_movimento').asdatetime) or (test_numero_stringa <> tabella_esa_01.fieldbyname('des_num_doc').asstring) or (test_cod_causale <> trim(tabella_esa_01.fieldbyname('cod_caumag').asstring)) then
    begin
      progressivo := progressivo + 1;
      crea_mmt(progressivo);
      test_data := tabella_esa_01.fieldbyname('dat_movimento').asdatetime;
      try
        test_numero_stringa := trim(tabella_esa_01.fieldbyname('des_num_doc').asstring);
      except
        test_numero_stringa := '';
      end;
      test_cod_causale := tabella_esa_01.fieldbyname('cod_caumag').asstring;
    end;

    crea_mmr;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_esa_01.Close;

  // arc.attivazione_trigger(arc.arcdit, false, true);
end;

procedure TCONVEREADY.crea_mmt(progressivo: integer);
begin

  tabella_01.append;
  tabella_01.fieldbyname('progressivo').asInteger := progressivo;
  tabella_01.fieldbyname('data_registrazione').asdatetime := tabella_esa_01.fieldbyname('dat_movimento').asdatetime;
  tabella_01.fieldbyname('tmo_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_caumag').asstring);
  tabella_01.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('cod_dep_mov_t').asstring;
  try
    tabella_01.fieldbyname('numero_documento').asInteger := tabella_esa_01.fieldbyname('des_num_doc').asInteger;
  except
    tabella_01.fieldbyname('numero_documento').asInteger := 0;
  end;

  try
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
  except
    tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_movimento').asdatetime;
  end;

  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('sig_serie_doc').asstring;

  if trim(tabella_esa_01.fieldbyname('ind_clifor_c').asstring) = 'C' then
  begin
    if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
    begin
      read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor_c').asstring);
      cli_for := archivio.fieldbyname('codice').asstring;
    end
    else
    begin
      assegna_codice_cli(tabella_esa_01.fieldbyname('cod_clifor_c').asstring, cli_for);
    end;

    tabella_01.fieldbyname('cfg_tipo').asstring := 'C';
    tabella_01.fieldbyname('cfg_codice').asstring := cli_for;

  end

  else if trim(tabella_esa_01.fieldbyname('ind_clifor_c').asstring) = 'F' then
  begin
    if (codice_nom_numerico = 'si') and (not v_codifica_fornitori.checked) then
    begin
      read_tabella(arc.arcdit, 'frn', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor_c').asstring);
      cli_for := archivio.fieldbyname('codice').asstring;
    end
    else
    begin
      assegna_codice_frn(tabella_esa_01.fieldbyname('cod_clifor_c').asstring, cli_for);
    end;

    tabella_01.fieldbyname('cfg_tipo').asstring := 'F';
    tabella_01.fieldbyname('cfg_codice').asstring := cli_for;
  end;

  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
  tabella_01.fieldbyname('ese_codice').asstring := tabella_esa_01.fieldbyname('num_esercizio').asstring;
  tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_valuta').asFloat;
  tabella_01.post;

  riga := 0;

end;

procedure TCONVEREADY.crea_mmr;
var
  campo_codice: string;
begin
  if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
  begin
    campo_codice := 'codice_alternativo';
  end
  else
  begin
    campo_codice := 'codice';
  end;

  if read_tabella(arc.arcdit, 'art', campo_codice, codice_articolo) then
  begin
    tabella_02.append;

    tabella_02.fieldbyname('progressivo').asFloat := tabella_01.fieldbyname('progressivo').asFloat;
    riga := riga + 1;
    tabella_02.fieldbyname('riga').asInteger := riga;

    if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
    begin
      tabella_02.fieldbyname('art_codice').asstring := archivio.fieldbyname('codice').asstring;
    end
    else
    begin
      tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
    end;

    tabella_02.fieldbyname('quantita').asFloat := tabella_esa_01.fieldbyname('qta_merce').asFloat;
    tabella_02.fieldbyname('prezzo').asFloat := tabella_esa_01.fieldbyname('prz_merce').asFloat;
    tabella_02.fieldbyname('importo').asFloat := tabella_esa_01.fieldbyname('val_merce_doc').asFloat;
    tabella_02.fieldbyname('importo_euro').asFloat := tabella_esa_01.fieldbyname('val_merce').asFloat;
    tabella_02.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('cod_dep_mov').asstring;

    if (tabella_02.fieldbyname('prezzo').asFloat = 0) and (tabella_02.fieldbyname('importo').asFloat > 0) then
    begin
      tabella_02.fieldbyname('prezzo').asFloat := tabella_esa_01.fieldbyname('val_merce_doc').asFloat / tabella_esa_01.fieldbyname('qta_merce').asFloat;
    end;

    read_tabella(arc.arcdit, 'tmo', 'codice', tabella_01.fieldbyname('tmo_codice').asstring);
    if archivio.fieldbyname('esistenza').asstring = 'incrementa' then
    begin
      tabella_02.fieldbyname('quantita_entrate').asFloat := tabella_esa_01.fieldbyname('qta_merce').asFloat;
      tabella_02.fieldbyname('quantita_uscite').asFloat := 0;
    end
    else if archivio.fieldbyname('esistenza').asstring = 'decrementa' then
    begin
      tabella_02.fieldbyname('quantita_entrate').asFloat := 0;
      tabella_02.fieldbyname('quantita_uscite').asFloat := tabella_esa_01.fieldbyname('qta_merce').asFloat;
    end
    else if archivio.fieldbyname('esistenza').asstring = 'ignora' then
    begin
      tabella_02.fieldbyname('quantita_entrate').asFloat := 0;
      tabella_02.fieldbyname('quantita_uscite').asFloat := 0;
    end;

    (* ?
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
 *)

    tabella_02.post;

  end;

end;

procedure TCONVEREADY.converti_ordini_clienti;
var
  i: integer;
  test_serie, test_numero_doc: string;
  test_progressivo: integer;
begin
  v_tabella_01.caption := 'ordini clienti';
  application.processmessages;

  query.sql.clear;
  query.sql.add('select * from esa_tag');
  query.sql.add('where');
  query.sql.add('esa_codice=:esa_codice');
  try
    query.open;
  except
    converti_tag;
  end;

  cancella_tabella('ovt');
  tabella_01.Close;
  tabella_01.tablename := 'ovt';
  tabella_01.open;

  cancella_tabella('ovr');
  tabella_02.Close;
  tabella_02.tablename := 'ovr';
  tabella_02.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM OR_ORDINIT');
  tabella_esa_01.sql.add('WHERE');
  tabella_esa_01.sql.add('cod_gruppo_tipodoc=:cod_gruppo_tipodoc');
  tabella_esa_01.sql.add('ORDER BY prg_ordine');
  tabella_esa_01.Parameters.ParamByName('cod_gruppo_tipodoc').Value := 'IMP';
  tabella_esa_01.open;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT * FROM OR_ORDINIR');
  tabella_esa_02.sql.add('WHERE');
  tabella_esa_02.sql.add('prg_ordine=:prg_ordine');
  tabella_esa_02.sql.add('ORDER BY prg_ordine_riga');

  test_data := 0;
  test_alfa := '';
  test_numero := 0;
  test_progressivo := 0;

  test_serie := '';
  test_numero_doc := '';

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin
    tabella_01_ds.DataSet := tabella_esa_01;
    crea_ovt;

    test_serie := tabella_esa_01.fieldbyname('sig_serie_doc').asstring;
    test_numero_doc := tabella_esa_01.fieldbyname('num_doc').asstring;

    test_progressivo := tabella_esa_01.fieldbyname('prg_ordine').asInteger;

    tabella_esa_02.Close;
    tabella_esa_02.Parameters.ParamByName('prg_ordine').Value := test_progressivo;
    tabella_esa_02.open;
    if not tabella_esa_02.eof then
    begin
      crea_ovr(test_progressivo);
    end;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_03.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;

end;

procedure TCONVEREADY.crea_ovt;
begin
  tabella_01.append;
  tabella_01.fieldbyname('progressivo').asInteger := tabella_esa_01.fieldbyname('prg_ordine').asInteger;
  tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
  tabella_01.fieldbyname('tdo_codice').asstring := 'ORDV';
  tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('sig_serie_doc').asstring;
  tabella_01.fieldbyname('numero_documento').asInteger := tabella_esa_01.fieldbyname('num_doc').asInteger;

  assegna_codice_cli(tabella_esa_01.fieldbyname('cod_clifor_c').asstring, cli_for);

  tabella_01.fieldbyname('cli_codice').asstring := cli_for;
  tabella_01.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('cod_dep_mov').asstring;
  if trim(tabella_esa_01.fieldbyname('cod_agente').asstring) = '' then
  begin
    tabella_01.fieldbyname('tag_codice').asstring := '0000'
  end
  else
  begin
    query.Close;
    query.ParamByName('esa_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_agente').asstring);
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

  tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_01.fieldbyname('cod_pag').asstring;
  tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_01.fieldbyname('cod_banca').asstring;
  tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_01.fieldbyname('sig_ccbanc').asstring;

  tabella_01.fieldbyname('tlv_codice').asstring := tabella_esa_01.fieldbyname('cod_listino').asstring;
  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
  tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_valuta').asFloat;

  tabella_01.fieldbyname('nostro_riferimento').asstring := tabella_esa_01.fieldbyname('des_riferimento').asstring;

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
  tabella_01.fieldbyname('tsp_codice').asstring := archivio.fieldbyname('tsp_codice').asstring;
  tabella_01.fieldbyname('tst_codice').asstring := archivio.fieldbyname('tst_codice').asstring;
  tabella_01.fieldbyname('addebito_spese_fattura').asstring := archivio.fieldbyname('addebito_spese_fattura').asstring;

  tabella_01.fieldbyname('tpo_codice').asstring := tabella_esa_01.fieldbyname('cod_porto').asstring;;
  read_tabella(arc.arcdit, 'tpo', 'codice', tabella_01.fieldbyname('tpo_codice').asstring);
  tabella_01.fieldbyname('addebito_spese_trasporto').asstring := archivio.fieldbyname('addebito').asstring;
  tabella_01.post;

  riga := 0;
end;

procedure TCONVEREADY.crea_ovr(progressivo: integer);
var
  codice_articolo: string;
begin

  while not tabella_esa_02.eof do
  begin

    codice_articolo := trim(tabella_esa_02.fieldbyname('cod_art').asstring);
    if (codice_articolo <> '') then
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

      tabella_02.fieldbyname('progressivo').asFloat := tabella_esa_02.fieldbyname('prg_ordine').asFloat;
      tabella_02.fieldbyname('riga').asInteger := tabella_esa_02.fieldbyname('prg_ordine_riga').asInteger;

      if (codice_articolo = '') then
      begin
        tabella_02.fieldbyname('art_codice').asstring := '';
      end
      else
      begin
        if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
        begin
          read_tabella(arc.arcdit, 'art', 'codice_alternativo', codice_articolo);
          tabella_02.fieldbyname('art_codice').asstring := archivio.fieldbyname('codice').asstring;
        end
        else
        begin
          tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
        end;
      end;

      tabella_02.fieldbyname('descrizione1').asstring := copy(trim(tabella_esa_02.fieldbyname('des_articolo_riga').asstring), 1, 40);
      tabella_02.fieldbyname('descrizione2').asstring := copy(trim(tabella_esa_02.fieldbyname('des_articolo_riga').asstring), 41, 40);

      tabella_02.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('cod_cpartita').asstring;
      tabella_02.fieldbyname('tiv_codice').asstring := tabella_esa_02.fieldbyname('cod_iva').asstring;

      tabella_02.fieldbyname('quantita').asFloat := arrotonda(tabella_esa_02.fieldbyname('qta_merce').asFloat + tabella_esa_02.fieldbyname('qta_merce_delta').asFloat);
      tabella_02.fieldbyname('quantita_evasa').asFloat := arrotonda(tabella_esa_02.fieldbyname('qta_merce_delta').asFloat);

      tabella_02.fieldbyname('prezzo').asFloat := tabella_esa_02.fieldbyname('prz_merce').asFloat;
      tabella_02.fieldbyname('tsm_codice').asFloat := tabella_esa_02.fieldbyname('prc_sconto1').asFloat;
      tabella_02.fieldbyname('tsm_codice').asFloat := tabella_esa_02.fieldbyname('prc_sconto2').asFloat;

      calcola_importo;
      tabella_02.fieldbyname('tma_codice').asstring := tabella_esa_02.fieldbyname('cod_dep_mov').asstring;

      if tabella_esa_02.fieldbyname('ind_stato_riga').asstring = 'E' then
        tabella_02.fieldbyname('situazione').asstring := 'evaso'
      else if tabella_esa_02.fieldbyname('ind_stato_riga').asstring = 'P' then
        tabella_02.fieldbyname('situazione').asstring := 'evaso parziale'
      else
        tabella_02.fieldbyname('situazione').asstring := 'inserito';

      tabella_02.post;
    end;

    tabella_esa_02.next;
  end; // while

end;

procedure TCONVEREADY.converti_fatture_vendita;
begin
  v_tabella_01.caption := 'fatture clienti';
  application.processmessages;

  cancella_tabella('fvr');
  tabella_02.Close;
  tabella_02.tablename := 'fvr';
  tabella_02.open;

  cancella_tabella('fvi');
  tabella_03.Close;
  tabella_03.tablename := 'fvi';
  tabella_03.open;

  cancella_tabella('fvt');
  tabella_01.Close;
  tabella_01.tablename := 'fvt';
  tabella_01.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT T.*, ');
  tabella_esa_01.sql.add('case when T.cod_tipodoc=''022'' then ''FAIV''  ');
  tabella_esa_01.sql.add('     when T.cod_tipodoc=''022ACC'' then ''FAAV''  ');
  tabella_esa_01.sql.add('     when T.cod_tipodoc=''023'' then ''FADV''  ');
  tabella_esa_01.sql.add('     when T.cod_tipodoc=''025'' then ''NC''  ');
  tabella_esa_01.sql.add('     else ''FAIV'' ');
  tabella_esa_01.sql.add('end tdo_codice  ');
  tabella_esa_01.sql.add('FROM MG_MOVMAGT T');
  tabella_esa_01.sql.add('WHERE');
  tabella_esa_01.sql.add('T.cod_gruppo_tipodoc=''FAV'' ');
  tabella_esa_01.sql.add('ORDER BY T.prg_movimento');
  tabella_esa_01.open;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT * FROM MG_MOVMAGR');
  tabella_esa_02.sql.add('WHERE');
  tabella_esa_02.sql.add('prg_movimento=:prg_movimento');
  tabella_esa_02.sql.add('ORDER BY num_riga');

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  tabella_01_ds.DataSet := tabella_esa_01;
  while not tabella_esa_01.eof do
  begin
    application.processmessages;

    crea_fvt;
    crea_fvi;
    crea_fvr(tabella_esa_01.fieldbyname('prg_movimento').asInteger);

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_03.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;
  tabella_esa_03.Close;
end;

procedure TCONVEREADY.crea_fvt;
var
  tcc_codice: string;
begin
  tabella_esa_03.Close;
  tabella_esa_03.sql.clear;
  tabella_esa_03.sql.add('SELECT * FROM VN_DOCUMENTO');
  tabella_esa_03.sql.add('WHERE');
  tabella_esa_03.sql.add('prg_movimento=:prg_movimento ');
  tabella_esa_03.Parameters.ParamByName('prg_movimento').Value := tabella_esa_01.fieldbyname('prg_movimento').asInteger;
  tabella_esa_03.open;

  if (codice_nom_numerico = 'si') and (not v_codifica_clienti.checked) then
  begin
    read_tabella(arc.arcdit, 'cli', 'codice_alternativo', tabella_esa_01.fieldbyname('cod_clifor_c').asstring);
    cli_for := archivio.fieldbyname('codice').asstring;
  end
  else
  begin
    assegna_codice_cli(tabella_esa_01.fieldbyname('cod_clifor_c').asstring, cli_for);
  end;

  read_tabella(arc.arcdit, 'cli', 'codice', cli_for);
  tcc_codice := archivio.fieldbyname('tcc_codice').asstring;

  tabella_01.append;
  tabella_01.fieldbyname('progressivo').asInteger := tabella_esa_01.fieldbyname('prg_movimento').asInteger;

  tabella_01.fieldbyname('cli_codice').asstring := cli_for;
  tabella_01.fieldbyname('tlv_codice').asstring := archivio.fieldbyname('tlv_codice').asstring;
  tabella_01.fieldbyname('tpo_codice').asstring := archivio.fieldbyname('tpo_codice').asstring;
  tabella_01.fieldbyname('tsp_codice').asstring := archivio.fieldbyname('tsp_codice').asstring;
  tabella_01.fieldbyname('tst_codice').asstring := archivio.fieldbyname('tst_codice').asstring;
  tabella_01.fieldbyname('addebito_spese_fattura').asstring := archivio.fieldbyname('addebito_spese_fattura').asstring;
  tabella_01.fieldbyname('addebito_spese_trasporto').asstring := archivio.fieldbyname('spese_manuali_trasporto').asstring;

  read_tabella(arc.arcdit, 'tdo', 'codice', tabella_esa_01.fieldbyname('tdo_codice').asstring);
  tabella_01.fieldbyname('tdo_codice').asstring := archivio.fieldbyname('codice').asstring;
  tabella_01.fieldbyname('tipo_documento').asstring := archivio.fieldbyname('tipo_documento').asstring;
  tabella_01.fieldbyname('tco_codice').asstring := archivio.fieldbyname('tco_codice').asstring;

  read_tabella(arc.arcdit, 'art', 'codice', arc.dit.fieldbyname('art_codice_spese_trasporto').asstring);
  read_tabella(arc.arcdit, 'cpv', 'tcc_codice;tca_codice', vararrayof([tcc_codice, archivio.fieldbyname('tca_codice').asstring]));
  tabella_01.fieldbyname('gen_codice_trasporto').asstring := archivio.fieldbyname('gen_codice').asstring;

  read_tabella(arc.arcdit, 'art', 'codice', arc.dit.fieldbyname('art_codice_bollo').asstring);
  tabella_01.fieldbyname('tiv_codice_spese_bollo').asstring := archivio.fieldbyname('tiv_codice_vendite').asstring;
  read_tabella(arc.arcdit, 'cpv', 'tcc_codice;tca_codice', vararrayof([tcc_codice, archivio.fieldbyname('tca_codice').asstring]));
  tabella_01.fieldbyname('gen_codice_bollo').asstring := archivio.fieldbyname('gen_codice').asstring;

  read_tabella(arc.arcdit, 'art', 'codice', arc.dit.fieldbyname('art_codice_spese_incasso').asstring);
  tabella_01.fieldbyname('tiv_codice_spese_incasso').asstring := archivio.fieldbyname('tiv_codice_vendite').asstring;
  read_tabella(arc.arcdit, 'cpv', 'tcc_codice;tca_codice', vararrayof([tcc_codice, archivio.fieldbyname('tca_codice').asstring]));
  tabella_01.fieldbyname('gen_codice_incasso').asstring := archivio.fieldbyname('gen_codice').asstring;

  read_tabella(arc.arcdit, 'art', 'codice', arc.dit.fieldbyname('art_codice_spese_extra').asstring);
  tabella_01.fieldbyname('tiv_codice_spese_extra').asstring := archivio.fieldbyname('tiv_codice_vendite').asstring;
  read_tabella(arc.arcdit, 'cpv', 'tcc_codice;tca_codice', vararrayof([tcc_codice, archivio.fieldbyname('tca_codice').asstring]));
  tabella_01.fieldbyname('gen_codice_spese_extra').asstring := archivio.fieldbyname('gen_codice').asstring;

  tabella_01.fieldbyname('data_documento').asdatetime := tabella_esa_01.fieldbyname('dat_doc').asdatetime;
  tabella_01.fieldbyname('numero_documento').asInteger := tabella_esa_01.fieldbyname('des_num_doc').asInteger;
  tabella_01.fieldbyname('ese_codice').asstring := tabella_esa_01.fieldbyname('num_esercizio').asstring;

  tabella_01.fieldbyname('tma_codice').asstring := tabella_esa_01.fieldbyname('cod_dep_mov').asstring;
  if trim(tabella_esa_01.fieldbyname('cod_agente').asstring) = '' then
    tabella_01.fieldbyname('tag_codice').asstring := '0000'
  else
  begin

    tabella_01.fieldbyname('tag_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_agente').asstring);
    (*
      query.close;
      query.parambyname('esa_codice').asstring := trim(tabella_esa_01.fieldbyname('cod_agente').asstring);
      query.open;
      if not query.eof then
      begin
      tabella_01.fieldbyname('tag_codice').asstring := query.fieldbyname('codice').asstring;
      end
      else
      begin
      tabella_01.fieldbyname('tag_codice').asstring := '0000';
      end;
 *)
  end;

  tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_03.fieldbyname('cod_pag').asstring;
  if tabella_esa_03.fieldbyname('cod_banca').asstring <> '' then
  begin
    tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_03.fieldbyname('cod_banca').asstring;
    tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_03.fieldbyname('cod_agenzia').asstring;
    if tabella_esa_03.fieldbyname('sig_ccbanc').asstring <> '' then
    begin
      tabella_01.fieldbyname('conto_corrente').asstring := tabella_esa_03.fieldbyname('sig_ccbanc').asstring;
    end;

  end;
  if tabella_esa_01.fieldbyname('cod_listino').asstring <> '' then
  begin
    tabella_01.fieldbyname('tlv_codice').asstring := tabella_esa_01.fieldbyname('cod_listino').asstring;
  end;
  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('cod_valuta').asstring;
  tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cmb_valuta').asFloat;

  tabella_01.fieldbyname('riferimento').asstring := tabella_esa_03.fieldbyname('des_riferimento').asstring;
  tabella_01.fieldbyname('note').asstring := tabella_esa_03.fieldbyname('des_annotazioni').asstring;

  read_tabella(arc.arcdit, 'tlv', 'codice', tabella_01.fieldbyname('tlv_codice').asstring);
  tabella_01.fieldbyname('listino_con_iva').asstring := archivio.fieldbyname('iva_inclusa').asstring;

  if tabella_esa_01.fieldbyname('cdn_sede_div').asstring <> '' then
  begin
    tabella_01.fieldbyname('indirizzo').asstring := tabella_esa_01.fieldbyname('cdn_sede_div').asstring;

    read_tabella(arc.arcdit, 'ind', 'cli_codice;indirizzo', vararrayof([tabella_01.fieldbyname('cli_codice').asstring, tabella_01.fieldbyname('indirizzo').asstring]));
    tabella_01.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione1').asstring;
    tabella_01.fieldbyname('descrizione1').asstring := archivio.fieldbyname('descrizione2').asstring;
    tabella_01.fieldbyname('via').asstring := archivio.fieldbyname('via').asstring;
    tabella_01.fieldbyname('cap').asstring := archivio.fieldbyname('cap').asstring;
    tabella_01.fieldbyname('citta').asstring := archivio.fieldbyname('citta').asstring;
  end;

  tabella_01.fieldbyname('importo_bollo').asFloat := tabella_esa_03.fieldbyname('val_bollo_ese').asFloat + tabella_esa_03.fieldbyname('val_bollo_art15').asFloat;

  if tabella_esa_03.fieldbyname('val_spese_tra').asFloat > 0 then
  begin
    tabella_01.fieldbyname('spese_manuali').asstring := 'si';
    tabella_01.fieldbyname('spese_manuali_trasporto').asstring := 'si';
    tabella_01.fieldbyname('importo_spese_trasporto').asFloat := tabella_esa_03.fieldbyname('val_spese_tra').asFloat
  end;

  if tabella_esa_03.fieldbyname('val_spese_inc').asFloat > 0 then
  begin
    tabella_01.fieldbyname('spese_manuali').asstring := 'si';
    tabella_01.fieldbyname('spese_manuali_incasso').asstring := 'si';
    tabella_01.fieldbyname('importo_spese_incasso').asFloat := tabella_esa_03.fieldbyname('val_spese_inc').asFloat;
  end;
  tabella_01.fieldbyname('importo_spese_extra').asFloat := tabella_esa_03.fieldbyname('val_spese_acc').asFloat;

  tabella_01.fieldbyname('importo_totale_imponibile').asFloat := tabella_esa_03.fieldbyname('val_tot_imp').asFloat;
  tabella_01.fieldbyname('importo_totale_imponibile_euro').asFloat := tabella_esa_03.fieldbyname('val_tot_imp').asFloat;
  tabella_01.fieldbyname('importo_totale_iva').asFloat := tabella_esa_03.fieldbyname('val_tot_iva').asFloat;
  tabella_01.fieldbyname('importo_totale').asFloat := tabella_esa_03.fieldbyname('val_tot_doc').asFloat;
  tabella_01.fieldbyname('importo_totale_euro').asFloat := tabella_esa_03.fieldbyname('val_tot_doc').asFloat;

  tabella_01.post;

  riga := 0;
end;

procedure TCONVEREADY.crea_fvi;
begin

  tabella_esa_03.Close;
  tabella_esa_03.sql.clear;
  tabella_esa_03.sql.add('SELECT * FROM VN_DOCUMENTO_IVA');
  tabella_esa_03.sql.add('WHERE');
  tabella_esa_03.sql.add('prg_movimento=:prg_movimento ');
  tabella_esa_03.Parameters.ParamByName('prg_movimento').Value := tabella_esa_01.fieldbyname('prg_movimento').asInteger;
  tabella_esa_03.open;
  while not tabella_esa_03.eof do
  begin
    if (tabella_esa_03.fieldbyname('cod_tipoomag').asstring = '') or ((tabella_esa_03.fieldbyname('cod_tipoomag').asstring <> '') and (tabella_esa_03.fieldbyname('val_tot_imp').asFloat > 0)) then
    begin
      try
        tabella_03.append;

        tabella_03.fieldbyname('progressivo').asInteger := tabella_esa_03.fieldbyname('prg_movimento').asInteger;
        if tabella_esa_03.fieldbyname('cod_tipoomag').asstring <> '' then
        begin
          tabella_03.fieldbyname('tipo_movimento').asstring := 'normale';
        end;
        tabella_03.fieldbyname('tiv_codice').asstring := tabella_esa_03.fieldbyname('cod_iva').asstring;
        tabella_03.fieldbyname('importo').asFloat := tabella_esa_03.fieldbyname('val_tot_imp').asFloat;
        tabella_03.fieldbyname('importo_imponibile_trasporto').asFloat := tabella_esa_03.fieldbyname('val_netto_merce').asFloat + tabella_esa_03.fieldbyname('val_spese_tra').asFloat;
        tabella_03.fieldbyname('importo_sconto_percentuale').asFloat := 0;
        tabella_03.fieldbyname('importo_sconto').asFloat := 0;
        tabella_03.fieldbyname('importo_sconto_cassa').asFloat := 0;
        tabella_03.fieldbyname('importo_trasporto').asFloat := tabella_esa_03.fieldbyname('val_spese_tra').asFloat;
        tabella_03.fieldbyname('importo_bollo').asFloat := tabella_esa_03.fieldbyname('val_bollo_ese').asFloat + tabella_esa_03.fieldbyname('val_bollo_art15').asFloat;
        tabella_03.fieldbyname('importo_incasso').asFloat := tabella_esa_03.fieldbyname('val_spese_inc').asFloat;
        tabella_03.fieldbyname('importo_imponibile').asFloat := tabella_esa_03.fieldbyname('val_tot_imp').asFloat;
        tabella_03.fieldbyname('importo_iva').asFloat := tabella_esa_03.fieldbyname('val_tot_iva').asFloat;
        tabella_03.post;
      except
      end;
    end;
    tabella_esa_03.next;
  end;
end;

procedure TCONVEREADY.crea_fvr(progressivo: integer);
var
  codice_articolo, campo_codice: string;
  descrizione1, descrizione2: string;
begin
  if (arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si') then
  begin
    campo_codice := 'codice_alternativo';
  end
  else
  begin
    campo_codice := 'codice';
  end;

  tabella_esa_02.Close;
  tabella_esa_02.Parameters.ParamByName('prg_movimento').Value := progressivo;
  tabella_esa_02.open;
  while not tabella_esa_02.eof do
  begin

    codice_articolo := trim(tabella_esa_02.fieldbyname('cod_art').asstring);
    if (codice_articolo <> '-M') and (codice_articolo <> '-D') and (codice_articolo <> '') then
    begin
      esiste_articolo := true;
      if (not read_tabella(arc.arcdit, 'art', campo_codice, codice_articolo)) then
      begin
        esiste_articolo := false;
      end
      else
      begin
        codice_articolo := archivio.fieldbyname('codice').asstring;
      end;
    end
    else
    begin
      esiste_articolo := true;
    end;

    if esiste_articolo then
    begin
      tabella_02.append;

      tabella_02.fieldbyname('progressivo').asFloat := tabella_01.fieldbyname('progressivo').asFloat;
      tabella_02.fieldbyname('riga').asInteger := tabella_esa_02.fieldbyname('num_riga').asInteger;

      if (codice_articolo = '-M') or (codice_articolo = '-D') then
      begin
        tabella_02.fieldbyname('art_codice').asstring := '';
      end
      else if codice_articolo = '' then
      begin
        if (tabella_esa_02.fieldbyname('qta_merce').asFloat > 0) or (tabella_esa_02.fieldbyname('prz_merce').asFloat > 0) then
        begin
          tabella_02.fieldbyname('art_codice').asstring := '000000001';
        end;
      end
      else
      begin
        tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
      end;

      arc.spezza_descrizione(tabella_esa_02.fieldbyname('des_articolo_riga').asstring, descrizione1, descrizione2, 40);

      tabella_02.fieldbyname('descrizione1').asstring := descrizione1;
      tabella_02.fieldbyname('descrizione2').asstring := descrizione2;
      // tabella_02.fieldbyname('note').asstring := tabella_esa_02.fieldbyname('OR__NOTE').asstring;
      tabella_02.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('cod_cpartita').asstring;
      tabella_02.fieldbyname('tiv_codice').asstring := tabella_esa_02.fieldbyname('cod_iva').asstring;
      tabella_02.fieldbyname('quantita').asFloat := tabella_esa_02.fieldbyname('qta_merce').asFloat;
      tabella_02.fieldbyname('prezzo').asFloat := tabella_esa_02.fieldbyname('prz_merce').asFloat;
      tabella_02.fieldbyname('importo').asFloat := tabella_esa_02.fieldbyname('val_merce_doc').asFloat;
      tabella_02.fieldbyname('importo_euro').asFloat := tabella_esa_02.fieldbyname('val_merce').asFloat;

      tabella_02.fieldbyname('tma_codice').asstring := tabella_01.fieldbyname('tma_codice').asstring;
      if tabella_esa_02.fieldbyname('cod_tipoomag').asstring <> '' then
      begin
        tabella_02.fieldbyname('tipo_movimento').asstring := 'omaggio';
      end;

      if tabella_esa_02.fieldbyname('ind_stato_riga').asstring = 'E' then
        tabella_02.fieldbyname('situazione').asstring := 'evaso'
      else if tabella_esa_02.fieldbyname('ind_stato_riga').asstring = 'P' then
        tabella_02.fieldbyname('situazione').asstring := 'evaso parziale'
      else
        tabella_02.fieldbyname('situazione').asstring := 'inserito';

      tabella_02.post;
    end;

    tabella_esa_02.next;
  end; // while

end;

procedure TCONVEREADY.converti_ordini_fornitori;
var
  i: integer;
  test_serie, test_numero_doc: string;
  test_progressivo: integer;
begin
  v_tabella_01.caption := 'ordini fornitori';
  application.processmessages;

  cancella_tabella('oat');
  tabella_01.Close;
  tabella_01.tablename := 'oat';
  tabella_01.open;

  cancella_tabella('oat');
  tabella_02.Close;
  tabella_02.tablename := 'oar';
  tabella_02.open;

  tabella_01_ds.DataSet := tabella_01;
  tabella_esa_01.Close;
  tabella_esa_01.sql.clear;
  tabella_esa_01.sql.add('SELECT * FROM OR_ORDINIT T');
  tabella_esa_01.sql.add('WHERE');
  tabella_esa_01.sql.add('ORDER BY T.prg_ordine');
  tabella_esa_01.open;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('SELECT * FROM OR_ORDINIRR');
  tabella_esa_02.sql.add('WHERE');
  tabella_esa_02.sql.add('R.prg_ordine=:prg_ordine');
  tabella_esa_02.sql.add('ORDER BY R.prg_ordine_riga');
  tabella_esa_02.open;

  test_data := 0;
  test_alfa := '';
  test_numero := 0;
  test_progressivo := 0;

  test_serie := '';
  test_numero_doc := '';

  query.sql.clear;
  query.sql.add('select codice from tsm where sconto_maggiorazione = ''sconto''');
  query.sql.add('and percentuale_01 = :percentuale_01 and percentuale_02 = :percentuale_02 limit 1');

  while not tabella_esa_01.eof do
  begin
    crea_oat;

    test_serie := tabella_esa_01.fieldbyname('sig_serie_doc').asstring;
    test_numero_doc := tabella_esa_01.fieldbyname('num_doc').asstring;

    test_numero := tabella_esa_01.fieldbyname('prg_ordine').asInteger;
    if tabella_esa_02.locate('progressivo', test_progressivo, []) then
    begin
      crea_oar(test_progressivo);
    end;

    tabella_esa_01.next;
  end;

  tabella_01.Close;
  tabella_02.Close;
  tabella_03.Close;
  tabella_esa_01.Close;
  tabella_esa_02.Close;

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

procedure TCONVEREADY.crea_oat;
begin
  tabella_01.append;

  tabella_01.fieldbyname('tipo_documento').asstring := 'ordine';
  tabella_01.fieldbyname('tda_codice').asstring := 'ORDA';
  tabella_01.fieldbyname('data_documento').asdatetime := Converti_data(tabella_esa_01.fieldbyname('data_ordine').asstring);
  tabella_01.fieldbyname('serie_documento').asstring := tabella_esa_01.fieldbyname('serie_documento').asstring;
  tabella_01.fieldbyname('numero_documento').asInteger := strtoint(trim(tabella_esa_01.fieldbyname('numero_documento').asstring));

  assegna_codice_frn(tabella_esa_01.fieldbyname('Cod_clifor_Contropar').asstring, cli_for);

  tabella_01.fieldbyname('frn_codice').asstring := cli_for;

  tabella_01.fieldbyname('tma_codice').asstring := '000'; // tabella_esa_01.fieldbyname('ortesmag').asstring;
  // tabella_01.fieldbyname('tag_codice').asstring := tabella_esa_01.fieldbyname('Codice_Agente').asstring;
  tabella_01.fieldbyname('tpa_codice').asstring := tabella_esa_01.fieldbyname('Codice_Pagamento').asstring;
  // tabella_01.fieldbyname('codice_abi').asstring := tabella_esa_01.fieldbyname('Codice_Banca').asstring;
  // tabella_01.fieldbyname('codice_cab').asstring := tabella_esa_01.fieldbyname('Codice_Agenzia').asstring;

  tabella_01.fieldbyname('tla_codice').asstring := tabella_esa_01.fieldbyname('num_listino').asstring;
  tabella_01.fieldbyname('tva_codice').asstring := tabella_esa_01.fieldbyname('codice_valuta').asstring;
  tabella_01.fieldbyname('cambio').asFloat := tabella_esa_01.fieldbyname('cambio').asFloat;
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

procedure TCONVEREADY.crea_oar(progressivo: integer);
var
  codice_articolo: string;
begin

  while not(tabella_esa_02.eof) do
  begin

    codice_articolo := trim(tabella_esa_02.fieldbyname('Codice_Articolo').asstring);

    if (codice_articolo <> '-M') and (codice_articolo <> '-D') then
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

      tabella_02.fieldbyname('progressivo').asFloat := tabella_01.fieldbyname('progressivo').asFloat;

      tabella_02.fieldbyname('riga').asInteger := tabella_esa_02.fieldbyname('Numero_riga').asInteger;

      if (codice_articolo = '-M') or (codice_articolo = '-D') then
      begin
        tabella_02.fieldbyname('art_codice').asstring := '';
      end
      else
      begin
        if arc.dit.fieldbyname('codice_articolo_numerico').asstring = 'si' then
        begin
          read_tabella(arc.arcdit, 'art', 'codice_alternativo', codice_articolo);
          tabella_02.fieldbyname('art_codice').asstring := archivio.fieldbyname('codice').asstring;
        end
        else
        begin
          tabella_02.fieldbyname('art_codice').asstring := codice_articolo;
        end;
      end;

      tabella_02.fieldbyname('descrizione1').asstring := tabella_esa_02.fieldbyname('Descrizione_Articolo').asstring;
      tabella_02.fieldbyname('descrizione2').asstring := tabella_esa_02.fieldbyname('Descrizione2Articolo').asstring;
      // tabella_02.fieldbyname('note').asstring := tabella_esa_02.fieldbyname('OR__NOTE').asstring;
      tabella_02.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('Contropartita_Ricavo').asstring;
      tabella_02.fieldbyname('tiv_codice').asstring := tabella_esa_02.fieldbyname('Codice_Iva').asstring;

      tabella_02.fieldbyname('quantita').asFloat := arrotonda(tabella_esa_02.fieldbyname('qta_movimento').asFloat);

      tabella_02.fieldbyname('prezzo').asFloat := tabella_esa_02.fieldbyname('Prezzo_Unitario').asFloat;
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

    tabella_esa_02.next;
  end; // while

end;

procedure TCONVEREADY.calcola_importo;
var
  tiv_codice: string;
  imponibile: double;
begin
  if not((tabella_02.fieldbyname('quantita').asFloat = 0) or (tabella_02.fieldbyname('prezzo').asFloat = 0)) then
  begin
    tabella_02.fieldbyname('importo').asFloat := arrotonda(tabella_02.fieldbyname('quantita').asFloat * tabella_02.fieldbyname('prezzo').asFloat * sconto(tabella_02.fieldbyname('tsm_codice').asstring) / 100);
  end;
  tabella_02.fieldbyname('importo_euro').asFloat := arrotonda(tabella_02.fieldbyname('importo').asFloat / tabella_01.fieldbyname('cambio').asFloat);

  tiv_codice := tabella_02.fieldbyname('tiv_codice').asstring;
  if read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring) then
  begin
    if tabella_01.fieldbyname('listino_con_iva').asstring = 'no' then
    begin
      tabella_02.fieldbyname('importo_iva').asFloat := arrotonda(tabella_02.fieldbyname('importo').asFloat * archivio.fieldbyname('percentuale').asFloat / 100);
      read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring);
      tabella_02.fieldbyname('importo_iva_euro').asFloat := arrotonda(tabella_02.fieldbyname('importo_iva').asFloat / tabella_01.fieldbyname('cambio').asFloat);
    end
    else
    begin
      imponibile := arrotonda(tabella_02.fieldbyname('importo').asFloat / (1 + archivio.fieldbyname('percentuale').asFloat / 100));
      tabella_02.fieldbyname('importo_iva').asFloat := arrotonda(tabella_02.fieldbyname('importo').asFloat - imponibile);
      read_tabella(arc.arcdit, 'tiv', 'codice', tabella_02.fieldbyname('tiv_codice').asstring);
      tabella_02.fieldbyname('importo_iva_euro').asFloat := arrotonda(tabella_02.fieldbyname('importo_iva').asFloat / tabella_01.fieldbyname('cambio').asFloat);
    end;
  end;
end;

procedure TCONVEREADY.assegna_codice_cli(codice_adhoc: string;

  var cli_for: string);
begin
  codice_adhoc := trim(codice_adhoc);
  if v_codifica_clienti.checked then
  begin
    if (codice_nom_numerico = 'si') then
    begin
      if v_codifica_clienti.checked then
      begin
        if trim(codice_adhoc) <> '' then
        begin
          try
            cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);
          except
            cli_for := '';
          end;
        end;
      end
      else
      begin
        cli_for := '';
      end;
    end
    else
    begin
      cli_for := trim(codice_adhoc);
    end;
  end
  else
  begin
    // cli_for := 'C' + codice_adhoc;
    cli_for := trim(codice_adhoc);
  end;
end;

procedure TCONVEREADY.assegna_codice_frn(codice_adhoc: string;

  var cli_for: string);
begin
  codice_adhoc := trim(codice_adhoc);
  if v_codifica_fornitori.checked then
  begin
    if codice_nom_numerico = 'si' then
    begin

      if trim(codice_adhoc) <> '' then
      begin
        try
          cli_for := setta_lunghezza(strtoint(codice_adhoc), 8, 0);
        except
          cli_for := '';
        end;
      end;

    end
    else
    begin
      cli_for := trim(codice_adhoc);
      if codice_nom_numerico = 'si' then
      begin

        if trim(codice_adhoc) <> '' then
          cli_for := setta_lunghezza(strtoint(codice_adhoc), 6, 0); // codice_adhoc;
      end;
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

procedure TCONVEREADY.FormClose(Sender: TObject;

  var Action: TCloseAction);
begin
  inherited;
  ADOEsatto.Connected := false;
end;

procedure TCONVEREADY.FormShow(Sender: TObject);
begin
  inherited;
  ADOEsatto.Connected := false;
  ADOEsatto.Connected := true;
end;

function TCONVEREADY.Converti_data(data: string): TDateTime;
begin
  try
    if length(data) = 8 then
      Result := StrToDate(copy(data, 7, 2) + FormatSettings.DateSeparator + copy(data, 5, 2) + FormatSettings.DateSeparator + copy(data, 1, 4))
    else
      Result := StrToDate('01/01/1900');
  except
    if copy(data, 5, 2) <> '2' then
    begin
      Result := StrToDate('30' + FormatSettings.DateSeparator + copy(data, 5, 2) + FormatSettings.DateSeparator + copy(data, 1, 4))
    end
    else
    begin
      Result := StrToDate('28' + FormatSettings.DateSeparator + copy(data, 5, 2) + FormatSettings.DateSeparator + copy(data, 1, 4))
    end;

  end;
end;

procedure TCONVEREADY.crea_tsm(sconto1, sconto2: double);
var
  ultimo_id: integer;
begin
  tsm.last;
  tsm_codice := tsm.fieldbyname('id').asInteger + 1;

  tsm.append;
  tsm.fieldbyname('codice').asstring := setta_lunghezza(tsm_codice, 4, 0);
  tsm.fieldbyname('descrizione').asstring := 'SCONTO ' + floattostr(sconto1);
  if sconto2 <> 0 then
  begin
    tsm.fieldbyname('descrizione').asstring := tsm.fieldbyname('descrizione').asstring + ' + ' + floattostr(sconto2);
  end;
  tsm.fieldbyname('sconto_maggiorazione').asstring := 'sconto';
  tsm.fieldbyname('percentuale_01').asFloat := sconto1;
  tsm.fieldbyname('percentuale_02').asFloat := sconto2;

  tsm.fieldbyname('percentuale_totale').asFloat := 100;
  if sconto2 <> 0 then
  begin
    tsm.fieldbyname('percentuale_totale').asFloat := (100 - sconto1) * (100 - sconto2) / 100;
  end
  else
  begin
    if sconto1 <> 0 then
    begin
      tsm.fieldbyname('percentuale_totale').asFloat := (100 - sconto1) / 1;
    end;
  end;

  tsm.post;
end;

procedure TCONVEREADY.converti_provvigioni;
begin
  v_tabella_01.caption := 'provvigioni agenti';
  application.processmessages;

  (*
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
 *)
end;

procedure TCONVEREADY.cancella_tabella(tabella: string);
begin
  query.sql.clear;
  query.sql.add('delete from ' + tabella);
  query.ExecSQL;

end;

procedure TCONVEREADY.crea_cpv;
var
  progressivo: word;
begin
  query.sql.clear;
  query.sql.add('delete from tca');
  query.sql.add('where codice <>''0'' ');
  query.ExecSQL;

  tca.open;

  query.Close;
  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('select distinct cod_cpart_ric');
  tabella_esa_02.sql.add('from MG_ARTBASE');
  tabella_esa_02.sql.add('where cod_cpart_ric is not null');
  tabella_esa_02.open;

  while not tabella_esa_02.eof do
  begin

    progressivo := progressivo + 1;

    tca.append;
    tca.fieldbyname('codice').asstring := setta_lunghezza(inttostr(progressivo), 4, true, '0');
    tca.fieldbyname('descrizione').asstring := 'CAT.CONTABILE DA ESA - RICAVO ' + tabella_esa_02.fieldbyname('cod_cpart_ric').asstring;
    tca.post;

    cpv.append;
    cpv.fieldbyname('tcc_codice').asstring := arc.dit.fieldbyname('tcc_codice_cli').asstring;
    cpv.fieldbyname('tca_codice').asstring := tca.fieldbyname('codice').asstring;
    cpv.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('cod_cpart_ric').asstring;
    cpv.fieldbyname('gen_codice_omaggi').asstring := tabella_esa_02.fieldbyname('cod_cpart_ric').asstring;
    cpv.fieldbyname('gen_codice_sconti').asstring := tabella_esa_02.fieldbyname('cod_cpart_ric').asstring;
    cpv.post;

    tabella_esa_02.next;
  end;

  query.Close;
  tca.Close;
end;

procedure TCONVEREADY.crea_cpa;
var
  progressivo: word;
begin
  query.sql.clear;
  query.sql.add('delete from taq');
  query.sql.add('where codice <>''0'' ');
  query.ExecSQL;

  taq.open;

  tabella_esa_02.Close;
  tabella_esa_02.sql.clear;
  tabella_esa_02.sql.add('select distinct cod_cpart_cos');
  tabella_esa_02.sql.add('from MG_ARTBASE');
  tabella_esa_02.sql.add('where cod_cpart_cos is not null');
  tabella_esa_02.open;

  while not tabella_esa_02.eof do
  begin

    progressivo := progressivo + 1;

    taq.append;
    taq.fieldbyname('codice').asstring := setta_lunghezza(inttostr(progressivo), 4, true, '0');
    taq.fieldbyname('descrizione').asstring := 'CAT.CONTABILE DA ESA - ACQUISTI ' + tabella_esa_02.fieldbyname('cod_cpart_cos').asstring;
    taq.post;

    cpa.append;
    cpa.fieldbyname('tcf_codice').asstring := arc.dit.fieldbyname('tcf_codice_frn').asstring;
    cpa.fieldbyname('taq_codice').asstring := taq.fieldbyname('codice').asstring;
    cpa.fieldbyname('gen_codice').asstring := tabella_esa_02.fieldbyname('cod_cpart_cos').asstring;
    cpa.fieldbyname('gen_codice_omaggi').asstring := tabella_esa_02.fieldbyname('cod_cpart_cos').asstring;
    cpa.fieldbyname('gen_codice_sconti').asstring := tabella_esa_02.fieldbyname('cod_cpart_cos').asstring;
    cpa.post;

    tabella_esa_02.next;
  end;

  tabella_esa_02.Close;
  tca.Close;
end;

initialization

registerclass(TCONVEREADY);

end.

  / * * * * * * Script per comando SelectTopNRows da SSMS * * * * * * / SELECT cp.prg_prinota, t.* FROM MG_MOVMAGT t left join[AZI_ENI].[dbo].CO_DOCUMENTI co on co.num_esercizio = t.num_esercizio and co.dat_doc = t.dat_doc and
  co.des_num_doc = t.des_num_doc left join[AZI_ENI].[dbo].CO_DOCUMENTI_PN cp on cp.prg_documento = co.prg_documento where t.cod_gruppo_tipodoc = 'FAV'
