unit GGMAIN;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DB, MyAccess, rzLabel, RzPanel, RzDBEdit, RzListVw, RzTreeVw, RzDBChk,
  RzRadChk, RzButton, RzSplit, RzCmboBx, RzPrgres,
  RzSpnEdt, RzShellDialogs, RzDBCmbo, raizeedit_go, MemDS, DBAccess, RzEdit;

type

  Tmain = class(TForm)
    ComboEdit1: TRzEdit_go;
    ComboEdit2: TRzEdit_go;
    ComboEdit3: TRzEdit_go;
    ComboEdit4: TRzEdit_go;
    ComboEdit5: TRzEdit_go;
    ComboEdit6: TRzEdit_go;
    ComboEdit7: TRzEdit_go;
    Label1: TRzlabel;
    Label2: TRzlabel;
    Label3: TRzlabel;
    ggg: tmyquery;
    dit: tmyquery;
    Label4: TRzlabel;
    Label5: TRzlabel;
    Label6: TRzlabel;
    Label7: TRzlabel;
    v_esegui_go: Trzbutton;
    GroupBox2: TGroupBox;
    v_cnvesa: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure v_cnvesaClick(Sender: TObject);
    procedure v_esegui_goClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    attesa: integer;
  end;

var
  main: Tmain;

implementation

{$r *.dfm}


uses DMARC;

procedure Tmain.FormCreate(Sender: TObject);
var
  i: integer;
  numero_serie_go, numero_serie: string;
begin
  Self.caption := 'CONVERSIONE eReady - GO 9.02.08 DEL 07/03/2015';

  // --------------------------------------------------------------------------------------------------
  // routine di assegnazione ambiente di lavoro
  // usa l'utente passato come parametro dal comando di esecuzione (es. ...\collegato.exe GO
  // --------------------------------------------------------------------------------------------------
  // --------------------------------------------------------------------------------------------------
  // routine di assegnazione ambiente di lavoro
  // usa l'utente passato come parametro dal comando di esecuzione (es. ...\collegato.exe GO
  // --------------------------------------------------------------------------------------------------
  attesa := 0;
  for i := 1 to system.paramcount do
  begin

    case i of
      1:
        begin
          utente := paramstr(i);
        end;
      2:
        begin
          attesa := strtoint(paramstr(i)) * 1000;
        end;

    end; // for

  end;

  utente := strtran(utente, '<utente>', '', true);
  utente := strtran(utente, '</utente>', '', true);
  utente := trim(utente);

  utente_passato_programma_collegato := utente;

  (*
    if attesa > 0 then
    begin
    timer1.interval := attesa;
    timer1.enabled := true;

    self.windowstate := wsminimized;
    end;
  *)


  arc.assegna_ambiente_lavoro_programma_collegato;
  // --------------------------------------------------------------------------------------------------
  // fine routine di assegnazione ambiente di lavoro
  // --------------------------------------------------------------------------------------------------

  inherited;

  // controllo se variabili ambiente settate correttamente
  ComboEdit1.text := ditta;
  ComboEdit2.text := esercizio;
  ComboEdit3.text := utente;
  ComboEdit4.text := datetostr(data_inizio);
  ComboEdit5.text := datetostr(data_fine);
  ComboEdit6.text := arc.dit.fieldbyname('descrizione1').asstring;
  ComboEdit7.text := arc.utn.fieldbyname('descrizione').asstring;

  // controllo se utente abilitato
  (*
    dit.close;
    dit.open;
    if dit['data_ora_creazione'] <> null then
    begin
    numero_serie := uppercase(md5print(md5string(dit.fieldbyname('data_ora_creazione').asstring)));
    numero_serie := copy(numero_serie, 1, 4) + '-' + copy(numero_serie, 5, 4) + '-' +
    copy(numero_serie, 9, 4) + '-' + copy(numero_serie, 13, 4);

    ggg.close;
    ggg.open;
    if ggg.fieldbyname('serial_number').asstring <> numero_serie then
    begin
    messaggio(000, 'installazione non autorizzata ad eseguire il programma');
    application.terminate;
    end;

    numero_serie_go := copy(uppercase(MD5Print(MD5String(ggg.fieldbyname('serial_number').asstring + 'GO'))), 1, 20);
    if ggg.fieldbyname('codice_attivazione').asstring <> numero_serie_go then
    begin
    messaggio(000, 'installazione non autorizzata ad eseguire il programma');
    application.terminate;
    end;
    ggg.close;
    end;
  *)

  if now > data_fine then
  begin
    messaggio(200, 'la data attuale è superiore alla data fine esercizio utilizzato');
  end;
end;

procedure Tmain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  if (key = vk_escape) and (shift = []) then
  begin
    close;
  end;
end;

procedure Tmain.v_esegui_goClick(Sender: TObject);
begin
  esegui_effettivo('.\GO.exe', '');
end;

procedure Tmain.v_cnvesaClick(Sender: TObject);
begin
  esegui_programma('CONVEREADY', Vararrayof(['', '']), TRUE);
end;

end.
