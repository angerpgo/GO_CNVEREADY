unit GGMENU;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, DBAccess, MyAccess, query_go, GGBASE,
  RzTabs, RzButton, Vcl.StdCtrls, Vcl.Mask, RzEdit, raizeedit_go, RzLabel,
  FRMDITTA, Vcl.ExtCtrls, RzPanel, RzCmboBx, Vcl.Buttons, RzSpnEdt,
  RzShellDialogs, Vcl.Grids, Vcl.DBGrids, RzDBGrid, Vcl.ComCtrls, MemDS;

type
  TMENUGG = class(TForm)
    Label1: TRzLabel;
    ComboEdit1: trzedit_go;
    ComboEdit6: trzedit_go;
    Label6: TRzLabel;
    Label2: TRzLabel;
    ComboEdit2: trzedit_go;
    ComboEdit4: trzedit_go;
    Label4: TRzLabel;
    ComboEdit5: trzedit_go;
    Label5: TRzLabel;
    query: TMyQuery_go;
    ComboEdit7: trzedit_go;
    Label7: TRzLabel;
    ComboEdit3: trzedit_go;
    Label3: TRzLabel;
    RzGroupBox2: TRzGroupBox;
    v_esporta_anagrafiche: TButton;
    Panel4: TRzPanel;
    Bevel1: TBevel;
    v_conferma: TRzBitBtn;
    v_esci: TRzBitBtn;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure v_confermaClick(Sender: TObject);
    procedure v_esciClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure v_esporta_anagraficheClick(Sender: TObject);
  protected
    schedulato: boolean;
  public
    programma_preferito: string;
    procedure taskbar_aggiungi_programma(nome_programma, descrizione_programma: string; nome_form: tbase);
    procedure taskbar_rimuovi_programma(nome_programma: string);
    procedure taskbar_attivazione(nome: string; attivazione: boolean);
    procedure gestione_programmi_preferiti;
  end;

var
  MENUGG: TMENUGG;

implementation

{$r *.dfm}


uses DMCLOUD, GGMESSAGGIO, DMARC;

procedure TMENUGG.FormCreate(Sender: TObject);
begin
  schedulato := false;
  assegna_parametri_lavoro;
end;

procedure TMENUGG.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = vk_escape) and (shift = []) then
  begin
    close;
  end;
end;

procedure TMENUGG.FormShow(Sender: TObject);
begin
  (*
    if cloud.arc_standard.connected = false then
    begin
    application.terminate;
    end;
  *)

  Self.caption := self.Caption + ' [Database ' + arc.arcdit.server + ' ]';
  comboedit1.text := ditta;
  comboedit2.text := esercizio;
  comboedit3.text := utente;
  comboedit4.text := datetostr(data_inizio);
  comboedit5.text := datetostr(data_fine);
  comboedit6.text := arc.dit.fieldbyname('descrizione1').asstring;
  comboedit7.text := arc.utn.fieldbyname('descrizione').asstring;

  if schedulato then
  begin
    close;
  end;
end;

procedure TMENUGG.taskbar_aggiungi_programma(nome_programma, descrizione_programma: string; nome_form: tbase);
begin
  //  presente perché referenziato da TFORMBASE
end;

procedure TMENUGG.taskbar_attivazione(nome: string; attivazione: boolean);
begin
  //  presente perché referenziato da TFORMBASE
end;

procedure TMENUGG.taskbar_rimuovi_programma(nome_programma: string);
begin
  //  presente perché referenziato da TFORMBASE
end;

procedure TMENUGG.v_confermaClick(Sender: TObject);
begin
  v_conferma.enabled := false;

  v_conferma.enabled := true;
end;

procedure TMENUGG.v_esciClick(Sender: TObject);
begin
  Close;
end;

procedure TMENUGG.v_esporta_anagraficheClick(Sender: TObject);
begin
  esegui_programma('CONVEREADY', '', true);
end;

procedure TMENUGG.gestione_programmi_preferiti;
begin
  //
end;

initialization

registerclass(tMENUGG);

end.
