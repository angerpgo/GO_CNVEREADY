unit GGLOGIN;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, ExtCtrls, DB, adscnnct, adsset,
  Mask, ToolWin, ppForms, idGlobalProtocols, DBCtrls, jpeg, shellapi,
  query_go, MyAccess, StrUtils, IniFiles, RzButton, Comobj, rzLabel, RzPanel,
  RzRadChk, raizeedit_go, MemDS, DBAccess, RzEdit, ZZACCESSO, RzSpnEdt;

type

  TLOGIN = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TRzlabel;
    Label2: TRzlabel;
    v_password: TRzEdit_go;
    v_utente: TRzEdit_go;
    Panel1: TRzpanel;
    Image1: TImage;
    v_conferma: TRzBitBtn;
    v_ricorda_utente: TRzCheckBox;
    v_esci: TRzBitBtn;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormCreate(Sender: TObject);
    procedure Image1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure v_confermaClick(Sender: TObject);
    procedure v_esciClick(Sender: TObject);
  protected
    accesso: taccesso;
  public
  end;

var
  LOGIN: TLOGIN;

implementation

{$r *.dfm}


procedure TLOGIN.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (key = vk_return) then
  begin
    key := 0;

    v_conferma.setfocus;
    v_confermaclick(v_conferma);
  end;

  if (key = vk_escape) and (shift = []) then
  begin
    close;
  end;
end;

procedure TLOGIN.FormShow(Sender: TObject);
begin

  //assegna nome programma solo su LOGIN e MENU
  //per gli altri programmi legge variabili precaricate
  self.hint := 'programma LOGIN : identificazione utente';
  self.caption := 'LOGIN   versione ' + accesso.versione;

  //carica immagine login
  if accesso.bitmap_login_go <> '' then
  begin
    image1.picture.loadfromfile(accesso.bitmap_login_go);
  end;

  //verifica presenza parametri
  v_ricorda_utente.checked := accesso.assegna_parametri_programma;

  v_utente.text := accesso.utente_os;
  v_password.text := accesso.password;

  if accesso.salta_login then
  begin
    v_confermaclick(v_conferma);
  end;
end;

procedure TLOGIN.v_confermaClick(Sender: TObject);
begin
  if accesso.verifica_accesso(v_utente.text, v_password.text) then
  begin
    //  assegna variabili utente globali
    accesso.assegna_var_utente(v_utente.text);

    //  aggiorna data ultimo accesso e stringhe utilizzo utente
    accesso.accesso_utente;

    //  crea utente database se necessario
    accesso.crea_user_mysql(v_utente.text);

    //  aggiorna go.cfg per attivare la spunta di recupero ultimo utente
    accesso.ricorda_ultimo_utente(v_ricorda_utente.checked);

    //  chiama programma menu
    accesso.attiva_screen_saver(true);
    try
      hide;
    except
    end;
    accesso.chiama_programma_successivo;
    accesso.attiva_screen_saver(false);

    close;
  end
  else
  begin
    accesso.mostra_risposta;
    v_utente.setfocus;
  end;
end;

procedure TLOGIN.v_esciClick(Sender: TObject);
begin
  close;
end;

procedure TLOGIN.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  freeandnil(accesso);

  //libera memoria utilizzata dalla form
  action := cafree;
end;

procedure TLOGIN.FormCreate(Sender: TObject);
begin
  accesso := taccesso.create;
end;

procedure TLOGIN.Image1Click(Sender: TObject);
begin
  shellexecute(application.handle, pchar('open'), pchar(accesso.sito_web_go), nil, nil, SW_SHOWNORMAL);
end;

end.
