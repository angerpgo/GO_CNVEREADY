unit FRMDITTA;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Mask,
  RzEdit, raizeedit_go, RzLabel;

type
  TDITTA = class(TFrame)
    RzLabel1: TRzLabel;
    v_codice: trzedit_go;
    v_descrizione1: trzedit_go;
    v_descrizione2: trzedit_go;
    Label1: TRzLabel;
    Label2: TRzLabel;
    Label3: TRzLabel;
    Label4: TRzLabel;
    Label5: TRzLabel;
    Label6: TRzLabel;
    Label7: TRzLabel;
    Label8: TRzLabel;
    Label9: TRzLabel;
    v_via: trzedit_go;
    v_cap: trzedit_go;
    v_localita: trzedit_go;
    v_provincia: trzedit_go;
    v_partita_iva: trzedit_go;
    v_codice_fiscale: trzedit_go;
    v_telefono: trzedit_go;
  private
  public
    function controllo_campi: boolean;
  end;

implementation

{$R *.dfm}

function TDITTA.controllo_campi: boolean;
begin
  result := true;

  if trim(v_codice.text) = '' then
  begin
    result := false;
  end;

  if trim(v_descrizione1.text) = '' then
  begin
    result := false;
  end;

  if trim(v_descrizione2.text) = '' then
  begin
//    result := false;
  end;

  if trim(v_via.text) = '' then
  begin
    result := false;
  end;

  if trim(v_cap.text) = '' then
  begin
    result := false;
  end;

  if trim(v_localita.text) = '' then
  begin
    result := false;
  end;

  if trim(v_provincia.text) = '' then
  begin
    result := false;
  end;

  if trim(v_partita_iva.text) = '' then
  begin
    result := false;
  end;

  if trim(v_codice_fiscale.text) = '' then
  begin
    result := false;
  end;

  if result = false then
  begin
    showmessage('tutti i campi della ditta devono essere compilati');
  end;
end;

end.
