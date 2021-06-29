unit ZZCALLPRG;

interface

procedure assegna_programma_personalizzato(programma_da_eseguire: string; var codice_archivio: variant;
  var programma_personalizzato: boolean);

implementation

uses variants;

procedure assegna_programma_personalizzato(programma_da_eseguire: string; var codice_archivio: variant;

  var programma_personalizzato: boolean);
begin
  programma_personalizzato := false;
end;

end.
