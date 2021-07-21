//
//  Versione 11.01.02
//
program GO_CNVEREADY;







uses
  Forms,
  ZZACCESSO in '..\..\..\go_1100\src\ZZACCESSO.pas',
  DMARC in 'DMARC.pas' {ARC: TDataModule},
  GGLOGIN in '..\..\..\go_1100\src\GGLOGIN.pas' {LOGIN},
  GGBASE in '..\..\..\go_1100\src\GGBASE.pas' {BASE},
  GGFORMBASE in '..\..\..\go_1100\src\GGFORMBASE.pas' {FORMBASE},
  GGGESARC in '..\..\..\go_1100\src\GGGESARC.pas' {GESARC},
  GGELABORA in '..\..\..\go_1100\src\GGELABORA.pas' {ELABORA},
  GGSTAMPA in '..\..\..\go_1100\src\GGSTAMPA.pas' {STAMPA},
  GGMAIL in '..\..\..\go_1100\src\GGMAIL.pas' {MAIL},
  ZZARROTONDAMENTO in '..\..\..\go_1100\src\ZZARROTONDAMENTO.pas',
  DMCLOUD in 'DMCLOUD.pas' {CLOUD: TDataModule},
  FRMDITTA in 'FRMDITTA.pas' {DITTA: TFrame},
  GGMENU in 'GGMENU.pas' {MENUGG},
  GGCONVEREADY in 'GGCONVEREADY.pas' {CONVEREADY};

{$r *.res}


begin
  Application.Initialize;
  Application.Initialize;
  Application.Title := 'Programma collegato';
  Application.CreateForm(TCLOUD, CLOUD);
  Application.CreateForm(TARC, ARC);
  Application.CreateForm(TLOGIN, LOGIN);
  Application.Run;

end.
