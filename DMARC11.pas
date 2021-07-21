unit DMARC11;

interface

uses
  System.SysUtils, System.Classes, System.ImageList, Vcl.ImgList, Vcl.Controls,
  cxGraphics;

type
  Tarc11 = class(TDataModule)
    immagine_16: TcxImageList;
    immagine_24: TcxImageList;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  arc11: Tarc11;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

end.
