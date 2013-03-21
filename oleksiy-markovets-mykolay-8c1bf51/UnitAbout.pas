unit UnitAbout;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, jpeg, ExtCtrls;

type
  TFormAbout = class(TForm)
    OkBitBtn: TBitBtn;
    Image1: TImage;
    Label3: TLabel;
    Label4: TLabel;
    procedure Label4Click(Sender: TObject);
    procedure OkBitBtnClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormAbout: TFormAbout;

implementation

{$R *.dfm}

procedure TFormAbout.Label4Click(Sender: TObject);
begin
//  ShellExecute(0, 0,"\"mailto:OliaMonka@ukr.net\"",0,0,SW_SHOWNORMAL);
//  WinExec('hh.exe Material\help.chm',SW_SHOW);
end;

procedure TFormAbout.OkBitBtnClick(Sender: TObject);
begin
Close;
end;

end.
