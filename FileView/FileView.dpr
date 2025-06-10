program FileView;

uses
  Vcl.Forms,
  mFileView in 'mFileView.pas' {fFileView};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TfFileView, fFileView);
  Application.Run;
end.
