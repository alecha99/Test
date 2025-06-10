unit mFileView;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.FileCtrl, Vcl.StdCtrls;

type
  TfFileView = class(TForm)
    pFileIntput: TPanel;
    cbDrive: TDriveComboBox;
    DirectoryList: TDirectoryListBox;
    FileList: TFileListBox;
    splFileDir: TSplitter;
    splView: TSplitter;
    pcView: TPageControl;
    tshTxt: TTabSheet;
    tshGraf: TTabSheet;
    lText: TListBox;
    Glyph: TImage;
    procedure FileListClick(Sender: TObject);
    procedure DirectoryListClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fFileView: TfFileView;

implementation
 uses System.StrUtils, Vcl.Imaging.jpeg, Vcl.Imaging.pngimage;
{$R *.dfm}

procedure TfFileView.DirectoryListClick(Sender: TObject);
begin
  FileListClick(nil);
end;

procedure TfFileView.FileListClick(Sender: TObject);
var
  ext: string;
begin
  if(FileList.ItemIndex < 0)then
    pcView.ActivePageIndex := -1
  else
    if MatchText(ExtractFileExt(FileList.FileName), ['.txt','.pas','.ini','.xml', '.bat','*.sql']) then
    begin
      pcView.ActivePageIndex := 0;
      lText.Items.LoadFromFile(FileList.FileName);
    end
    else
    begin
      if MatchText(ExtractFileExt(FileList.FileName), ['.bmp','.jpg','.png']) then
      begin
        pcView.ActivePageIndex := 1;
        ext := ExtractFileExt(FileList.FileName);
        Glyph.Picture.LoadFromFile(FileList.FileName);
      end
      else
        pcView.ActivePageIndex := -1;
    end;
  tshGraf.TabVisible := pcView.ActivePageIndex = 1;
  tshTxt.TabVisible := pcView.ActivePageIndex = 0;
end;

end.
