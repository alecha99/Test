object fFileView: TfFileView
  Left = 0
  Top = 0
  Caption = #1044#1077#1084#1086#1085#1089#1090#1088#1072#1094#1080#1103' FileView'
  ClientHeight = 806
  ClientWidth = 1197
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poOwnerFormCenter
  TextHeight = 15
  object splView: TSplitter
    Left = 401
    Top = 0
    Width = 6
    Height = 806
  end
  object pFileIntput: TPanel
    Left = 0
    Top = 0
    Width = 401
    Height = 806
    Align = alLeft
    Caption = 'pFileIntput'
    TabOrder = 0
    object splFileDir: TSplitter
      Left = 1
      Top = 344
      Width = 399
      Height = 12
      Cursor = crVSplit
      Align = alBottom
    end
    object cbDrive: TDriveComboBox
      Left = 1
      Top = 1
      Width = 399
      Height = 21
      Align = alTop
      DirList = DirectoryList
      TabOrder = 0
      ExplicitLeft = 49
      ExplicitTop = -1
    end
    object DirectoryList: TDirectoryListBox
      Left = 1
      Top = 22
      Width = 399
      Height = 322
      Align = alClient
      FileList = FileList
      TabOrder = 1
      OnClick = DirectoryListClick
      ExplicitLeft = -1
      ExplicitTop = 16
    end
    object FileList: TFileListBox
      Left = 1
      Top = 356
      Width = 399
      Height = 449
      Align = alBottom
      ItemHeight = 15
      Mask = '*.txt;*.pas;*.ini;*.xml;*.bat;*.sql;*.bmp;*.jpg;*.png;'
      TabOrder = 2
      OnClick = FileListClick
    end
  end
  object pcView: TPageControl
    Left = 407
    Top = 0
    Width = 790
    Height = 806
    ActivePage = tshGraf
    Align = alClient
    TabOrder = 1
    ExplicitLeft = 404
    ExplicitWidth = 793
    object tshTxt: TTabSheet
      Caption = #1058#1077#1082#1089#1090#1086#1074#1099#1081
      object lText: TListBox
        Left = 0
        Top = 0
        Width = 782
        Height = 776
        Align = alClient
        ItemHeight = 15
        TabOrder = 0
        ExplicitLeft = 184
        ExplicitTop = 240
        ExplicitWidth = 121
        ExplicitHeight = 97
      end
    end
    object tshGraf: TTabSheet
      Caption = #1048#1079#1086#1073#1088#1072#1078#1077#1085#1080#1077
      ImageIndex = 1
      object Glyph: TImage
        Left = 0
        Top = 0
        Width = 782
        Height = 776
        Align = alClient
        ExplicitLeft = 24
        ExplicitTop = 24
        ExplicitWidth = 745
        ExplicitHeight = 737
      end
    end
  end
end
