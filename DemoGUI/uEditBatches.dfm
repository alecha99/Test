object fEditBatches: TfEditBatches
  Left = 0
  Top = 0
  ClientHeight = 132
  ClientWidth = 445
  Caption = #1056#1077#1076#1072#1082#1090#1086#1088' '#1087#1086#1079#1080#1094#1080#1080' '#1087#1088#1080#1093#1086#1076#1072
  OnShow = UniFormShow
  BorderStyle = bsDialog
  OldCreateOrder = False
  MonitoredKeys.Keys = <>
  Font.Charset = RUSSIAN_CHARSET
  Font.Name = 'Times New Roman'
  PixelsPerInch = 96
  TextHeight = 14
  object UniDBEdit1: TUniDBEdit
    Left = 84
    Top = 8
    Width = 248
    Height = 22
    Hint = ''
    DataField = 'Name'
    DataSource = ds
    TabOrder = 0
    ReadOnly = True
  end
  object dtBegin: TUniDateTimePicker
    Left = 84
    Top = 68
    Width = 248
    Height = 25
    Hint = ''
    DateTime = 45673.000000000000000000
    DateFormat = 'dd/MM/yyyy'
    TimeFormat = 'HH:mm:ss'
    TabOrder = 1
  end
  object UniDBEdit2: TUniDBEdit
    Left = 276
    Top = 40
    Width = 56
    Height = 22
    Hint = ''
    DataField = 'Short_NAME'
    DataSource = ds
    TabOrder = 2
    ReadOnly = True
  end
  object UniDBEdit3: TUniDBEdit
    Left = 84
    Top = 40
    Width = 186
    Height = 22
    Hint = ''
    DataField = 'ED_NAME'
    DataSource = ds
    TabOrder = 3
    ReadOnly = True
  end
  object UniLabel1: TUniLabel
    Left = 8
    Top = 8
    Width = 70
    Height = 14
    Hint = ''
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
    TabOrder = 4
  end
  object UniLabel2: TUniLabel
    Left = 15
    Top = 40
    Width = 63
    Height = 14
    Hint = ''
    Caption = #1045#1076#1077#1085#1080#1094#1072' '#1080#1079#1084'.'
    TabOrder = 5
  end
  object UniLabel3: TUniLabel
    Left = 10
    Top = 72
    Width = 69
    Height = 14
    Hint = ''
    Caption = #1044#1072#1090#1072' '#1087#1086#1089#1090#1072#1074#1082#1080
    TabOrder = 6
  end
  object UniLabel4: TUniLabel
    Left = 46
    Top = 107
    Width = 32
    Height = 14
    Hint = ''
    Caption = #1050#1086#1083'-'#1074#1086
    TabOrder = 7
  end
  object UniLabel5: TUniLabel
    Left = 196
    Top = 107
    Width = 24
    Height = 14
    Hint = ''
    Caption = #1062#1077#1085#1072
    TabOrder = 8
  end
  object UniBitBtn1: TUniBitBtn
    Left = 352
    Top = 8
    Width = 85
    Height = 33
    Hint = ''
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
    ModalResult = 1
    TabOrder = 9
  end
  object UniBitBtn2: TUniBitBtn
    Left = 352
    Top = 88
    Width = 85
    Height = 33
    Hint = ''
    Caption = #1054#1090#1084#1077#1085#1080#1090#1100
    ModalResult = 2
    TabOrder = 10
  end
  object ePrice: TUniNumberEdit
    Left = 226
    Top = 102
    Width = 106
    Hint = ''
    TabOrder = 11
    FieldLabelFont.Charset = RUSSIAN_CHARSET
    FieldLabelFont.Name = 'Times New Roman'
    DecimalSeparator = ','
  end
  object eQuantity: TUniNumberEdit
    Left = 84
    Top = 102
    Width = 106
    Hint = ''
    TabOrder = 12
    FieldLabelFont.Charset = RUSSIAN_CHARSET
    FieldLabelFont.Name = 'Times New Roman'
    DecimalPrecision = 4
    DecimalSeparator = ','
  end
  object ds: TDataSource
    AutoEdit = False
    DataSet = UniMainModule.fdSelect
    Left = 8
    Top = 96
  end
end
