object fEditOrdersPos: TfEditOrdersPos
  Left = 0
  Top = 0
  ClientHeight = 119
  ClientWidth = 418
  Caption = #1042#1074#1086#1076' '#1082#1086#1083'-'#1074#1086' '#1087#1086' '#1087#1077#1088#1077#1084#1077#1097#1077#1085#1080#1102
  OnShow = UniFormShow
  BorderStyle = bsDialog
  OldCreateOrder = False
  MonitoredKeys.Keys = <>
  PixelsPerInch = 96
  TextHeight = 13
  object eQuantity: TUniNumberEdit
    Left = 84
    Top = 89
    Width = 106
    Hint = ''
    TabOrder = 0
    FieldLabelFont.Charset = RUSSIAN_CHARSET
    FieldLabelFont.Name = 'Times New Roman'
    DecimalPrecision = 4
    DecimalSeparator = ','
  end
  object UniBitBtn2: TUniBitBtn
    Left = 324
    Top = 78
    Width = 85
    Height = 33
    Hint = ''
    Caption = #1054#1090#1084#1077#1085#1080#1090#1100
    ModalResult = 2
    TabOrder = 1
  end
  object UniBitBtn1: TUniBitBtn
    Left = 207
    Top = 78
    Width = 85
    Height = 33
    Hint = ''
    Caption = #1055#1088#1080#1084#1077#1085#1080#1090#1100
    ModalResult = 1
    TabOrder = 2
  end
  object UniLabel5: TUniLabel
    Left = 146
    Top = 40
    Width = 24
    Height = 14
    Hint = ''
    Caption = #1062#1077#1085#1072
    TabOrder = 3
  end
  object UniLabel4: TUniLabel
    Left = 46
    Top = 94
    Width = 32
    Height = 14
    Hint = ''
    Caption = #1050#1086#1083'-'#1074#1086
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
  object UniLabel1: TUniLabel
    Left = 8
    Top = 12
    Width = 70
    Height = 14
    Hint = ''
    Caption = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
    TabOrder = 6
  end
  object UniDBEdit2: TUniDBEdit
    Left = 84
    Top = 36
    Width = 56
    Height = 22
    Hint = ''
    DataField = 'ED'
    DataSource = ds
    TabOrder = 7
    ReadOnly = True
  end
  object UniDBEdit1: TUniDBEdit
    Left = 84
    Top = 8
    Width = 325
    Height = 22
    Hint = ''
    DataField = 'PRODUCT_NAME'
    DataSource = ds
    TabOrder = 8
    ReadOnly = True
  end
  object UniDBEdit4: TUniDBEdit
    Left = 176
    Top = 36
    Width = 81
    Height = 22
    Hint = ''
    DataField = 'Price'
    DataSource = ds
    TabOrder = 9
  end
  object UniLabel3: TUniLabel
    Left = 263
    Top = 40
    Width = 43
    Height = 13
    Hint = ''
    Caption = #1054#1089#1090#1072#1090#1086#1082
    TabOrder = 10
  end
  object UniDBEdit3: TUniDBEdit
    Left = 312
    Top = 36
    Width = 97
    Height = 22
    Hint = ''
    DataField = 'QUANTITY'
    DataSource = ds
    TabOrder = 11
  end
  object ds: TDataSource
    AutoEdit = False
    DataSet = UniMainModule.fdSelect2
    Left = 16
    Top = 64
  end
end
