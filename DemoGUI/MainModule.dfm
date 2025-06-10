object UniMainModule: TUniMainModule
  OldCreateOrder = False
  MonitoredKeys.Keys = <>
  Height = 508
  Width = 799
  object FDConnection: TFDConnection
    Params.Strings = (
      'User_Name=SYSDBA'
      'Password=masterkey'
      'Port=3050'
      'RoleName=rdb$admin'
      'CharacterSet=WIN1251'
      'Database=C:\DelphiXE10\DemoGUI\db\DEMODB.FDB'
      'DriverID=FB')
    LoginPrompt = False
    Left = 32
    Top = 40
  end
  object FDTransactionRW: TFDTransaction
    Connection = FDConnection
    Left = 24
    Top = 105
  end
  object fdProduct: TFDQuery
    IndexFieldNames = 'Name;Short_NAME'
    DetailFields = 'ID;Name;Short_NAME;IS_FLOAT'
    Connection = FDConnection
    Transaction = FDTransactionRW
    SQL.Strings = (
      
        ' select pr.ID, pr."Name", ed."Short_NAME", ed.is_float, pr."id_e' +
        'd"'
      '           from "SPR_Product_Name" pr'
      '             left join spr_ed ed on ed.id=pr."id_ed"')
    Left = 152
    Top = 40
    object fdProductID: TLargeintField
      DisplayWidth = 5
      FieldName = 'ID'
      Origin = 'ID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Required = True
      Visible = False
    end
    object fdProductName: TStringField
      DisplayLabel = ' '#1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
      DisplayWidth = 48
      FieldName = 'Name'
      Origin = 'Name'
      Required = True
      Size = 150
    end
    object fdProductShort_NAME: TStringField
      Alignment = taCenter
      AutoGenerateValue = arDefault
      DisplayLabel = #1045#1076'. '#1080#1079#1084'.'
      DisplayWidth = 11
      FieldName = 'Short_NAME'
      Origin = 'Short_NAME'
      ProviderFlags = []
      ReadOnly = True
      Size = 5
    end
    object fdProductIS_FLOAT: TStringField
      AutoGenerateValue = arDefault
      FieldName = 'IS_FLOAT'
      Origin = 'IS_FLOAT'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
      FixedChar = True
      Size = 1
    end
    object fdProductid_ed: TIntegerField
      FieldName = 'id_ed'
      Origin = '"id_ed"'
      Required = True
      Visible = False
    end
  end
  object fdGoods: TFDQuery
    IndexFieldNames = 'Name;Beg_Date;Quantity;Price;Short_NAME'
    DetailFields = 'Beg_Date;ID;IS_FLOAT;Name;Price;Product_ID;Quantity;Short_NAME'
    Connection = FDConnection
    Transaction = FDTransactionRW
    FormatOptions.AssignedValues = [fvFmtDisplayDate, fvFmtEditNumeric]
    FormatOptions.FmtDisplayDate = 'dd/mm/yyyy'
    FormatOptions.FmtEditNumeric = '### ### ### ##0,0###'
    SQL.Strings = (
      
        '  select bg.ID, "Product_ID", bg."Quantity", bg."Price", bg."Beg' +
        '_Date", '
      '     product."Name", ed."Short_NAME", ed.is_float'
      '  from "Batches_goods" bg'
      
        '  left join "SPR_Product_Name" product on bg."Product_ID" = prod' +
        'uct.id'
      '  left join spr_ed ed on ed.id=product."id_ed"'
      'where bg."Quantity">0'
      'order by product."Name",  bg."Beg_Date",  bg."Quantity"')
    Left = 152
    Top = 89
    object fdGoodsName: TStringField
      Alignment = taCenter
      AutoGenerateValue = arDefault
      DisplayLabel = ' '#1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077
      DisplayWidth = 45
      FieldName = 'Name'
      Origin = '"Name"'
      ProviderFlags = []
      ReadOnly = True
      FixedChar = True
      Size = 150
    end
    object fdGoodsBeg_Date: TDateField
      Alignment = taCenter
      DisplayLabel = #1044#1072#1090#1072' '#1087#1088#1080#1093#1086#1076#1072
      DisplayWidth = 14
      FieldName = 'Beg_Date'
      Origin = '"Beg_Date"'
      Required = True
      DisplayFormat = 'dd/mm/yyyy'
    end
    object fdGoodsPrice: TBCDField
      DisplayLabel = #1062#1077#1085#1072
      DisplayWidth = 11
      FieldName = 'Price'
      Origin = '"Price"'
      Required = True
      Precision = 18
      Size = 2
    end
    object fdGoodsQuantity: TBCDField
      DisplayLabel = #1050#1086#1083'-'#1074#1086
      DisplayWidth = 9
      FieldName = 'Quantity'
      Origin = '"Quantity"'
      Required = True
      Precision = 18
    end
    object fdGoodsShort_NAME: TStringField
      AutoGenerateValue = arDefault
      DisplayLabel = #1045#1076'. '#1080#1079#1084'.'
      DisplayWidth = 8
      FieldName = 'Short_NAME'
      Origin = '"Short_NAME"'
      ProviderFlags = []
      ReadOnly = True
      Size = 5
    end
    object fdGoodsID: TLargeintField
      FieldName = 'ID'
      Origin = 'ID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Required = True
      Visible = False
      EditFormat = '### ### ### ##0.0###'
    end
    object fdGoodsProduct_ID: TLargeintField
      FieldName = 'Product_ID'
      Origin = '"Product_ID"'
      Required = True
      Visible = False
      EditFormat = '### ### ### ##0.0###'
    end
    object fdGoodsIS_FLOAT: TStringField
      AutoGenerateValue = arDefault
      FieldName = 'IS_FLOAT'
      Origin = 'IS_FLOAT'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
      FixedChar = True
      Size = 1
    end
  end
  object fdED: TFDQuery
    DetailFields = 'ID;NAME;Short_NAME'
    Connection = FDConnection
    FetchOptions.AssignedValues = [evRecsMax]
    FetchOptions.RecsMax = 100
    SQL.Strings = (
      'select ID, NAME, "Short_NAME"'
      '   from SPR_ED'
      'order by name')
    Left = 152
    Top = 136
    object fdEDID: TIntegerField
      FieldName = 'ID'
      Origin = 'ID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Required = True
    end
    object fdEDNAME: TStringField
      FieldName = 'NAME'
      Origin = 'NAME'
      Required = True
      Size = 50
    end
    object fdEDShort_NAME: TStringField
      FieldName = 'Short_NAME'
      Origin = '"Short_NAME"'
      Required = True
      Size = 5
    end
  end
  object fdSetQuery: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      
        ' select pr.ID, pr."Name", ed."Short_NAME", ed.is_float, ed.name ' +
        'as ed_name'
      '           from "SPR_Product_Name" pr'
      '             left join spr_ed ed on ed.id=pr."id_ed"')
    Left = 152
    Top = 192
  end
  object fdSelect: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      
        '   select pr.ID, pr."Name", ed."Short_NAME", ed.is_float, ed.nam' +
        'e as ed_name'
      '           from "SPR_Product_Name" pr'
      '             left join spr_ed ed on ed.id=pr."id_ed"')
    Left = 344
    Top = 296
  end
  object fdOrder: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      '   select'
      '        Od.ID, Od."Name", Od."Date",'
      
        '        sum(coalesce((OG."Quantity" * b."Price"),0)) as order_su' +
        'mm'
      '   from "Order" as Od'
      '   left join "Order_Goods" as OG on Od.id = OG."ID_Goods"'
      '   left join "Batches_goods" as b on b.id = OG."ID_Batches"'
      '   group by Od.ID, Od."Name", Od."Date"')
    Left = 208
    Top = 48
    object fdOrderID: TLargeintField
      DisplayWidth = 5
      FieldName = 'ID'
      Origin = 'ID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Required = True
    end
    object fdOrderName: TStringField
      DisplayLabel = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1079#1072#1082#1072#1079#1072
      DisplayWidth = 27
      FieldName = 'Name'
      Origin = '"Name"'
      Required = True
      Size = 100
    end
    object fdOrderDate: TDateField
      DisplayLabel = #1044#1072#1090#1072
      DisplayWidth = 12
      FieldName = 'Date'
      Origin = '"Date"'
      Required = True
      DisplayFormat = 'dd/mm/yyyy'
    end
    object fdOrderORDER_SUMM: TFMTBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1057#1091#1084#1084#1072' '#1079#1072#1082#1072#1079#1072
      DisplayWidth = 15
      FieldName = 'ORDER_SUMM'
      Origin = 'ORDER_SUMM'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
      Size = 6
    end
  end
  object fdResult: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      'select bg.ID, bg."Product_ID", bg."Quantity", bg."Price",'
      '    (bg."Quantity" * bg."Price")  summ_in,'
      '   ( bg."Quantity"  - ost.outq) QuantityRes,'
      '   (( bg."Quantity"  - ost.outq) *  bg."Price")  SummRes,'
      
        '   product."Name" as Product_Name, ed."Short_NAME" as Ed, ed.is_' +
        'float'
      'from "Batches_goods" bg'
      '  left join ('
      '    select'
      '      sum(coalesce(og."Quantity", 0)) outq,'
      '      bg."Product_ID",'
      '      bg.id'
      '    from  "Batches_goods" bg'
      '       left join  "Order_Goods" og on bg.id = og."ID_Batches"'
      '    group by bg.id, bg."Product_ID"  ) ost on ost.id = bg.id'
      
        '  left join "SPR_Product_Name" product on bg."Product_ID" = prod' +
        'uct.id'
      '  left join spr_ed ed on ed.id=product."id_ed"'
      'Where  (bg."Quantity" - ost.outq) > 0')
    Left = 208
    Top = 96
    object fdResultID: TLargeintField
      FieldName = 'ID'
      Origin = 'ID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Required = True
      Visible = False
    end
    object fdResultPRODUCT_NAME: TStringField
      AutoGenerateValue = arDefault
      DisplayLabel = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1090#1086#1074#1072#1088#1072
      DisplayWidth = 37
      FieldName = 'PRODUCT_NAME'
      Origin = '"Name"'
      ProviderFlags = []
      ReadOnly = True
      Size = 150
    end
    object fdResultProduct_ID: TLargeintField
      FieldName = 'Product_ID'
      Origin = '"Product_ID"'
      Required = True
      Visible = False
    end
    object fdResultQuantity: TBCDField
      FieldName = 'Quantity'
      Origin = '"Quantity"'
      Required = True
      Visible = False
      Precision = 18
    end
    object fdResultPrice: TBCDField
      DisplayWidth = 8
      FieldName = 'Price'
      Origin = '"Price"'
      Required = True
      Precision = 18
      Size = 2
    end
    object fdResultSUMM_IN: TFMTBCDField
      AutoGenerateValue = arDefault
      FieldName = 'SUMM_IN'
      Origin = 'SUMM_IN'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
      Precision = 18
      Size = 6
    end
    object fdResultQUANTITYRES: TBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1054#1089#1090#1072#1090#1086#1082
      DisplayWidth = 18
      FieldName = 'QUANTITYRES'
      Origin = 'QUANTITYRES'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
    end
    object fdResultSUMMRES: TFMTBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1057#1091#1084#1084#1072' '#1086#1089#1090#1072#1090#1082#1072
      DisplayWidth = 21
      FieldName = 'SUMMRES'
      Origin = 'SUMMRES'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
      Size = 6
    end
    object fdResultED: TStringField
      AutoGenerateValue = arDefault
      DisplayLabel = #1045#1076#1080#1085#1080#1094#1072' '#1080#1079#1084'.'
      DisplayWidth = 14
      FieldName = 'ED'
      Origin = '"Short_NAME"'
      ProviderFlags = []
      ReadOnly = True
      Size = 5
    end
    object fdResultIS_FLOAT: TStringField
      AutoGenerateValue = arDefault
      FieldName = 'IS_FLOAT'
      Origin = 'IS_FLOAT'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
      FixedChar = True
      Size = 1
    end
  end
  object fdDataResult: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      'Select *'
      ''
      'from'
      '('
      'select'
      '    1 as ind'
      
        '    ,Sum(tt.QuantityRes * tt."Price") / sum(tt.QuantityRes)  "Pr' +
        'ice"'
      '    ,tt."Product_ID"'
      '    ,tt.Product_Name'
      '    ,tt.Ed'
      '    ,sum(SummRes) SummRes'
      '    ,sum(QuantityRes) QuantityRes'
      'from'
      '('
      'select bg.ID, bg."Product_ID", bg."Quantity", bg."Price",'
      '    (bg."Quantity" * bg."Price")  summ_in,'
      '   ( bg."Quantity"  - ost.outq) QuantityRes,'
      '   (( bg."Quantity"  - ost.outq) *  bg."Price")  SummRes,'
      
        '   product."Name" as Product_Name, ed."Short_NAME" as Ed, ed.is_' +
        'float'
      'from "Batches_goods" bg'
      '  left join ('
      '    select'
      '      sum(coalesce(og."Quantity", 0)) outq,'
      '      bg."Product_ID",'
      '      bg.id'
      '    from  "Batches_goods" bg'
      '       left join  "Order_Goods" og on bg.id = og."ID_Batches"'
      '    group by bg.id, bg."Product_ID"  ) ost on ost.id = bg.id'
      
        '  left join "SPR_Product_Name" product on bg."Product_ID" = prod' +
        'uct.id'
      '  left join spr_ed ed on ed.id=product."id_ed"'
      'Where (bg."Quantity" - ost.outq) > 0) tt'
      'group by tt."Product_ID", tt.Product_Name, tt.Ed'
      'having sum(tt.QuantityRes) > 0'
      'union'
      ' select'
      '   2 as ind,'
      '   null, null'
      '   ,'#39#1057#1091#1084#1084#1072' '#1086#1073#1097#1077#1075#1086' '#1086#1089#1090#1072#1090#1082#1072#39
      '   ,null'
      '   ,(    select'
      
        '      sum((bg."Quantity" - coalesce(og."Quantity", 0))* bg."Pric' +
        'e") outq'
      '    from  "Batches_goods" bg'
      '       left join  "Order_Goods" og on bg.id = og."ID_Batches"'
      '       )'
      '   ,null'
      ' from rdb$database'
      ')tt'
      'order by tt.ind, tt.Product_Name')
    Left = 208
    Top = 152
    object fdDataResultPRODUCT_NAME: TStringField
      AutoGenerateValue = arDefault
      DisplayLabel = '  '#1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1087#1086#1079#1080#1094#1080#1080
      DisplayWidth = 49
      FieldName = 'PRODUCT_NAME'
      Origin = 'PRODUCT_NAME'
      ProviderFlags = []
      ReadOnly = True
      Size = 150
    end
    object fdDataResultIND: TIntegerField
      AutoGenerateValue = arDefault
      FieldName = 'IND'
      Origin = 'IND'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
    end
    object fdDataResultProduct_ID: TLargeintField
      AutoGenerateValue = arDefault
      FieldName = 'Product_ID'
      Origin = '"Product_ID"'
      ProviderFlags = []
      ReadOnly = True
      Visible = False
    end
    object fdDataResultED: TStringField
      AutoGenerateValue = arDefault
      DisplayLabel = #1045#1076#1077#1085#1080#1094#1072' '#1080#1079#1084'.'
      DisplayWidth = 12
      FieldName = 'ED'
      Origin = 'ED'
      ProviderFlags = []
      ReadOnly = True
      Size = 5
    end
    object fdDataResultQUANTITYRES: TBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1058#1077#1082#1091#1097#1080#1081' '#1086#1089#1090#1072#1090#1086#1082
      DisplayWidth = 19
      FieldName = 'QUANTITYRES'
      Origin = 'QUANTITYRES'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
    end
    object fdDataResultPrice: TFMTBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1057#1088'. '#1094#1077#1085#1072' '#1086#1089#1090#1072#1090#1082#1072
      DisplayWidth = 19
      FieldName = 'Price'
      Origin = '"Price"'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
      Size = 10
    end
    object fdDataResultSUMMRES: TFMTBCDField
      AutoGenerateValue = arDefault
      DisplayLabel = #1057#1091#1084#1084#1072' '#1086#1089#1090#1072#1090#1082#1072
      DisplayWidth = 19
      FieldName = 'SUMMRES'
      Origin = 'SUMMRES'
      ProviderFlags = []
      ReadOnly = True
      Precision = 18
      Size = 6
    end
  end
  object DataSource1: TDataSource
    DataSet = fdOrder
    Left = 288
    Top = 224
  end
  object fdOrder_Goods: TFDQuery
    MasterSource = DataSource1
    MasterFields = 'ID'
    Connection = FDConnection
    FetchOptions.AssignedValues = [evCache]
    FetchOptions.Cache = [fiBlobs, fiMeta]
    SQL.Strings = (
      'select'
      '    p."Name" as ProductName'
      '   ,og."ID_Batches"'
      '   ,bg."Price"'
      '   ,og."Quantity"'
      '   ,ed."Short_NAME" as ed_ism'
      '   ,og."ID_Goods"'
      'from "Order_Goods"  OG'
      'left join "Batches_goods" bg on bg.id = og."ID_Batches"'
      'left join  "SPR_Product_Name" p on p.id =bg."Product_ID"'
      'left join  spr_ed ed on ed.id = p."id_ed"'
      'where og."ID_Goods" =:ID')
    Left = 200
    Top = 280
    ParamData = <
      item
        Name = 'ID'
        DataType = ftLargeint
        ParamType = ptInput
        Value = 1
      end>
    object fdOrder_GoodsPRODUCTNAME: TStringField
      DisplayLabel = #1053#1072#1080#1084#1077#1085#1086#1074#1072#1085#1080#1077' '#1090#1086#1074#1072#1088#1072
      DisplayWidth = 31
      FieldName = 'PRODUCTNAME'
      Origin = 'PRODUCTNAME'
      Size = 150
    end
    object fdOrder_GoodsID_Batches: TLargeintField
      DisplayWidth = 15
      FieldName = 'ID_Batches'
      Origin = '"ID_Batches"'
      Required = True
      Visible = False
    end
    object fdOrder_GoodsPrice: TBCDField
      DisplayLabel = #1062#1077#1085#1072
      DisplayWidth = 10
      FieldName = 'Price'
      Origin = '"Price"'
      Precision = 18
      Size = 2
    end
    object fdOrder_GoodsQuantity: TBCDField
      DisplayLabel = #1050#1086#1083'-'#1074#1086
      DisplayWidth = 11
      FieldName = 'Quantity'
      Origin = '"Quantity"'
      Required = True
      Precision = 18
    end
    object fdOrder_GoodsED_ISM: TStringField
      DisplayLabel = #1045#1076'.'
      DisplayWidth = 7
      FieldName = 'ED_ISM'
      Origin = 'ED_ISM'
      Size = 5
    end
    object fdOrder_GoodsID_Goods: TLargeintField
      DisplayWidth = 15
      FieldName = 'ID_Goods'
      Origin = '"ID_Goods"'
      Required = True
      Visible = False
    end
  end
  object fdSelect2: TFDQuery
    Connection = FDConnection
    SQL.Strings = (
      '  select'
      '     bg.ID'
      
        '    ,(bg."Quantity" - sum(coalesce(og."Quantity", 0))) as Quanti' +
        'ty'
      '    ,bg."Price"'
      '    ,product."Name" as Product_Name'
      '    ,ed."Short_NAME" as Ed'
      '    ,ed.is_float as  ed_f'
      '  from "Batches_goods" bg'
      
        '      left join "SPR_Product_Name" product on bg."Product_ID" = ' +
        'product.id'
      '      left join spr_ed ed on ed.id=product."id_ed"'
      '      left join "Order_Goods" og on(og."ID_Batches" = bg.id)'
      '  group by'
      '     bg.ID'
      '    ,bg."Price"'
      '    ,bg."Quantity"'
      '    ,product."Name"'
      '    ,ed."Short_NAME"'
      '    ,ed.is_float'
      '  having (bg."Quantity" - sum(coalesce(og."Quantity", 0))>0)')
    Left = 392
    Top = 304
  end
end
