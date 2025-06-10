unit MainModule;

interface

uses
  uniGUIMainModule, SysUtils, Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.FB,
  FireDAC.Phys.FBDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client;

type
  TUniMainModule = class(TUniGUIMainModule)
    FDConnection: TFDConnection;
    FDTransactionRW: TFDTransaction;
    fdProduct: TFDQuery;
    fdProductID: TLargeintField;
    fdProductName: TStringField;
    fdProductShort_NAME: TStringField;
    fdProductIS_FLOAT: TStringField;
    fdGoods: TFDQuery;
    fdGoodsID: TLargeintField;
    fdGoodsProduct_ID: TLargeintField;
    fdGoodsQuantity: TBCDField;
    fdGoodsPrice: TBCDField;
    fdGoodsBeg_Date: TDateField;
    fdGoodsName: TStringField;
    fdGoodsShort_NAME: TStringField;
    fdGoodsIS_FLOAT: TStringField;
    fdProductid_ed: TIntegerField;
    fdED: TFDQuery;
    fdEDID: TIntegerField;
    fdEDNAME: TStringField;
    fdEDShort_NAME: TStringField;
    fdSetQuery: TFDQuery;
    fdSelect: TFDQuery;
    fdOrder: TFDQuery;
    fdOrderID: TLargeintField;
    fdOrderName: TStringField;
    fdOrderDate: TDateField;
    fdOrderORDER_SUMM: TFMTBCDField;
    fdResult: TFDQuery;
    fdResultID: TLargeintField;
    fdResultProduct_ID: TLargeintField;
    fdResultQuantity: TBCDField;
    fdResultPrice: TBCDField;
    fdResultSUMM_IN: TFMTBCDField;
    fdResultQUANTITYRES: TBCDField;
    fdResultSUMMRES: TFMTBCDField;
    fdResultPRODUCT_NAME: TStringField;
    fdResultED: TStringField;
    fdResultIS_FLOAT: TStringField;
    fdDataResult: TFDQuery;
    fdDataResultIND: TIntegerField;
    fdDataResultPrice: TFMTBCDField;
    fdDataResultProduct_ID: TLargeintField;
    fdDataResultPRODUCT_NAME: TStringField;
    fdDataResultED: TStringField;
    fdDataResultSUMMRES: TFMTBCDField;
    fdDataResultQUANTITYRES: TBCDField;
    DataSource1: TDataSource;
    fdOrder_Goods: TFDQuery;
    fdOrder_GoodsPRODUCTNAME: TStringField;
    fdOrder_GoodsID_Batches: TLargeintField;
    fdOrder_GoodsPrice: TBCDField;
    fdOrder_GoodsQuantity: TBCDField;
    fdOrder_GoodsED_ISM: TStringField;
    fdOrder_GoodsID_Goods: TLargeintField;
    fdSelect2: TFDQuery;
  private

  public
    function GetGenerator(const ANameGen: String): Largeint;
  end;

function UniMainModule: TUniMainModule;

implementation

{$R *.dfm}

uses
  UniGUIVars, ServerModule, uniGUIApplication;

function UniMainModule: TUniMainModule;
begin
  Result := TUniMainModule(UniApplication.UniMainModule)
end;

{ TUniMainModule }

function TUniMainModule.GetGenerator(const ANameGen: String): Largeint;
begin
  fdSetQuery.Close;
  fdSetQuery.SQL.Text := Format('select gen_id("%0:s", 1) as id from rdb$database',[ANameGen]);
  fdSetQuery.Prepare;
  fdSetQuery.Open;
  Result := fdSetQuery.FieldByName('ID').AsLargeint;
  fdSetQuery.Close;
end;

initialization
  RegisterMainModuleClass(TUniMainModule);
end.
