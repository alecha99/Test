unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, uniGUITypes, uniGUIAbstractClasses,
  uniGUIClasses, uniGUIRegClasses, uniGUIForm, uniPanel, uniPageControl,
  uniGUIBaseClasses, uniBasicGrid, uniDBGrid, Data.DB, uniLabel, uniButton,
  uniBitBtn, uniEdit, uniDBEdit, Vcl.ExtCtrls, uniMultiItem, uniComboBox,
  uniDBComboBox, uniDBLookupComboBox, uniGUImJSForm, unimPanel, uniImageList,
  uniDateTimePicker, uniSplitter;

type
  TEditRecord =(terNone, terEdit, terNew);
  TTableRecord =(ttrProduct, ttrBatches, ttrOrder);
  TMainForm = class(TUniForm)
    pgProgram: TUniPageControl;
    tshGoods: TUniTabSheet;
    tshOrders: TUniTabSheet;
    dsProduct: TDataSource;
    dsGoods: TDataSource;
    pProduct: TUniPanel;
    pEditProduct: TUniPanel;
    dsED: TDataSource;
    eNameProduct: TUniEdit;
    pbtProduct: TUniPanel;
    gProduct: TUniDBGrid;
    UniDBEdit1: TUniDBEdit;
    UniLabel1: TUniLabel;
    UniLabel2: TUniLabel;
    UniLabel3: TUniLabel;
    btSaveProduct: TUniBitBtn;
    btCancelProduct: TUniBitBtn;
    btAddProduct: TUniBitBtn;
    btEditProduct: TUniBitBtn;
    cbED: TUniComboBox;
    gBatches: TUniDBGrid;
    ImageList: TUniImageList;
    pOrder: TUniPanel;
    pEditOrder: TUniPanel;
    eNameOrder: TUniEdit;
    UniLabel4: TUniLabel;
    UniLabel5: TUniLabel;
    btSaveOrder: TUniBitBtn;
    btCancelOrder: TUniBitBtn;
    pbtOrder: TUniPanel;
    UniDBEdit2: TUniDBEdit;
    UniLabel6: TUniLabel;
    btOrderNew: TUniBitBtn;
    gOrders: TUniDBGrid;
    UniDBEdit3: TUniDBEdit;
    lLabelOrderID: TUniLabel;
    dtOrder: TUniDateTimePicker;
    dsOrder: TDataSource;
    gOstatok: TUniDBGrid;
    dsOstatok: TDataSource;
    tshOstatok: TUniTabSheet;
    btRefresh: TUniBitBtn;
    dsResult: TDataSource;
    UniDBGrid3: TUniDBGrid;
    gRasxod: TUniDBGrid;
    UniSplitter1: TUniSplitter;
    dsOrder_Goods: TDataSource;
    procedure btEditProductClick(Sender: TObject);
    procedure btAddProductClick(Sender: TObject);
    procedure btSaveProductClick(Sender: TObject);
    procedure btCancelProductClick(Sender: TObject);
    procedure cbEDChange(Sender: TObject);
    procedure UniDBGrid1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure gBatchesDropRowsEvent(SrcGrid, DstGrid: TUniDBGrid;
      Rows: TUniBookmarkList; Params: TUniDragDropParams; var Handled: Boolean);
    procedure UniFormShow(Sender: TObject);
    procedure btOrderNewClick(Sender: TObject);
    procedure UniBitBtn2Click(Sender: TObject);
    procedure btSaveOrderClick(Sender: TObject);
    procedure btCancelOrderClick(Sender: TObject);
    procedure pgProgramChange(Sender: TObject);
    procedure btRefreshClick(Sender: TObject);
    procedure gRasxodDropRowsEvent(SrcGrid, DstGrid: TUniDBGrid;
      Rows: TUniBookmarkList; Params: TUniDragDropParams; var Handled: Boolean);
  private
    fStatus: TEditRecord;
    fIdRec: Int64;
    fTableEdit: TTableRecord;
    procedure ShowCallBackBatches(Sender: TComponent; AResult:Integer);
    procedure ShowCallBackOrdersPos(Sender: TComponent; AResult:Integer);
    procedure SetEdit(const AEdit: TEditRecord; const ATableRec: TTableRecord);
  public
    { Public declarations }
  end;

function MainForm: TMainForm;

implementation

{$R *.dfm}

uses
  uniGUIVars, MainModule, uniGUIApplication, uEditBatches, uEditOrdersPos;

function MainForm: TMainForm;
begin
  Result := TMainForm(UniMainModule.GetFormInstance(TMainForm));
end;

procedure TMainForm.btAddProductClick(Sender: TObject);
begin
  fIdRec := -1;
  eNameProduct.Text := '';
  SetEdit(terNew, ttrProduct);
  dsED.DataSet.FindFirst;
  cbEd.Items.Clear;
  while not dsED.DataSet.Eof do
  begin
    cbEd.Items.Add(dsED.DataSet.FieldByName('Name').AsString);
    dsED.DataSet.Next;
  end;
  if cbEd.Items.Count > 0 then
    cbEd.ItemIndex := 0;
end;

procedure TMainForm.btCancelOrderClick(Sender: TObject);
begin
  SetEdit(terNone, ttrOrder);
end;

procedure TMainForm.btCancelProductClick(Sender: TObject);
begin
  SetEdit(terNone, ttrProduct);
end;

procedure TMainForm.btEditProductClick(Sender: TObject);
var
  i: Integer;
begin
  Self.btAddProductClick(Sender);
  if (dsProduct.DataSet.RecordCount > 0) then
  begin
    fIdRec := dsProduct.DataSet.FieldByName('ID').AsLargeInt;
    eNameProduct.Text := dsProduct.DataSet.FieldByName('Name').AsString;
    dsED.DataSet.Locate('id', dsProduct.DataSet.FieldByName('id_ed').AsInteger, [loCaseInsensitive, loPartialKey]);
    SetEdit(terEdit, ttrProduct);
    for i := 0 to Pred(cbEd.Items.Count) do
    begin
      if AnsiSameStr(cbEd.Items.Strings[i], dsED.DataSet.FieldByName('Name').AsString) then
      begin
        cbEd.ItemIndex := i;
        Break;
      end;
    end;
  end;
end;

procedure TMainForm.btOrderNewClick(Sender: TObject);
begin
  fIdRec := -1;
  eNameOrder.Text := '';
  dtOrder.DateTime := Date;
  SetEdit(terNew, ttrOrder);
end;

procedure TMainForm.btRefreshClick(Sender: TObject);
begin
  dsOrder.DataSet.Refresh;
  dsOstatok.DataSet.Refresh;
end;

procedure TMainForm.btSaveProductClick(Sender: TObject);
begin
  if (fStatus = terNew)then
    fIdRec := UniMainModule.GetGenerator('GEN_SPR_Product_Name_ID');
  UniMainModule.fdSetQuery.Close;
  UniMainModule.fdSetQuery.Sql.Text :=
  'update or insert into "SPR_Product_Name" (ID, "Name", "id_ed")' +
  ' values (:ID, :"Name", :"id_ed")' +
  ' matching (ID)';
  UniMainModule.fdSetQuery.Prepare;
  with UniMainModule.fdSetQuery do
  begin
    ParamByName('ID').AsLargeInt := fIdRec;
    ParamByName('Name').AsString := eNameProduct.text;
    ParamByName('id_ed').AsInteger := dsEd.DataSet.FieldByName('id').asInteger;
    UniMainModule.fdSetQuery.ExecSQL;
  end;
  dsProduct.DataSet.Close;
  dsProduct.DataSet.Open;
  dsProduct.DataSet.Locate('id', fIdRec,[loCaseInsensitive, loPartialKey]);

  btCancelProductClick(Sender);
end;

procedure TMainForm.btSaveOrderClick(Sender: TObject);
begin
  fIdRec := UniMainModule.GetGenerator('GEN_Order_ID');
  UniMainModule.fdSetQuery.Close;
  UniMainModule.fdSetQuery.Sql.Text :=
  'insert into "Order" (ID, "Name", "Date")'+
  '    values (:ID, :"Name", :"Date")';

  UniMainModule.fdSetQuery.Prepare;
  with UniMainModule.fdSetQuery do
  begin
    ParamByName('ID').AsLargeInt := fIdRec;
    ParamByName('Name').AsString := eNameOrder.text;
    ParamByName('Date').AsDate := dtOrder.DateTime;
    ExecSQL;
  end;
  dsOrder.DataSet.Close;
  dsOrder.DataSet.Open;
  dsOrder.DataSet.Locate('id', fIdRec,[loCaseInsensitive, loPartialKey]);

  btCancelOrderClick(Sender);
end;

procedure TMainForm.cbEDChange(Sender: TObject);
begin
  dsED.DataSet.Locate('Name', TUniComboBox(Sender).Text, [loCaseInsensitive, loPartialKey]);
end;

procedure TMainForm.gBatchesDropRowsEvent(SrcGrid, DstGrid: TUniDBGrid;
  Rows: TUniBookmarkList; Params: TUniDragDropParams; var Handled: Boolean);
begin
  UniMainModule.fdSelect.Close;
  UniMainModule.fdSelect.Sql.Text :=
   ' select pr.ID, pr."Name", ed."Short_NAME", ed.is_float, ed.name as ed_name ' +
   '        from "SPR_Product_Name" pr  ' +
   '          left join spr_ed ed on ed.id=pr."id_ed"' +
   'Where pr.ID = :ID';
  UniMainModule.fdSelect.Prepare;
  with UniMainModule.fdSelect do
  begin
    ParamByName('ID').AsLargeInt := SrcGrid.DataSource.DataSet.FieldByName('ID').AsLargeInt;
    Open;
  end;
  fEditBatches.ShowModal(ShowCallBackBatches);
  Handled := True;
end;

procedure TMainForm.gRasxodDropRowsEvent(SrcGrid, DstGrid: TUniDBGrid;
  Rows: TUniBookmarkList; Params: TUniDragDropParams; var Handled: Boolean);
begin
  UniMainModule.fdSelect2.Close;
  UniMainModule.fdSelect2.Sql.Text :=
   'select ' + sLineBreak +
   '  bg.ID ' + sLineBreak +
   ' ,(bg."Quantity" - sum(coalesce(og."Quantity", 0))) as Quantity ' + sLineBreak +
   ' ,bg."Price"  ' + sLineBreak +
   ' ,product."Name" as Product_Name  ' + sLineBreak +
   ' ,ed."Short_NAME" as Ed ' + sLineBreak +
   ' ,ed.is_float as  ed_f  ' + sLineBreak +
   'from "Batches_goods" bg ' + sLineBreak +
   '   left join "SPR_Product_Name" product on bg."Product_ID" = product.id  ' + sLineBreak +
   '   left join spr_ed ed on ed.id=product."id_ed" ' + sLineBreak +
   '   left join "Order_Goods" og on(og."ID_Batches" = bg.id) ' + sLineBreak +
   'where (bg.id = :ID) ' + sLineBreak +
   'group by   ' + sLineBreak +
   '   bg.ID   ' + sLineBreak +
   '  ,bg."Price"  ' + sLineBreak +
   '  ,bg."Quantity" ' + sLineBreak +
   '  ,product."Name"   ' + sLineBreak +
   '  ,ed."Short_NAME"   ' + sLineBreak +
   '  ,ed.is_float       ' + sLineBreak +
   'having (bg."Quantity" - sum(coalesce(og."Quantity", 0))>0)';
  UniMainModule.fdSelect2.Prepare;
  with UniMainModule.fdSelect2 do
  begin
    ParamByName('ID').AsLargeInt := SrcGrid.DataSource.DataSet.FieldByName('ID').AsLargeInt;
    Open;
  end;
  fEditOrdersPos.ShowModal(ShowCallBackOrdersPos);
  Handled := True;
end;

procedure TMainForm.pgProgramChange(Sender: TObject);
begin
  if pgProgram.ActivePage = tshOrders then
    UniMainModule.fdResult.Refresh
  else
    if pgProgram.ActivePage = tshOstatok then
      dsResult.DataSet.Refresh;
end;

procedure TMainForm.SetEdit(const AEdit: TEditRecord; const ATableRec: TTableRecord);
begin
  fStatus := AEdit;
  if AEdit in [terEdit, terNew] then
  begin
    fTableEdit := ATableRec;
    case ATableRec of
      ttrProduct:
        Self.ActiveControl := eNameProduct;
      ttrOrder:
        Self.ActiveControl := eNameOrder;
    end
  end;
  gProduct.Enabled := (AEdit = terNone);
  tshGoods.Enabled := (AEdit = terNone) or (ATableRec in [ttrProduct, ttrBatches]);
  tshOrders.Enabled := (AEdit = terNone) or (ATableRec = ttrOrder);
  pEditProduct.Visible := not(AEdit = terNone)and(ATableRec = ttrProduct);
  pbtProduct.Visible := (AEdit = terNone);
  pEditOrder.Visible := not(AEdit = terNone)and(ATableRec = ttrOrder);
  pbtOrder.Visible := (AEdit = terNone);
end;

procedure TMainForm.ShowCallBackBatches(Sender: TComponent; AResult: Integer);
var
  ff: TfEditBatches;
  Id, pid: Int64;
begin
  if AResult = mrOk then
  begin
    ff := nil;
    if Sender.InheritsFrom(TfEditBatches) then
      ff:= TfEditBatches(Sender);
    if Assigned(ff) then
    begin
      pid := UniMainModule.fdSelect.FieldByName('id').AsLargeInt;
      id := UniMainModule.GetGenerator('GEN_Batches_goods_ID');
      UniMainModule.fdSetQuery.Close;
      UniMainModule.fdSetQuery.SQL.Clear;
      UniMainModule.fdSetQuery.Sql.Text :=
       'insert into "Batches_goods" (ID, "Product_ID", "Quantity", "Price", "Beg_Date")' +
       'values(:ID, :"Product_ID", :"Quantity", :"Price", :"Beg_Date");';
      UniMainModule.fdSetQuery.Prepare;
      with UniMainModule.fdSetQuery do
      begin
        ParamByName('ID').AsLargeInt := id;
        ParamByName('Product_ID').AsLargeInt := pid;
        ParamByName('Quantity').Value := ff.eQuantity.Value;
        ParamByName('Price').Value := ff.ePrice.Value;
        ParamByName('Beg_Date').AsDate := Round(ff.dtBegin.DateTime);
        ExecSQL;
      end;
      gBatches.DataSource.DataSet.Close;
      gBatches.DataSource.DataSet.Open;
      gBatches.DataSource.DataSet.Locate('id', Id, []);
    end;
  end;
end;

procedure TMainForm.ShowCallBackOrdersPos(Sender: TComponent; AResult: Integer);
var
  ff: TfEditOrdersPos;
  Id, pid: Int64;
begin
  if AResult = mrOk then
  begin
    ff := nil;
    if Sender.InheritsFrom(TfEditOrdersPos) then
      ff:= TfEditOrdersPos(Sender);

    if Assigned(ff) then
    begin
      if ff.eQuantity.Value > UniMainModule.fdSelect2.FieldByName('Quantity').value then
         ShowMessage('¬ы пытаетесь списать кол-во превышающее остаток!')
      else
      begin
        pid := UniMainModule.fdSelect2.FieldByName('id').AsLargeInt;
        id :=  dsOrder.DataSet.FieldByName('ID').AsLargeInt;
        UniMainModule.fdSetQuery.Close;
        UniMainModule.fdSetQuery.SQL.Clear;

        UniMainModule.fdSetQuery.Sql.Text :=
         'update or insert into "Order_Goods"("ID_Goods", "ID_Batches", "Quantity") '+
         '     values (:"ID_Goods", :"ID_Batches", :"Quantity")' +
         ' matching ("ID_Goods", "ID_Batches");';
        UniMainModule.fdSetQuery.Prepare;
        with UniMainModule.fdSetQuery do
        begin
          ParamByName('ID_Batches').AsLargeInt := pid;
          ParamByName('ID_Goods').AsLargeInt := id;
          ParamByName('Quantity').Value := ff.eQuantity.Value;
          ExecSQL;
        end;
        dsOrder.DataSet.Refresh;
        dsOrder_Goods.DataSet.Refresh;
        dsOstatok.DataSet.Refresh;
        dsOrder_Goods.DataSet.Locate('ID_Batches', pid, []);
      end;
    end;
  end;
end;

procedure TMainForm.UniBitBtn2Click(Sender: TObject);
begin
  SetEdit(terNone, ttrOrder);
end;

procedure TMainForm.UniDBGrid1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if (Button = mbLeft) and (ssAlt in Shift) then
    (Sender as TControl).BeginDrag(false);
end;

procedure TMainForm.UniFormShow(Sender: TObject);
begin
  pgProgram.ActivePageIndex := 0;
end;

initialization
  RegisterAppFormClass(TMainForm);

end.
