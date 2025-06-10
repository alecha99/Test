unit uEditBatches;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, uniGUITypes, uniGUIAbstractClasses,
  uniGUIClasses, uniGUIForm, Data.DB, uniGUIBaseClasses, uniEdit, uniDBEdit,
  uniLabel, uniDBText, uniButton, uniBitBtn, uniDateTimePicker;

type
  TfEditBatches = class(TUniForm)
    ds: TDataSource;
    UniDBEdit1: TUniDBEdit;
    dtBegin: TUniDateTimePicker;
    UniDBEdit2: TUniDBEdit;
    UniDBEdit3: TUniDBEdit;
    UniLabel1: TUniLabel;
    UniLabel2: TUniLabel;
    UniLabel3: TUniLabel;
    UniLabel4: TUniLabel;
    UniLabel5: TUniLabel;
    UniBitBtn1: TUniBitBtn;
    UniBitBtn2: TUniBitBtn;
    ePrice: TUniNumberEdit;
    eQuantity: TUniNumberEdit;
    procedure UniFormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

function fEditBatches: TfEditBatches;

implementation
 uses System.Math, MainModule, uniGUIApplication;
{$R *.dfm}


function fEditBatches: TfEditBatches;
begin
  Result := TfEditBatches(UniMainModule.GetFormInstance(TfEditBatches));
end;

procedure TfEditBatches.UniFormShow(Sender: TObject);
begin
  dtBegin.DateTime := Date;
  eQuantity.DecimalPrecision := IfThen(ds.DataSet.FieldByName('is_float').AsString = 'F', 0, 4);
end;

end.
