unit uEditOrdersPos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, uniGUITypes, uniGUIAbstractClasses,
  uniGUIClasses, uniGUIForm, Data.DB, uniDBEdit, uniLabel, uniButton, uniBitBtn,
  uniEdit, uniGUIBaseClasses, uniDateTimePicker;

type
  TfEditOrdersPos = class(TUniForm)
    eQuantity: TUniNumberEdit;
    UniBitBtn2: TUniBitBtn;
    UniBitBtn1: TUniBitBtn;
    UniLabel5: TUniLabel;
    UniLabel4: TUniLabel;
    UniLabel2: TUniLabel;
    UniLabel1: TUniLabel;
    UniDBEdit2: TUniDBEdit;
    UniDBEdit1: TUniDBEdit;
    ds: TDataSource;
    UniDBEdit4: TUniDBEdit;
    UniLabel3: TUniLabel;
    UniDBEdit3: TUniDBEdit;
    procedure UniFormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

function fEditOrdersPos: TfEditOrdersPos;

implementation

{$R *.dfm}

uses
  MainModule, uniGUIApplication,  System.Math;

function fEditOrdersPos: TfEditOrdersPos;
begin
  Result := TfEditOrdersPos(UniMainModule.GetFormInstance(TfEditOrdersPos));
end;

procedure TfEditOrdersPos.UniFormShow(Sender: TObject);
begin
  eQuantity.DecimalPrecision := IfThen(ds.DataSet.FieldByName('ed_f').AsString = 'F', 0, 4);
end;

end.
