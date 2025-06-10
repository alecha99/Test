unit FrameModules;

interface
  {$WARN SYMBOL_PLATFORM OFF}
  {$I Directives.inc}
uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.UITypes, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, IBX.IBDatabase, IBX.IBCustomDataSet,
  IBX.IBQuery{$IfDef UPDATE}, FrameRolesDB{$EndIf}, Vcl.ExtCtrls, Vcl.StdCtrls, VirtualTrees, VTIHelper,
  System.Generics.Defaults, System.Generics.Collections, ComUtils, ObjModule,
  Data.DB, Vcl.ComCtrls, Vcl.Buttons, System.ImageList, Vcl.ImgList, Vcl.Imaging.jpeg,
  Vcl.Imaging.pngimage, Vcl.ExtDlgs, Vcl.Dialogs, Vcl.Grids, Vcl.DBGrids;


type
  TRecNode =(trNoneN, trBaseN, trModuleN, trVersN);
  TListType = (tltFull, tltAction, tltNoAction);
  TfModulesUser = class(TFrame)
    pRolesDB: TPanel;
    pModuleEdit: TPanel;
    VTI: TVirtualStringTree;
    pModuleData: TPanel;
    qSelect: TIBQuery;
    TransactionRead: TIBTransaction;
    pgInfoFile: TPageControl;
    tshFileInfo: TTabSheet;
    tshRollInf: TTabSheet;
    mRollBack: TMemo;
    mFunctional: TMemo;
    pData: TPanel;
    pbtm: TPanel;
    sbEditCancel: TSpeedButton;
    pRedData: TPanel;
    sbInsertSave: TSpeedButton;
    sbRollback: TSpeedButton;
    pInfoFileName: TPanel;
    pFileName: TPanel;
    eFileName: TEdit;
    lFileName: TLabel;
    pExt: TPanel;
    lExt: TLabel;
    cbExt: TComboBox;
    phelpDispl: TPanel;
    pCatalog: TPanel;
    lCatalog: TLabel;
    eCatalog: TEdit;
    pDescriptor: TPanel;
    lDescriptor: TLabel;
    eDescriptor: TEdit;
    lNPP: TLabel;
    eNPP: TEdit;
    pVersion: TPanel;
    lVersion: TLabel;
    eVersion: TEdit;
    pNPP: TPanel;
    pInfo: TPanel;
    pSize: TPanel;
    lSize: TLabel;
    eSize: TEdit;
    pDateFrom: TPanel;
    lDateFrom: TLabel;
    eDateFrom: TEdit;
    pDateTo: TPanel;
    lDateTo: TLabel;
    eDateTo: TEdit;
    pDopInfo: TPanel;
    pImage: TPanel;
    lImage: TLabel;
    Image: TImage;
    pStatusUsers: TPanel;
    pStatus: TPanel;
    lStatus: TLabel;
    cbStatus: TComboBox;
    pUpdate: TPanel;
    lUpdate: TLabel;
    cbUpdate: TComboBox;
    FileOpenDialog: TOpenDialog;
    sbAddLibrarry: TSpeedButton;
    ImageList: TImageList;
    sbAddVersion: TSpeedButton;
    pCustomer: TPanel;
    lImageText: TLabel;
    sbRoles: TSpeedButton;
    pDopStart: TPanel;
    cbOneStart: TCheckBox;
    cbAllUser: TCheckBox;
    gCustomer: TDBGrid;
    dsCoustomer: TDataSource;
    cbNoExportRegion: TCheckBox;
    procedure VTIResize(Sender: TObject);
    procedure VTICanSplitterResizeColumn(Sender: TVTHeader; P: TPoint;
      Column: TColumnIndex; var Allowed: Boolean);
    procedure sbInsertSaveClick(Sender: TObject);
    procedure sbEditCancelClick(Sender: TObject);
    procedure VTIChange(Sender: TBaseVirtualTree; Node: PVirtualNode);
    procedure lImageTextClick(Sender: TObject);
    procedure sbAddLibrarryClick(Sender: TObject);
    procedure sbRolesClick(Sender: TObject);
    procedure sbAddVersionClick(Sender: TObject);
    procedure sbRollbackClick(Sender: TObject);
    procedure VTIChecking(Sender: TBaseVirtualTree; Node: PVirtualNode;
      var NewState: TCheckState; var Allowed: Boolean);
    procedure pCustomerResize(Sender: TObject);
    procedure VTIDblClick(Sender: TObject);
  private
    fDataBase: TIBDataBase;
    fIsReadOnly: Boolean;
    fHelperVTI: THelperVTI;
    fListModule: TListModule;
    fIsSelected: Boolean;
    fFontVTI: TFont;
    {$IfDef UPDATE}
     fTypeModule: TListType;
     FListNodeSel: TList;
     fRolesDB: TfRolesDB;
     procedure CreateRolesDB;
     function GetIsListRoles: Boolean;
     procedure SetListRoles(const AVisibleList: Boolean);
     procedure SetTypeModule(const ATypeModule: TListType);
    {$EndIf}
    procedure GetBMPFromFile(out ABitMap: TMemoryStream);
    procedure SetReadOnly(const AReadOnly: Boolean);
    function OpenNewFile(const AModuleID: Integer = 0): Boolean;
    {$IfDef UPDATE}
    function AddLibFile(const AModuleID: Integer): Integer;
    function ADDVersionFile(ARecModul: TRecModule;out APathFile: String): TRecVersion;
    {$EndIf}
    procedure FileReadOnly(const AReadOnly: Boolean);
    procedure ChangeModule(const ARead: Boolean);
    procedure GetActiveRec(const ARecNode: TRecNode; AObject: TComponent);
    procedure LoadDataNodes(const ARecIDSel: Integer =-1; const ALibSection: Boolean = False);
    procedure SetIsSelected(const ASelected: Boolean);
    function GetObjFromVTI: TComponent;
  public
    constructor Create(AOwner: TComponent; ADataBase: TIBDataBase;
     {$IfNDef UPDATE} const AListRoles: array of string
     {$Else} const AbaseId: Integer; const ATypeModule: TListType = tltFull{$EndIf}); reintroduce; overload;
    destructor Destroy; override;
    procedure ReoladData;
    procedure ExpandListModule;
    procedure Collapce(const AExpand: Boolean = False);
    property ListModule: TListModule read fListModule;
    property IsReadOnly: Boolean read fIsReadOnly Write SetReadOnly;
    property IsSelected: Boolean Read fIsSelected Write SetIsSelected default False;
    {$IfDef UPDATE}property IsListRoles: Boolean read GetIsListRoles write SetListRoles;
     property TypeModule: TListType read fTypeModule write SetTypeModule;
    {$EndIf}
  end;

implementation
 uses
     VCLRoutine, uThread, System.Math, System.StrUtils
    ,ProcesssObjAbs, VirtualTrees.Types{$IfDef UPDATE}, InfoVersion{$EndIf};
{$R *.dfm}

{ TfModulesUser }
procedure TfModulesUser.Collapce(const AExpand: Boolean);
var
  tmp: TVTIRecBase;
begin
  Self.fHelperVTI.FullCollapce(AExpand);
  tmp:= Self.fHelperVTI.GetNodeBase(2);
  if Assigned(tmp) then
    Self.VTI.FullCollapse(tmp.Node);
  tmp:= Self.fHelperVTI.GetNodeBase(1);
  if Assigned(tmp) then
  begin
    Self.VTI.FullCollapse(tmp.Node);
    Self.VTI.Expanded[tmp.Node] := True;
  end;
end;

constructor TfModulesUser.Create(AOwner: TComponent; ADataBase: TIBDataBase;
    {$IfNDef UPDATE} const AListRoles: array of string {$Else}
    const ABaseID: Integer; const ATypeModule: TListType{$EndIf});
begin
  inherited Create(AOwner);
  Self.fFontVTI := nil;
  Self.FileReadOnly(False);
  Self.fFontVTI := TFont.Create;
  Self.fFontVTI.Color := clYellow;
  Self.fFontVTI.Height := Self.VTI.Font.Height;
  Self.fFontVTI.Charset := Self.VTI.Font.Charset;
  Self.fFontVTI.Name := Self.VTI.Font.Name;
  Self.fFontVTI.Size := 9;//Self.VTI.Font.Size;
  Self.fFontVTI.Style := Self.VTI.Font.Style;
  Self.cbNoExportRegion.Caption := ' Запрет на '#13#10' распростратение'#13#10' в Регионах';
  Self.fHelperVTI := nil;
  {$IfDef UPDATE}
    Self.fRolesDB := nil;
    Self.fTypeModule:= ATypeModule;
  {$Else}
    Self.pRolesDB.Visible:= False;
  {$EndIf}
  Self.fDataBase := ADataBase;
  {$IfDef UPDATE}Self.CreateRolesDB;{$EndIf}
  Self.TransactionRead.DefaultDatabase := Self.fDataBase;
  Self.qSelect.Database := Self.fDataBase;
  Self.qSelect.Transaction := Self.TransactionRead;
  Self.fListModule := TListModule.Create(Self, ADataBase{$IfNDef UPDATE}, AListRoles{$Else}, ABaseID{$EndIf});
  try
    Self.LoadDataNodes;
    Self.ChangeModule(False);
    Self.GetActiveRec(trNoneN, nil);
  except
    on e :Exception do
    begin
      FreeAndNil(Self.fListModule);
      FreeAndNil(Self.fHelperVTI);
      Self.pModuleEdit.Visible := False;
      Self.pRolesDB.Visible := False;
      MessageBox(0, PChar(e.ClassName + ': ' +  e.Message),
              PChar('Ошибка загрузки данных'), MB_OK + MB_ICONERROR);
    end;
  end;
end;

destructor TfModulesUser.Destroy;
begin
  FreeAndNil(Self.fListModule);
  FreeAndNil(Self.fHelperVTI);
  FreeAndNil(Self.fFontVTI);
  {$IfDef UPDATE}
    FreeAndNil(Self.fRolesDB);
    FreeAndNil(Self.FListNodeSel);
  {$EndIf}
  inherited;
end;

procedure TfModulesUser.ExpandListModule;
var
  tmp: TVTIRecBase;
begin
  tmp:= Self.fHelperVTI.GetNodeBase(2);
  if Assigned(tmp) then
    Self.VTI.FullCollapse(tmp.Node);
  tmp:= Self.fHelperVTI.GetNodeBase(1);
  if Assigned(tmp) then
  begin
    Self.VTI.FullCollapse(tmp.Node);
    Self.VTI.Expanded[tmp.Node] := True;
  end;
end;

{$IfDef UPDATE}
function TfModulesUser.AddLibFile(const AModuleID: Integer): Integer;
var
  fn: String;
  i: Integer;
begin
  Result := -1;
  // Ищем новый  файл если он есть таковой
  Self.FileOpenDialog.FileName := '';
  Self.FileOpenDialog.InitialDir :=
     MyPObj.GetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(ParamStr(0)));
  Self.FileOpenDialog.Filter := //'Программа "exe"|*.exe|
    'Библиотека "dll"|*.dll|Файлы конфигураций|*.ini; *.cfg|Все файлы|*.*';
  if Self.FileOpenDialog.Execute(Self.Handle)and Assigned(Self.fListModule) then
  begin
    fn := Self.FileOpenDialog.FileName;

    MyPObj.SetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(fn));
    With Self do
    begin
      {$IfDef UPDATE}
        Result := fListModule.CreateLibFile(fn, AModuleID);
      {$EndIf}
      if Assigned(fListModule.RecRed)then
      begin
        with Self.fListModule  do
        begin
          eVersion.Text := RecRed.Version;
          eDescriptor.Text := RecRed.Descr;
          eFileName.Text := RecRed.FileName;
          cbExt.ItemIndex := -1;
          for i := 0 to Pred(cbExt.Items.Count)do
          begin
            if SameText(cbExt.Items.Strings[i], RecRed.Ext)then
            begin
              cbExt.ItemIndex := i;
              Break;
            end;
          end;
          if(cbExt.ItemIndex < 0)then
          begin
            cbExt.Items.Add(RecRed.Ext);
            cbExt.ItemIndex := Pred(cbExt.Items.Count);
          end;
          eSize.Text := MyPObj.GetStrOfSize(RecRed.Size, true) + ' ';
          cbUpdate.ItemIndex := RecRed.Update;
          eNPP.Text := IntToStr(RecRed.NPP);
          mFunctional.Clear;
          eDateFrom.Text := FormatDateTime('dd.mm.yyyy hh:nn', RecRed.DateFrom);
          Result := 0;
          Self.cbAllUser.Visible := True;
          Self.cbOneStart.Visible := True;
        end;
      end
      else
        Result := IfThen(Result = 0, -1, Result);
    end;
  end;
end;

function TfModulesUser.ADDVersionFile(ARecModul: TRecModule; out APathFile: String): TRecVersion;
var
  ver, s, s2: String;
  rc: TRecVersion;
  sr : TSearchRec;
  sz, sm: Int64;
  pr: Extended;
  b: Boolean;
begin
  Result := nil;
  Self.FileOpenDialog.FileName := '';
  Self.FileOpenDialog.InitialDir :=
     MyPObj.GetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(ParamStr(0)));

  if ARecModul.GetFlagParams(tfpOsn) then
    Self.FileOpenDialog.Filter := 'Программа "exe"|*.exe|Ресурс "bin"|*.bin'
  else
    Self.FileOpenDialog.Filter := 'Обновленная версия файла|'
      + ARecModul.FileName + '.' + LowerCase(ARecModul.Ext);

  if Self.FileOpenDialog.Execute(Self.Handle) then
  begin
    APathFile := Self.FileOpenDialog.FileName;
    ver := MyPObj.GetVersionFile(Self.FileOpenDialog.FileName);
    s2 := ChangeStrChars(ver, cntDelimNum);
    s := ChangeStrChars(ARecModul.Version, cntDelimNum);
    try
      if FindFirst(APathFile, faAnyFile, sr ) = 0 then
        sz := Int64(sr.FindData.nFileSizeHigh) shl Int64(32) + Int64(sr.FindData.nFileSizeLow)
      else
        sz := 0;
    finally
      FindClose(sr) ;
    end;
    if(Length(ChangeStrChars(s2, cntDelimNum)) < 1)then
      ver := '0.0.0.' + IntToStr(StrToInt(ChangeStrChars(ARecModul.Version, cntDelimNum))+ 1);

    {$Region ' Проверка входных параметров '}
    pr := MyPObj.RoundMath((sz - ARecModul.Size) / (ARecModul.Size div 100), 2);
    if(pr > 20)then
    begin
      if MessageBox(0, PChar(
       Format('Размер выбранного файла больше на %0:s %%, ' +
       'возможно вы загружаете версию "Debug" или модуль иной битности,'+
       'что крайне не рекомендовано!' + sLineBreak + sLineBreak +
       'Есть необходимость поднять версию'+ sLineBreak +'  именно выбранным модулем?',[
        MyPObj.CrashSumToString(Double(pr),'.',' ', 3, 2, true)])),
        PChar('Контроль размера файла'),
        MB_YESNO + MB_ICONWARNING + MB_DEFBUTTON2
        )= idNo
      then
        Exit;
    end;
    if(pr < -20)then
    begin
      if MessageBox(0, PChar(
       Format('Размер выбранного файла меньше на %0:s %%, ' +
       'возможно вы загружаете версию "Debug" что крайне не рекомендовано!' + sLineBreak + sLineBreak +
       'Есть необходимость поднять версию'+ sLineBreak +' именно выбранным модулем?',[
        MyPObj.CrashSumToString(Double(pr * -1),'.',' ', 3, 2, true)])),
        PChar('Контроль размера файла'),
        MB_YESNO + MB_ICONWARNING + MB_DEFBUTTON2
        ) = idNo
      then
        Exit;
    end;
    if(sz < 1) then
    begin
      MessageBox(0, PChar('Размер выбранного файла равен 0'),
        PChar('Ошибка выбора файла'), MB_OK + MB_ICONERROR);
      Exit;
    end;
    sm := MyPObj.СheckVersionFile(ARecModul.Version, ver);
    if (sm < 0)then
    begin
      MessageBox(0, PChar(Format('Версия нового файла "%0:s" ниже текущей версии "%1:s"',[ver, ARecModul.Version])),
        PChar('Поднимите версию выше "' + ARecModul.Version +'"'), MB_OK + MB_ICONERROR);
      Exit;
    end;
    if (sm = 0)then
    begin
      MessageBox(0, PChar(Format('Версия нового файла "%0:s" и текущей версии эдентичны',[ver, ARecModul.Version])),
        PChar('Поднимите версию файла выше "' + ARecModul.Version +'"'), MB_OK + MB_ICONERROR);
      Exit;
    end;
    b := False;
    rc:= nil;
    for rc in ARecModul.ListVers do
    begin
      if SameText(ver, rc.Version) then
      begin
        b := True;
        Break;
      end;
    end;
    if b then
    begin
      MessageBox(0, PChar(Format('Версия нового файла "%0:s" соответствует обновлению №%1:d',[ver, rc.NPP])),
        PChar('Поднимите версию файла выше "' + ARecModul.Version +'"'), MB_OK + MB_ICONERROR);
      Exit;
    end;
    {$EndRegion}
    sm := 0;
    for rc in ARecModul.ListVers do
    begin
      if(rc.NPP > sm)then
        sm := rc.NPP;
    end;
    sm := sm + 1;
    Result := TRecVersion.Create(Self.fListModule, ARecModul, sm);
    Result.SetFlagParams(tfpvActive, True);
    Result.SetFlagParams(tfpvRollBack, False);
    Result.Version := ver;
    Result.DateFrom := MyPObj.GetDateFile(APathFile, 0);
    Result.Size := sz;
    Result.FileName := StringReplace(ver,'.','-',[rfReplaceAll, rfIgnoreCase]);
    Result.FileName := ExtractFileName(ChangeFileExt(APathFile, ''))+ '('+ Result.FileName +')';
    Result.FileName := ARecModul.FileName + '.' + LowerCase(ARecModul.Ext)  + '\' + Result.FileName + ExtractFileExt(APathFile);
  end;
end;
{$EndIf}

procedure TfModulesUser.ChangeModule(const ARead: Boolean);
begin
  with Self do
  begin
    sbInsertSave.ImageIndex := IfThen(ARead, 3, 1);
    sbEditCancel.ImageIndex := IfThen(ARead, 4, 2);
    sbRoles.Visible := not ARead ;
    sbAddLibrarry.Enabled := not ARead;
    sbAddLibrarry.Visible := not ARead;
    sbAddVersion.Enabled := not ARead;
    sbRollback.Enabled := not ARead;
    pDopStart.Enabled := ARead;
    pStatusUsers.Enabled := ARead;
    sbEditCancel.Hint := IfThen(not ARead,
            'Редактировать текущую запись',
            'Отменить редактор');
    sbInsertSave.Hint := IfThen(not ARead,
            'Добавить новый модуль',
            'Сохранить в базу');
    if Assigned(fListModule) then
    begin
      if Assigned(fListModule.RecRed) then
      begin
        pExt.Enabled := ARead and fListModule.RecRed.GetFlagParams(tfpOsn);
        eCatalog.ReadOnly := not(ARead or fListModule.RecRed.GetFlagParams(tfpOsn));
        eDescriptor.ReadOnly := not(ARead and
            fListModule.RecRed.GetFlagParams(tfpOsn)and
            (Length(Trim(fListModule.RecRed.Descr)) > 0)
            );
      end
      else
      begin
        pExt.Enabled := False;
        eCatalog.ReadOnly := True;
        eDescriptor.ReadOnly := True;
      end;
    end
    else
    begin
      pExt.Enabled := False;
      eCatalog.ReadOnly := True;
      eDescriptor.ReadOnly := True;
    end;
    pFileName.Enabled := True;
    eFileName.ReadOnly := True;
    eDescriptor.ReadOnly := True;

    pUpdate.Enabled :=  ARead;
//    pRedData.Enabled := ARead;
    mRollBack.ReadOnly := not ARead;
    mFunctional.ReadOnly := not ARead;
    VTI.Enabled := not ARead;
    {$IfDef UPDATE}
    fRolesDB.IsReadOnly := True;
    pRolesDB.Enabled := not ARead;
    {$EndIf}
    if ARead then
      sbEditCancel.Enabled := ARead;
  end;
end;

{$IfDef UPDATE}
procedure TfModulesUser.CreateRolesDB;
begin
  if Assigned(Self.fDataBase) then
  begin
    if Self.fDataBase.Connected then
    begin
      Self.fRolesDB := TfRolesDB.Create(Self, Self.fDatabase);
      Self.fRolesDB.Parent:= Self.pRolesDB;
      Self.fRolesDB.Align:= alClient;
      Self.fRolesDB.IsReadOnly := True;
      Self.fRolesDB.ColorIsChecked := False;
      Self.fRolesDB.sbFilterDBOff.Visible := False;
      Self.fRolesDB.sbFilterDBOn.Visible := False;
    end;
  end;
end;
{$EndIf}

procedure TfModulesUser.FileReadOnly(const AReadOnly: Boolean);
begin
  With Self do
  begin
    pExt.Enabled := not AReadOnly;
    pUpdate.Enabled := not AReadOnly;
    pStatusUsers.Enabled := not AReadOnly;
    eCatalog.ReadOnly := AReadOnly;
  end;
end;

procedure TfModulesUser.GetActiveRec(const ARecNode: TRecNode; AObject: TComponent);

  function IsRoleModule(AModule: TRecModule; const ARoleID: Integer): Boolean;
  var
    osn: Boolean;
    {$IfDef UPDATE}
      i: Integer;
    {$EndIf}
  begin
     osn := AModule.GetFlagParams(tfpOsn);
    if not Osn or AModule.GetFlagParams(tfpAllUser) then
      Result := True
    else
    begin
      Result := False;
      {$IfDef UPDATE}
      for i := 0 to Pred(AModule.ListRole.Count) do
      begin
        if(AModule.ListRole[i].ID = ARoleID) then
        begin
          Result := True;
          Break;
        end;
      end;
      {$EndIf}
    end;
  end;

var
  s: String;
  i: Integer;
  {$IfDef UPDATE}
    role: Integer;
  {$EndIf}
begin
  try
    with Self do
    begin
      cbAllUser.Visible := False;
      cbOneStart.Visible := False;
      sbRoles.Enabled := False;
      case ARecNode of
        trNoneN, trBaseN:
        {$Region ' Пустые базовые теги '}
        begin
          lNPP.Caption := '  Node';
          {$IfDef UPDATE}
          if pRolesDB.Visible then
          begin
            if fRolesDB.gRolesDB.DataSource.DataSet.Active then
            begin
              if(fRolesDB.gRolesDB.DataSource.DataSet.RecordCount > 0)then
              begin
                role := Self.fRolesDB.gRolesDB.DataSource.DataSet.FieldByName('recid').AsInteger;
                for i := 0 to Pred(fRolesDB.ListRole.Count) do
                  fRolesDB.ListRole[i].CodeColor := IfThen(Assigned(AObject) and( TRecBase(AObject).CodeNode < 3) , 1, 0);
                Self.fRolesDB.gRolesDB.DataSource.DataSet.Last;
                Self.fRolesDB.gRolesDB.DataSource.DataSet.First;
                Self.fRolesDB.gRolesDB.DataSource.DataSet.Locate('recid', role,[loCaseInsensitive, loPartialKey]);
              end;
            end;
          end;
          {$EndIf}
          cbNoExportRegion.Visible := false;
          eNPP.Text := IfThen(Assigned(AObject), 'Есть ', 'Нет ');
          eCatalog.Clear;
          if Assigned(AObject) then
          begin
            if(TRecBase(AObject).CodeNode < 3)then
              eFileName.Text := 'Список = ' + IntToStr(TRecBase(AObject).CodeNode)
            else
              eFileName.Text := 'Список отображения истории версий';
            eCatalog.Clear;
            TRecBase(AObject).GetText(1, s);
            eDescriptor.Text := s;
            TRecBase(AObject).GetText(5, s);
            Self.mFunctional.Text :=  s;
          end
          else
          begin
            eFileName.Clear;
            eDescriptor.Clear;
            Self.mFunctional.Clear;
          end;
          cbExt.ItemIndex := -1;
          cbUpdate.ItemIndex := -1;
          cbAllUser.Checked := False;
          cbOneStart.Visible := False;
          pImage.Visible := False;
          cbStatus.ItemIndex := -1;
          eSize.Clear;
          eDateFrom.Clear;
          eDateTo.Clear;
          eVersion.Clear;
          tshRollInf.TabVisible := False;
          sbAddLibrarry.Visible := False;
          sbAddVersion.Enabled := False;
          sbRollback.Enabled := False;
          sbRoles.Enabled := False;
          sbEditCancel.Enabled := False;
        end;
        {$EndRegion}
         trModuleN:
        {$Region ' Теги модулей и библиотек '}
        begin
          {$IfDef UPDATE}
          if pRolesDB.Visible then
          begin
            if fRolesDB.gRolesDB.DataSource.DataSet.Active then
            begin
              if(fRolesDB.gRolesDB.DataSource.DataSet.RecordCount > 0)then
              begin
                role := Self.fRolesDB.gRolesDB.DataSource.DataSet.FieldByName('recid').AsInteger;
                for i := 0 to Pred(fRolesDB.ListRole.Count) do
                  fRolesDB.ListRole[i].CodeColor := IfThen(IsRoleModule(TRecModule(AObject),
                    fRolesDB.ListRole[i].ID), 1, 2);
                Self.fRolesDB.gRolesDB.DataSource.DataSet.Last;
                Self.fRolesDB.gRolesDB.DataSource.DataSet.First;
                Self.fRolesDB.gRolesDB.DataSource.DataSet.Locate('recid',role,[loCaseInsensitive, loPartialKey]);
              end;
            end;
          end;
          {$EndIf}
          lNPP.Caption := '  Запуск №';
          if TRecModule(AObject).GetFlagParams(tfpOsn) then
            eNPP.Text := IntToStr(IfThen(Assigned(AObject), TRecModule(AObject).NPP, -1))
          else
            eNPP.Text := 'По запросу ';
          eCatalog.Text := TRecModule(AObject).FileDir;
          eFileName.Text := TRecModule(AObject).FileName;
          sbRoles.Enabled := not TRecModule(AObject).GetFlagParams(tfpAllUser) and
                            TRecModule(AObject).GetFlagParams(tfpOsn);
          eDescriptor.Text := TRecModule(AObject).Descr;
          cbNoExportRegion.Visible := true;
          cbNoExportRegion.Checked := TRecModule(AObject).GetFlagParams(tfpNoExportRegion);
          tshRollInf.TabVisible := TRecModule(AObject).Active > 1;
          if(TRecModule(AObject).Active > 1)then
          begin
            mRollBack.Lines.Text := TRecModule(AObject).InfoClose;
            pgInfoFile.ActivePageIndex := 1;
          end;
          mFunctional.Text :=  TRecModule(AObject).Info;
          cbAllUser.Checked := TRecModule(AObject).GetFlagParams(tfpAllUser);
          cbOneStart.Checked := TRecModule(AObject).GetFlagParams(tfpOneStart);
          eSize.Text := MyPObj.GetStrOfSize(TRecModule(AObject).Size, True);
          eDateFrom.Text := FormatDateTime('dd.mm.yyyy hh:nn', TRecModule(AObject).DateFrom);
          lDateTo.Caption := 'Блокировка';
          eDateTo.Text := 'Без ограничений';
          if (TRecModule(AObject).Active > 1) then
            eDateTo.Text := FormatDateTime('dd.mm.yyyy hh:nn', TRecModule(AObject).DateTo);
          eVersion.Text := TRecModule(AObject).Version;
          s := AnsiLowerCase(TRecModule(AObject).Ext);
          cbExt.ItemIndex := -1;
          for i := 0 to Pred(cbExt.Items.Count)do
          begin
            if SameText(cbExt.Items.Strings[i], s)then
            begin
              cbExt.ItemIndex := i;
              Break;
            end;
          end;
          if(cbExt.ItemIndex < 0)then
          begin
            cbExt.Items.Add(s);
            cbExt.ItemIndex := Pred(cbExt.Items.Count);
          end;
          if TRecModule(AObject).GetFlagParams(tfpOsn) then
          begin
            Self.lImageText.Visible := not Assigned(TRecModule(AObject).Image);
            Self.Image.Visible := Assigned(TRecModule(AObject).Image);
            if Assigned(TRecModule(AObject).Image) then
            begin
              TRecModule(AObject).Image.Seek(0, soBeginning);
              Image.Picture.LoadFromStream(TRecModule(AObject).Image);
            end;
          end;
          lImage.Visible := TRecModule(AObject).GetFlagParams(tfpOsn);
          lImageText.Visible := TRecModule(AObject).GetFlagParams(tfpOsn)and not Assigned(TRecModule(AObject).Image);
          pImage.Visible := TRecModule(AObject).GetFlagParams(tfpOsn);
          cbUpdate.ItemIndex := TRecModule(AObject).Update;
          cbStatus.ItemIndex := TRecModule(AObject).Active;
          sbAddVersion.Enabled := not TRecModule(AObject).IsRedOnly and (TRecModule(AObject).Active < 2);
          sbAddLibrarry.Visible := TRecModule(AObject).GetFlagParams(tfpOsn)and (TRecModule(AObject).Active < 2);
          sbRollback.Enabled := TRecModule(AObject).Active < 2;
          sbRollback.ImageIndex := 8;
          sbEditCancel.Enabled := not TRecModule(AObject).IsRedOnly{ and (TRecModule(AObject).Active < 2)};
          cbAllUser.Visible := TRecModule(AObject).GetFlagParams(tfpOsn) and (TRecModule(AObject).Active < 2);
          cbOneStart.Visible := cbAllUser.Visible;
          lVersion.Caption := '  Текущая версия';
        end;
        {$EndRegion}
        trVersN:
        begin
          cbNoExportRegion.Visible := true;
          sbEditCancel.Enabled := False;
          sbAddLibrarry.Visible := False;
          sbRoles.Enabled := False;
          sbAddVersion.Enabled:= False;
          sbRollback.Enabled := True;
          sbRollback.ImageIndex := 6;
          eNPP.Text := IntToStr(TRecVersion(AObject).NPP);
          lNPP.Caption := '  Версия №';
          eFileName.Text := TRecVersion(AObject).FileName;
          tshRollInf.TabVisible := TRecVersion(AObject).GetFlagParams(tfpvRollBack);
          cbNoExportRegion.Checked := TRecModule(AObject).GetFlagParams(tfpNoExportRegion);
          mFunctional.Lines.Text := TRecVersion(AObject).Info;
          if TRecVersion(AObject).GetFlagParams(tfpvRollBack) then
          begin
            mRollBack.Lines.Text := TRecVersion(AObject).InfRollBack;
            pgInfoFile.ActivePageIndex := 1;
            sbRollback.Enabled := False;
          end;
          cbExt.ItemIndex := -1;
          lVersion.Caption := '  Номер версии';
          eVersion.Text := TRecVersion(AObject).Version;
          eCatalog.Text := ' Каталоги родителей';
          eDescriptor.Text := '  Иформация о версии файла ';
          eSize.Text :=  MyPObj.GetStrOfSize(TRecVersion(AObject).Size, True);
          eDateFrom.Text := FormatDateTime('dd.mm.yyyy hh:nn', TRecVersion(AObject).DateFrom);
          eDateTo.Text := ' Порядок версий';
          if TRecVersion(AObject).GetFlagParams(tfpvRollBack) then
            eDateTo.Text := FormatDateTime(' dd.mm.yyyy hh:nn', TRecVersion(AObject).DateTo);
          lImage.Visible := False;
          lImageText.Visible := False;
          Image.Visible := False;
          cbAllUser.Visible := False;
          cbOneStart.Visible := False;
          cbStatus.ItemIndex := -1;
          cbUpdate.ItemIndex := -1;
        end;
      end;
    end;
  except
    on e: Exception do
     VCL.Dialogs.MessageDlg(
        e.ClassName + ': ' +e.Message +'({0EA65701-49D9-437F-9BA9-A741D52BE809} .GetActiveRec)'
          ,mtError , [mbIgnore], 0);
  end;
end;

procedure TfModulesUser.GetBMPFromFile(out ABitMap: TMemoryStream);
var
  bmp: TBitmap;
  jpg: TJPEGImage;
  png: TPngImage;
  ms: TMemoryStream;
begin
  ABitMap := nil;
  Self.FileOpenDialog.InitialDir :=
     MyPObj.GetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(ParamStr(0)));
  Self.FileOpenDialog.Filter := 'Файлы изображений (.jpg, .jpeg, .bmp, .png)|*.jpg;*.bmp;*.jpeg;*.png';

  if not FileOpenDialog.Execute then
    Exit;

  bmp:= TBitmap.Create;

  ms := TMemoryStream.Create;
  try
    if AnsiSameText(LowerCase(ExtractFileExt(FileOpenDialog.FileName)),'.bmp') then
     bmp.LoadFromFile(FileOpenDialog.FileName)
    else
    if AnsiSameText(LowerCase(ExtractFileExt(FileOpenDialog.FileName)),'.jpg')or
       AnsiSameText(LowerCase(ExtractFileExt(FileOpenDialog.FileName)),'.jpeg')
    then
    begin
      jpg :=TJPEGImage.Create;
      try
        jpg.LoadFromFile(FileOpenDialog.FileName);
        bmp.Assign(jpg);
      finally
        FreeAndNil(jpg);
      end;
    end
    else
    if AnsiSameText(LowerCase(ExtractFileExt(FileOpenDialog.FileName)),'.png') then
    begin
      png := TPngImage.Create;
      try
        png.LoadFromFile(FileOpenDialog.FileName);
        bmp.Assign(png);
      finally
        FreeAndNil(png);
      end;
    end;
    bmp.SaveToStream(ms);
    if(ms.Size > 0)then
    begin
      Self.Image.Picture.Assign(bmp);
      Self.lImageText.Visible := False;
      Self.Image.Visible := True;
      ms.Seek(0, soBeginning);
      ABitMap:= TMemoryStream.Create;
      ABitMap.LoadFromStream(ms);
    end;
  finally
    ms.Free;
    bmp.Free;
  end;
end;

{$IfDef UPDATE}
function TfModulesUser.GetIsListRoles: Boolean;
begin
  Result := Assigned(Self.fRolesDB);
end;
{$EndIf}

function TfModulesUser.GetObjFromVTI: TComponent;
var
   tmp: pVTIRecBase;
   obj: TComponent;
begin
  Result := nil;
  if(Self.VTI.GetFirstSelected(true)^.States = [vsOnFreeNodeCallRequired])then

  else
  begin
    tmp := Self.VTI.GetNodeData(Self.VTI.GetFirstSelected(true));
    if tmp^.InheritsFrom(TRecBase) then
          Self.GetActiveRec(trBaseN, tmp^)
    else
    begin
      obj := tmp^.DataObject;
      if Assigned(obj)then
        Result := obj;
    end;
  end;
end;

procedure TfModulesUser.lImageTextClick(Sender: TObject);
var
  ms: TMemoryStream;
begin
  if Assigned(fListModule) then
  begin
    if Assigned(Self.fListModule.RecRed) then
    begin
      if Self.fListModule.RecRed.GetFlagParams(tfpOsn)then
      begin
        Self.GetBMPFromFile(ms);
        try
          if Assigned(ms) then
            Self.fListModule.RecRed.Image := ms;
        finally
          FreeAndNil(ms);
        end;
      end;
    end;
  end;
end;

procedure TfModulesUser.LoadDataNodes(const ARecIDSel: Integer =-1; const ALibSection: Boolean = False);
var
  rcv, rcl: TRecModulVTI;
  rv: TRecVersionVTI;
  rb: TRecBase;
  i: integer;
  c: Integer;
 {$IfDef UPDATE} bb: Boolean; {$EndIf}
  NodeSel: PVirtualNode;
begin
  NodeSel := nil;
  if Assigned(Self.fHelperVTI) then
    FreeAndNil(Self.fHelperVTI);
  Self.VTI.BeginUpdate;
  try

    Self.fHelperVTI := THelperVTI.Create(Self, Self.VTI);
    Self.fHelperVTI.MaxLinesCount := 4;
    Self.fHelperVTI.NoDublicateNode := False;
    Self.fHelperVTI.CellNodeFrame := True;
    Self.fHelperVTI.VTIDisplayText := 0;
    Self.fHelperVTI.SelColorBrush := clBlack;// clMoneyGreen;
    Self.fHelperVTI.SelFont := Self.fFontVTI;// clMoneyGreen;
    rb := TRecBase.Create(1,Self.VTI.Font, nil);
    Self.fHelperVTI.AddBase(0, rb);
    rb := TRecBase.Create(2,Self.VTI.Font, nil);
    Self.fHelperVTI.AddBase(0, rb);
    {$IfDef UPDATE}
      Self.FListNodeSel:= TList.Create;
    {$EndIf}

    for i := 0 to Pred(Self.fListModule.ListAPP.Count) do
    begin
      rcv := TRecModulVTI.Create(Self.fListModule.ListAPP[i], Self.VTI.Font);
      case Self.fListModule.ListAPP[i].Active of
        0: rcv.BrushColor := 9697271;
        1: rcv.BrushColor := 15191717;//clSkyBlue;//15269887;// clSkyBlue;
      else
        rcv.BrushColor := 14409725;
      end;
      if(Self.fListModule.ListAPP[i].Update < 1)then
          rcv.BrushColor := 14409725;
      Self.fHelperVTI.AddData(1, rcv);
      {$IfDef UPDATE}
       // Скрываем ненужный хлам
       case Self.fTypeModule of
         tltAction:
           bb := (Self.fListModule.ListAPP[i].Update < 1) or (Self.fListModule.ListAPP[i].Active > 1);
         tltNoAction:
           bb := (Self.fListModule.ListAPP[i].Update > 0) and (Self.fListModule.ListAPP[i].Active < 2);
       else
         bb:= False;
       end;
       if bb and Assigned(rcv.Node) then
         Self.VTI.IsVisible[rcv.Node]:= False;

       if //not Self.fListModule.ListAPP[i].GetFlagParams(tfpNoExportRegion)
         //(Self.fListModule.ListAPP[i].Active < 2) and
         (Self.fListModule.ListAPP[i].Update > 0)
       then
         Self.FListNodeSel.Add(rcv);
      {$EndIf}

      if (ARecIDSel > 0)and not(ALibSection) then
        if(Self.fListModule.ListAPP[i].RecID = ARecIDSel)then
          NodeSel := rcv.Node;
      {$IfDef UPDATE}
      for c := 0 to Pred(Self.fListModule.ListAPP[i].ListLib.Count) do
      begin
        rcl := TRecModulVTI.Create(Self.fListModule.ListAPP[i].ListLib[c], Self.VTI.Font);
        case Self.fListModule.ListAPP[i].Listlib[c].Active of
          0: rcl.BrushColor := 9697271;
          1: rcl.BrushColor := 16571269;// clSkyBlue;
        else
          rcl.BrushColor := 14409725;
        end;
        if(Self.fListModule.ListAPP[i].ListLib[c].Update < 1)then
          rcl.BrushColor := 14409725;
        Self.fHelperVTI.AddData(rcv, rcl);
        // Скроем мусор
        case Self.fTypeModule of
          tltAction:
            bb := (Self.fListModule.ListAPP[i].ListLib[c].Update < 1) or
                        (Self.fListModule.ListAPP[i].ListLib[c].Active > 1);
//          tltNoAction:
//            bb := (Self.fListModule.ListAPP[i].ListLib[c].Update > 0) and
//                        (Self.fListModule.ListAPP[i].ListLib[c].Active < 2);
        else
          bb:= False;
        end;
        if bb and Assigned(rcl.Node) then
        begin
          Self.VTI.IsVisible[rcl.Node]:= False;
        end;
      end;
      {$EndIf}
      rb:= TRecBase.Create(3, Self.VTI.Font, nil);
      rb.BrushColor := clMedGray;
      Self.fHelperVTI.AddData(rcv, rb);
      for c := Pred(Self.fListModule.ListAPP[i].ListVers.Count) downto 0 do
      begin
        rv := TRecVersionVTI.Create(Self.fListModule.ListAPP[i].ListVers[c], Self.VTI.Font);
        if Self.fListModule.ListAPP[i].ListVers[c].GetFlagParams(tfpvRollBack) then
          rv.BrushColor := 14409725
        else
        begin
          if AnsiSameText(Self.fListModule.ListAPP[i].ListVers[c].Version, Self.fListModule.ListAPP[i].Version) then
            rv.BrushColor := 9043313//clInfoBk;//clCream;
          else
            rv.BrushColor := clCream;
        end;
        Self.fHelperVTI.AddData(rb, rv);
      end;
    end;
    for i := 0 to Pred(Self.fListModule.Listlib.Count) do
    begin
      rcl := TRecModulVTI.Create(Self.fListModule.Listlib[i], Self.VTI.Font);
      case Self.fListModule.Listlib[i].Active of
        0: rcl.BrushColor := 9697271;
        1: rcl.BrushColor := 16571269;// clSkyBlue;
      else
        rcl.BrushColor := 14409725;
      end;
      if(Self.fListModule.Listlib[i].Update < 1)then
          rcl.BrushColor := 14409725;

      {$IfDef UPDATE}
      case Self.fTypeModule of
        tltAction:
          bb := (Self.fListModule.Listlib[i].Update < 1) or (Self.fListModule.Listlib[i].Active > 1);
        tltNoAction:
          bb := (Self.fListModule.Listlib[i].Update > 0) and (Self.fListModule.Listlib[i].Active < 2);
      else
        bb:= False;
      end;
      {$EndIf}

      Self.fHelperVTI.AddData(2, rcl);
      {$IfDef UPDATE}
      if (Self.fListModule.Listlib[i].Active < 2)and
         (Self.fListModule.Listlib[i].Update > 0)
      then
        Self.FListNodeSel.Add(rcl);
      if bb and Assigned(rcl.Node) then
      begin
        Self.VTI.IsVisible[rcl.Node]:= False;
      end;
      {$EndIF}
      if (ARecIDSel > 0)and(ALibSection) then
        if(Self.fListModule.Listlib[i].RecID = ARecIDSel)then
          NodeSel := rcl.Node;

      rb:= TRecBase.Create(3, Self.VTI.Font, nil);
      rb.BrushColor := clMedGray;
      Self.fHelperVTI.AddData(rcl, rb);
      for c := Pred(Self.fListModule.Listlib[i].ListVers.Count) downto 0 do
      begin
        rv := TRecVersionVTI.Create(Self.fListModule.Listlib[i].ListVers[c], Self.VTI.Font);
        if Self.fListModule.Listlib[i].ListVers[c].GetFlagParams(tfpvRollBack) then
          rv.BrushColor := 14409725
        else
        begin
          if AnsiSameText(Self.fListModule.Listlib[i].ListVers[c].Version, Self.fListModule.Listlib[i].Version) then
            rv.BrushColor := 9043313//clInfoBk;//clCream;
          else
            rv.BrushColor := clCream;
        end;
        Self.fHelperVTI.AddData(rb, rv);
      end;
    end;
    // Открываем нужный Node
    if Assigned(NodeSel)then
    begin
      Self.VTI.FullExpand(NodeSel.Parent);
        Self.VTI.FocusedNode := NodeSel;
        Self.VTI.Selected[NodeSel]:= true;
    end;
  finally
    Self.VTI.EndUpdate;
  end;
end;

function TfModulesUser.OpenNewFile(const AModuleID: Integer): Boolean;
{$IfDef UPDATE}
var
  fn: String;
  i: Integer;
{$EndIf}
begin
  Result := False;
  {$IfDef UPDATE}
  // Ищем новый  файл если он есть таковой
  Self.FileOpenDialog.FileName := '';
  Self.FileOpenDialog.InitialDir :=
     MyPObj.GetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(ParamStr(0)));
  Self.FileOpenDialog.Filter := 'Программа "exe"|*.exe';//|Библиотека "dll"|*.dll|Файлы конфигураций|*.ini;*.cfg|Все файлы|*.*';
  if Self.FileOpenDialog.Execute(Self.Handle) then
  begin
    fn := Self.FileOpenDialog.FileName;

    MyPObj.SetIniParams('ModulUpdateServer', 'InPathFile', ExtractFilePath(fn));
    With Self do
    begin
      fListModule.CreateRedFile(fn, AModuleID);
      if Assigned(fListModule.RecRed) then
      begin
        with Self.fListModule  do
        begin
          eVersion.Text := RecRed.Version;
          eDescriptor.Text := RecRed.Descr;
          eFileName.Text := RecRed.FileName;
          cbExt.ItemIndex := -1;
          for i := 0 to Pred(cbExt.Items.Count)do
          begin
            if SameText(cbExt.Items.Strings[i], RecRed.Ext)then
            begin
              cbExt.ItemIndex := i;
              Break;
            end;
          end;
          if(cbExt.ItemIndex < 0)then
          begin
            cbExt.Items.Add(RecRed.Ext);
            cbExt.ItemIndex := Pred(cbExt.Items.Count);
          end;
          eSize.Text := MyPObj.GetStrOfSize(RecRed.Size, true) + ' ';
          cbUpdate.ItemIndex := RecRed.Update;
          eNPP.Text := IntToStr(RecRed.NPP);
          mFunctional.Clear;
          eDateFrom.Text := FormatDateTime('dd.mm.yyyy hh:nn', RecRed.DateFrom);
          Result := True;
          Self.cbAllUser.Visible := True;
          Self.cbOneStart.Visible := True;
        end;
      end;
    end;
  end;
  {$EndIf}
end;

procedure TfModulesUser.pCustomerResize(Sender: TObject);
begin
  ResizeCustomGrid(Self.gCustomer, 0);
  Self.gCustomer.Columns[0].Width := Self.gCustomer.Columns[0].Width -1;
end;

procedure TfModulesUser.ReoladData;
begin
  Self.VTI.BeginUpdate;
  try
    Self.VTI.FullCollapse();
    {$IfDef UPDATE}
      FreeAndNil(Self.FListNodeSel);
    {$EndIf}
  //  Self.fHelperVTI.ClearDataNode;
    FreeAndNil(Self.fHelperVTI);
  finally
    Self.VTI.EndUpdate;
  end;
  Self.fListModule.ReoladFile;
  Self.ChangeModule(False);
  Self.LoadDataNodes(1, False);
end;

procedure TfModulesUser.sbAddLibrarryClick(Sender: TObject);
 var
   obj: TComponent;
   {$IfDef UPDATE}
   i: Integer;
   {$EndIf}
begin
  With Self do
  begin
    obj := GetObjFromVTI;
    if Assigned(obj)then
    begin
      if obj.InheritsFrom(TRecModule)then
      begin
        pgInfoFile.ActivePageIndex := 0;
        pgInfoFile.Pages[1].TabVisible := False;
        pImage.Visible := False;
        {$IfDef UPDATE}
        i := AddLibFile(TRecModule(obj).RecID);
        if( i = 0)then
        begin
          fListModule.RecRed.SetFlagParams(tfpOsn, False);
          ChangeModule(True);
          eCatalog.ReadOnly := True;
          cbAllUser.Visible := False;
          cbOneStart.Visible := False;
        end
        else
        begin
          if( i > 0)then
          begin
            VTI.BeginUpdate;
            try
              VTI.FullCollapse();
              fHelperVTI.ClearDataNode;
              ChangeModule(False);
              LoadDataNodes(i, True);
            finally
              VTI.EndUpdate;
            end;
          end;
        end;
        {$EndIf}
      end;
    end;
  end;
end;

procedure TfModulesUser.sbAddVersionClick(Sender: TObject);
{$IfDef UPDATE}
var
   obj: TComponent;
   PathFile: String;
   RecVers: TRecVersion;
   NewRecID: Integer;
   Lib: Boolean;
{$EndIf}
begin
{$IfDef UPDATE}
  With Self do
  begin
    obj := GetObjFromVTI;
    if Assigned(obj)then
    begin
      if obj.InheritsFrom(TRecModule)then
      begin
        RecVers := ADDVersionFile(TRecModule(obj), PathFile);
        try
          if Assigned(RecVers)then
          begin
            fVersionInfo := TfVersionInfo.Create(Self, 'Добавлен функционал в новую версию');
            try
              if fVersionInfo.ShowModal = mrOk then
              begin
                RecVers.Info := fVersionInfo.mInfo.Lines.Text;
                Lib := not RecVers.BaseModul.GetFlagParams(tfpOsn);
                NewRecID := RecVers.BaseModul.RecID;
                if Self.fListModule.SaveVersion(RecVers, PathFile)then
                begin
                  Self.VTI.FullCollapse();
                  Self.FListNodeSel.Clear;
//                  Self.fHelperVTI.ClearDataNode;
                  Self.LoadDataNodes(NewRecID, Lib);
                  Self.ChangeModule(False);
                end;
              end;
            finally
              FreeAndNil(fVersionInfo);
            end;
          end;
        finally
          FreeAndNil(RecVers);
        end;
      end;
    end;
  end;
{$EndIf}
end;

procedure TfModulesUser.sbEditCancelClick(Sender: TObject);
 var
   obj: TComponent;
begin
  if Assigned(Self.fListModule.RecRed) then
  begin
    FreeAndNil(Self.fListModule.RecRed);
    Self.ChangeModule(False);
    Self.VTIChange(Self.VTI, Self.VTI.GetFirstSelected(true));
  end
  else
  begin
    With Self do
    begin
      if not (sbRoles.ImageIndex = 5) then
      begin
        sbRoles.ImageIndex := 5;
        VTI.Enabled := (sbRoles.ImageIndex = 5) ;
        sbAddLibrarry.Visible := (sbRoles.ImageIndex = 5);
        {$IfDef UPDATE}
        fRolesDB.IsReadOnly := (sbRoles.ImageIndex = 5);
        fRolesDB.ColorIsChecked := (sbRoles.ImageIndex = 3);
        {$EndIf}
        sbEditCancel.ImageIndex := IfThen((sbRoles.ImageIndex = 5), 4, 2);
        sbInsertSave.Enabled := (sbRoles.ImageIndex = 5);
        sbAddVersion.Enabled := (sbRoles.ImageIndex = 5);
        sbRollback.Enabled := (sbRoles.ImageIndex = 5);
        Self.ChangeModule(False);
      end
      else
      begin
        obj := GetObjFromVTI;
        if Assigned(obj)then
        begin
          if obj.InheritsFrom(TRecModule)then
          begin
            if Assigned(obj)then
            begin
              if obj.InheritsFrom(TRecModule)then
              begin
                pgInfoFile.ActivePageIndex := 0;
                pgInfoFile.Pages[1].TabVisible := False;
                pImage.Visible := TRecModule(obj).GetFlagParams(tfpOsn);
                {$IfDef UPDATE}
                Self.fListModule.CreateRedFile(TRecModule(obj));
                if Assigned(Self.fListModule.RecRed) then
                  Self.ChangeModule(True);
                {$EndIf}
              end;
            end;
          end;
        end;
      end;
    end;
  end;
end;

procedure TfModulesUser.sbInsertSaveClick(Sender: TObject);
var
  NewRecID: Integer;
  Lib: Boolean;
begin
  with Self do
  begin
    if Assigned(fListModule.RecRed) then
    begin
      if Length(Trim(mFunctional.Lines.Text))= 0 then
      begin
        MessageBox(0, PChar('Не заполнен  "Основной функционал", с таким подходом идите ка вы в лес!'),
            PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
        Exit;
      end
      else
      begin
        if(cbExt.ItemIndex < 0)then
        begin
          MessageBox(0, PChar('Не указано конечное расширение файла, нет расширения нет сохранности!'),
              PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
          Exit;
        end
        else
        begin
          if(cbUpdate.ItemIndex < 0)then
          begin
            MessageBox(0, PChar('Не указан способ идентификации файла, и как это понимать?'),
                PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
            Exit;
          end
          else
          begin
            if(cbStatus.ItemIndex < 0)then
            begin
              MessageBox(0, PChar('Не указан способ идентификации файла, сами искать будете?'),
                  PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
              Exit;
            end
            else
            begin
              if not Assigned(fListModule.RecRed.Image)and fListModule.RecRed.GetFlagParams(tfpOsn) then
              begin
                MessageBox(0, PChar('Без иконки выбора, а стреляете тоже в слепую?'),
                  PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
                Exit;
              end
              else
              begin
                if (cbStatus.ItemIndex = 2)and(sbRollback.Tag = 0) then
                begin
                  MessageBox(0, PChar('Откат/Закытие версий и программ имеет отдельную операцию'),
                    PChar('Ошибка заполнения данных'), MB_OK + MB_ICONERROR);
                  Exit;
                end
                else
                begin
                  Lib := not fListModule.RecRed.GetFlagParams(tfpOsn);
                  fListModule.RecRed.Info := mFunctional.Lines.Text;
                  fListModule.RecRed.Update := cbUpdate.ItemIndex;
                  fListModule.RecRed.Active := cbStatus.ItemIndex;
                  fListModule.RecRed.SetFlagParams(tfpNoExportRegion, cbNoExportRegion.Checked);
                  fListModule.RecRed.Ext := cbExt.Items.Strings[cbExt.ItemIndex];
                  if fListModule.RecRed.GetFlagParams(tfpOsn) then
                  begin
                    fListModule.RecRed.SetFlagParams(tfpAllUser, cbAllUser.Checked);
                    fListModule.RecRed.SetFlagParams(tfpOneStart, cbOneStart.Checked);
                    fListModule.RecRed.FileDir := eCatalog.Text;
                  end
                  else
                  begin
                    fListModule.RecRed.SetFlagParams(tfpAllUser, False);
                    fListModule.RecRed.FileDir := '';
                  end;
                  if fListModule.SaveModul(NewRecID)then
                  begin
                    VTI.BeginUpdate;
                    try
                      VTI.FullCollapse();
                      {$IfDef UPDATE}
                        Self.FListNodeSel.Clear;
                      {$EndIf}
                      Self.fHelperVTI.ClearDataNode;
                      Self.LoadDataNodes(NewRecID, Lib);
                      Self.ChangeModule(False);
                    finally
                      VTI.EndUpdate;
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
      end;
    end
    else
    begin
      pgInfoFile.ActivePageIndex := 0;
      pgInfoFile.Pages[1].TabVisible := False;
      Image.Picture := nil;
      lImage.Visible := True;
      lImageText.Visible := True;
      pImage.Visible := True;
      if OpenNewFile then
      begin
        fListModule.RecRed.SetFlagParams(tfpOsn, True);
        ChangeModule(True);
        cbAllUser.Visible := True;
        cbOneStart.Visible := True;
      end
      else
      begin
        VTIChange(VTI, VTI.GetFirstSelected(true));
      end;
    end;
  end;
end;

procedure TfModulesUser.sbRolesClick(Sender: TObject);
var
  arrRole: array of string;
  {$IfDef UPDATE}i, c: integer;{$EndIf}
  obj: TComponent;
begin
  With Self do
  begin
    obj := GetObjFromVTI;
    if Assigned(obj)then
    begin
      if obj.InheritsFrom(TRecModule)then
      begin
        if(sbRoles.ImageIndex = 3)then
        begin
          {$IfDef UPDATE}
          SetLength(arrRole, 0);
          for i := 0 to Pred(fRolesDB.ListRole.Count) do
          begin
            if fRolesDB.ListRole[i].Checket then
            begin
              SetLength(arrRole, Length(arrRole) + 1);
              arrRole[High(arrRole)] := fRolesDB.ListRole[i].RoleDB;
            end;
          end;
          {$EndIf}
          if fListModule.SetRoleList(Integer(TWinControl(Sender).Tag), arrRole) then
          begin
            sbRoles.ImageIndex := 5;
            VTI.BeginUpdate;
            try
              VTI.FullCollapse();
              {$IfDef UPDATE}
                FListNodeSel.Clear;
              {$EndIf}
              fHelperVTI.ClearDataNode;
              fListModule.ReoladFile;
              ChangeModule(False);
              LoadDataNodes(Integer(TWinControl(Sender).Tag), False);
              TWinControl(Sender).Tag := -1;
            finally
              VTI.EndUpdate;
            end;
          end;
        end
        else
        begin
          if TRecModule(Obj).GetFlagParams(tfpOsn)and
             not TRecModule(Obj).GetFlagParams(tfpAllUser)
          then
          begin
            sbRoles.ImageIndex := 3;
            TWinControl(Sender).Tag := TRecModule(Obj).RecID;
            {$IfDef UPDATE}
            i := fRolesDB.qRolesDB.FieldByName('recid').AsInteger;
            fRolesDB.qRolesDB.First;
            while not fRolesDB.qRolesDB.Eof do
            begin
              fRolesDB.qRolesDB.FieldByName('CheckBox').Value := False;
              for c := 0 to Pred(TRecModule(Obj).ListRole.Count) do
              begin
                if(TRecModule(Obj).ListRole[c].ID = fRolesDB.qRolesDB.FieldByName('recid').AsInteger)then
                begin
                  fRolesDB.qRolesDB.FieldByName('CheckBox').Value := True;
                  Break;
                end;
              end;
              fRolesDB.qRolesDB.Next;
            end;
            fRolesDB.qRolesDB.Locate('recid', i,[loCaseInsensitive, loPartialKey]);
            {$EndIf}
          end;
        end;
        VTI.Enabled := (sbRoles.ImageIndex = 5);
        sbAddLibrarry.Visible := (sbRoles.ImageIndex = 5);
        {$IfDef UPDATE}
        fRolesDB.IsReadOnly := (sbRoles.ImageIndex = 5);
        fRolesDB.ColorIsChecked := (sbRoles.ImageIndex = 3);
        {$EndIf}
        sbEditCancel.ImageIndex := IfThen((sbRoles.ImageIndex = 3), 4, 2);
        sbInsertSave.Enabled := (sbRoles.ImageIndex = 5);
        sbAddVersion.Enabled := (sbRoles.ImageIndex = 5);
        sbRollback.Enabled := (sbRoles.ImageIndex = 5);
        if (sbRoles.ImageIndex = 5) then
          Self.ChangeModule(False);
        sbRollback.Hint :=IfThen((sbRoles.ImageIndex = 5),
            'Выбрать разрешения для запуска модуля',
            'Активировать выбранный список разрешений');
        sbEditCancel.Hint := IfThen((sbRoles.ImageIndex = 5),
            'Редактировать текущую запись',
            'Отменить выбор списка разрешений');
      end;
    end;
  end;
end;

procedure TfModulesUser.sbRollbackClick(Sender: TObject);
var
  obj: TComponent;
  {$IfDef UPDATE}
    NewRecID: Integer;
    Lib: Boolean;
  {$EndIf}
begin
  With Self do
  begin
    obj := GetObjFromVTI;
    if Assigned(obj)then
    begin
      {$IfDef UPDATE}
      if obj.InheritsFrom(TRecModule)then
      begin
        fVersionInfo := TfVersionInfo.Create(Self, 'Причина отказа от использования модуля');
        try
          if fVersionInfo.ShowModal = mrOk then
          begin
            if fListModule.BlockModule(TRecModule(obj), fVersionInfo.mInfo.Lines.Text)then
            begin
              Lib := not TRecModule(obj).GetFlagParams(tfpOsn);
              NewRecID := TRecModule(obj).RecID;
              Self.FListNodeSel.Clear;
              Self.fHelperVTI.ClearDataNode;
              Self.LoadDataNodes(NewRecID, Lib);
              Self.ChangeModule(False);
            end;
          end;
        finally
          FreeAndNil(fVersionInfo);
        end;
      end
      else
      begin
        if obj.InheritsFrom(TRecVersion)then
        begin
          fVersionInfo := TfVersionInfo.Create(Self, 'Причина отката версии');
          try
            if fVersionInfo.ShowModal = mrOk then
            begin
              if fListModule.BlockVersModule(TRecVersion(obj), fVersionInfo.mInfo.Lines.Text)then
              begin
                Lib := not TRecVersion(obj).BaseModul.GetFlagParams(tfpOsn);
                NewRecID := TRecVersion(obj).BaseModul.RecID;
                Self.FListNodeSel.Clear;
                Self.fHelperVTI.ClearDataNode;
                Self.ChangeModule(False);
                Self.LoadDataNodes(NewRecID, Lib);
              end;
            end;
          finally
            FreeAndNil(fVersionInfo);
          end;
        end;
      end;
      {$EndIf}
    end;
  end;
end;

procedure TfModulesUser.SetIsSelected(const ASelected: Boolean);
{$IfDef UPDATE}
var
  i: Integer;
{$EndIf}
begin
  Self.fIsSelected := ASelected;
  {$IfDef UPDATE}
  if ASelected then
  begin
    Self.ChangeModule(False);
    Self.fListModule.IsBlockTable := False;
  end
  else
  begin
    Self.fListModule.IsBlockTable := True;
  end;
  for i := 0 to Pred(Self.FListNodeSel.Count) do
  begin
    if Assigned(Self.FListNodeSel[i]) then
    begin
      if ASelected then
        TRecModulVTI(Self.FListNodeSel[i]).Node^.CheckType := ctCheckBox
      else
        TRecModulVTI(Self.FListNodeSel[i]).Node^.CheckType := ctNone;
    end;
    Self.fRolesDB.Visible := not ASelected;
  end;
  {$EndIf}
end;

{$IfDef UPDATE}
procedure TfModulesUser.SetListRoles(const AVisibleList: Boolean);
begin
  if not(AVisibleList = Self.GetIsListRoles)then
  begin
    if Assigned(Self.fRolesDB)then
      FreeAndNil(Self.fRolesDB)
    else
      Self.CreateRolesDB;
  end;
end;

procedure TfModulesUser.SetTypeModule(const ATypeModule: TListType);
var
  tmp: TVTIRecBase;
begin
  if not(Self.fTypeModule = ATypeModule)then
  begin
    Self.fTypeModule := ATypeModule;
    Self.ReoladData;
    tmp:= Self.fHelperVTI.GetNodeBase(2);
    if Assigned(tmp) then
      Self.VTI.FullCollapse(tmp.Node);
    tmp:= Self.fHelperVTI.GetNodeBase(1);
    if Assigned(tmp) then
    begin
      Self.VTI.FullCollapse(tmp.Node);
      Self.VTI.Expanded[tmp.Node] := True;
    end;
  end;
end;
{$EndIf}

procedure TfModulesUser.SetReadOnly(const AReadOnly: Boolean);
begin
  Self.fIsReadOnly := AReadOnly;
  Self.pbtm.Visible := not Self.fIsReadOnly;
  if Assigned(Self.fListModule) then
  Begin
    {$IfDef UPDATE} Self.fRolesDB.IsReadOnly := True; {$EndIf}
    Self.fListModule.IsBlockTable := not AReadOnly;
    if AReadOnly then
      Self.FileReadOnly(AReadOnly);
    if not AReadOnly then
      Self.VTI.Header.Columns[0].Options :=
        (Self.VTI.Header.Columns[0].Options - [coVisible])
    else
      Self.VTI.Header.Columns[0].Options :=
          (Self.VTI.Header.Columns[0].Options + [coVisible]);
    Self.VTIResize(VTI);
  End;
end;

procedure TfModulesUser.VTICanSplitterResizeColumn(Sender: TVTHeader; P: TPoint;
  Column: TColumnIndex; var Allowed: Boolean);
begin
  Self.VTI.Header.Columns[5].Width :=
      Self.VTI.Width -(
      Ifthen(
      coVisible in Self.VTI.Header.Columns[0].Options,
          Self.VTI.Header.Columns[0].Width, 0) +
      Self.VTI.Header.Columns[1].Width +
      Self.VTI.Header.Columns[2].Width +
      Self.VTI.Header.Columns[3].Width +
      Self.VTI.Header.Columns[4].Width +
      21);
  if(Self.VTI.Header.Columns[5].Width <= 150) then
    Self.VTI.Header.Columns[Column].Width :=
          150 - Self.VTI.Header.Columns[5].Width;
end;

procedure TfModulesUser.VTIChange(Sender: TBaseVirtualTree; Node: PVirtualNode);
 var
   tmp: pVTIRecBase;
   obj: TComponent;
begin
  try
    tmp := Sender.GetNodeData(Node);
    if Assigned(tmp) then
    begin
      if(vsDisabled in Node^.States)or([vsOnFreeNodeCallRequired] = Node^.States)then
        Self.GetActiveRec(trNoneN, nil)
      else
      begin
        if tmp^.InheritsFrom(TRecBase) then
          Self.GetActiveRec(trBaseN, tmp^)
        else
        begin
          obj := tmp^.DataObject;
          if Assigned(obj) then
          begin
            if obj.InheritsFrom(TRecModule) then
              Self.GetActiveRec(trModuleN, obj)
            else
              if obj.InheritsFrom(TRecVersion) then
                Self.GetActiveRec(trVersN, obj)
              else
                Self.GetActiveRec(trNoneN, tmp^);
          end
          else
            Self.GetActiveRec(trNoneN, nil);
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      Self.GetActiveRec(trNoneN, nil);
     {$IfDef Debug}
      VCL.Dialogs.MessageDlg(
        e.ClassName + ': ' +e.Message +'({4A3CF076-7456-481C-B1E6-BCB72F4DB369}.VTIChange)'
          ,mtError , [mbIgnore], 0);
      {$EndIf}
    end;
  end;
end;

procedure TfModulesUser.VTIChecking(Sender: TBaseVirtualTree;
  Node: PVirtualNode; var NewState: TCheckState; var Allowed: Boolean);
var
   tmp: pVTIRecBase;
   obj: TComponent;
begin
  tmp := Sender.GetNodeData(Node);
  if Assigned(tmp) then
  begin
    if tmp^.InheritsFrom(TRecModulVTI) then
    begin
      obj := tmp^.DataObject;
      if obj.InheritsFrom(TRecModule) then
       TRecModule(obj).Checked := (NewState = csCheckedNormal);
    end;
  end;
end;

procedure TfModulesUser.VTIDblClick(Sender: TObject);
//var
//   NodeArray: TNodeArray;
begin
//  NodeArray := VTI.GetSortedSelection(False);
//  Self.VTI.IsVisible[NodeArray[0]]:= False;
end;

procedure TfModulesUser.VTIResize(Sender: TObject);
begin
  if Assigned(Self.fListModule) then
  begin
    Self.VTI.Header.Columns[5].Width :=
      Self.VTI.Width -(
      Ifthen(
        coVisible in Self.VTI.Header.Columns[0].Options,
            Self.VTI.Header.Columns[0].Width, 0) +
      Self.VTI.Header.Columns[1].Width +
      Self.VTI.Header.Columns[2].Width +
      Self.VTI.Header.Columns[3].Width +
      Self.VTI.Header.Columns[4].Width +
      21);
  end;
end;
end.
