unit ObjModule;
{$WARN SYMBOL_PLATFORM OFF}
{$WARN UNIT_DEPRECATED OFF}
{$I Directives.inc}
interface
  uses Winapi.Windows, Vcl.Forms, {$ifDef Debug} Vcl.Dialogs, {$EndIf}System.SysUtils,  System.Classes, Vcl.Graphics,
       IBX.IBDatabase, IBX.IBCustomDataSet, IBX.IBQuery, Data.DB, IBX.IBBlob{$IfDef UPDATE}, FrameRolesDB{$EndIf}, IBX.IBScript,
       VTIHelper, System.Generics.Defaults, System.Generics.Collections, ComUtils;

{$Region ' Основной список скриптов '}
const
  SQLModul:array[1..15]of string =(
          // 1. Работа с таблицами модулей
          {$IfDef UPDATE}
          'select m.*, mp.parent' + slineBreak +
          'from spr_module as m ' + slineBreak +
          'join spr_module_pack mp on(mp.recid = m.recid) ' + slineBreak +
          'where not (m.recid = 0)' + slineBreak +
          'order by mp.parent, m.recid, m.npp'
          {$Else}
          'select m.*, mp.parent' + slineBreak +
          'from spr_module as m ' + slineBreak +
          'join spr_module_pack mp on(mp.recid = m.recid) ' + slineBreak +
          'where not (m.recid = 0)and(mp.parent = 0)' + slineBreak +
          'and(m.isactive < 2)and(m.date_to is null)and(m.date_from <= current_timestamp)' + slineBreak +
          'and((m.isall_users =1)or(m.recid in(' + slineBreak +
          'select mm.recid from spr_module as mm join spr_module_role mr on'+ slineBreak +
          '(mr.recid = mm.recid)and(mr.isactive =1)and(mr.date_to is null)and' + slineBreak +
          '(mr.user_role in(%0:s))' + slineBreak +
          ')))' + slineBreak +
          'order by mp.parent, m.recid, m.npp'
          {$EndIf}
          , //2. Блокируем 4 таблицы селектом
          'select max(m.recid) recid, max(m.npp) npp' + slineBreak +
          'from spr_module m' + slineBreak +
          ' join spr_module_pack mp on (mp.recid = m.recid)' + slineBreak +
          ' join spr_module_vers mv on (mv.recid = m.recid)' + slineBreak +
          ' left join spr_module_role mr on (mr.recid = m.recid)'
          , //3. Обновление модуля
          'update or insert into spr_module (recid, npp, module_name, name_file, catalog_upd, name_ext, isall_users, isonestart,' + slineBreak +
          'isactive, date_from, date_to, actualversion, isupdate, sizefile, info, image, no_export_regions)' + slineBreak +
          'values (:recid, :npp, :module_name, :name_file, :catalog_upd, :name_ext, :isall_users, :isonestart, :isactive,' + slineBreak +
          '        :date_from, :date_to, :actualversion, :isupdate, :sizefile, :info, :image, :no_export_regions)' + slineBreak +
          'matching (recid)'
          ,// 4. пакетами
          'update or insert into spr_module_pack (recid, parent)values(:recid, :parent)matching (recid, parent)'
          ,// 5. версионной
          'select ' + slineBreak +
          '   recid,' + slineBreak +
          '   npp,   ' + slineBreak +
          '   version, ' + slineBreak +
          '   filename, ' + slineBreak +
          '   info_vers, ' + slineBreak +
          '   info_rollback,' + slineBreak +
          '   sizefile, ' + slineBreak +
          '   isactive, ' + slineBreak +
          '   isrollback,' + slineBreak +
          '   date_from,  ' + slineBreak +
          '   date_to ' + slineBreak +
          'from spr_module_vers ' + slineBreak +
          'where recid in(<:RecList:>)' + slineBreak +
          'order by recid, npp desc'
          , // 6.
          'update or insert into spr_module_vers (recid, npp, version, filename, info_vers, info_rollback, sizefile, isactive,' + slineBreak +
          ' isrollback, user_red, role_red, date_from, date_to) ' + slineBreak +
          'values (:recid, :npp, :version, :filename, :info_vers, :info_rollback, :sizefile, :isactive, :isrollback,  ' + slineBreak +
          'current_user, current_role, :date_from, :date_to)matching (recid, npp)'
          ,// 7. Работа с ролями
          'select  ' + slineBreak +
          '      ur.recid   ' + slineBreak +
          '     ,mr.recid mRecId ' + slineBreak +
          '     ,ur.user_role ' + slineBreak +
          '     ,mr.isactive  ' + slineBreak +
          '     ,mr.date_from ' + slineBreak +
          '     ,mr.date_to   ' + slineBreak +
          'from spr_module_role mr ' + slineBreak +
          '  join spr_role ur on  ' + slineBreak +
          '       (ur.user_role = mr.user_role) and ' + slineBreak +
          '       (ur.isactive = 1) and ' + slineBreak +
          '       not(ur.date_from > current_timestamp) and ' + slineBreak +
          '       not(coalesce(ur.date_to, current_timestamp) < current_timestamp)' + slineBreak +
          'where(mr.recid in(<:RecMList:>)) and  ' + slineBreak +
          '     (mr.isactive = 1) and ' + slineBreak +
          '     not(mr.date_from > current_timestamp) and ' + slineBreak +
          'not(coalesce(mr.date_to, current_timestamp) < current_timestamp)'
          , // 8.
          'update or insert into spr_module_role (recid, user_role, isactive, date_from, date_to)'+ slineBreak +
          'values (<:recid:>, <:user_role:>, 1, current_timestamp, null)'+ slineBreak +
          'matching (recid, user_role);'
          , // 9.
          'update spr_module_role'+ slineBreak +
          '   set isactive = 0,'+ slineBreak +
          '       date_to = current_timestamp'+ slineBreak +
          'where (recid = :recid) and' + slineBreak +
          '     not(user_role in(<:user_role:>))and' + slineBreak +
          '     (isactive = 1)',
          //10
          'select m.*, mp.parent' + slineBreak +
          'from spr_module as m ' + slineBreak +
          'join spr_module_pack mp on(mp.recid = m.recid)' + slineBreak +
          'where (m.recid in(%0:s))or(mp.parent in(%0:s))' + slineBreak +
          'and(m.isactive=1)and(m.date_to is null)',
          // 11
          'update spr_module m set m.sizefile = :size, m.actualversion = :version where m.recid =:recid',
          // 12 Откаты/Закрытие модулей
          'update spr_module m set m.isactive = 2, m.date_to = current_timestamp, m.infoclose = :infoclose where m.recid =:recid',
          // 13 Получить последню валидную версию
          'select * ' + slineBreak +
          ' from spr_module_vers mvv ' + slineBreak +
          ' where(mvv.recid = :recid) and ' + slineBreak +
          '  mvv.npp =( ' + slineBreak +
          '  select Max(NPP) ' + slineBreak +
          '  from spr_module_vers mv  ' + slineBreak +
          '  where (mv.recid= mvv.recid)and ' + slineBreak +
          '        (mv.isactive = 1)and ' + slineBreak +
          '        (mv.date_to is null)and ' + slineBreak +
          '        not(mv.npp = :npp))',
          // 14 Откаты / Закрытие версий
          'update spr_module_vers ' + slineBreak +
          'set ' + slineBreak +
          '    info_rollback = :info_rollback,' + slineBreak +
          '    isactive = 0,' + slineBreak +
          '    isrollback = 1,' + slineBreak +
          '    user_red = current_user,' + slineBreak +
          '    role_red = current_role,' + slineBreak +
          '    date_to = current_timestamp' + slineBreak +
          'where (recid = :recid)and' + slineBreak +
          '      (npp = :npp)',
          // 15 Сохраняем лог проверок
          'insert into spr_module_exp_log (data_export, date_to, logdata)' + slineBreak +
          'values (:data_export, current_timestamp, :logdata)'
          );
{$EndRegion}

type
  TRecVersion = class;
  TFlagsParams = (tfpOsn = 1, tfpOneStart = 2, tfpAllUser =3, tfpIsLoad = 4, tfpIsRead = 5, tfpNoRead = 6, tfpNoExportRegion = 7);
  TFlagsParamsVer = (tfpvActive = 1, tfpvRollBack = 2);

  TListCheckFile = class(TComponent)
  strict private
    fList : TList<WideString>;
  public
    constructor Create(AOwner: TComponent); reintroduce; override;
    destructor Destroy; override;
    procedure AddPathFile(const AFilePath: WideString);
    function CheckFilePath(const AFilePath: WideString): Boolean;
  end;

  TRecBase = class(TVTIRecBase)
  private
    fNameNode: String;
    fInfo: String;
  published
    procedure GetText(Column: integer; var CellText: String); override;
  public
    constructor Create(const ACodeNode: DWord; const ABaseDisplFont: TFont;
            ADataObject: TComponent = nil;
             const AClearData: Boolean = false); reintroduce; overload;
  end;
  
  TRecModule = class(TComponent)
  private
    // Для модулей порядковый номер
    fRecID: Integer;
    fNPP: Integer;
    fSize: Int64;
    fInfo: String;
    fInfoClose: String;
    fFileName,
    fFileDir,
    fDescr,
    fExt: String;
    fDateFrom,
    fDateTo : TDateTime;
    // 0: не проверяем (Просто информация к размышлению)
    // 1: проверяем на версию (Бинарники)
    // 2: проверяем на наличие (Обычно конфиги)
    fUpdate: Byte;
    // 0: Только создание,
    // 1: Aктивен,
    // 3: Деактивирован
    fActive: Byte;
    // Картинка / ярлык
    fImage: TMemoryStream;
    fVersion: String;
    // 1: Основной модуль/Библиотека,
    // 2: Первый старт
    // 3: Для всех пользователей(Уровень доступности)
    // 4: Загружен/Новый(Создан)
    // 5: Изменен
    // 6: Нельзя редактировать
    fFlags: int64;
    fListVers: TObjectList<TRecVersion>;
    fListLib: TObjectList<TRecModule>;
    fChecked: Boolean;
    {$IfDef UPDATE}fListRole: TObjectList<TRecRole>; {$EndIf}
    function GetRed: Boolean;
    procedure SetRed(const Value: Boolean);
    procedure SetImage(const AImage: TMemoryStream);
    procedure SetChecked(const AChecked: Boolean);
    procedure FreeList;
  public
    constructor Create(AOwner: TComponent); reintroduce; override;
    destructor Destroy; override;
    function GetFlagParams(const AParams: TFlagsParams): Boolean;
    procedure SetFlagParams(const AParams: TFlagsParams; const AValue: Boolean);
    procedure Assign(ARec: TRecModule); reintroduce; overload;
    property RecID: Integer read fRecID;
    property ListLib: TObjectList<TRecModule> read fListLib;
    property InfoClose: String read fInfoClose write fInfoClose;
    {$IfDef UPDATE}
    property ListRole: TObjectList<TRecRole> read fListRole;{$EndIf}
    property ListVers: TObjectList<TRecVersion> read fListVers;
    property NPP: Integer read fNPP write fNPP;
    property Size: Int64 read fSize write fSize;
    property Info: String read fInfo write fInfo;
    property FileName:String read fFileName write fFileName;
    property FileDir:String read fFileDir write fFileDir;
    property Descr:String read fDescr write fDescr;
    property Ext:String read fExt write fExt;
    property DateFrom: TDateTime read fDateFrom write fDateFrom;
    property DateTo: TDateTime read fDateTo write fDateTo;
    property Update: Byte read fUpdate write fUpdate;
    property Active: Byte read fActive write fActive;
    property Image: TMemoryStream read fImage write SetImage;
    property Version: String read fVersion write fVersion;
    property IsRedOnly: Boolean read GetRed write SetRed;
    property Checked: Boolean read fChecked write SetChecked default False;
    function Status: String;
  end;

  TRecRed = class(TRecModule)
  private
    fPathFile: String;
    fParent: Integer;
  public
    constructor Create(AOwner: TComponent; APathFile: String; AParent: Integer); reintroduce; overload;
    constructor Create(AOwner: TComponent; ARecModul: TRecModule); reintroduce; overload;
    procedure LoadDataFile;
    property Parent: Integer read fParent write fParent;
  end;

  TRecModulVTI = class(TVTIRecBase)
  public
    constructor Create(ADataObject: TRecModule; const ABaseDisplFont: TFont); reintroduce; overload;
    procedure GetText(Column: integer; var CellText: String); override;
  end;

  TRecVersion = class(TComponent)
  private
    fNPP: Integer;
    fSize: Int64;
    fDateFrom,
    fDateTo: TDateTime;
    fInfo: String;
    fInfRollBack: String;
    fFileName: String;
    fVersion: String;
    fFlags: Int64;
    fModul: TRecModule;
  public
    constructor Create(AOwner: TComponent; ABaseModul: TRecModule;
      const ANPP: Integer); reintroduce; overload;
    function GetFlagParams(const AParamsVer: TFlagsParamsVer): Boolean;
    procedure SetFlagParams(const AParamsVer: TFlagsParamsVer;
      const AValue: Boolean);
    procedure Assign(ARec: TRecVersion); reintroduce; overload;
    property NPP: Integer read fNPP;
    property Version: String read fVersion write fVersion;
    property DateFrom:TDateTime Read fDateFrom write fDateFrom;
    property DateTo:TDateTime Read fDateTo write fDateTo;
    property Size: Int64 read fSize write fSize;
    property Info: String read fInfo write fInfo;
    property InfRollBack: String read fInfRollBack write fInfRollBack;
    property BaseModul: TRecModule read fModul;
    property FileName:String read fFileName write fFileName;
  end;

  TRecVersionVTI = class(TVTIRecBase)
  public
    constructor Create(ADataObject: TRecVersion; const ABaseDisplFont: TFont); reintroduce; overload;
    procedure GetText(Column: integer; var CellText: String); override;
  end;

  TListModule = class(TComponent)
  private
    fListAPP: TObjectList<TRecModule>;
    fListlib: TObjectList<TRecModule>;
    fRecRed: TRecRed;
    fDataBase: TIBDataBase;
    qSelect: TIBQuery;
    qBlock: TIBQuery;
    fRTranaction,
    fRWTransaction: TIBTransaction;
    fListLog: TStringList;
    fSmallLog: TStringList;
    {$IfDef UPDATE}
      fError: Boolean;
      NDir,
      fPathUpd,
      fUserRW,
      fPassRW: WideString;
      fBaseID: Integer;
    {$Else}
      fLisRolesDB: array of string;
    {$EndIf}
      function LoadFromBase(const AFullVersion: Boolean = True): Boolean;
    {$IfDef UPDATE}
    procedure LoadModuleRoles(const ARecList: array of integer);
    function GetParamUPD: Boolean;
    {$EndIf}
    procedure LoadFullVersion(const ARecList: array of integer);
    procedure SetNppBlock;
    function GetModulByRecID(const ARecID: Integer): TRecModule;
    procedure SetBlock(const ABlock: Boolean);
    function GetBlock: Boolean;
    function UpdateOrInsertM(ARecRed: TRecRed): Boolean;
    function UIParent(ARecRed: TRecRed): Boolean;
    function GetObjModul(const ARecID: Integer;
      AList: TObjectList<TRecModule>): TRecModule;
    function Init(ADataBase: TIBDataBase): Boolean;
    procedure FreeList;
    procedure SetLogStr(const AStr: String);
    {$IfDef UPDATE}
    function CheckBasePath(ARecModule: TRecModule;
        ARecVers: TRecVersion; const ATmpDir: WideString;
        ACheckList: TListCheckFile;
        out ALogs: WideString;const ACheckRes: Boolean = True;
        const APathUPD: WideString = ''; const ABaseID:Integer = 0): Boolean;
    function CheckFile(ARec: TRecModule;out ACheck: Boolean;
         const ATmpDir: WideString; ACheckList: TListCheckFile): WideString;
    {$EndIf}
  public
    {$IfDef UPDATE}
    constructor Create(AOwner: TComponent; ADataBase: TIBDataBase; const ABaseID: Integer; const ASmallLog: Boolean = False); reintroduce; overload;
    {$Else}
    constructor Create(AOwner: TComponent; ADataBase: TIBDataBase; const AListRoles: array of string); reintroduce; overload;
    class function OpenSelectModuleRoles(ASelect: TIBQuery;
      AListRoles: array of string): Boolean; static;
    {$EndIf}
    destructor Destroy; override;
    function SetRoleList(const AModuleRecID: Integer;
      const AListRole: array of string; const ACheckModule: Boolean = true): Boolean;
    function SaveModul(out ARecID: Integer): Boolean;
    {$IfDef UPDATE}
    procedure CreateRedFile(ARecModul: TRecModule); overload;
    function CreateLibFile(APathFile: String; AParent: Integer): Integer;
    procedure CreateRedFile(APathFile: String; AParent: Integer); overload;
    function SaveVersion(AVersion: TRecVersion;
      const ANewPathFile: WideString;const ANewVersion: Boolean = True): Boolean;
    function BlockModule(ARecModule: TRecModule;const ATextRoll: WideString): Boolean;
    function BlockVersModule(ARecVers: TRecVersion;const ATextRoll: WideString): Boolean;
    {$EndIf}
    property ListAPP: TObjectList<TRecModule> read fListAPP;
    property Listlib: TObjectList<TRecModule> read fListlib;
    property IsBlockTable: Boolean read GetBlock write SetBlock;
    property RecRed: TRecRed read fRecRed write fRecRed;
    procedure ReoladFile;
    {$IfDef UPDATE}
    function isValidFile(const ATmpDir: WideString;out ALog:
         WideString; ACheckList: TListCheckFile;
        const ASaveLog: Boolean = True): Boolean;
    procedure SetLogToDB(const ALog: WideString; const ADataStart: TDateTime);
    function ExtractDirFile(const AExtractFolder: WideString;
      const AOneIsCheck: Boolean; out ALogs: WideString; out ACount: Integer;
      ACheckFile: TListCheckFile;
      const ACheckRes: Boolean = True;
      const APathUPD:WideString = '';
      const ABaseID:Integer = 0): Boolean;

     property Error: Boolean read fError;
     property Logs: TStringList read fSmallLog;
     property PathUpd: WideString read fPathUpd;
     property QueryBlock:TIBQuery read qBlock;
    {$EndIf}
  end;

implementation
uses System.Variants, System.Math, System.StrUtils, uThread, ProcesssObjAbs, UpdConst, u_Progress;

{ TListCheckFile }
procedure TListCheckFile.AddPathFile(const AFilePath: WideString);
begin
  if not Self.CheckFilePath(AFilePath) then
    Self.fList.Add(AFilePath);
end;

function TListCheckFile.CheckFilePath(const AFilePath: WideString): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Pred(Self.fList.Count) do
  begin
    if AnsiSameText(AFilePath, Self.fList.Items[i]) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

constructor TListCheckFile.Create(AOwner: TComponent);
begin
  inherited;
  Self.fList := TList<WideString>.Create;
end;

destructor TListCheckFile.Destroy;
begin
  FreeAndNil(Self.fList);
  inherited;
end;

{ TListModule }
function TListModule.Init(ADataBase: TIBDataBase): Boolean;
begin
  Result := False;
  with Self do
  begin
    fListLog := TStringList.Create;
    fSmallLog := TStringList.Create;
    fRecRed := nil;
    fListAPP := nil;
    fListLib := nil;
    fDataBase := ADataBase;
    SetLogStr('Создание основного листа модулей/релизов');
    if Assigned(fDataBase) then
    begin
      if fDataBase.Connected then
      begin
        SetLogStr('Связь с БД прошла успешно');
        fRTranaction := TIBTransaction.Create(Self);
        fRWTransaction := TIBTransaction.Create(Self);
        fRTranaction.Params.Clear;
        fRTranaction.Params.Add('read_committed');
        fRTranaction.Params.Add('rec_version');
        fRTranaction.Params.Add('nowait');
        fRTranaction.DefaultDatabase := fDataBase;
        fRWTransaction.Params.Clear;
        fRWTransaction.Params.Add('write');
        {$IfDef UPDATE}
        // Транзакция с блокировкой целевых таблиц
        fRWTransaction.Params.Add('wait');
        fRWTransaction.Params.Add('consistency');
        {$EndIf}
        fRWTransaction.DefaultDatabase := fDataBase;
        qSelect := TIBQuery.Create(Self);
        qSelect.Database := fDataBase;
        qSelect.Transaction := fRTranaction;
        {$IfDef UPDATE}
        try
          if not Self.GetParamUPD then
            raise Exception.Create('Не удалось получить параметры сетевой директории обновления!');
        except
          on e: Exception do
          begin
            if SameText(ParamStr(3), Self.fDataBase.DatabaseName) then
              if
                MessageBox(0,
                  PChar('Не удалось получить параметры сетевой директории обновления: '
                  + sLineBreak + sLineBreak + e.Message),
                  PChar('Продолжить загрузку'), MB_OKCANCEL + MB_ICONERROR)=IDOK
              then

              else
                raise Exception.Create(e.Message)
            else
              raise Exception.Create(e.Message);
          end;
        end;
        {$EndIf}
        fListAPP := TObjectList<TRecModule>.Create(
          IComparer<TRecModule>(
            function(const Left, Rigth: TRecModule): Integer
            begin
              Result := CompareValue(Left.NPP, Rigth.NPP);
              if(Result = 0) then
                Result := CompareValue(Left.RecID, Rigth.RecID);
              if(Result = 0) then
                Result := CompareText(Left.FileName, Rigth.FileName);
            end), False);
        fListLib := TObjectList<TRecModule>.Create(
          IComparer<TRecModule>(
            function(const Left, Rigth: TRecModule): Integer
            begin
              Result := CompareValue(Left.RecID, Rigth.RecID);
              if(Result <>0) then
                Result := CompareText(Left.FileName, Rigth.FileName);
            end), False);
        Result := True;
      end;
    end;
  end;
end;

{$IfDef UPDATE}
function TListModule.isValidFile(const ATmpDir: WideString;
    out ALog: WideString; ACheckList: TListCheckFile;
    const ASaveLog: Boolean = True): Boolean;
var
  stl: TStringList;
   rec: TRecModule;
   b: Boolean;
   dst: TDateTime;
   s: String;
begin
  stl := TStringList.Create;
  try
    dst := Now;
    Result := True;
    MyPObj.DeleteDir(ATmpDir, True);

    if DirectoryExists(ATmpDir) then
       RemoveDir(ATmpDir);

    if DirectoryExists(ATmpDir) then
    begin
      stl.Add(Format(' Ошибка удаления временной директории  %0:s',[ATmpDir]));
      Result := False;
    end
    else
    begin
      {$IfDef Debug}
      if IsFileDebugPas then
      begin
        stl.Add('  Вилидация пропущена наличием debug.pas');
      end
      else
      {$EndIf}
      begin
        stl.Add(StringOfChar('-',20) + ' Основные файлы ' + StringOfChar('-',64));
        for rec in Self.ListAPP do
        begin
          stl.Add(FormatDateTime('  hh:nn:ss.zzz | ', Now) + 'RecID: ' + IntToStr(rec.RecID));
          s := Self.CheckFile(rec, b, ATmpDir, ACheckList);
          stl.Add(s);
          if not b then
            Result := False;
          Application.ProcessMessages;
        end;
        stl.Add(StringOfChar('-',20) + ' Подгружаемые файлы ' + StringOfChar('-',64));
        for rec in Self.Listlib do
        begin
          stl.Add(FormatDateTime('  hh:nn:ss.zzz | ', Now) + 'RecID: ' + IntToStr(rec.RecID));
          stl.Add(Self.CheckFile(rec, b, ATmpDir, ACheckList));
          if not b then
            Result := False;
          Application.ProcessMessages;
        end;
      end;
    end;
    if ASaveLog then
      Self.SetLogToDB(stl.Text, dst);

    if Result then
    begin
      {$ifDef Release}
      MyPObj.DeleteDir(ATmpDir, True);
      if DirectoryExists(ATmpDir) then
       RemoveDir(ATmpDir);
      {$EndIf}
    end;
  finally
    ALog := stl.Text;
    FreeAndNil(stl);
  end;
end;

procedure TListModule.SetLogToDB(const ALog: WideString; const ADataStart: TDateTime);
var
  ss : TStringStream;
begin
  qBlock.Close;
  if qBlock.Transaction.InTransaction then
    qBlock.Transaction.Rollback;
  try
    if not qBlock.Transaction.InTransaction then
      qBlock.Transaction.StartTransaction;
    qBlock.SQl.Clear;
    qBlock.SQl.Text := SQLModul[15];
    qBlock.GenerateParamNames := True;
    qBlock.Prepare;
    qBlock.ParamByName('data_export').AsDateTime := ADataStart;
    ss := TStringStream.Create(string(ALog));
    try
      ss.Seek(0, soBeginning);
      qBlock.ParamByName('logdata').LoadFromStream(ss, ftBlob);
      if not uThread.OpenExecDS(qBlock, 'Сохранение статистики (Логирования)', 'Запись данных', 'ExecSQL') then
        raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text);
      if qBlock.Transaction.InTransaction then
        qBlock.Transaction.Commit;
    finally
      FreeAndNil(ss);
    end;
  except
    on e : Exception do
    begin
      MessageBox(0, PChar(e.ClassName +': ' + e.Message),
        PChar('Ошибка сохранения данных статистики'), MB_OK + MB_ICONERROR);
    end;
  end;
  if qBlock.Transaction.InTransaction then
      qBlock.Transaction.Rollback;
    SetNppBlock;
end;

constructor TListModule.Create(AOwner: TComponent; ADataBase: TIBDataBase; const ABaseID: Integer; const ASmallLog: boolean);
begin
  inherited Create(AOwner);
  with Self do
  begin
    Self.fBaseID := ABaseID;
    try
      if not Self.Init(ADataBase) then
        FreeAndNil(fListAPP)
      else
        if not LoadFromBase then
          FreeAndNil(fListAPP);
    except
      on e: Exception do
      begin
        Self.fError := True;
        fSmallLog.Add(e.ClassName + ': ' + e.Message);
        if not ASmallLog then
          raise Exception.Create(Self.fListLog.Text)
        else
          raise Exception.Create(Self.fSmallLog.Text)
      end;
    end;
  end;
end;

function TListModule.BlockModule(ARecModule: TRecModule;const ATextRoll: WideString): Boolean;
var
  ss: TStringStream;
begin
  Result := False;
  if Assigned(ARecModule)then
  begin
    with Self do
    begin
      if qBlock.Transaction.InTransaction then
        qBlock.Transaction.Rollback;
      try
        qBlock.Close;
        qBlock.SQl.Text := SQLModul[12];
        qBlock.GenerateParamNames := True;
        qBlock.Prepare;
        qBlock.ParamByName('recid').AsInteger := ARecModule.RecID;
        ss := TStringStream.Create(ATextRoll);
        try
          ss.Seek(0, soBeginning);
            qBlock.ParamByName('infoclose').LoadFromStream(ss, ftBlob);
          if not uThread.OpenExecDS(qBlock, 'Блокировка использования модуля', 'Сохранение данных') then
            raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text)
          else
          begin
            ARecModule.fActive := 2;
            ARecModule.fInfoClose := ATextRoll;
            ARecModule.DateTo := Now;
            Result := True;
            if qBlock.Transaction.InTransaction then
              qBlock.Transaction.Commit;
          end;
        finally
          FreeAndNil(ss)
        end;
      except
        on e: Exception do
        begin
          Result := False;
          MessageBox(0, PChar(e.Message),
              PChar('Ошибка сохранение данных'), MB_OK + MB_ICONERROR);
        end
      end;
      if qBlock.Transaction.InTransaction then
        qBlock.Transaction.Rollback;
      SetNppBlock;
    end;
  end;
end;

function TListModule.BlockVersModule(ARecVers: TRecVersion;const ATextRoll: WideString): Boolean;
var
  Vers: WideString;
  Size: Int64;
  ss: TStringStream;
begin
  Result := False;
  with Self do
  begin
    if Assigned(ARecVers)then
    begin
      if qBlock.Transaction.InTransaction then
        qBlock.Transaction.Rollback;
      try
        qBlock.Close;
        qBlock.SQl.Text := SQLModul[13];
        qBlock.GenerateParamNames := True;
        qBlock.Prepare;
        qBlock.ParamByName('recid').AsInteger := ARecVers.fModul.RecID;
        qBlock.ParamByName('npp').AsInteger := ARecVers.npp;
        if not uThread.OpenExecDS(qBlock, 'Получение следующей версии', 'Запрос данных') then
          raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text);

        qBlock.FetchAll;
        if not(qBlock.RecordCount = 1) then
          raise Exception.Create('Не найдено ни одной активной версии что бы заменить текущую!');

        vers := qBlock.FieldByName('version').AsString;
        Size := qBlock.FieldByName('sizefile').AsInteger;
        qBlock.Close;
        qBlock.SQl.Text := SQLModul[11];
        qBlock.GenerateParamNames := True;
        qBlock.Prepare;
        qBlock.ParamByName('recid').AsInteger := ARecVers.fModul.RecID;
        qBlock.ParamByName('size').AsInteger := Size;
        qBlock.ParamByName('version').AsString := vers;
        if not uThread.OpenExecDS(qBlock, 'Активация следующей версии', 'Обновление данных', 'ExecSQL') then
          raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text);
        qBlock.Close;
        qBlock.SQl.Text := SQLModul[14];
        qBlock.GenerateParamNames := True;
        qBlock.Prepare;
        qBlock.ParamByName('recid').AsInteger := ARecVers.fModul.RecID;
        qBlock.ParamByName('npp').AsInteger := ARecVers.NPP;
        ss := TStringStream.Create(ATextRoll);
        try
          ss.Seek(0, soBeginning);
          qBlock.ParamByName('info_rollback').LoadFromStream(ss, ftBlob);
          if not uThread.OpenExecDS(qBlock, 'Закрытие строки версии', 'Обновление данных', 'ExecSQL') then
            raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text);
          if qBlock.Transaction.InTransaction then
            qBlock.Transaction.Commit;
          ARecVers.fInfRollBack := ATextRoll;
          ARecVers.DateTo := Now;
          ARecVers.SetFlagParams(tfpvRollBack, True);
          ARecVers.fModul.Version := Vers;
          ARecVers.fModul.Size := Size;
          Result := True;
        finally
          FreeAndNil(ss)
        end;
      except
        on e: Exception do
        begin
          Result := False;
          MessageBox(0, PChar(e.Message),
              PChar('Ошибка сохранение данных'), MB_OK + MB_ICONERROR);
        end
      end
    end;
    if qBlock.Transaction.InTransaction then
      qBlock.Transaction.Rollback;
    SetNppBlock;
  end;
end;
{$Else}
constructor TListModule.Create(AOwner: TComponent; ADataBase: TIBDataBase;
  const AListRoles: array of string);
var
  i: Integer;
begin
  inherited Create(AOwner);
  try
    with Self do
    begin
      if Self.Init(ADataBase) then
      begin
        SetLength(Self.fLisRolesDB, Length(AListRoles));
        for i := Low(AListRoles) to High(AListRoles) do
          Self.fLisRolesDB[i] := AListRoles[i];
        SetLogStr(Format('Лист ограничений содержит %0:d элементов',[Length(AListRoles)]));
        if not LoadFromBase then
          FreeAndNil(fListAPP);
      end;
    end;
  except
    on e: Exception do
    begin
      SetLogStr(Format('Ошибка загрузки данных Error: %0:s',[e.message]));
      if MessageBox(0, PChar(e.message + slineBreak + 'Показать лог загрузки ?'),
                  PChar('Ошибка загрузки данных'), MB_OKCANCEL + MB_ICONERROR) = IDOK
      then
      begin
        MessageBox(0, PChar(Self.fListLog.Text),
                  PChar('Лог загрузки данных'), MB_OK + MB_ICONERROR);
      end;
    end;
  end;
end;
{$EndIf}

{$IfDef UPDATE}
function TListModule.CheckBasePath(ARecModule: TRecModule;
     ARecVers: TRecVersion; const ATmpDir: WideString;
     ACheckList: TListCheckFile;
     out ALogs: WideString; const ACheckRes: Boolean = True;
     const APathUPD: WideString = ''; const ABaseID:Integer = 0): Boolean;
var
  s, sv, ss, sz: WideString;
  size: Int64;
  srh: TSearchRec;
  frmProgress: TfProgress;
  bb: Boolean;
begin
  {$Region ' Проверка файла в директории '}
  with Self do
  begin
    Result := False;
    bb:= True;
    if Assigned(ACheckList) then
    begin
      if ACheckList.CheckFilePath(fPathUpd + ARecVers.fFileName) then
      begin
        ALogs := '  Файл был ранее валидирован с положительным результатом';
        bb := False;
        Result := True;
      end;
    end;

    if Assigned(ACheckList)and not bb then
    begin
      if not ACheckList.CheckFilePath(ATmpDir + ARecVers.fFileName) then
      begin
        ALogs := '  Файл Отсутствует в индивидуальном хранидище';
        bb := True;
        Result := False;
      end;
    end;

    if bb then
    begin
      if not uThread.OpenDir(ifThen(ACheckRes, fPathUpd, APathUPD) , True)then
        raise Exception.Create(Format('Базовая директория "%0:s" хранения обновлений недоступна', [fPathUpd]));
      Application.ProcessMessages;
      s := ExtractFilePath(ifThen(ACheckRes, fPathUpd, APathUPD) + ARecVers.fFileName);
      s := IncludeTrailingPathDelimiter(s);
      Application.ProcessMessages;
      if not uThread.OpenDir(s, True)then
        if not ForceDirectories(s) then
          raise Exception.Create(Format('Не удалось создать директорию "%0:s" хранения файла ресурса', [s]));
      Application.ProcessMessages;
      s := fPathUpd + ARecVers.fFileName;

      if not FileExists(s) then
        raise Exception.Create(Format('Ошибка наличия файла "%0:s"', [s]))
      else
      begin
        sv := ATmpDir + ARecVers.FileName;
        sz := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0))) + 'TmpReserve'  + IfThen(ABaseID >0, IntToStr(ABaseID),'');
        sz := IncludeTrailingPathDelimiter(sz) + ARecVers.FileName;

        frmProgress := TfProgress.Create(nil);
        frmProgress.Caption := 'Копирование файла';
        frmProgress.LabelProcess.Caption := 'Проверка ресурса';
        frmProgress.Show;
        Application.ProcessMessages;
        try
          Result :=  False;
          if FileExists(sv) then
          begin
            if not MyPObj.CopyFile(sv, sz, True, ss, frmProgress.ProgressBar1, frmProgress.LabelProcess) then
               raise Exception.Create(' Файл на целевом сервере найден, но его извлечение привело к ошибке');
            try
              if FindFirst(sz, faAnyFile, srh ) = 0 then
                Size := Int64(srh.FindData.nFileSizeHigh) shl Int64(32) + Int64(srh.FindData.nFileSizeLow)
              else
                raise Exception.Create('  Файл на целевом сервере найден и изъят, но его наличие не подтвердилось');
            finally
              FindClose(srh) ;
            end;
            ALogs := '  Файл найден в целевой директории';
            s := MyPObj.GetVersionFile(sz);
            if (Length(Trim(s)) = 0) then
              s := ARecModule.Version;
            Result := SameText(s, ARecVers.fVersion)and(Size = ARecVers.Size);
            if Result and Assigned(ACheckList) then
            begin
              ACheckList.AddPathFile(fPathUpd + ARecVers.fFileName);
              ACheckList.AddPathFile(ATmpDir + ARecVers.fFileName);
            end;
            ALogs := ALogs + sLineBreak + Format('  Проверка сообщает об %0:s состоянии',[ifthen(Result,'актуальном','невалидном')]);

            if not Result then
            begin
              if not DeleteFile(sv)then
                 raise Exception.Create(ALogs + sLineBreak + '  Ошибка удаление файла несоответствующего шаблону');
            end;
          end;

          if not Result then
          begin
            s := ifThen(ACheckRes, fPathUpd, APathUPD) + ARecVers.fFileName;
            if not MyPObj.CopyFile(s, sv, True, ss, frmProgress.ProgressBar1, frmProgress.LabelProcess) then
            begin
              if(Length(ALogs) > 0)then
                  ALogs := ALogs + sLineBreak;
              ALogs := ALogs  + Format('  Ошибка сохранения файла "%0:s", с сообщением (%1:s)', [s, ss]);
              raise Exception.Create(ALogs);
            end;
            try
              if FindFirst(sv, faAnyFile, srh ) = 0 then
                Size := Int64(srh.FindData.nFileSizeHigh) shl Int64(32) + Int64(srh.FindData.nFileSizeLow)
              else
              begin
                if(Length(ALogs) > 0)then
                  ALogs := ALogs + sLineBreak;
                ALogs := ALogs + Format('  Ошибка наличия файла "%0:s" (FindFirst)', [sv]);
                raise Exception.Create(ALogs);
              end;
            finally
              FindClose(srh) ;
            end;

            if ACheckRes then
            begin
              s := MyPObj.GetVersionFile(sv);
              if (Length(Trim(s)) = 0) then
                s := ARecModule.Version;

              if not SameText(s, ARecVers.fVersion)then
              begin
                if(Length(ALogs) > 0)then
                  ALogs := ALogs + sLineBreak;
                ALogs := ALogs + Format('  Версии файлов не совпадают искомый "%0:s"/ "1:s"', [ARecVers.Version, s]);
                raise Exception.Create(ALogs);
              end;
              if not (Size = ARecVers.Size)then
              begin
                if(Length(ALogs) > 0)then
                  ALogs := ALogs + sLineBreak;
                ALogs := ALogs + Format('  Файлы отличаются размером на "%0:s"', [MyPObj.GetStrOfSize( ARecVers.fSize - size, true)]);
                raise Exception.Create(ALogs);
              end;
            end;
            Result := True;
            if(Length(ALogs) > 0)then
              ALogs := ALogs + sLineBreak;
            ALogs := ALogs + '  Файл успешно скопирован в целевую директорию';
          end;

        finally
          FreeAndNil(frmProgress);
        end;
      end;
    end;
  end;
  {$EndRegion}
end;

function TListModule.CheckFile(ARec: TRecModule; out ACheck: Boolean;
     const ATmpDir: WideString; ACheckList: TListCheckFile): WideString;
var
  vers: TRecVersion;
  i: Integer;
  s: WideString;
  sl: TStringList;
begin
  sl:= TStringList.Create;
  try
    sl.Add('Имя файла: ' + ARec.FileDir + '/' + ARec.FileName + '.' +ARec.Ext);
    sl.Add('Descriptor: ' + ARec.Descr);
    sl.Add('Version: ' + ARec.Version);
    sl.Add('Размер: ' + MyPObj.GetStrOfSize(ARec.Size, True));
    vers := nil;
    for i := 0 to Pred(ARec.fListVers.Count) do
    begin
      if SameText(ARec.fListVers[i].Version, ARec.Version)then
      begin
        vers := ARec.fListVers[i];
        Application.ProcessMessages;
      end;
    end;
    if not Assigned(vers) then
    begin
      sl.Add('Не найдена информация место положения файла');
      ACheck:= False;
    end
    else
    begin
      s := '';
      sl.Add('Путь к файлу: ' + vers.fFileName);
      try
        if(ARec.Active = 1)then
        begin
          ACheck := Self.CheckBasePath(ARec, vers, ATmpDir, ACheckList, s);
          sl.Add(s);
        end
        else
        begin
          sl.Add('Данный ресурс не подлежит проверке');
          ACheck := True;
        end;
        if ACheck then
          if(ARec.Active = 1)then
            sl.Add('Успешная проверка соответствия ресурса в хранилище');
      except
        on e: Exception do
        begin
          sl.Add(e.Message);
          ACheck := False;
        end;
      end;
    end;
    for i := 0 to Pred(sl.Count) do
    begin
      if not((Copy(sl.Strings[i],1,3) = '===')or
        (Copy(sl.Strings[i],1,3) = '---')or
        (Copy(sl.Strings[i],1,3) = '+++'))
      then
        sl.Strings[i] := '  ' + Trim(sl.Strings[i]);
    end;
    Result := sl.Text;
  finally
    FreeAndNil(sl);
  end;
end;
{$EndIf}

{$IfDef UPDATE}
function TListModule.CreateLibFile(APathFile: String; AParent: Integer): Integer;
var
  rec, recIdx: TRecModule;
  i, idx: Integer;
begin
  Result := 0;
  With Self do
  begin
    fRecRed := nil;
    if(AParent < 1) then
      MessageBox(0, PChar('Передан нуливой RecID  основного модуля {34EC3D8F-AC97-4F58-BE79-58BDA40BDC56}'),
        PChar('Ошибка Разработчика'), MB_OK + MB_ICONERROR)
    else
    begin
      fRecRed := TRecRed.Create(Self, APathFile, AParent);
      fRecRed.SetFlagParams(tfpOsn, False);
      if(Length(Trim(fRecRed.Version)) = 0)then
        fRecRed.Version := '0.0.0.0';
      if(Length(Trim(fRecRed.Descr)) = 0)then
        fRecRed.Descr := 'Отсутствует информационный параметр "FileDescription"';
      recIdx := nil;
      for rec in Self.fListAPP do
      begin
        if AnsiSameText( AnsiLowerCase(rec.FileName + '.' + rec.fExt),
           AnsiLowerCase(fRecRed.FileName + '.' + fRecRed.Ext))
        then
        begin
          recIdx := rec;
          Break;
        end;
      end;
      if Assigned(recIdx) then
      begin
        MessageBox(0, PChar('Подключаемая библиотека зарегестрирована автономным модулем, выбор следует делать внимательнее'),
        PChar('Ошибка выбранного модуля'), MB_OK + MB_ICONERROR);
        FreeAndNil(Self.fRecRed);
      end
      else
      begin
        for recIdx in Self.fListAPP do
        begin
          if(recIdx.fRecID = AParent)then
          begin
            for rec in recIdx.ListLib do
            begin
              if AnsiSameText( AnsiLowerCase(rec.FileName + '.' + rec.fExt),
                 AnsiLowerCase(fRecRed.FileName + '.' + fRecRed.Ext))
              then
              begin
                MessageBox(0, PChar('Подключаемая библиотека зарегестрирована к этому модулю !'),
                   PChar('Повторная регистрация библиотеки'), MB_OK  + MB_ICONERROR);
                FreeAndNil(fRecRed);
                Break;
              end;
            end;
          end;
        end;
        if Assigned(fRecRed)then
        begin
          for rec in Self.fListlib do
          begin
            if AnsiSameText( AnsiLowerCase(rec.FileName + '.' + rec.fExt),
               AnsiLowerCase(fRecRed.FileName + '.' + fRecRed.Ext))
            then
            begin
              if MessageBox(0, PChar('Подключаемая библиотека зарегестрирована ранее, подключить ее и к этому модулю?'),
                PChar('Уточняющая информация'), MB_OKCANCEL  + MB_ICONINFORMATION) = IDOK
              then
              begin
                if qBlock.Transaction.InTransaction then
                  qBlock.Transaction.Rollback;
                try
                  qBlock.Close;
                  fRecRed.fRecID := rec.RecID;
                  Result := rec.RecID;
                  if UIParent(fRecRed) then
                    if qBlock.Transaction.InTransaction then
                      qBlock.Transaction.Commit;
                  recIdx := TRecModule.Create(nil);
                  recIdx.Assign(rec);
                 for i := 0 to Pred(ListAPP.Count)do
                  begin
                    if ListAPP.Items[i].RecID = fRecRed.Parent then
                    begin
                      if not ListAPP.Items[i].fListLib.BinarySearch(recIdx, idx) then
                      begin
                        ListAPP.Items[i].fListLib.Insert(idx, recIdx);
                        recIdx := nil;
                      end;
                      Break;
                    end;
                  end;
                  if Assigned(recIdx)then
                    FreeAndNil(recIdx);
                finally
                  if qBlock.Transaction.InTransaction then
                    qBlock.Transaction.Rollback;
                  SetNppBlock;
                end;
              end;
              FreeAndNil(fRecRed);
              Break;
            end;
          end;
        end;
      end;
    end;
  end;
end;

procedure TListModule.CreateRedFile(APathFile: String; AParent: Integer);
var
  rec, recIdx: TRecModule;
begin
  Self.fRecRed := TRecRed.Create(Self, APathFile, AParent);
  with Self do
  begin
    recIdx := nil;
    for rec in Self.fListAPP do
    begin
      if AnsiSameText( AnsiLowerCase(rec.FileName + '.' + rec.fExt),
         AnsiLowerCase(fRecRed.FileName + '.' + fRecRed.Ext))
      then
      begin
        recIdx := rec;
        Break;
      end;
    end;
    if not Assigned(recIdx) then
    begin
      for rec in Self.fListlib do
      begin
        if AnsiSameText( AnsiLowerCase(rec.FileName + '.' + rec.fExt),
           AnsiLowerCase(fRecRed.FileName + '.' + fRecRed.Ext))
        then
        begin
          recIdx := rec;
          Break;
        end;
      end;
    end;
  end;
  if Assigned(recIdx)then
  begin
    MessageBox(0, PChar(
        'Файл "' + Self.RecRed.FileName + '" зарегестрирован в' + sLineBreak +
        'системе обновлений, добавление разных но одноименных файлов ' +
        'создаст путаницу версий. Во избежании рецидива ошибок данная операция исключена !'),
    PChar('Ошибка'), MB_OK + MB_ICONERROR);
     FreeAndNil(Self.fRecRed);
  end;
end;

procedure TListModule.CreateRedFile(ARecModul: TRecModule);
begin
  Self.fRecRed := TRecRed.Create(Self, ARecModul);
end;

function TListModule.ExtractDirFile(const AExtractFolder: WideString;
      const AOneIsCheck: Boolean; out ALogs: WideString; out ACount: Integer;
      ACheckFile: TListCheckFile;
      const ACheckRes: Boolean = True;
      const APathUPD:WideString = '';
      const ABaseID:Integer = 0): Boolean;
var
  sl: TStringList;
  i: Integer;

  function ExctractFile(ARec: TRecModule; out FileExt: String): Boolean;
  var
    s: WideString;
    i: Integer;
    vers: TRecVersion;
  begin
    sl.Add(StringOfChar('-' ,120));
    sl.Add(Format('%0:s) File: %1:s  vers: %2:s Size: %3:s',[
      StringOfChar('0', 4 - Length(IntToStr(ARec.RecID))) + IntToStr(ARec.RecID),
      ARec.FileName + '.' + ARec.Ext,
      ARec.Version, MyPObj.GetStrOfSize(ARec.Size, True)]));
    inc(ACount);
    vers := nil;
    for i := 0 to Pred(ARec.fListVers.Count) do
    begin
      if SameText(ARec.fListVers[i].Version, ARec.Version)then
      begin
        vers := ARec.fListVers[i];
        Application.ProcessMessages;
      end;
    end;
    if not Assigned(vers) then
    begin
      sl.Add('Не найдена информация место положения актуальной версии файла');
      Result := False;
    end
    else
    begin
      try
        FileExt :=  AExtractFolder + vers.fFileName;

        if not Assigned(ACheckFile) then
            Result := Self.CheckBasePath(ARec, vers, AExtractFolder, ACheckFile, s, ACheckRes, APathUPD, ABaseID)
        else
           if not ACheckFile.CheckFilePath(FileExt)then
             Result := Self.CheckBasePath(ARec, vers, AExtractFolder, ACheckFile, s, ACheckRes, APathUPD, ABaseID)
           else
           begin
             Result := True;
             s := 'В Валидации файла нет необходимость, она пройдена ранее';
           end;
        sl.Add(s);
      except
        on e: Exception do
        begin
          sl.Add(e.Message);
          Result := False;
        end;
      end;
    end;
  end;

var
  rec: TRecModule;
  dt: int64;
  ss ,sp: String;
begin
  sl := TStringList.Create;
  try
    Result := True;
    ACount := 0;
    dt := GetTickCount64;
    if ACheckRes then
    begin
      if DirectoryExists(AExtractFolder) then
       RemoveDir(AExtractFolder);

      if not MyPObj.DeleteDir(AExtractFolder)then
      begin
        sl.Add(Format('Ошибка удаления директории %0:s временного хранения файлов', [AExtractFolder]));
        Result := False;
        Exit;
      end;
    end;
    if Result then
    begin
      sl.Add(FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz | ', Now)  + 'Начало копирования');
      for rec in Self.fListAPP do
      begin
        if rec.Checked or not AOneIsCheck then
        begin
          if ExctractFile(rec, sp) then
          begin
            if FileExists(sp)then
              sl.Add('Файл успешно скопирован в локальную директорию')
            else
              sl.Add('Валидация файла подтверждена')
          end
          else
          begin
            sl.Add('Ошибка копирования файла');
            Result := False;
            Break;
          end;
        end;
      end;
    end;
    if Result then
    begin
      for rec in Self.fListlib do
      begin
        if rec.Checked or not AOneIsCheck then
        begin
          if not ExctractFile(rec, sp) then
          begin
            sl.Add('Ошибка копирования файла');
            Result := False;
            Break;
          end;
        end;
      end;
    end;
    sl.Add(StringOfChar('-', 120));
    dt := (GetTickCount - dt) div 1000;
    sl.Add(Format ('Потрачено всего: %0:d мин. %1:d сек.',[ (dt div 60), (dt mod 60) ]));
    ss := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0))) + 'TmpReserve' + IfThen(ABaseID >0, IntToStr(ABaseID),'');
    ss := IncludeTrailingPathDelimiter(ss);
    if DirectoryExists(ss)then
      if not MyPObj.DeleteDir(ss)then
        sl.Add(Format('Ошибка удаления директории %0:s проверки файлов', [ss]));

    for i := 0 to Pred(sl.Count) do
    begin
      if not((Copy(sl.Strings[i],1,3) = '===')or
        (Copy(sl.Strings[i],1,3) = '---')or
        (Copy(sl.Strings[i],1,3) = '+++'))
      then
        sl.Strings[i] := '  ' + Trim(sl.Strings[i]);
    end;
  finally
    ALogs := sl.Text;
    FreeAndNil(sl);
  end;
end;
{$EndIf}

destructor TListModule.Destroy;
begin
  Self.FreeList;
  FreeAndNil(Self.fListAPP);
  FreeAndNil(Self.fListlib);
  FreeAndNil(Self.fListLog);
  FreeAndNil(Self.fSmallLog);
  inherited;
end;

procedure TListModule.FreeList;
var
  app: TRecModule;
begin
  if Assigned(Self.fListAPP) then
  begin
    while not(Self.fListAPP.Count =0) do
    begin
      app:= Self.fListAPP.Items[Pred(Self.fListAPP.Count)];
      try
        Self.fListAPP.Items[Pred(Self.fListAPP.Count)] := nil;
        Self.fListAPP.Delete(Pred(Self.fListAPP.Count));
      finally
        FreeAndNil(app);
      end;
    end;
  end;
  if Assigned(Self.fListlib) then
  begin
    while not(Self.fListlib.Count =0) do
    begin
      app:= Self.fListlib.Items[Pred(Self.fListlib.Count)];
      try
        Self.fListlib.Delete(Pred(Self.fListlib.Count));
      finally
        FreeAndNil(app);
      end;
    end;
  end;
end;

function TListModule.GetBlock: Boolean;
begin
  Result := Assigned(Self.qBlock);
end;

function TListModule.GetModulByRecID(const ARecID: Integer): TRecModule;
var
  rec: TRecModule;
begin
  Result := nil;
  for rec in Self.fListAPP do
  begin
    if(rec.RecID = ARecID)then
    begin
      Result := rec;
      Break;
    end;
  end;
end;

function TListModule.GetObjModul(const ARecID: Integer;
  AList: TObjectList<TRecModule>): TRecModule;
var
  mr: TRecModule;
begin
  Result := nil;
  for mr in AList do
  begin
    if(mr.RecID = ARecID)then
    begin
      Result := mr;
      Break;
    end;
  end;
end;

{$IfDef UPDATE}
function TListModule.GetParamUPD: Boolean;
var
  UserR, PassR: WideString;
  recDir: TRecConnection;
  DirPath: String;
begin
  With Self do
  begin
    qSelect.Close;
    Result := UpdConst.GetParamUpdate(qSelect, fPathUpd,
    UserR, PassR,
    fUserRW, fPassRW);
    if Result then
    begin
      SetLogStr(Format('Даные сетевой дирекории получены "%0:s"',[fPathUpd]));
      recDir.NPath := fPathUpd;
      recDir.User := fUserRW;
      recDir.Password := fPassRW;
      recDir.Code := 'DirUserUPDRW' + IntToStr(Self.fBaseID);
      MyPObj.ADDNetResource(recDir);
      DirPath :=IncludeTrailingPathDelimiter( IncludeTrailingPathDelimiter(fPathUpd) + cntDirUpdPath);
      SetLogStr(Format('Даные сетевой дирекории обновлены "%0:s"',[DirPath]));
      u_Progress.StartFormThread('Подключение сетевого каталога', 'Инициализация');
      try
        if Copy(Trim(DirPath), 1, 2)='\\' then
        begin
          NDir :=  recDir.NPath;
          Result := MyPObj.NetConnection('DirUserUPDRW' + IntToStr(Self.fBaseID), True);
          if not Result then
            SetLogStr('Ошибка: ' + MyPObj.GetErrorNetDir('DirUserUPDRW' + IntToStr(Self.fBaseID)) +
               ' {183926B7-5541-4605-B22E-92C2F24CEDA2}' )
          else
            Result := DirectoryExists(fPathUpd);
        end
        else
        begin
          NDir :=  recDir.NPath;
          Result := DirectoryExists(DirPath);
        end;
      except
        on e:Exception do
        begin
          Result := False;
          SetLogStr('Ошибка ' + e.ClassName + ': ' + e.Message + ' {DE1CC777-5796-43EE-B18D-7D4A2D366A00}' );
        end;
      end;
      u_Progress.FormDestroy(nil);

      if Result then
      begin
        u_Progress.StartFormThread('Проверка каталога устаревших фалов', 'Инициализация');
        SetLogStr(Format('Подключение к сетевой директории "%0:s" установлено',[DirPath]));
        if DirectoryExists(IncludeTrailingPathDelimiter(fPathUpd) + 'TmpOLD\') then
        begin
          try
            if MyPObj.DeleteDir(WideString(IncludeTrailingPathDelimiter(fPathUpd) + 'TmpOLD\'), True) then
              SetLogStr(Format('Сетевая директория "%0:s" очищена',[IncludeTrailingPathDelimiter(fPathUpd) + 'TmpOLD\']));
          except
            on e: Exception do
            begin
              SetLogStr(Format('Ошибка очистки сетевой директории "%0:s" Error: %1:s ',[IncludeTrailingPathDelimiter(fPathUpd) + 'TmpOLD\',
              e.ClassName + ': ' + e.Message]));
            end;
          end;
        end;
        u_Progress.FormDestroy(nil);
      end;

      if not Result then
        raise Exception.Create(Format('Директория хранения версий "%0:s" не доступна !', [DirPath]) + ' ' + RLS_EXEPT)
      else
        fPathUpd := DirPath;
    end;
  end;
end;
{$EndIf}

function TListModule.LoadFromBase(const AFullVersion: Boolean): Boolean;
var
  mRec, recLib, ParentRec: TRecModule;
  Stream: TStream;//TIBDSBlobStream;
  index: Integer;
  arV: Array of integer;
begin
  fRecRed:= nil;
  Stream := nil;
  SetLength(arV, 0);
  with Self do
  begin
    FreeList;
    qSelect.Close;
    if Self.qSelect.Transaction.InTransaction then
      Self.qSelect.Transaction.Rollback;
    if not Self.fRWTransaction.InTransaction then
      Self.fRWTransaction.StartTransaction;
    try
      qSelect.SQL.Clear;
      qSelect.Params.Clear;
      {$Region ' Оснвной запрос к базе '}
      if AFullVersion then
      begin
        {$IfDef UPDATE}
          qSelect.SQl.Text := SQLModul[1];
        {$Else}
           OpenSelectModuleRoles(qSelect, fLisRolesDB);
        {$EndIf}
      end
      else
      begin

      end;
      {$EndRegion}
      try
        qSelect.GenerateParamNames := True;
        qSelect.Prepare;
        if uThread.OpenExecDS(qSelect, 'Загрузка базовых модулей', 'Получение данных') then
        begin
          qSelect.FetchAll;
          SetLogStr(Format('Получен список модулей/Librarry %0:d записей', [qSelect.RecordCount]));
          qSelect.First;
          while not qSelect.Eof do
          begin
            mRec := TRecModule.Create(nil);
            mRec.fNPP := qSelect.FieldByName('npp').AsInteger;
            mRec.fRecID := qSelect.FieldByName('recid').AsInteger;
            mRec.fDescr := qSelect.FieldByName('module_name').AsString;
            mRec.fFileName := qSelect.FieldByName('name_file').AsString;
            mRec.fFlags := 0;
            mRec.SetFlagParams(tfpOsn, qSelect.FieldByName('parent').AsInteger = 0);
            mRec.SetFlagParams(tfpNoExportRegion, qSelect.FieldByName('no_export_regions').AsInteger = 1);
            mRec.SetFlagParams(tfpOneStart, qSelect.FieldByName('isonestart').AsInteger =1);
            mRec.SetFlagParams(tfpAllUser, qSelect.FieldByName('isall_users').AsInteger =1);
            mRec.SetFlagParams(tfpIsLoad, True);
            mRec.SetFlagParams(tfpIsRead, True);
            mRec.fSize := qSelect.FieldByName('sizefile').AsInteger;
            mRec.fUpdate := qSelect.FieldByName('isupdate').AsInteger;
            mRec.fVersion := qSelect.FieldByName('actualversion').AsString;
            mRec.fActive := qSelect.FieldByName('isactive').AsInteger;
            mRec.fExt := qSelect.FieldByName('name_ext').AsString;
            // Относительный путь для программы
            mRec.fFileDir := qSelect.FieldByName('catalog_upd').AsString;
            // Основная информация для модуля
            if not qSelect.FieldByName('info').IsNull then
              mRec.fInfo := (qSelect.FieldByName('info') as TBlobField).AsString;

            if(mRec.fActive > 1)then
            begin
              if not qSelect.FieldByName('InfoClose').IsNull then
                mRec.fInfoClose := (qSelect.FieldByName('InfoClose') as TBlobField).AsString;
              if(Length(Trim(mRec.fInfoClose))= 0)then
                mRec.fInfoClose := 'Информация о причине блокировки модуля, не внесена !';
            end;
            // Основная иконка/картинка для модуля
            if not qSelect.FieldByName('image').IsNull then
            begin
              mRec.fImage := TMemoryStream.Create;
              Stream := qSelect.CreateBlobStream(qSelect.FieldByName('image'), bmRead);
              Stream.Seek(0, soBeginning);
              try
                mRec.fImage.LoadFromStream(Stream);
              finally
                FreeAndNil(Stream);
              end;
            end;
            mRec.fDateFrom := qSelect.FieldByName('date_from').AsDateTime;
            if not qSelect.FieldByName('date_to').IsNull then
                mRec.fDateTo := qSelect.FieldByName('date_to').AsDateTime;

            if(qSelect.FieldByName('parent').AsInteger = 0)then
            begin
              if not fListAPP.BinarySearch(mRec, index)then
              begin
                fListAPP.Insert(index, mRec);
                SetLength(arV, Length(arV) + 1);
                arV[High(arV)] := mRec.RecID;
              end;
              mRec.IsRedOnly := False;
            end
            else
            begin
              ParentRec := Self.GetModulByRecID(qSelect.FieldByName('parent').AsInteger);
              recLib := TRecModule.Create(nil);
              recLib.Assign(mRec);
              mRec.IsRedOnly := False;
              if not Listlib.BinarySearch(mRec, index)then
                Listlib.Insert(index, recLib)
              else
                FreeAndNil(recLib);

              if Assigned(ParentRec)then
              begin
                if not ParentRec.ListLib.BinarySearch(mRec, index)then
                begin
                  ParentRec.ListLib.Insert(index, mRec);
                  mRec.IsRedOnly := True;
                  SetLength(arV, Length(arV) + 1);
                  arV[High(arV)] := mRec.RecID;
                end
                else
                  FreeAndNil(mRec);
              end
              else
                FreeAndNil(mRec);
            end;
            qSelect.Next;
          end;
          if AFullVersion then
          begin
            LoadFullVersion(arV);
            {$IfDef UPDATE} LoadModuleRoles(arV);{$EndIf}
          end;
          Result := True;
        end
        else
        begin
          raise Exception.Create(RLS_EXEPT + sLineBreak + qSelect.SQl.Text);
        end;
      except
        on e:Exception do
        begin
          SetLogStr(Format('Обнаружена ошибка в процедуре LoadFromBase Error: %0:s', [e.ClassName + ': ' +e.Message]));
          raise Exception.Create(e.ClassName + ': ' +e.Message);
        end;
      end;
    finally
      qSelect.Close;
      if qSelect.Transaction.InTransaction then
        qSelect.Transaction.Rollback;
    end;
  end;
end;

procedure TListModule.LoadFullVersion(const ARecList: array of integer);
var
  fil: String;
  i, idx: Integer;
  recv: TRecVersion;
  mr: TRecModule;
begin
  mr := nil;
  SetLogStr('Загрузка списка версий');
  if(Length(ARecList) > 0)then
  begin
    for i := Low(ARecList) to High(ARecList)do
    begin
      fil := fil + IntToStr(ARecList[i]);
      if not(i = High(ARecList))then
        fil := fil + ', ';
    end;
    with Self do
    begin
      qSelect.Close;
      try
        qSelect.Close;
        qSelect.SQL.Clear;
        qSelect.Params.Clear;
        qSelect.SQl.Text := StringReplace(SQLModul[5], '<:RecList:>', fil,[rfReplaceAll, rfIgnoreCase]) ;
        qSelect.Prepare;
        if not uThread.OpenExecDS(qSelect, 'Загрузка истории версий', 'Получение данных') then
          raise Exception.Create(RLS_EXEPT + sLineBreak + qSelect.SQl.Text);
        qSelect.FetchAll;
        SetLogStr(Format('Получен список версий %0:d записей', [qSelect.RecordCount]));
        qSelect.First;
        while not qSelect.Eof do
        begin
          if not Assigned(mr)then
            mr := GetObjModul(qSelect.FieldByName('recid').AsInteger, ListAPP)
          else
            if not (mr.RecID = qSelect.FieldByName('recid').AsInteger)then
               mr := GetObjModul(qSelect.FieldByName('recid').AsInteger, ListAPP);

          if not Assigned(mr)then
            mr := GetObjModul(qSelect.FieldByName('recid').AsInteger, Listlib)
          else
            if not (mr.RecID = qSelect.FieldByName('recid').AsInteger)then
               mr := GetObjModul(qSelect.FieldByName('recid').AsInteger, Listlib);

          if Assigned(mr){$IfNDEF UPDATE} and(mr.fListVers.Count < 11){$EndIf} then
          begin
            recV := TRecVersion.Create(nil, mr, qSelect.FieldByName('NPP').AsInteger);
            recV.fSize := qSelect.FieldByName('sizefile').AsInteger;
            recv.fVersion := qSelect.FieldByName('version').AsString;
            recv.fFileName := qSelect.FieldByName('filename').AsString;
            recv.SetFlagParams(tfpvActive, qSelect.FieldByName('isactive').AsInteger = 1);
            recv.SetFlagParams(tfpvRollBack, qSelect.FieldByName('isrollback').AsInteger = 1);
            recv.fDateFrom := qSelect.FieldByName('date_from').AsDateTime;
            recv.fDateTo := qSelect.FieldByName('date_to').AsDateTime;
            recv.fInfo := (qSelect.FieldByName('info_vers') as TBlobField).AsString;
            if recv.GetFlagParams(tfpvRollBack) then
            begin
              if not qSelect.FieldByName('info_rollback').IsNull then
                recv.fInfRollBack := (qSelect.FieldByName('info_rollback') as TBlobField).AsString;
              if(Length(Trim(recv.fInfRollBack))= 0)then
               recv.fInfRollBack := 'Информация о причине отката версии необнаружена';
            end;
            if not mr.fListVers.BinarySearch(recv, idx)then
              mr.fListVers.Insert(idx, recv);
          end;
          qSelect.Next;
        end;
        qSelect.Close;
      except
        on e: Exception do
        begin
          SetLogStr(Format('Обнаружена ошибка в процедуре LoadFullVersion Error: %0:s', [e.Message ]));
          raise Exception.Create(e.Message + ' {86C514D3-391C-4DFA-B4ED-D0CB5447ECCE}');
        end;
      end;
    end;
  end;
end;

{$IfDef UPDATE}
procedure TListModule.LoadModuleRoles(const ARecList: array of integer);
var
  fil: String;
  i, idx: Integer;
  recR: TRecRole;
  mr: TRecModule;
begin
  SetLogStr(Format('Запущена процедура загрузки ограничений "%0:s"', ['LoadModuleRoles']));
  if(Length(ARecList) > 0)then
  begin
    for i := Low(ARecList) to High(ARecList)do
    begin
      fil := fil + IntToStr(ARecList[i]);
      if not(i = High(ARecList))then
        fil := fil + ', ';
    end;
    with Self do
    begin
      qSelect.Close;
      try
        qSelect.SQL.Clear;
        qSelect.Params.Clear;
        qSelect.SQl.Text := StringReplace(SQLModul[7], '<:RecMList:>', fil,[rfReplaceAll, rfIgnoreCase]);
        qSelect.Prepare;
        if not uThread.OpenExecDS(qSelect, 'Загрузка ограничений модулей', 'Получение данных') then
          raise Exception.Create(RLS_EXEPT + sLineBreak + qSelect.SQl.Text);
        mr := nil;
        qSelect.FetchAll;
        SetLogStr(Format('Список ограничений вернул "%0:d" записи', [qSelect.RecordCount]));
        qSelect.First;
        while not qSelect.Eof do
        begin
          if not Assigned(mr)then
            mr := GetObjModul(qSelect.FieldByName('mrecid').AsInteger, ListAPP)
          else
            if not (mr.RecID = qSelect.FieldByName('mrecid').AsInteger)then
               mr := GetObjModul(qSelect.FieldByName('mrecid').AsInteger, ListAPP);
          if Assigned(mr)then
          begin
            if(mr.RecID = qSelect.FieldByName('mrecid').AsInteger)then
            begin
              recR := TRecRole.Create(nil,
              qSelect.FieldByName('recid').AsInteger,
              qSelect.FieldByName('user_role').AsString,
              qSelect.FieldByName('user_role').AsString,
              True, False);
              if not mr.fListRole.BinarySearch(recR, idx)then
                mr.fListRole.Insert(idx, recR);
            end;
          end;
          qSelect.Next;
        end;
        qSelect.Close;
      except
        on e: Exception do
        begin
          SetLogStr(Format('Найдена ошибка в процедуре "LoadModuleRoles" Error: %0:s', [e.Message]));
          raise Exception.Create(e.Message + ' {8D05F830-7C06-4DBA-A4E8-723F8DA26217}');
        end;
      end;
    end;
  end;
end;
{$Else}

class function TListModule.OpenSelectModuleRoles(ASelect: TIBQuery;
  AListRoles: array of string): Boolean;
var
  arV: Array of integer;
  i: integer;
  s: String;
begin
  Result := False;
  for i := Low(AListRoles) to High(AListRoles) do
  begin
    s := s  + Char(39) + AListRoles[i]+ Char(39);
    if i < High(AListRoles) then
      s := s  +  ', ';
  end;
  try
    ASelect.SQl.Text := Format(SQLModul[1],[s]);
    ASelect.Prepare;
    if not uThread.OpenExecDS(ASelect, 'Загрузка базовых модулей', 'Получение данных') then
      raise Exception.Create(RLS_EXEPT + sLineBreak + ASelect.SQl.Text);
    ASelect.FetchAll;
    SetLength(arV, ASelect.RecordCount);
    ASelect.First;
    while not ASelect.Eof do
    begin
      arV[ASelect.RecNo - 1] := ASelect.FieldByName('recid').AsInteger;
      ASelect.Next;
    end;
    s := '';
    if(Length(arV) = 0)then
      s :='null'
    else
    begin
      for i := Low(arV) to High(arV) do
      begin
        s := s + IntToStr(arV[i]);
        if i < High(arV) then
          s := s  +  ', ';
      end;
    end;
    SetLength(arV, 0);
    ASelect.Close;
    ASelect.SQL.Clear;
    ASelect.Params.Clear;
    ASelect.SQl.Text := Format(SQLModul[10],[s]);
  finally
    SetLength(arV, 0);
  end;
end;
{$EndIf}

procedure TListModule.ReoladFile;
begin
  with Self do
  begin
    LoadFromBase;
    FreeAndNil(fRecRed);
  end;
end;

function TListModule.SaveModul(out ARecID: Integer): Boolean;
var
  npp: Integer;
begin
  ARecID := -1;
  Result := False;
  with Self do
  begin
    if not Assigned(qBlock) then
      raise Exception.Create('Error: ListModule.SaveModul {022B5E8A-D692-4140-A24F-95877FFC0AEE}');

    if( RecRed.RecID < 0)then
    begin
      ARecID := qBlock.FieldByName('recid').AsInteger + 1;
      npp := qBlock.FieldByName('npp').AsInteger + 1;
    end
    else
    begin
      ARecID := RecRed.RecID;
      npp := RecRed.NPP;
    end;
    RecRed.fRecID := ARecID;
    RecRed.NPP := npp;
    if qBlock.Transaction.InTransaction then
      qBlock.Transaction.Rollback;
    qBlock.Close;
    try
      if UpdateOrInsertM(fRecRed)then
      begin
        if qBlock.Transaction.InTransaction then
          qBlock.Transaction.Commit;
        Result := False;
        if Assigned(fRecRed) then
        begin
          Result := True;
          FreeAndNil(fRecRed);
        end;
      end;
    finally
      if qBlock.Transaction.InTransaction then
        qBlock.Transaction.Rollback;
      SetNppBlock;
    end;
  end;
end;

{$IfDef UPDATE}
function TListModule.SaveVersion(AVersion: TRecVersion;
    const ANewPathFile: WideString; const ANewVersion: Boolean): Boolean;
var
  ss, sr: TStringStream;
  idx: Integer;
  Rec: TRecVersion;
  s, sm, ssx: WideString;
  frmProgress: TfProgress;
  srh : TSearchRec;
  Size: Int64;
begin
  {$Region ' Копирование файла в целевую директорию '}
  with Self do
  begin
    {$IfDef Debug}
      {$Message ' Тут должно быть копирование файла обязательно '}
    {$EndIf}
    if not uThread.OpenDir(fPathUpd, True)then
      raise Exception.Create(Format('Базовая директория "%0:s" хранения обновлений недоступна', [fPathUpd]));

    s := ExtractFilePath(fPathUpd + AVersion.fFileName);
    s := IncludeTrailingPathDelimiter(s);
    if not uThread.OpenDir(s, True)then
      if not ForceDirectories(s) then
        raise Exception.Create(Format('Не удалось создать директорию "%0:s" хранения файла ресурса', [s]));

    s := fPathUpd + AVersion.fFileName;

    if not FileExists(s) then
    begin
      frmProgress := TfProgress.Create(nil);
      frmProgress.Caption := 'Копирование файла';
      frmProgress.Show;
      try
        if not MyPObj.CopyFile(ANewPathFile, s, True, sm, frmProgress.ProgressBar1, frmProgress.LabelProcess) then
          raise Exception.Create(Format('Ошибка сохранения файла "%0:s", с сообщением (%1:s)', [s, sm]));
      finally
        FreeAndNil(frmProgress);
      end;
    end
    else
    begin
      Size := 0;
      try
        if FindFirst(s, faAnyFile, srh ) = 0 then
          Size := Int64(srh.FindData.nFileSizeHigh) shl Int64(32) + Int64(srh.FindData.nFileSizeLow);
      finally
        FindClose(srh) ;
      end;
      try
        sm := MyPObj.GetVersionFile(s);
      except
        on e: Exception do
          sm := ComUtils.GUIDStr;
      end;
      if not(SameText(sm, AVersion.fVersion)and(Size = AVersion.Size))then
      begin
        ssx := IncludeTrailingPathDelimiter(fPathUpd);
        ssx := ExtractFilePath(Copy(ssx, 1, Length(ssx) -1));
        ssx := IncludeTrailingPathDelimiter(ssx) + 'TmpOLD\' + ExtractFilePath(AVersion.fFileName) ;
        if MyPObj.FileMove(WideString(s), WideString(ssx), true) then
        begin
          frmProgress := TfProgress.Create(nil);
          frmProgress.Caption := 'Копирование файла';
          frmProgress.Show;
          try
            if not MyPObj.CopyFile(WideString(ANewPathFile), s, True, sm, frmProgress.ProgressBar1, frmProgress.LabelProcess) then
              raise Exception.Create(Format('Ошибка сохранения файла "%0:s", с сообщением (%1:s)', [s, sm]));
          finally
            FreeAndNil(frmProgress);
          end;
        end
        else
          raise Exception.Create('Ошибка перемещения устаревшего файла');
      end;
    end;
  end;
  {$EndRegion}

  if ANewVersion then
    if qBlock.Transaction.InTransaction then
      qBlock.Transaction.Rollback;

  qBlock.Close;
  try
    qBlock.SQL.Clear;
    qBlock.Params.Clear;
    qBlock.SQl.Text := SQLModul[11];
    qBlock.GenerateParamNames := True;
    qBlock.Prepare;
    qBlock.ParamByName('recid').AsInteger := AVersion.fModul.RecID;
    qBlock.ParamByName('size').AsInteger := AVersion.fSize;
    qBlock.ParamByName('version').AsString := AVersion.Version;
    if not uThread.OpenExecDS(qBlock, 'Изменение версии модуля', 'Запись данных','ExecSQL') then
        raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text)
    else
    begin
      qBlock.Close;
      qBlock.SQL.Clear;
      qBlock.Params.Clear;
      qBlock.SQl.Text := SQLModul[6];
      qBlock.GenerateParamNames := True;
      qBlock.Prepare;
      qBlock.ParamByName('recid').AsInteger := AVersion.fModul.RecID;
      qBlock.ParamByName('npp').AsInteger := AVersion.fNPP;
      qBlock.ParamByName('sizefile').AsInteger := AVersion.fSize;
      qBlock.ParamByName('isactive').AsInteger := Ifthen( AVersion.GetFlagParams(tfpvActive), 1, 0);
      qBlock.ParamByName('isrollback').AsInteger := Ifthen( AVersion.GetFlagParams(tfpvRollBack), 1, 0);
      qBlock.ParamByName('date_from').AsDateTime := AVersion.fDateFrom;
      qBlock.ParamByName('version').AsString := AVersion.Version;
      qBlock.ParamByName('filename').AsString := AVersion.fFileName;
      if AVersion.GetFlagParams(tfpvRollBack) then
        qBlock.ParamByName('date_to').AsDateTime := AVersion.fDateTo
      else
        qBlock.ParamByName('date_to').Clear;
      ss := TStringStream.Create(AVersion.Info);
      sr := nil;
      try
        ss.Seek(0, soBeginning);
        qBlock.ParamByName('info_vers').LoadFromStream(ss, ftBlob);
        if AVersion.GetFlagParams(tfpvRollBack) then
        begin
          sr.Seek(0, soBeginning);
          qBlock.ParamByName('info_rollback').LoadFromStream(sr, ftBlob);
        end
        else
          qBlock.ParamByName('info_rollback').Clear;
        if not uThread.OpenExecDS(qBlock, 'Сохранение версии модуля', 'Запись данных','ExecSQL') then
          raise Exception.Create(RLS_EXEPT + sLineBreak + qBlock.SQl.Text)
        else
        begin
          if ANewVersion then
          begin
            if qBlock.Transaction.InTransaction then
              qBlock.Transaction.Commit;

            Rec := TRecVersion.Create(nil, AVersion.fModul, AVersion.NPP);
            Rec.Assign(AVersion);
            AVersion.fModul.Version := Rec.Version;
            AVersion.fModul.Size := Rec.Size;
            AVersion.fModul.DateFrom := Rec.DateFrom;
            if not AVersion.fModul.ListVers.BinarySearch(Rec, idx) then
              AVersion.fModul.ListVers.Insert(idx, Rec)
            else
              FreeAndNil(Rec);
            Result := True;
          end
          else
            Result := True;
        end;
      finally
        FreeAndNil(ss);
        FreeAndNil(sr);
        if ANewVersion then
        begin
          if qBlock.Transaction.InTransaction then
            qBlock.Transaction.Rollback;
          SetNppBlock;
        end;
      end;
    end;
  except
    on e: Exception do
    begin
      Result := False;
      if ANewVersion then
      begin
        if qBlock.Transaction.InTransaction then
          qBlock.Transaction.Rollback;
        SetNppBlock;
      end;
      MessageBox(0, PChar(
              e.ClassName + ': ' + e.Message),
          PChar('Ошибка сохранения данных'), MB_OK + MB_ICONERROR);
    end;
  end;
end;
{$EndIf}

procedure TListModule.SetBlock(const ABlock: Boolean);
begin
  with Self do
  begin
    qBlock := TIBQuery.Create(Self);
    qBlock.Database := fDataBase;
    qBlock.Transaction := fRWTransaction;
    SetNppBlock;
  end;
end;

procedure TListModule.SetLogStr(const AStr: String);
begin
  Self.fListLog.Add('');
  Self.fListLog.Add(FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz | ', Now) + AStr);
  Self.fSmallLog.Add(FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz | ', Now) + AStr);
  Self.fListLog.Add(StringOfChar('=', 120));
end;

procedure TListModule.SetNppBlock;
begin
  With Self do
  begin
    qBlock.Close;
    if not qBlock.Transaction.InTransaction then
       qBlock.Transaction.StartTransaction;
    qBlock.SQL.Clear;
    qBlock.Params.Clear;
    qBlock.SQl.Text := SQLModul[2];
    qBlock.Open;
  end;
end;

function TListModule.SetRoleList(const AModuleRecID: Integer;
  const AListRole: array of string; const ACheckModule: Boolean): Boolean;
var
  lr, ln, s, ss: String;
  i: Integer;
  exec: TIBScript;
  APP: TRecModule;
begin
  APP := nil;
  With Self do
  begin
    try
      lr := '';
      if ACheckModule then
      begin
        APP := nil;
        for i := 0 to Pred(ListAPP.Count) do
        begin
          if ListAPP[i].RecID = AModuleRecID then
          begin
            APP := ListAPP[i];
            Break;
          end;
        end;
        if not Assigned(APP) then
          raise Exception.Create('Не найден целевой модуль для назначения ролей {7AF16D76-CD21-4904-AC65-3B0224D6588A}');
        if not APP.GetFlagParams(tfpOsn) then
          raise Exception.Create('Модуль не является ограниченной {1ED46FAE-8DAE-484C-BFCD-C7E6ADDCE039}');
        if APP.GetFlagParams(tfpAllUser) then
          raise Exception.Create('Модуль имеет флаг доступен всем {F678F773-7AC4-4347-939B-FA15F104FAC5}');
        ss := StringReplace(SQLModul[8],'<:recid:>', IntToStr(APP.RecID),[rfReplaceAll, rfIgnoreCase]);
      end
      else
        ss := StringReplace(SQLModul[8],'<:recid:>', IntToStr(AModuleRecID),[rfReplaceAll, rfIgnoreCase]);


      if Length(AListRole)= 0 then
        lr := Char(39) +'NoUser_Role' +Char(39)
      else
      begin
        for i := Low(AListRole) to High(AListRole) do
        begin
          s := StringReplace(ss,'<:user_role:>', Char(39) + AListRole[i] + Char(39),[rfReplaceAll, rfIgnoreCase]);
          ln := ln + s;
          lr := lr + Char(39) + AListRole[i] + Char(39);
          if i < High(AListRole) then
          begin
            lr := lr + ', ';
            ln := ln + sLineBreak;
          end;
        end;
      end;
      qBlock.Close;
      if not qBlock.Transaction.InTransaction then
         qBlock.Transaction.StartTransaction;
      {$Region ' Обновление списка разрешений '}
      qBlock.SQL.Clear;
      qBlock.GenerateParamNames := True;
      qBlock.SQL.Text := StringReplace(SQLModul[9], '<:user_role:>', lr,[rfReplaceAll, rfIgnoreCase]);
      qBlock.Prepare;
      if ACheckModule then
        qBlock.ParamByName('recid').AsInteger := APP.RecID
      else
        qBlock.ParamByName('recid').AsInteger := AModuleRecID;
      if not uThread.OpenExecDS(qBlock, 'Блокируем оставшиеся роли', 'Сохранение данных','ExecSQL') then
        raise Exception.Create(RLS_EXEPT);
      if qBlock.Transaction.InTransaction then
         qBlock.Transaction.Commit;
      if not qBlock.Transaction.InTransaction then
         qBlock.Transaction.StartTransaction;
      qBlock.Close;
      // Добавляем обноляем роли
      if(Length(AListRole) > 0) then
      begin
        exec := TIBScript.Create(Self);
        try
          With exec do
          begin
            Database := Self.fDataBase;
            Transaction := qBlock.Transaction;
            Script.Text := ln;
            Terminator := ';';
            if not ValidateScript then
              raise Exception.Create('Script добавления ролей не прошел валидацию {CC6C2131-0972-44D4-9399-BA57CB2E482A}' );
            if not uThread.OpenExecDS(exec, 'Добавление разрешений Модулю для ролей', 'Сохранение данных','ExecuteScript') then
              raise Exception.Create(RLS_EXEPT);
            if qBlock.Transaction.InTransaction then
             qBlock.Transaction.Commit;
          end;
        finally
          FreeAndNil(exec);
        end;
      end;
      Result := True;
      {$EndRegion}
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.ClassName + ': ' +  e.Message),
              PChar('Ошибка сохранение данных'), MB_OK + MB_ICONERROR);
        Result := False;
      end;
    end;
    SetNppBlock;
  end;
end;

function TListModule.UIParent(ARecRed: TRecRed): Boolean;
begin
  if not ARecRed.GetFlagParams(tfpIsLoad) then
  begin
    qBlock.Close;
    qBlock.SQL.Clear;
    try
      qBlock.GenerateParamNames := True;
      qBlock.SQL.Text := SQLModul[4];
      qBlock.Prepare;
      qBlock.ParamByName('recid').AsInteger := RecRed.RecID;
      qBlock.ParamByName('parent').AsInteger := RecRed.Parent;
      if uThread.OpenExecDS(qBlock, 'Объеденяем модули в пакет', 'Сохранение данных','ExecSQL') then
        Result := True
      else
        raise Exception.Create(RLS_EXEPT);
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.ClassName + ': ' +  e.Message),
              PChar('Ошибка сохранение данных'), MB_OK + MB_ICONERROR);
        Result := False;
      end;
    end;
  end
  else
    Result := True;
end;

function TListModule.UpdateOrInsertM(ARecRed: TRecRed): Boolean;
var
  ss: TStringStream;
  rv, rvn: TRecVersion;
  rm, rd: TRecModule;
  i, idx: Integer;
begin
  {$Region ' Добавление нового файла, либо обновления текущего'}
  With Self do
  begin
    try
      qBlock.SQL.Clear;
      qBlock.GenerateParamNames := True;
      qBlock.SQL.Text := SQLModul[3];
      qBlock.Prepare;
      qBlock.ParamByName('recid').AsInteger := RecRed.RecID;
      qBlock.ParamByName('npp').AsInteger := RecRed.NPP;
      qBlock.ParamByName('module_name').AsString := RecRed.Descr;
      qBlock.ParamByName('no_export_regions').AsInteger :=
        IfThen(RecRed.GetFlagParams(tfpNoExportRegion), 1, 0);
      qBlock.ParamByName('name_file').AsString := RecRed.FileName;
      qBlock.ParamByName('name_ext').AsString := RecRed.Ext;
      qBlock.ParamByName('sizefile').AsInteger := RecRed.Size;
      qBlock.ParamByName('isupdate').AsInteger := RecRed.Update;
      qBlock.ParamByName('isactive').AsInteger := RecRed.Active;
      qBlock.ParamByName('isonestart').AsInteger :=
        IfThen(RecRed.GetFlagParams(tfpOsn) and RecRed.GetFlagParams(tfpOneStart), 1,0);
      qBlock.ParamByName('isall_users').AsInteger :=
        IfThen(RecRed.GetFlagParams(tfpOsn)and RecRed.GetFlagParams(tfpAllUser), 1,0);
      qBlock.ParamByName('catalog_upd').AsString := RecRed.FileDir;
      qBlock.ParamByName('actualversion').AsString := RecRed.Version;
      qBlock.ParamByName('date_from').AsDateTime := RecRed.DateFrom;
      if(RecRed.DateTo = 0)or not(RecRed.Active < 2)then
        qBlock.ParamByName('date_to').Clear
      else
        qBlock.ParamByName('date_to').AsDateTime := IfThen((RecRed.DateTo = 0), Now, RecRed.DateTo);

      if Assigned(RecRed.Image)then
      begin
        RecRed.Image.Seek(0, soBeginning);
        qBlock.ParamByName('image').LoadFromStream(RecRed.Image, ftBlob);
      end
      else
        qBlock.ParamByName('image').Clear;
      ss := TStringStream.Create(RecRed.Info);
      try
        ss.Seek(0, soBeginning);
        qBlock.ParamByName('info').LoadFromStream(ss, ftBlob);
        Result := uThread.OpenExecDS(qBlock, 'Обновление данных модуля', 'Сохранение данных','ExecSQL');
         if not Result then
           raise Exception.Create(RLS_EXEPT)
         else
         begin
           if not RecRed.GetFlagParams(tfpIsLoad) then
           begin
             {$Region ' Создаем версию для нового файла  '}
             rv := TRecVersion.Create(nil, RecRed, 1);
             rv.fInfo := 'Первая версия файла, автоматически создается при добавление в список';
             rv.fDateFrom := RecRed.DateFrom;
             rv.fSize := RecRed.Size;
             rv.fVersion := RecRed.Version;
             rv.fFileName := StringReplace(RecRed.Version,'.','-',[rfReplaceAll, rfIgnoreCase]);
             rv.fFileName := ExtractFileName(ChangeFileExt(RecRed.fPathFile, ''))+ '('+ rv.fFileName +')';
             rv.fFileName := RecRed.FileName + '.' + LowerCase(RecRed.fExt)  + '\' + rv.fFileName + ExtractFileExt(RecRed.fPathFile);
             rv.SetFlagParams(tfpvActive, True);
             rv.SetFlagParams(tfpvRollBack, False);
             {$EndRegion}
             try
               {$IfDef UPDATE}
                try
                  Result := Self.SaveVersion(rv, RecRed.fPathFile, False);
                except
                  on e: Exception do
                  begin
                    FreeAndNil(rv);
                    raise Exception.Create(e.Message);
                  end
                end;
               {$EndIf}
               if Result then
               begin
                 rm := TRecModule.Create(nil);
                 rm.Assign(ARecRed);
                 rm.SetFlagParams(tfpIsLoad, True);
                 rm.SetFlagParams(tfpIsRead, True);
                 rvn := TRecVersion.Create(nil, rm, rv.NPP);
                 rvn.Assign(rv);
                 if rm.GetFlagParams(tfpOsn) then
                 begin
                   if Self.fListAPP.BinarySearch(rm, i) then
                   begin
                     Self.fListAPP.Items[i].Assign(rm);
                     FreeAndNil(rm);
                     FreeAndNil(rvn);
                   end
                   else
                   begin
                     Self.fListAPP.Insert(i, rm);
                     rm.fListVers.Add(rvn);
                   end;
                 end
                 else
                 begin
                   rd := TRecModule.Create(nil);
                   rd.Assign(rm);
                   if Self.fListlib.BinarySearch(rm, i) then
                   begin
                     Self.fListlib.Items[i].Assign(rm);
                     FreeAndNil(rm);
                   end
                   else
                   begin
                     Self.fListlib.Insert(i, rm);
                     rm.fListVers.Add(rvn);
                   end;
                   for i := 0 to Pred(Self.fListAPP.Count)do
                   begin
                     if(Self.fListAPP[i].fRecID = RecRed.Parent)then
                     begin
                       if not Self.fListAPP[i].fListLib.BinarySearch(rd, idx)then
                       begin
                         Self.fListAPP[i].fListLib.Insert(idx, rd);
                         rd := nil;
                       end
                       else
                         FreeAndNil(rd);
                       Break;
                     end;
                   end;
                   if Assigned(rd) then
                     FreeAndNil(rd);
                 end;
               end;
             finally
               FreeAndNil(rv);
             end;
           end
           else
           begin
             if RecRed.GetFlagParams(tfpOsn) then
             begin
               if Self.fListAPP.BinarySearch(ARecRed, i) then
               begin
                 Self.fListAPP[i].Assign(ARecRed);
                 Self.fListAPP[i].SetFlagParams(tfpIsLoad, True);
                 Self.fListAPP[i].SetFlagParams(tfpIsRead, True);
               end;
             end
             else
             begin
               if Self.fListlib.BinarySearch(ARecRed, i) then
                 Self.fListlib[i].Assign(ARecRed);
               for i := 0 to Pred(Self.fListAPP.Count)do
               begin
                 for rd in Self.fListAPP[i].ListLib do
                 begin
                   if(rd.fRecID = ARecRed.RecID)then
                   begin
                     rd.Assign(ARecRed);
                     Break;
                   end;
                 end;
               end;
             end;
           end;
         end;
      finally
        FreeAndNil(ss);
        qBlock.Close;
      end;
    except
      on e: Exception do
      begin
        MessageBox(0, PChar(e.ClassName + ': ' +  e.Message),
              PChar('Ошибка сохранение данных'), MB_OK + MB_ICONERROR);
        Result := False;
      end;
    end;
    if Result then
      Result := UIParent(ARecRed);
  end;
  {$EndRegion}
end;

{ TRecModule }
procedure TRecModule.Assign(ARec: TRecModule);
begin
  with Self do
  begin
    fRecID := ARec.RecID;
    fNPP := ARec.NPP;
    fSize := ARec.Size;
    fInfo := ARec.Info;
    fFileName := ARec.FileName;
    fFileDir := ARec.FileDir;
    fDescr := ARec.Descr;
    fExt:= ARec.Ext;
    fDateFrom :=ARec.DateFrom;
    fDateTo := ARec.DateTo;
    fUpdate := ARec.Update;
    fActive := ARec.Active;
    fFlags := ARec.fFlags;
    fInfoClose := ARec.InfoClose;
    Checked := ARec.Checked;
    fVersion := ARec.Version;
    FreeAndNil(fImage);
    if Assigned(ARec.Image) then
    begin
      fImage := TMemoryStream.Create;
      ARec.Image.Seek(0, soBeginning);
      fImage.LoadFromStream(ARec.Image);
    end;
  end;
end;

constructor TRecModule.Create(AOwner: TComponent);
begin
  inherited;
  Self.fInfo := '';
  Self.fImage := nil;
  Self.fRecID := -1;
  with Self do
  begin
    fListVers := TObjectList<TRecVersion>.Create(
      IComparer<TRecVersion>(
        function(const Left, Rigth: TRecVersion): Integer
        begin
          Result := CompareValue(Left.NPP, Rigth.NPP);
        end), False);

    fListLib := TObjectList<TRecModule>.Create(
      IComparer<TRecModule>(
        function(const Left, Rigth: TRecModule): Integer
        begin
          Result := CompareValue(Left.RecID, Rigth.RecID);
          if(Result <>0) then
           Result := CompareText(Left.FileName, Rigth.FileName);
        end), False);
    {$IfDef UPDATE}
    fListRole:= TObjectList<TRecRole>.Create(
      IComparer<TRecRole>(
        function(const Left, Rigth: TRecRole): Integer
        begin
          Result := CompareValue(Left.ID, Rigth.ID);
        end), False);
    {$EndIf}
  end;
end;

destructor TRecModule.Destroy;
begin
  Self.FreeList;
  FreeAndNil(Self.fListVers);
  FreeAndNil(Self.fListLib);
  FreeAndNil(Self.fImage);
  {$IfDef UPDATE}FreeAndNil(Self.fListRole); {$EndIf}
  inherited;
end;

procedure TRecModule.FreeList;
var
  rec: TRecVersion;
  recm: TRecModule;
  {$IfDef UPDATE} Recr: TRecRole;{$EndIf}
begin
  while not(Self.fListLib.Count =0) do
  begin
    recm := Self.fListlib.Items[Pred(Self.fListlib.Count)];
    try
      Self.fListlib.Items[Pred(Self.fListlib.Count)] := nil;
      Self.fListLib.Delete(Pred(Self.fListLib.Count));
    finally
      FreeAndNil(recm);
    end;
  end;
  while not(Self.fListVers.Count =0) do
  begin
    rec := Self.fListVers.Items[Pred(Self.fListVers.Count)];
    try
      Self.fListVers.Items[Pred(Self.fListVers.Count)] := nil;
      Self.fListVers.Delete(Pred(Self.fListVers.Count));
    finally
      FreeAndNil(rec);
    end;
  end;
  {$IfDef UPDATE}
  while not(Self.fListRole.Count = 0) do
  begin
    Recr := Self.fListRole.Items[Pred(Self.fListRole.Count)];
    try
      Self.fListRole.Items[Pred(Self.fListRole.Count)] := nil;
      Self.fListRole.Delete(Pred(Self.fListRole.Count));
    finally
      FreeAndNil(Recr);
    end;
  end;
  {$EndIf}
end;

function TRecModule.GetFlagParams(const AParams: TFlagsParams): Boolean;
begin
  Result := MyPObj.GetFlagIdx(Self.fFlags, Byte(AParams));
end;

function TRecModule.GetRed: Boolean;
begin
  Result := MyPObj.GetFlagIdx(Self.fFlags, 6);
end;

procedure TRecModule.SetChecked(const AChecked: Boolean);
begin
  Self.fChecked := AChecked;
end;

procedure TRecModule.SetFlagParams(const AParams: TFlagsParams;
  const AValue: Boolean);
begin
  Self.fFlags:= MyPObj.SetFlagIdx(Self.fFlags, Byte(AParams), AValue);
end;

procedure TRecModule.SetImage(const AImage: TMemoryStream);
begin
  if Assigned(AImage)and Self.GetFlagParams(tfpOsn)then
  begin
    if(AImage.Size > 0)then
    begin
      FreeAndNil(Self.fImage);
      Self.fImage := TMemoryStream.Create;
      AImage.Seek(0, soBeginning);
      Self.Image.LoadFromStream(AImage);
    end;
  end;
end;

procedure TRecModule.SetRed(const Value: Boolean);
begin
  Self.SetFlagParams(tfpNoRead, Value);
end;

function TRecModule.Status: String;
begin
  case Self.fActive of
    0: Result := 'Информация';
    1: Result := 'Активен';
    2: Result := 'Деактивирован';
  end;
  if not(Self.fActive = 1 ) then
    Result := Result + sLineBreak + 'Без проверки'
  else
  begin
    if MyPObj.GetFlagIdx(Self.fFlags, 1) then
    begin
      Result := Result + sLineBreak +IfThen(Self.GetFlagParams(tfpOneStart),'Первый старт','Доп. выбор');
      Result := Result + sLineBreak +IfThen(Self.GetFlagParams(tfpAllUser),'Доступен всем','Ограничен ролью');
    end
    else
      Result := Result + sLineBreak + 'Подгружаем';
  end;
end;

{ TRecBase }
constructor TRecBase.Create(const ACodeNode: DWord; const ABaseDisplFont: TFont;
  ADataObject: TComponent; const AClearData: Boolean);
begin
  Inherited Create(ACodeNode,ABaseDisplFont, ADataObject, AClearData);
  Self.ColumnHeigth := 2;
  case ACodeNode of
    1:
    begin
      Self.fNameNode := {$IfDef DEBUG}'Основные модули'{$Else}'Доступные модули'{$EndIf};
      Self.fInfo :=
      'Файлы программ ".exe"'+sLineBreak +
      'Файлы доп. модулей ".bin" и т.п.';
    end;
    2:
    begin
      Self.fNameNode := {$IfDef DEBUG}'Подгружаемые файлы'{$Else}'Вспомогательные файлы'{$EndIf};
      Self.fInfo :=
      'Библиотеки ".dll",'+sLineBreak +
      'Файлы конфигураций ".ini",'+sLineBreak +
      'Файлы картинок ".jpg" и т.п.';
    end;
    3:
    begin
      Self.fNameNode := 'История версий';
      Self.fInfo := '';
    end
  else
  end;
end;

procedure TRecBase.GetText(Column: integer; var CellText: String);
begin
  inherited;
  case Column of
    1: CellText := Self.fNameNode;
    5: CellText := Self.fInfo;
  else
    CellText := '';
  end;
end;

{ TRecModulVTI }
constructor TRecModulVTI.Create(ADataObject: TRecModule; const ABaseDisplFont: TFont);
begin
  Inherited Create(ADataObject.fRecID,  ABaseDisplFont, ADataObject);
  Self.ColumnHeigth := 1;
end;

procedure TRecModulVTI.GetText(Column: integer; var CellText: String);
var
  obj: TRecModule;
begin
  inherited;
  if Assigned(Self)then
  if Assigned(Self.DataObject)then
  begin
    obj :=  TRecModule(Self.DataObject);
    case Column of
      1: CellText := ifthen(
         (Length(Trim(obj.FileDir)) > 0)and(MyPObj.GetFlagIdx(obj.fFlags, 1)),
             IncludeTrailingPathDelimiter(obj.FileDir),'' ) +
              obj.FileName + '.' + obj.fExt;
      2: CellText := obj.fDescr  ;
      3: CellText := obj.Status;
      4: CellText :=
                 'Версия: ' + obj.Version + sLineBreak +
                 'Дата: ' + FormatDateTime('ddd. dd.mm.yyyy hh:nn', obj.fDateFrom) + sLineBreak +
                 'Размер: ' + MyPObj.GetStrOfSize(obj.fSize, true);

      5: CellText := IfThen(obj.fActive < 2, obj.Info, obj.fInfoClose);
    else
      CellText := '';
    end;
  end;
end;

{ TRecRed }

constructor TRecRed.Create(AOwner: TComponent; APathFile: String; AParent: Integer);
begin
  inherited Create(AOwner);
  with Self do
  begin
    fRecID := 0;
    fPathFile := APathFile;
    fParent := AParent;
    fRecID := -1;
    Self.LoadDataFile;
  end;
end;

constructor TRecRed.Create(AOwner: TComponent; ARecModul: TRecModule);
begin
  inherited Create(AOwner);
  Self.Assign(ARecModul);
  Self.fFlags:= MyPObj.SetFlagIdx(ARecModul.fFlags, 5, True);
end;

procedure TRecRed.LoadDataFile;
var
  sr : TSearchRec;
begin
  with Self do
  begin
    fRecID := -1;
    fVersion := MyPObj.GetVersionFile(fPathFile);
    // В случае вновь созданного файла загружаем данные
    if Self.fRecID < 0 then
    begin
      fDescr := MyPObj.GetFileInfoFromParams(fPathFile);
      fFileName := ExtractFileName(ChangeFileExt(fPathFile, ''));
      fExt := Copy(ExtractFileExt(fPathFile), 2);
      fUpdate := IfThen(Length(Trim(Version))> 0, 1, 0);
      fNPP := -1;
      fDateFrom := MyPObj.GetDateFile(fPathFile, 1);
      fDateTo := 0;
      fActive := 1;
    end;
    try
      if FindFirst(fPathFile, faAnyFile, sr ) = 0 then
        fSize := Int64(sr.FindData.nFileSizeHigh) shl Int64(32) + Int64(sr.FindData.nFileSizeLow);
    finally
      FindClose(sr) ;
    end;
  end;
end;

{ TRecVersion }
procedure TRecVersion.Assign(ARec: TRecVersion);
begin
  with Self do
  begin
    fSize:= ARec.fSize;
    fDateFrom := ARec.DateFrom;
    fDateTo:= ARec.DateTo;
    fInfo := ARec.fInfo;
    fInfRollBack := ARec.fInfRollBack;
    fFileName := ARec.fFileName;
    fVersion := ARec.fVersion;
    fFlags := ARec.fFlags;
  end;
end;

constructor TRecVersion.Create(AOwner: TComponent; ABaseModul: TRecModule; const ANPP: Integer);
begin
  inherited Create(AOwner);
  Self.fModul := ABaseModul;
  Self.fNPP := ANPP;
  Self.fFlags := 0;
end;

function TRecVersion.GetFlagParams(const AParamsVer: TFlagsParamsVer): Boolean;
begin
  Result := MyPObj.GetFlagIdx(Self.fFlags, Byte(AParamsVer));
end;

procedure TRecVersion.SetFlagParams(const AParamsVer: TFlagsParamsVer;
  const AValue: Boolean);
begin
  Self.fFlags:= MyPObj.SetFlagIdx(Self.fFlags, Byte(AParamsVer), AValue);
end;

{ TRecVersionVTI }

constructor TRecVersionVTI.Create(ADataObject: TRecVersion;
  const ABaseDisplFont: TFont);
begin
  inherited Create(ADataObject.NPP,  ABaseDisplFont, ADataObject);
  Self.ColumnHeigth := 1;
end;

procedure TRecVersionVTI.GetText(Column: integer; var CellText: String);
var
  obj: TRecVersion;
begin
  inherited;
  if Assigned(Self)then
  begin
    if Assigned(Self.DataObject)then
    begin
      obj :=  TRecVersion(Self.DataObject);
      case Column of
        1: CellText := 'Версия №' + IntToStr(obj.NPP) ;
        2:
        begin
          CellText := 'Порядок версий';
          if obj.GetFlagParams(tfpvRollBack) then
            CellText := 'Версия снята с релиза' + sLineBreak +  'от ' + FormatDateTime('DDD. dd.mm.yyyy hh:nn', obj.fDateTo)
          else
            CellText := 'Активен в порядке версий';
        end;
        3:
        begin
          if obj.GetFlagParams(tfpvActive)and not obj.GetFlagParams(tfpvRollBack)then
            CellText := 'Рабочий'
          else
            if not obj.GetFlagParams(tfpvActive)then
              CellText := 'Информационый'
            else
              if obj.GetFlagParams(tfpvRollBack)then
                CellText := 'Закрыт/Откат';
          if SameText(Obj.fVersion, obj.fModul.Version) then
            CellText := CellText + sLineBreak + ifThen( obj.GetFlagParams(tfpvRollBack), 'Внимание', 'Актуален')
          else
            CellText := CellText + sLineBreak +
                  ifThen( obj.GetFlagParams(tfpvRollBack), 'Снят с релиза', 'Заменён релизом');
        end;
        4: CellText :=
                   'Версия: ' + obj.fVersion + sLineBreak +
                   'Дата: ' + FormatDateTime('ddd. dd.mm.yyyy hh:nn', obj.fDateFrom) + sLineBreak +
                   'Размер: ' + MyPObj.GetStrOfSize(obj.fSize, true);
        5: CellText := ifThen(obj.GetFlagParams(tfpvRollBack), obj.fInfRollBack, obj.Info);
      else
        CellText := '';
      end;
    end;
  end;
end;

end.
