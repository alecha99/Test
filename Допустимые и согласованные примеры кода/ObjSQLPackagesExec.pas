unit ObjSQLPackagesExec;

{$WARN SYMBOL_PLATFORM OFF}
{$WARN UNIT_DEPRECATED OFF}
{$I Directives.inc}
interface
  uses Winapi.Windows, Vcl.Forms, Vcl.Dialogs, System.SysUtils,  System.Classes, Vcl.Graphics,
       IBX.IBDatabase, IBX.IBCustomDataSet, IBX.IBQuery, Data.DB, IBX.IBBlob, IBX.IBScript,
       VTIHelper,{ System.Generics.Defaults, System.Generics.Collections,} ComUtils, UserSessionDB,
       VirtualTrees, CustomDM, ObjSQLPackages;
{$Region ' Основной список скриптов '}
const
  SQLPacketExec:array[0..6]of string =(
  '  with vers as(  ' + sLineBreak +
  '            select coalesce(max(recid), 0) recid, coalesce(max(npp_vers),0) ver, sqlscript_id, pack_recid ' + sLineBreak +
  '            from sys_sqlpackege_vers pack where pack.pack_recid = :pack_recid group by pack.sqlscript_id, pack.pack_recid), ' + sLineBreak +
  '  vMax as(select vers.recid, vers.ver, max(v.npp_vers) rec_verSQL ,vers.sqlscript_id ' + sLineBreak +
  '            from vers  ' + sLineBreak +
  '              join sys_sqlpackege_vers v on(v.pack_recid = vers.pack_recid)and(v.is_update = 1)and(vers.sqlscript_id=v.sqlscript_id) ' + sLineBreak +
  '             group by vers.recid, vers.ver, vers.sqlscript_id ) ' + sLineBreak +
  '    select vMax.recid, vMax.ver as npp_vers, vMax.sqlscript_id,' + sLineBreak +
  '      v.SCRIPTS,   ' + sLineBreak +
  '      v.ROLLBACK_SCR ' + sLineBreak +
  '    from vMax join sys_sqlpackege_vers v on(vMax.sqlscript_id = v.sqlscript_id)and(v.is_update = 1)and(v.npp_vers=vMax.rec_verSQL) ',
  '  select' + sLineBreak +
  '    coalesce((select max(recid) from sys_sqlpackege_vers), 0) + 1 recid, ' + sLineBreak +
  '    coalesce((select max(npp_vers) from sys_sqlpackege_vers v where v.pack_recid =:pack_recid), 0) + 1 npp_vers' + sLineBreak +
  '  from rdb$database',
  'insert into sys_sqlpackege_vers (pack_recid, sqlscript_id, npp_sql, npp_vers, object_recid, scripts, rollback_scr, charterm, recid, is_update) ' + sLineBreak +
  '  values (:pack_recid, :sqlscript_id, :npp_sql, :npp_vers, :object_recid, :scripts, :rollback_scr, :charterm, :recid, :is_update)' ,
  'insert into sys_sqlscript_exec (execid, recid_pack, date_start, user_start, user_full, user_role, role_info, user_win, mashine_ip, ' + sLineBreak + 
                                  ' mashine_win, user_sys, role_exec, session_id, recid_vers, is_rollback)' + sLineBreak +
                          'values (:execid, :recid_pack, :date_start, :user_start, :user_full, :user_role, :role_info, :user_win, :mashine_ip, ' + sLineBreak +
                                 ' :mashine_win, :user_sys, :role_exec, :session_id, :recid_vers, :is_rollback)',
  'select coalesce(max(execid), 0) + 1 RecID from rdb$database  join sys_sqlscript_exec on 1=1',
  'update sys_sqlscript_exec set  ' + sLineBreak +
  '     date_stopped = :date_stopped, ' + sLineBreak +
  '     message_txt = :message_txt ' + sLineBreak +
  ' where (execid = :execid)',
  'insert into sys_sqlpack_exec_db (baseid, pack_recid, exec_recid, sql_recid, sql_vers, result, log_str, date_start, date_stop, is_rollback) ' + sLineBreak +
  '                         values (:baseid, :pack_recid, :exec_recid, :sql_recid, :sql_vers, :result, :log_str, :date_start, :date_stop, :is_rollback)'
  );
{$EndRegion}
Type
   TPacketExec = class;
   TRecVers = class(TComponent)
   private
     fRecID,
     fVers: Integer;
     fTechnoDM: TfCustomDM;
     fPacket: TPacketExec;
     fBasePacket: TRecSqlPack;
     fSession: TUserSessionDB;
     fNewRecord: Boolean;
     fUpdateVers: array of integer;
     procedure AddVersion;
     function CheckActualVers(AQuery: TIBQuery): Boolean;
     procedure GetRecIDVersion;
     procedure Prepare;
   public
    constructor Create(APacket: TPacketExec; ATechnoDM: TfCustomDM; ASession: TUserSessionDB; ABasePacket: TRecSqlPack); reintroduce; overload;
    property RecID: Integer read fRecID;
    property Vers: Integer read fVers;
   end;

   TPacketExec = class(TComponent)
   private
     fUserExec,
     fRoleExec: String;
     fTimeStart,
     fTimeFinish: TDateTime;
     fTechnoDM: TfCustomDM;
     fRecID: Integer;
     fPacketBase: TRecSqlPack;
     fFullPacket:TRecSqlPack;
     fRecVers: TRecVers;
     fSession: TUserSessionDB;
     fNoValidTxt: string;
     fIs_Rollback: Boolean;
     function GetRecIDExecute: Integer;
     function ExecPack(aRegionDM: TfCustomDM; out AIsBreak: Boolean): String;
     procedure ParceError(Sender: TObject; Error, SQLText: string; LineIndex: Integer);
     procedure SetDBLOGS(const ABegDate: TDateTime; const ABaseID, AExecRLS: Integer;
           AScriptExec: TRecSqlScript; const ALogs: string);
   public
     constructor Create(AOwner: TComponent; APacketBase, AFullPacket: TRecSqlPack; ATechnoDM: TfCustomDM; ASession: TUserSessionDB); reintroduce; overload;
     destructor Destroy; override;
     procedure Prepare(const aIs_Rollback: Boolean);
     procedure Finish(const ALog: string);
     function BaseExec(const ABaseID: Integer; const APathBase, ABaseName, ALogBase: String; var AIsBreak: Boolean): String;
     property UserExec: String read fUserExec write fUserExec;
     property RoleExec: String read fRoleExec write fRoleExec;
     property Is_Rollback: Boolean read fIs_Rollback write fIs_Rollback;
     property RecVers: TRecVers read fRecVers;
     property TimeStart: TDateTime read fTimeStart;
     property TimeFinish: TDateTime read fTimeFinish;
   end;

implementation
uses uThread, System.StrUtils, System.Math;
{ TPacketExec }
constructor TPacketExec.Create(AOwner: TComponent; APacketBase, AFullPacket: TRecSqlPack; ATechnoDM: TfCustomDM; ASession: TUserSessionDB);
begin
  inherited Create(AOwner);
  with Self do
  begin
    fTechnoDM := ATechnoDM;
    fPacketBase := APacketBase;
    fSession := ASession;
    fFullPacket := AFullPacket;
    fRecVers := TRecVers.Create(Self, fTechnoDM, fSession, fFullPacket);
  end;
end;

destructor TPacketExec.Destroy;
begin
  FreeAndNil(Self.fRecVers);
  inherited;
end;

procedure TPacketExec.ParceError(Sender: TObject; Error, SQLText: string;
  LineIndex: Integer);
begin
  Self.fNoValidTxt := Self.fNoValidTxt + sLineBreak +
     Format('  Firebird Error: Строка: "%0:d" Message: "%1:s", В тексте : ' + sLineBreak + '    %2:s', [LineIndex, Error, SQLText]);
end;

function TPacketExec.ExecPack(aRegionDM: TfCustomDM; out AIsBreak: Boolean): String;
var
  RS: TRecSqlScript;
  ScriptExec: TIBScript;
  rls: SmallInt;
  ss: string;
  bb: Boolean;
  dt: TDateTime;
begin
  bb := False;
  AIsBreak := False;
  with Self do
  begin
    for RS in fPacketBase.ListScript do
    begin
      if AIsBreak then
        Break;
      fNoValidTxt := '';
      dt := Now;
      ss := '  Скрипт №' + IntToStr(RS.NPP) + ' / ';
      if(Length(Trim(RS.SQLScriptRollBack))=0)and Self.fIs_Rollback then
      begin
        ss := ss + '  Скрипт пропущен: нет тела скрипта "Отката/отмены"';
        Continue;
      end;
      ScriptExec:= TIBScript.Create(Self);
      rls := 0;
      try
        ScriptExec.Database := aRegionDM.Database;
        ScriptExec.Transaction := aRegionDM.TransactionWrite;
        if ScriptExec.Transaction.InTransaction then
          ScriptExec.Transaction.Commit;
        ScriptExec.Terminator := RS.Term;
        ScriptExec.Script.Text := IfThen(Self.fIs_Rollback, RS.SQLScriptRollBack, RS.SQLScript);
        ScriptExec.OnParseError := ParceError;
        bb := ScriptExec.ValidateScript;
        ss := ss + '  Валидация: ' + IfThen(bb, 'пройдена','не пройдена');
        if not bb then
          ss := ss + '  ' + Self.fNoValidTxt
        else
        begin
          inc(rls);
          ss := ss  + sLineBreak +
             '  Подготовка к выполнению: ' + FormatDateTime('dd.mm.yyyy hh:nn:ss', Now);
          try
            if not uThread.OpenExecDS(ScriptExec, RS.SQInfo, 'Выполнение скрипта','ExecuteScript',True, True, True) then
              raise Exception.Create(RLS_EXEPT);
            ss := ss + sLineBreak + '  Скрипт выполнен успешно: ' + FormatDateTime('dd.mm.yyyy hh:nn:ss', Now);
            inc(rls);
          except
            on e: Exception do
            begin
              bb := False;
              AIsBreak := uThread.isThreadBreak;
              ss := ss + sLineBreak + '  Ошибка выполнения: ' + IfThen(AIsBreak, 'Процесс прерван пользователем', e.Message);
              if ScriptExec.Transaction.InTransaction then
                ScriptExec.Transaction.Rollback;
            end;
          end;
        end;
      finally
        if ScriptExec.Transaction.InTransaction then
          if bb then
            ScriptExec.Transaction.Commit;
        Inc(rls);
        if bb then
          ss := ss + sLineBreak + '  .DataBase.Transaction.Commit;'
        else
          ss := ss + sLineBreak + '  .DataBase.Transaction.Rollback;';
        Self.SetDBLOGS(dt, aRegionDM.BaseID, rls, RS, ss);
        FreeAndNil(ScriptExec);
      end;
      Result := Result + IfThen(Result > '',sLineBreak + sLineBreak,'') + ss;
    end;
  end;
end;

procedure TPacketExec.Finish(const ALog: string);
var
  q: TIBQuery;
  ss: TStringStream;
begin
  with Self do
  begin
    fTimeFinish:= Now;

    fRecVers.Prepare;
    {$Region 'Создаем запись истории запусков'}
    q := Self.fTechnoDM.OpenSQL(SQLPacketExec[5], True, False);
    try
      with fPacketBase do
      begin
        fTechnoDM.SetParamByName(q, 'date_stopped', fTimeFinish);
        fTechnoDM.SetParamByName(q, 'execid', fRecID);
        ss := TStringStream.Create(ALog);
        ss.Seek(0, soBeginning);
        q.ParamByName('message_txt').LoadFromStream(ss, ftBlob);

        if not uThread.OpenExecDS(q, 'Обновление данных "SQL Packet"', 'Сохранение данных','ExecSQL') then
          raise Exception.Create(RLS_EXEPT);
        q.Close;
      end;
      if Self.fTechnoDM.TransactionWrite.InTransaction then
          Self.fTechnoDM.TransactionWrite.Commit;
    finally
      FreeAndNil(q);
      FreeAndNil(ss);
    end;
    {$EndRegion}
  end;
end;

function TPacketExec.BaseExec(const ABaseID: Integer; const APathBase, ABaseName, ALogBase: String; var AIsBreak: Boolean): String;
var
  dmDb: TfCustomDM;
  bb: Boolean;
begin
  try
    if SameStr(fUserExec, 'SYSDBA')then
      dmDb := TfCustomDM.Create(Self, APathBase, ABaseID, True, 1, True)
    else
    begin
      dmDb := TfCustomDM.Create(Self, APathBase, fUserExec, Self.fTechnoDM.UserPSW, fRoleExec, 1, AIsBreak);
      dmDb.BaseID := ABaseID;
    end;
    AIsBreak := dmDb.IsBreak;
  except
    on e:Exception do
    begin
      AIsBreak := dmDb.IsBreak;
      if AIsBreak then
        dmDb.ExceptDB := 'Ожидание связи с сервером прервано пользователем';
    end;
  end;
  try
    bb:= dmDb.Active;
    if Self.fPacketBase.IS_WAIT and bb then
    begin
      dmDb.TransactionWrite.Active := False;
      dmDb.TransactionWrite.Params.Clear;
      dmDb.TransactionWrite.Params.Add('write');
      dmDb.TransactionWrite.Params.Add('consistency');
      dmDb.TransactionWrite.Active := True;
    end;

    Result := '  ' + IntToStr(ABaseID) + ': ' + Trim(ABaseName) + IfThen(bb, ' / связь установлена',' ошибка связи');
    if not bb then
      Result := Result + sLineBreak + ('   ' + IfThen(AIsBreak, 'Основной процесс прерван пользователем', dmDb.ExceptDB ))
    else
      Result := Result + sLineBreak + Self.ExecPack(dmDb, AIsBreak);
  finally
    dmDb.Active := False;
    FreeAndNil(dmDb);
  end;
end;

function TPacketExec.GetRecIDExecute: Integer;
var
  q: TIBQuery;  
begin
  {$Region 'Получим номер новой версии '}
  q := Self.fTechnoDM.OpenSQL(SQLPacketExec[4]);
  try
    if not uThread.OpenExecDS(q, 'Получение последовательного номера', 'Получение данных','ExecSQL') then
       raise Exception.Create(RLS_EXEPT);
    Result := q.FieldByName('RecID').AsInteger;
  finally
    FreeAndNil(q);
  end;
  {$EndRegion}
end;

procedure TPacketExec.Prepare(const aIs_Rollback: Boolean);
var
  q: TIBQuery;
begin
  with Self do
  begin
    fTimeStart:= Now;
    fRecID := GetRecIDExecute;
    fRecVers.Prepare;
    fIs_Rollback := aIs_Rollback;
    {$Region 'Создаем запись истории запусков'}
    q := Self.fTechnoDM.OpenSQL(SQLPacketExec[3], True, False);
    try
      with fPacketBase do
      begin
        fTechnoDM.SetParamByName(q, 'recid_pack', fPacketBase.RecID);
        fTechnoDM.SetParamByName(q, 'date_start', fTimeStart);
        fTechnoDM.SetParamByName(q, 'user_start', fSession.UserDeb);
        fTechnoDM.SetParamByName(q, 'user_full', fSession.UserInfo);
        fTechnoDM.SetParamByName(q, 'user_role', fSession.RoleDeb);
        fTechnoDM.SetParamByName(q, 'role_info', fSession.RoleInfo);
        fTechnoDM.SetParamByName(q, 'user_win', fSession.UserOS);
        fTechnoDM.SetParamByName(q, 'mashine_ip', fSession.ComputerIP);
        fTechnoDM.SetParamByName(q, 'mashine_win', fSession.ComputerName);
        fTechnoDM.SetParamByName(q, 'user_sys', fUserExec);
        fTechnoDM.SetParamByName(q, 'role_exec', fRoleExec);
        fTechnoDM.SetParamByName(q, 'session_id', fTechnoDM.SessionID);
        fTechnoDM.SetParamByName(q, 'recid_vers', fRecVers.RecID);
        fTechnoDM.SetParamByName(q, 'execid', fRecID);
        fTechnoDM.SetParamByName(q, 'is_rollback', IfThen(aIs_Rollback, 1, 0));

        if not uThread.OpenExecDS(q, 'Обновление данных "SQL Packet"', 'Сохранение данных','ExecSQL') then
          raise Exception.Create(RLS_EXEPT);
        q.Close;
      end;
      if Self.fTechnoDM.TransactionWrite.InTransaction then
          Self.fTechnoDM.TransactionWrite.Commit;
    finally
      FreeAndNil(q);
    end;
    {$EndRegion}
  end;
end;

procedure TPacketExec.SetDBLOGS(const ABegDate: TDateTime; const ABaseID, AExecRLS: Integer;
           AScriptExec: TRecSqlScript; const ALogs: string);
var
  q: TIBQuery;
  ss: TStringStream;
begin
  with Self do
  begin
    {$Region 'Создаем запись истории запусков'}
    q := Self.fTechnoDM.OpenSQL(SQLPacketExec[6], True, False);
    try
      fTechnoDM.SetParamByName(q, 'baseid', ABaseID);
      fTechnoDM.SetParamByName(q, 'pack_recid', Self.fPacketBase.RecId);
      fTechnoDM.SetParamByName(q, 'exec_recid', Self.fRecID);
      fTechnoDM.SetParamByName(q, 'sql_recid', AScriptExec.RecID);
      fTechnoDM.SetParamByName(q, 'sql_vers', Self.fRecVers.fRecID);
      fTechnoDM.SetParamByName(q, 'result', AExecRLS);
      fTechnoDM.SetParamByName(q, 'date_start', ABegDate);
      fTechnoDM.SetParamByName(q, 'date_stop', Now);
      fTechnoDM.SetParamByName(q, 'is_rollback', IfThen(Self.Is_Rollback, 1, 0));
      ss := TStringStream.Create(ALogs);
      ss.Seek(0, soBeginning);
      q.ParamByName('log_str').LoadFromStream(ss, ftBlob);
      if not uThread.OpenExecDS(q, 'Сохраняем статистику по скрипту', 'Сохранение данных','ExecSQL') then
          raise Exception.Create(RLS_EXEPT);
        q.Close;
    finally
      if Self.fTechnoDM.TransactionWrite.InTransaction then
          Self.fTechnoDM.TransactionWrite.Commit;
      FreeAndNil(q);
      FreeAndNil(ss);
    end;
    {$EndRegion}
  end;
end;

{ TRecVers }
constructor TRecVers.Create(APacket: TPacketExec; ATechnoDM: TfCustomDM; ASession: TUserSessionDB; ABasePacket: TRecSqlPack);
begin
  inherited Create(nil);
  with Self do
  begin
    fTechnoDM := ATechnoDM;
    fBasePacket := ABasePacket;
    fPacket := APacket;
    fVers := 0;
    fSession := ASession;
    GetRecIDVersion;
  end;
end;

procedure TRecVers.AddVersion;
var
  q: TIBQuery;
begin
  with Self do
  begin
    fNewRecord := True;
    {$Region 'Создаем новую модификацию пакета'}
    q := Self.fTechnoDM.OpenSQL(SQLPacketExec[1]);
    fTechnoDM.SetParamByName(q, 'pack_recid', fPacket.fFullPacket.RecID);
    try
      if not uThread.OpenExecDS(q, 'Генерация порядкогого номера', 'Получение данных') then
        raise Exception.Create(RLS_EXEPT + sLineBreak + q.SQl.Text)
      else
      begin
        fRecID := q.FieldByName('recid').AsInteger;
        fVers := q.FieldByName('npp_vers').AsInteger;
      end;
    finally
      FreeAndNil(q);
    end;
    {$EndRegion}
  end;
end;

function TRecVers.CheckActualVers(AQuery: TIBQuery): Boolean;

  function GetRecSQL(const ARecSQL: Integer): TRecSqlScript;
  var
    i: Integer;
  begin
    Result := nil;
    for i := 0 to Pred(Self.fPacket.fFullPacket.ListScript.Count) do
      if(Self.fPacket.fFullPacket.ListScript.Items[i].RecID = ARecSQL)then
      begin
        Result := Self.fPacket.fFullPacket.ListScript.Items[i];
        Break;
      end;
  end;

var
  ss, sr: String;
  tmp: TRecSqlScript;
begin
  with Self do
  begin
    SetLength(fUpdateVers, 0);
    Result := AQuery.RecordCount = Self.fPacket.fFullPacket.ListScript.Count;
    begin
      while not AQuery.Eof do
      begin
        with Self.fPacket.fFullPacket do
        begin
          tmp := GetRecSQL(AQuery.FieldByName('sqlscript_id').AsInteger);
          if not Assigned(tmp) then
            Result := False
          else
          begin
            if not AQuery.FieldByName('scripts').IsNull then
              ss:= (AQuery.FieldByName('scripts') as TBlobField).AsString
            else
              ss := '';
            if not AQuery.FieldByName('rollback_scr').IsNull then
              sr:= (AQuery.FieldByName('rollback_scr') as TBlobField).AsString
            else
              sr := '';
            if (
              (AnsiCompareStr(tmp.SQLScript, ss)= 0) and
              (AnsiCompareStr(tmp.SQLScriptRollBack, sr) = 0))
            then
            begin
              SetLength(fUpdateVers, Length(fUpdateVers) + 1 );
              fUpdateVers[High(fUpdateVers)] := AQuery.FieldByName('sqlscript_id').AsInteger;
            end
            else
              Result := False;
          end;
        end;
        AQuery.Next;
      end;
    end;
  end;
end;

Procedure TRecVers.GetRecIDVersion;
var
  q: TIBQuery;
begin
  {$Region 'Получаем последнюю модификацию пакета'}
  q := Self.fTechnoDM.OpenSQL(SQLPacketExec[0]);
  try
    Self.fTechnoDM.SetParamByName(q, 'pack_recid', Self.fPacket.fPacketBase.RecID);
    if not uThread.OpenExecDS(q, 'Поиск версии пакета', 'Получение данных') then
      raise Exception.Create(RLS_EXEPT + sLineBreak + q.SQl.Text)
    else
    begin
      if(q.RecordCount > 0)then
      begin
        if Self.CheckActualVers(q) then
        begin
          Self.fRecID := q.FieldByName('RecID').AsInteger;
          Self.fVers := q.FieldByName('npp_vers').AsInteger;
        end
        else
          Self.AddVersion;
      end
      else
        Self.AddVersion;
    end;
  finally
    FreeAndNil(q);
  end;
  {$EndRegion}
end;

procedure TRecVers.Prepare;

  function GetUpdateSQL(const aRecIDSQL: Integer): Boolean;
  var
    i: Integer;
  begin
    Result := True;
    with Self do
    begin
      for i := Low(fUpdateVers) to High(fUpdateVers) do
      begin
        if fUpdateVers[i]= aRecIDSQL then
        begin
          Result := False;
          Break;
        end;
      end;
    end;
  end;

var
  rec: TRecSqlScript;
  q: TIBQuery;
  ss, sr: TStringStream;
begin
  sr := nil;
  ss := nil;
  with Self do
  begin
    if fNewRecord then
    begin
      {$Region 'Сохраняем версию пакета'}
      with fPacket.fFullPacket do
      begin
        for rec in ListScript do
        begin
          q := Self.fTechnoDM.OpenSQL(SQLPacketExec[2], True, True);
          try
            fTechnoDM.SetParamByName(q, 'pack_recid', fPacket.fFullPacket.RecID);
            fTechnoDM.SetParamByName(q, 'sqlscript_id', rec.RecID);
            fTechnoDM.SetParamByName(q, 'npp_sql', rec.NPP);
            fTechnoDM.SetParamByName(q, 'npp_vers', fVers);
            fTechnoDM.SetParamByName(q, 'is_update', ifThen(GetUpdateSQL(rec.RecID), 1, 0));
            if(rec.OBJ_RecID < 1)then
              q.ParamByName('object_recid').AsInteger
            else
              fTechnoDM.SetParamByName(q, 'object_recid', rec.OBJ_RecID);
            fTechnoDM.SetParamByName(q, 'recid', Self.fRecID);
            fTechnoDM.SetParamByName(q, 'charterm', rec.Term);
            ss := TStringStream.Create(rec.SQLScript);
            ss.Seek(0, soBeginning);
            q.ParamByName('scripts').LoadFromStream(ss, ftBlob);
            if(Length(Trim(rec.SQLScriptRollBack)) = 0)then
              q.ParamByName('rollback_scr').Clear
            else
            begin
              sr := TStringStream.Create(rec.SQLScriptRollBack);
              sr.Seek(0, soBeginning);
              q.ParamByName('rollback_scr').LoadFromStream(sr, ftBlob);
            end;
            if not uThread.OpenExecDS(q, 'Обновление данных "SQL Packet"', 'Сохранение данных','ExecSQL') then
             raise Exception.Create(RLS_EXEPT);
            q.Close;
          finally
            FreeAndNil(sr);
            FreeAndNil(ss);
            FreeAndNil(q);
            fNewRecord := False;
          end;
        end;
      end;
      SetLength(fUpdateVers, 0);
    end;
    {$EndRegion}
  end;
end;

end.
