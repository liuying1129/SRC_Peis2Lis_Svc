unit UPeis2Lis;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, SvcMgr, Dialogs,
  Registry,Inifiles, ExtCtrls, DB, ADODB,ActiveX, LYTray, Menus;

type
  TPeis2Lis = class(TService)
    Timer1: TTimer;
    ADOConn_Lis: TADOConnection;
    ADOConn_His: TADOConnection;
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    procedure ServiceAfterInstall(Sender: TService);
    procedure ServiceCreate(Sender: TObject);
    procedure ServiceStart(Sender: TService; var Started: Boolean);
    procedure ServiceStop(Sender: TService; var Stopped: Boolean);
    procedure Timer1Timer(Sender: TObject);
    procedure N1Click(Sender: TObject);
  private
    { Private declarations }
    procedure ReadIni;
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  Peis2Lis: TPeis2Lis;

implementation

var
  AppPath:string;
  Connstr_Lis:string;
  Connstr_His:string;
  UpdateFreq:integer;
  
{$R *.DFM}
function ShowOptionForm(const pCaption,pTabSheetCaption,pItemInfo,pInifile:Pchar):boolean;stdcall;external 'OptionSetForm.dll';

//编写NT服务程序。New->Service Application

//安装服务，在命令行中执行:Insert_Fph2Rdmx.exe /install
//卸载服务，在命令行中执行:Insert_Fph2Rdmx.exe /uninstall

procedure SetDescription(const AClassName:string;const ADescription:string);
//增加服务描述
var
  vReg:TRegistry;
begin
  vReg:=TRegistry.Create;
  vReg.RootKey:=HKEY_LOCAL_MACHINE;
  vReg.OpenKey('\SYSTEM\CurrentControlSet\Services\'+AClassName,True);
  vReg.WriteString('Description',ADescription);
  vReg.CloseKey;
  vReg.Free;
end;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  Peis2Lis.Controller(CtrlCode);
end;

function TPeis2Lis.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TPeis2Lis.ServiceAfterInstall(Sender: TService);
begin
  SetDescription(Name,'Peis医嘱转换为Lis申请单服务');
end;

procedure TPeis2Lis.ServiceCreate(Sender: TObject);
var
  buf: array[0..MAX_PATH] of Char;
  hinst: HMODULE;
begin
  hinst:=GetModuleHandle('SvcPeis2Lis.exe');
  GetModuleFileName(hinst,buf,MAX_PATH);
  AppPath:=strpas(buf);
  
  LYTray1.Hint:=Name+'设置';
end;

procedure TPeis2Lis.ReadIni;
var
  configini:tinifile;
begin
  CONFIGINI:=TINIFILE.Create(ChangeFileExt(AppPath,'.ini'));

  Connstr_Lis:=configini.ReadString('设置','Lis连接字符串','');
  Connstr_His:=configini.ReadString('设置','Peis连接字符串','');
  UpdateFreq:=configini.ReadInteger('设置','扫描中间表频率',1);
  if UpdateFreq<=0 then UpdateFreq:=1;
  
  configini.Free;
end;

procedure TPeis2Lis.ServiceStart(Sender: TService; var Started: Boolean);
begin
  ReadIni;

  CoInitialize(nil);//必须，否则 ADOConnection1.Connected := true; 总是报错

  ADOConn_His.Connected := false;
  ADOConn_His.ConnectionString := ConnStr_His;
  try
    ADOConn_His.Connected := true;
    LogMessage('连接数据库成功:'+ConnStr_His,EVENTLOG_INFORMATION_TYPE);
  except
    LogMessage('连接数据库失败:'+ConnStr_His,EVENTLOG_ERROR_TYPE);
    //exit;//OraQuery1.ExecSQL时会自动打开ADOConnection1，故不用exit.
  end;

  ADOConn_Lis.Connected := false;
  ADOConn_Lis.ConnectionString := ConnStr_Lis;
  try
    ADOConn_Lis.Connected := true;
    LogMessage('连接数据库成功:'+ConnStr_Lis,EVENTLOG_INFORMATION_TYPE);
  except
    LogMessage('连接数据库失败:'+ConnStr_Lis,EVENTLOG_ERROR_TYPE);
    //exit;//OraQuery1.ExecSQL时会自动打开ADOConnection1，故不用exit.
  end;
  
  Timer1.Interval:=UpdateFreq*60*1000;
  Timer1.Enabled:=true;
end;

procedure TPeis2Lis.ServiceStop(Sender: TService; var Stopped: Boolean);
begin
  Timer1.Enabled:=false;
  CoUninitialize;
end;

procedure TPeis2Lis.Timer1Timer(Sender: TObject);
var
  adotemp11,adotemp22,adotemp33,adotemp44,adotemp55,adotemp66,adotemp77,adotemp88,adotemp99,adotemp111,adotemp222,adotemp333,adotemp444:tadoquery;

  Insert_Identity:string;
  scombin_id:string;//申请单工作组
  sID:STRING;
  RecNum:integer;

  ID_Patient ,ID_PatientArchive,Age,F_Registered,F_FeeCharged,F_Paused ,	F_Transfered_IP ,	F_UseCodeHiden:integer;
  BirthDate,DateRegister:TDatetime;
  AgeOfReal:real;
  StrIDPatient ,	PatientCode ,	PatientCodeHiden ,	PatientCardNo ,	IDCardNo ,	PatientName :string;
  PatientArchiveNo ,	PatientRequestNo ,	Input_Code ,	Org_Name ,	Org_Depart ,	Sex 	 :string;
  AgeUnit ,	Marriage ,	DoctorReg ,	ExamType_Name :string;
  ParsedSuiteAndFI ,ParsedSuiteAndFILab :string;
  ID_PatientFeeItem ,ID_Depart, {ID_ExamFeeItem,}F_Back_Transfered ,	F_ResultTransfered:integer;
  F_Returned:boolean;
  FeeItemRequestNo ,	Depart_Name ,	TransfterTarget ,ExamFeeItem_Name :string;
  ExamFeeItem_Code ,	LabType_Name ,	LabType_Code  :string;
  LabSampleTime:TDatetime;//采样时间
begin
  Timer1.Enabled:=false;

  //怪了，数据库停掉后重启,tadoquery不能自动连接
  //if not ADOConnection1.Connected then//服务停后，该判断都为真
  ADOConn_His.Connected := false;
  try
    ADOConn_His.Open;
    //LogMessage('重新连接数据库成功'+ADOConnection1.ConnectionString,EVENTLOG_INFORMATION_TYPE);//每次都要连接，故不用写了
  except
    on E:Exception do
    begin
      LogMessage(E.Message+'.重新连接数据库失败:'+ADOConn_His.ConnectionString,EVENTLOG_ERROR_TYPE);
      Timer1.Enabled:=true;
      exit;
    end;
  end;

  ADOConn_Lis.Connected := false;
  try
    ADOConn_Lis.Open;
    //LogMessage('重新连接数据库成功'+ADOConnection1.ConnectionString,EVENTLOG_INFORMATION_TYPE);//每次都要连接，故不用写了
  except
    on E:Exception do
    begin
      LogMessage(E.Message+'.重新连接数据库失败:'+ADOConn_Lis.ConnectionString,EVENTLOG_ERROR_TYPE);
      Timer1.Enabled:=true;
      exit;
    end;
  end;
  //===========================}

  adotemp11:=tadoquery.Create(nil);
  adotemp11.Connection:=ADOConn_His;
  adotemp11.Close;
  adotemp11.SQL.Clear;
  adotemp11.SQL.Text:='select IP.ID_Patient ,	IP.StrIDPatient ,	IP.PatientCodeHiden ,	IP.PatientCardNo ,	IP.IDCardNo ,	IP.PatientName ,	'+
                      'IP.ID_PatientArchive ,	IP.PatientArchiveNo ,	IP.PatientRequestNo ,	IP.Input_Code ,	IP.Org_Name ,	IP.Org_Depart ,	IP.Sex ,	IP.BirthDate ,	'+
                      'IP.Age ,	IP.AgeUnit ,	IP.AgeOfReal ,	IP.Marriage ,	IP.F_Registered ,	IP.DateRegister ,	IP.DoctorReg ,	IP.ExamType_Name ,'+//	IP.F_FeeCharged ,	
                      'IP.F_UseCodeHiden ,	IP.ParsedSuiteAndFI ,IP.ParsedSuiteAndFILab , '+//IP.F_Paused ,	IP.F_Transfered AS F_Transfered_IP,	
                      'IPFI.ID_PatientFeeItem ,IPFI.PatientCode ,		IPFI.FeeItemRequestNo ,	IPFI.ID_Depart ,	IPFI.Depart_Name ,	IPFI.TransfterTarget ,	IPFI.ID_ExamFeeItem ,	IPFI.ExamFeeItem_Name ,	'+
                      'IPFI.ExamFeeItem_Code ,	IPFI.LabType_Name ,	IPFI.LabType_Code ,	IPFI.F_Back_Transfered,	IPFI.F_Returned,IPFI.LabSampleTime '+//,	IPFI.F_ResultTransfered   
                      ' from IntPatient IP '+
                      ' inner join IntPatientFeeItem IPFI '+
                      ' on IP.ID_Patient=IPFI.ID_Patient '+
                      ' AND IPFI.F_LabSampled=1 '+
                      ' AND IPFI.TransfterTarget=''LIS'' '+
                      ' and IPFI.LabSampleTime+0.8>getdate() ';
  try
    adotemp11.Open;
  except
    on E:Exception do
    begin
      LogMessage('查询PEIS的申请单时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
      adotemp11.Free;
      Timer1.Enabled:=true;
      exit;
    end;
  end;
  if adotemp11.RecordCount<=0 then
  begin
    adotemp11.Free;
    Timer1.Enabled:=true;
    exit;
  end;

  while not adotemp11.Eof do
  begin
    ID_Patient:=adotemp11.fieldbyname('ID_Patient').AsInteger;
    StrIDPatient:=adotemp11.fieldbyname('StrIDPatient').AsString;
    PatientCode:=adotemp11.fieldbyname('PatientCode').AsString;
    PatientCodeHiden:=adotemp11.fieldbyname('PatientCodeHiden').AsString;
    PatientCardNo:=adotemp11.fieldbyname('PatientCardNo').AsString;
    IDCardNo:=adotemp11.fieldbyname('IDCardNo').AsString;
    PatientName:=adotemp11.fieldbyname('PatientName').AsString; 
    //ID_PatientArchive:=adotemp11.fieldbyname('ID_PatientArchive').AsInteger;
    PatientArchiveNo:=adotemp11.fieldbyname('PatientArchiveNo').AsString;
    PatientRequestNo:=adotemp11.fieldbyname('PatientRequestNo').AsString;
    Input_Code:=adotemp11.fieldbyname('Input_Code').AsString;
    Org_Name:=adotemp11.fieldbyname('Org_Name').AsString;
    Org_Depart:=adotemp11.fieldbyname('Org_Depart').AsString;
    Sex:=adotemp11.fieldbyname('Sex').AsString;
    //BirthDate:=adotemp11.fieldbyname('BirthDate').AsDateTime;
    Age:=adotemp11.fieldbyname('Age').AsInteger;
    AgeUnit:=adotemp11.fieldbyname('AgeUnit').AsString;
    //AgeOfReal:=adotemp11.fieldbyname('AgeOfReal').AsFloat;
    Marriage:=adotemp11.fieldbyname('Marriage').AsString;
    //F_Registered:=adotemp11.fieldbyname('F_Registered').AsInteger;
    DateRegister:=adotemp11.fieldbyname('DateRegister').AsDateTime;
    DoctorReg:=adotemp11.fieldbyname('DoctorReg').AsString;
    ExamType_Name:=adotemp11.fieldbyname('ExamType_Name').AsString;
    //F_FeeCharged:=adotemp11.fieldbyname('F_FeeCharged').AsInteger;
    //F_Paused:=adotemp11.fieldbyname('F_Paused').AsInteger;
    //F_Transfered_IP:=adotemp11.fieldbyname('F_Transfered_IP').AsInteger;
    //F_UseCodeHiden:=adotemp11.fieldbyname('F_UseCodeHiden').AsInteger;
    ParsedSuiteAndFI:=adotemp11.fieldbyname('ParsedSuiteAndFI').AsString;
    ParsedSuiteAndFILab:=adotemp11.fieldbyname('ParsedSuiteAndFILab').AsString;
    ID_PatientFeeItem:=adotemp11.fieldbyname('ID_PatientFeeItem').AsInteger;
    FeeItemRequestNo:=adotemp11.fieldbyname('FeeItemRequestNo').AsString;
    //ID_Depart:=adotemp11.fieldbyname('ID_Depart').AsInteger;
    Depart_Name:=adotemp11.fieldbyname('Depart_Name').AsString;
    TransfterTarget:=adotemp11.fieldbyname('TransfterTarget').AsString;
    //ID_ExamFeeItem:=adotemp11.fieldbyname('ID_ExamFeeItem').AsInteger;
    ExamFeeItem_Name:=adotemp11.fieldbyname('ExamFeeItem_Name').AsString;
    ExamFeeItem_Code:=adotemp11.fieldbyname('ExamFeeItem_Code').AsString;
    LabType_Name:=adotemp11.fieldbyname('LabType_Name').AsString;
    LabType_Code:=adotemp11.fieldbyname('LabType_Code').AsString;
    //F_Back_Transfered:=adotemp11.fieldbyname('F_Back_Transfered').AsInteger;
    //F_Returned:=adotemp11.fieldbyname('F_Returned').AsBoolean;
    //F_ResultTransfered:=adotemp11.fieldbyname('F_ResultTransfered').AsInteger; 
    LabSampleTime:=adotemp11.fieldbyname('LabSampleTime').AsDateTime;

    {if F_Paused=1 then//PEIS禁用申请单//马工说，F_Paused不用管，基本上不会出现这种情况
    begin
      adotemp77:=tadoquery.Create(nil);
      adotemp77.Connection:=ADOConn_Lis;
      adotemp77.Close;
      adotemp77.SQL.Clear;
      adotemp77.SQL.Text:=' delete from chk_valu_his where Surem1='''+PatientCode+''' ';
      try
        adotemp77.ExecSQL;
      except
        on E:Exception do
        begin
          LogMessage('禁用PEIS申请单ID'+PatientCode+'失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
        end;
      end;
      adotemp77.Free;
      adotemp11.next;
      continue;
    end;//}
    
    //判断该申请单ID的病人申请单是否存在start
    adotemp55:=tadoquery.Create(nil);
    adotemp55.Connection:=ADOConn_Lis;
    adotemp55.Close;
    adotemp55.SQL.Clear;
    adotemp55.SQL.Text:=' select cch.unid from chk_con_his cch where cch.His_Unid='''+inttostr(ID_Patient)+''' ';
    try
      adotemp55.Open;
    except
      on E:Exception do
      begin
        LogMessage('在LIS主表中查询PEIS申请单ID'+inttostr(ID_Patient)+'时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
        adotemp55.Free;
        adotemp11.next;
        continue;
      end;
    end;
    Insert_Identity:=adotemp55.fieldbyname('unid').AsString;
    adotemp55.Free;

    if Insert_Identity='' then
    begin
      adotemp333:=tadoquery.Create(nil);
      adotemp333.Connection:=ADOConn_Lis;
      adotemp333.Close;
      adotemp333.SQL.Clear;
      adotemp333.SQL.Text:=' select cvh.pkunid from chk_valu_his cvh where cvh.Surem1='''+inttostr(ID_Patient)+''' ';
      try
        adotemp333.Open;
      except
        on E:Exception do
        begin
          LogMessage('在LIS从表中查询PEIS申请单ID'+inttostr(ID_Patient)+'时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp333.Free;
          adotemp11.next;
          continue;
        end;
      end;
      Insert_Identity:=adotemp333.fieldbyname('pkunid').AsString;
      adotemp333.Free;
    end;
    //判断该申请单ID的病人申请单是否存在stop

    if Insert_Identity='' then
    begin
      //根据申请科室查询申请单工作组start
      adotemp99:=tadoquery.Create(nil);
      adotemp99.Connection:=ADOConn_Lis;
      adotemp99.Close;
      adotemp99.SQL.Clear;
      adotemp99.SQL.Text:='SELECT * FROM CommCode cc where TypeName=''体检团体'' and ID='''+Org_Name+''' ';
      try
        adotemp99.Open;
      except
        on E:Exception do
        begin
          LogMessage('根据体检团体'+Org_Name+'查询申请单工作组失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp99.Free;
          adotemp11.next;
          continue;
        end;
      end;
      scombin_id:=trim(adotemp99.fieldbyname('ReMark').AsString);
      adotemp99.Free;
      //根据申请科室查询申请单工作组stop

      adotemp33:=tadoquery.Create(nil);
      adotemp33.Connection:=ADOConn_Lis;
      adotemp33.Close;
      adotemp33.SQL.Clear;
      //不插入PeIS的样本类型，由拆分存储过程去处理（拿组合项目的默认样本类型）
      adotemp33.SQL.Add('insert into chk_con_his (patientname,sex,age,report_date,bedno,His_Unid,check_doctor,deptname,combin_id,Caseno,diagnose,Diagnosetype,typeflagcase,WorkCompany,WorkDepartment,ifMarry) values '+
                          ' ('''+PatientName+''','''+Sex+''','''+inttostr(Age)+AgeUnit+''',:DateRegister,'''+''+''','''+inttostr(ID_Patient)+''','''+DoctorReg+''','''+'体检科'+''','''+scombin_id+''','''+PatientCode+''','''+''+''',''常规'',''正常'','''+Org_Name+''','''+Org_Depart+''','''+Marriage+''') ');
      adotemp33.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
      adotemp33.Parameters.ParamByName('DateRegister').Value:=DateRegister;
      try
        adotemp33.Open;
      except
        on E:Exception do
        begin
          LogMessage('向LIS中插入PEIS基本信息'+inttostr(ID_Patient)+'时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp33.Free;
          adotemp11.next;
          continue;
        end;
      end;
      Insert_Identity:=adotemp33.fieldbyname('Insert_Identity').AsString;
      adotemp33.Free;
    end;

    //获取申请单元ID的LIS对照start
    adotemp66:=tadoquery.Create(nil);
    adotemp66.Connection:=ADOConn_Lis;
    adotemp66.Close;
    adotemp66.SQL.Clear;
    adotemp66.SQL.Text:='select c.id from HisCombItem hci,combinitem c '+
                        ' where hci.CombUnid=c.unid and hci.ExtSystemId=''PEIS'' and hci.HisItem='''+ExamFeeItem_Code+''' ';
    try
      adotemp66.Open;
    except
      on E:Exception do
      begin
        LogMessage('在LIS中取PEIS申请单元ID'+ExamFeeItem_Code+'的对照关系时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
        adotemp66.Free;
        adotemp11.next;
        continue;
      end;
    end;
    if adotemp66.RecordCount<=0 then
    begin
      LogMessage('PEIS申请单元ID'+ExamFeeItem_Code+'在LIS中无对照关系',EVENTLOG_ERROR_TYPE	);
      adotemp66.Free;
      adotemp11.next;
      continue;
    end;

    while not adotemp66.Eof do
    begin
      sID:=adotemp66.fieldbyname('id').AsString;

      {if F_Returned then//PeIS退项.PEIS只有在采集样本前才能被退项，故不需要了
      begin
        adotemp88:=tadoquery.Create(nil);
        adotemp88.Connection:=ADOConn_Lis;
        adotemp88.Close;
        adotemp88.SQL.Clear;
        adotemp88.SQL.Text:=' delete from chk_valu_his where pkcombin_id='''+sID+''' and Surem1='''+inttostr(ID_Patient)+''' ';
        try
          adotemp88.ExecSQL;
        except
          on E:Exception do
          begin
            LogMessage('取消PEIS申请单'+inttostr(ID_Patient)+',申请单元ID'+ExamFeeItem_Code+'('+sID+')失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
          end;
        end;
        adotemp88.Free;
        adotemp66.next;
        continue;
      end;//}

      adotemp44:=tadoquery.Create(nil);
      adotemp44.Connection:=ADOConn_Lis;
      adotemp44.Close;
      adotemp44.SQL.Clear;
      adotemp44.SQL.Text:='select count(*) as RecNum from '+
                          ' chk_valu_his cvh where cvh.Surem1='''+inttostr(ID_Patient)+''' '+
                          ' and cvh.pkcombin_id='''+sID+''' ';
      try
        adotemp44.Open;
      except
        on E:Exception do
        begin
          LogMessage('判断PEIS申请单'+inttostr(ID_Patient)+',申请单元ID'+ExamFeeItem_Code+'('+sID+')在LIS是否存在时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp44.Free;
          adotemp66.next;
          continue;
        end;
      end;
      RecNum:=adotemp44.FieldByName('RecNum').AsInteger;
      adotemp44.Free;

      if RecNum<=0 then
      begin
        adotemp22:=tadoquery.Create(nil);
        adotemp22.Connection:=ADOConn_Lis;
        adotemp22.Close;
        adotemp22.SQL.Clear;
        adotemp22.SQL.Text:='insert into chk_valu_his (pkunid,pkcombin_id,Surem1,Surem2,Urine1,Urine2,TakeSampleTime) values ('''+Insert_Identity+''','''+sID+''','''+inttostr(ID_Patient)+''','''+ExamFeeItem_Code+''','''+PatientCode+''','''+LabType_Code+''',:TakeSampleTime) ';
        adotemp22.Parameters.ParamByName('TakeSampleTime').Value:=LabSampleTime;
        try
          adotemp22.ExecSQL;
        except
          on E:Exception do
          begin
            LogMessage('向LIS中插入PEIS申请单元ID'+ExamFeeItem_Code+'('+sID+')时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
            adotemp22.Free;
            adotemp66.next;
            continue;
          end;
        end;
        adotemp22.Free;

        //回写“已读取”标志。该标志对PEIS无任何控制作用，只是让PEIS知道该条记录已被读走
        adotemp444:=tadoquery.Create(nil);
        adotemp444.Connection:=ADOConn_His;
        adotemp444.Close;
        adotemp444.SQL.Clear;
        adotemp444.SQL.Text:='UPDATE IntPatientFeeItem SET F_Back_Transfered=1 WHERE ID_PatientFeeItem='+inttostr(ID_PatientFeeItem);
        try
          adotemp444.ExecSQL;
        except
          on E:Exception do
          begin
            LogMessage('Peis收费项目主键'+inttostr(ID_PatientFeeItem)+',修改其读取标志F_Back_Transfered失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
            adotemp444.Free;
            adotemp66.next;
            continue;
          end;
        end;
        adotemp444.Free;
      end;

      adotemp66.Next;
    end;
    adotemp66.Free;
    //获取申请单元ID的LIS对照stop

    adotemp11.Next;
  end;
  adotemp11.Free;

  //拆分申请单
  adotemp222:=tadoquery.Create(nil);
  adotemp222.Connection:=ADOConn_Lis;
  adotemp222.Close;
  adotemp222.SQL.Clear;
  adotemp222.SQL.Text:='dbo.pro_SplitRequestBill';
  try
    adotemp222.ExecSQL;
  except
    on E:Exception do
    begin
      LogMessage('拆分PEIS申请单元时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
    end;
  end;
  adotemp222.Free;

  //合并申请单
  adotemp111:=tadoquery.Create(nil);
  adotemp111.Connection:=ADOConn_Lis;
  adotemp111.Close;
  adotemp111.SQL.Clear;
  adotemp111.SQL.Text:='dbo.pro_MergeRequestBill';
  try
    adotemp111.ExecSQL;
  except
    on E:Exception do
    begin
      LogMessage('合并PEIS申请单元时失败:'+E.Message,EVENTLOG_ERROR_TYPE	);
    end;
  end;
  adotemp111.Free;  

  Timer1.Enabled:=true;
end;

procedure TPeis2Lis.N1Click(Sender: TObject);
var
  ss:string;
begin
    ss:='Lis连接字符串'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
        'Peis连接字符串'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
        '扫描中间表频率'+#2+'Edit'+#2+#2+'1'+#2+'单位:分钟'+#2+#3;

  if ShowOptionForm('','设置',Pchar(ss),Pchar(ChangeFileExt(AppPath,'.ini'))) then
	  ReadIni;
end;

end.
