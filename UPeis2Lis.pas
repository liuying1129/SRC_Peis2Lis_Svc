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

//��дNT�������New->Service Application

//��װ��������������ִ��:Insert_Fph2Rdmx.exe /install
//ж�ط�������������ִ��:Insert_Fph2Rdmx.exe /uninstall

procedure SetDescription(const AClassName:string;const ADescription:string);
//���ӷ�������
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
  SetDescription(Name,'Peisҽ��ת��ΪLis���뵥����');
end;

procedure TPeis2Lis.ServiceCreate(Sender: TObject);
var
  buf: array[0..MAX_PATH] of Char;
  hinst: HMODULE;
begin
  hinst:=GetModuleHandle('SvcPeis2Lis.exe');
  GetModuleFileName(hinst,buf,MAX_PATH);
  AppPath:=strpas(buf);
  
  LYTray1.Hint:=Name+'����';
end;

procedure TPeis2Lis.ReadIni;
var
  configini:tinifile;
begin
  CONFIGINI:=TINIFILE.Create(ChangeFileExt(AppPath,'.ini'));

  Connstr_Lis:=configini.ReadString('����','Lis�����ַ���','');
  Connstr_His:=configini.ReadString('����','Peis�����ַ���','');
  UpdateFreq:=configini.ReadInteger('����','ɨ���м��Ƶ��',1);
  if UpdateFreq<=0 then UpdateFreq:=1;
  
  configini.Free;
end;

procedure TPeis2Lis.ServiceStart(Sender: TService; var Started: Boolean);
begin
  ReadIni;

  CoInitialize(nil);//���룬���� ADOConnection1.Connected := true; ���Ǳ���

  ADOConn_His.Connected := false;
  ADOConn_His.ConnectionString := ConnStr_His;
  try
    ADOConn_His.Connected := true;
    LogMessage('�������ݿ�ɹ�:'+ConnStr_His,EVENTLOG_INFORMATION_TYPE);
  except
    LogMessage('�������ݿ�ʧ��:'+ConnStr_His,EVENTLOG_ERROR_TYPE);
    //exit;//OraQuery1.ExecSQLʱ���Զ���ADOConnection1���ʲ���exit.
  end;

  ADOConn_Lis.Connected := false;
  ADOConn_Lis.ConnectionString := ConnStr_Lis;
  try
    ADOConn_Lis.Connected := true;
    LogMessage('�������ݿ�ɹ�:'+ConnStr_Lis,EVENTLOG_INFORMATION_TYPE);
  except
    LogMessage('�������ݿ�ʧ��:'+ConnStr_Lis,EVENTLOG_ERROR_TYPE);
    //exit;//OraQuery1.ExecSQLʱ���Զ���ADOConnection1���ʲ���exit.
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
  scombin_id:string;//���뵥������
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
begin
  Timer1.Enabled:=false;

  //���ˣ����ݿ�ͣ��������,tadoquery�����Զ�����
  //if not ADOConnection1.Connected then//����ͣ�󣬸��ж϶�Ϊ��
  ADOConn_His.Connected := false;
  try
    ADOConn_His.Open;
    //LogMessage('�����������ݿ�ɹ�'+ADOConnection1.ConnectionString,EVENTLOG_INFORMATION_TYPE);//ÿ�ζ�Ҫ���ӣ��ʲ���д��
  except
    on E:Exception do
    begin
      LogMessage(E.Message+'.�����������ݿ�ʧ��:'+ADOConn_His.ConnectionString,EVENTLOG_ERROR_TYPE);
      Timer1.Enabled:=true;
      exit;
    end;
  end;

  ADOConn_Lis.Connected := false;
  try
    ADOConn_Lis.Open;
    //LogMessage('�����������ݿ�ɹ�'+ADOConnection1.ConnectionString,EVENTLOG_INFORMATION_TYPE);//ÿ�ζ�Ҫ���ӣ��ʲ���д��
  except
    on E:Exception do
    begin
      LogMessage(E.Message+'.�����������ݿ�ʧ��:'+ADOConn_Lis.ConnectionString,EVENTLOG_ERROR_TYPE);
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
                      'IP.Age ,	IP.AgeUnit ,	IP.AgeOfReal ,	IP.Marriage ,	IP.F_Registered ,	IP.DateRegister ,	IP.DoctorReg ,	IP.ExamType_Name ,	IP.F_FeeCharged ,	'+
                      'IP.F_Paused ,	IP.F_Transfered AS F_Transfered_IP,	IP.F_UseCodeHiden ,	IP.ParsedSuiteAndFI ,IP.ParsedSuiteAndFILab , '+
                      'IPFI.ID_PatientFeeItem ,IPFI.PatientCode ,		IPFI.FeeItemRequestNo ,	IPFI.ID_Depart ,	IPFI.Depart_Name ,	IPFI.TransfterTarget ,	IPFI.ID_ExamFeeItem ,	IPFI.ExamFeeItem_Name ,	'+
                      'IPFI.ExamFeeItem_Code ,	IPFI.LabType_Name ,	IPFI.LabType_Code ,	IPFI.F_Back_Transfered,	IPFI.F_Returned ,	IPFI.F_ResultTransfered   '+
                      ' from IntPatient IP '+
                      ' inner join IntPatientFeeItem IPFI '+
                      ' on IP.ID_Patient=IPFI.ID_Patient '+
                      ' AND IP.F_FeeCharged=1 '+
                      ' AND IPFI.TransfterTarget=''LIS'' '+
                      ' and IP.DateRegister+0.8>getdate() ';
  try
    adotemp11.Open;
  except
    on E:Exception do
    begin
      LogMessage('��ѯPEIS�����뵥ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
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
    F_Returned:=adotemp11.fieldbyname('F_Returned').AsBoolean;
    //F_ResultTransfered:=adotemp11.fieldbyname('F_ResultTransfered').AsInteger; 

    {if F_Paused=1 then//PEIS�������뵥//��˵��F_Paused���ùܣ������ϲ�������������
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
          LogMessage('����PEIS���뵥ID'+PatientCode+'ʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
        end;
      end;
      adotemp77.Free;
      adotemp11.next;
      continue;
    end;//}
    
    //�жϸ����뵥ID�Ĳ������뵥�Ƿ����start
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
        LogMessage('��LIS�����в�ѯPEIS���뵥ID'+inttostr(ID_Patient)+'ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
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
          LogMessage('��LIS�ӱ��в�ѯPEIS���뵥ID'+inttostr(ID_Patient)+'ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp333.Free;
          adotemp11.next;
          continue;
        end;
      end;
      Insert_Identity:=adotemp333.fieldbyname('pkunid').AsString;
      adotemp333.Free;
    end;
    //�жϸ����뵥ID�Ĳ������뵥�Ƿ����stop

    if Insert_Identity='' then
    begin
      //����������Ҳ�ѯ���뵥������start
      adotemp99:=tadoquery.Create(nil);
      adotemp99.Connection:=ADOConn_Lis;
      adotemp99.Close;
      adotemp99.SQL.Clear;
      adotemp99.SQL.Text:='SELECT * FROM CommCode cc where TypeName=''�������'' and ID='''+Org_Name+''' ';
      try
        adotemp99.Open;
      except
        on E:Exception do
        begin
          LogMessage('�����������'+Org_Name+'��ѯ���뵥������ʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp99.Free;
          adotemp11.next;
          continue;
        end;
      end;
      scombin_id:=trim(adotemp99.fieldbyname('ReMark').AsString);
      adotemp99.Free;
      //����������Ҳ�ѯ���뵥������stop

      adotemp33:=tadoquery.Create(nil);
      adotemp33.Connection:=ADOConn_Lis;
      adotemp33.Close;
      adotemp33.SQL.Clear;
      //������PeIS���������ͣ��ɲ�ִ洢����ȥ�����������Ŀ��Ĭ���������ͣ�
      adotemp33.SQL.Add('insert into chk_con_his (patientname,sex,age,report_date,bedno,His_Unid,check_doctor,deptname,combin_id,Caseno,diagnose,Diagnosetype,typeflagcase,WorkCompany,WorkDepartment,ifMarry) values '+
                          ' ('''+PatientName+''','''+Sex+''','''+inttostr(Age)+AgeUnit+''',:DateRegister,'''+''+''','''+inttostr(ID_Patient)+''','''+DoctorReg+''','''+'����'+''','''+scombin_id+''','''+PatientCode+''','''+''+''',''����'',''����'','''+Org_Name+''','''+Org_Depart+''','''+Marriage+''') ');
      adotemp33.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
      adotemp33.Parameters.ParamByName('DateRegister').Value:=DateRegister;
      try
        adotemp33.Open;
      except
        on E:Exception do
        begin
          LogMessage('��LIS�в���PEIS������Ϣ'+inttostr(ID_Patient)+'ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
          adotemp33.Free;
          adotemp11.next;
          continue;
        end;
      end;
      Insert_Identity:=adotemp33.fieldbyname('Insert_Identity').AsString;
      adotemp33.Free;
    end;

    //��ȡ���뵥ԪID��LIS����start
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
        LogMessage('��LIS��ȡPEIS���뵥ԪID'+ExamFeeItem_Code+'�Ķ��չ�ϵʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
        adotemp66.Free;
        adotemp11.next;
        continue;
      end;
    end;
    if adotemp66.RecordCount<=0 then
    begin
      LogMessage('PEIS���뵥ԪID'+ExamFeeItem_Code+'��LIS���޶��չ�ϵ',EVENTLOG_ERROR_TYPE	);
      adotemp66.Free;
      adotemp11.next;
      continue;
    end;

    while not adotemp66.Eof do
    begin
      sID:=adotemp66.fieldbyname('id').AsString;

      if F_Returned then//PeIS����
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
            LogMessage('ȡ��PEIS���뵥'+inttostr(ID_Patient)+',���뵥ԪID'+ExamFeeItem_Code+'('+sID+')ʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
          end;
        end;
        adotemp88.Free;
        adotemp66.next;
        continue;
      end;

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
          LogMessage('�ж�PEIS���뵥'+inttostr(ID_Patient)+',���뵥ԪID'+ExamFeeItem_Code+'('+sID+')��LIS�Ƿ����ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
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
        adotemp22.SQL.Text:='insert into chk_valu_his (pkunid,pkcombin_id,Surem1,Surem2,Urine1,Urine2) values ('''+Insert_Identity+''','''+sID+''','''+inttostr(ID_Patient)+''','''+ExamFeeItem_Code+''','''+PatientCode+''','''+LabType_Code+''') ';
        try
          adotemp22.ExecSQL;
        except
          on E:Exception do
          begin
            LogMessage('��LIS�в���PEIS���뵥ԪID'+ExamFeeItem_Code+'('+sID+')ʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
            adotemp22.Free;
            adotemp66.next;
            continue;
          end;
        end;
        adotemp22.Free;

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
            LogMessage('Peis�շ���Ŀ����'+inttostr(ID_PatientFeeItem)+',�޸����ȡ��־F_Back_Transferedʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
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
    //��ȡ���뵥ԪID��LIS����stop

    adotemp11.Next;
  end;
  adotemp11.Free;

  //������뵥
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
      LogMessage('���PEIS���뵥Ԫʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
    end;
  end;
  adotemp222.Free;

  //�ϲ����뵥
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
      LogMessage('�ϲ�PEIS���뵥Ԫʱʧ��:'+E.Message,EVENTLOG_ERROR_TYPE	);
    end;
  end;
  adotemp111.Free;  

  Timer1.Enabled:=true;
end;

procedure TPeis2Lis.N1Click(Sender: TObject);
var
  ss:string;
begin
    ss:='Lis�����ַ���'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
        'Peis�����ַ���'+#2+'DBConn'+#2+#2+'1'+#2+#2+#3+
        'ɨ���м��Ƶ��'+#2+'Edit'+#2+#2+'1'+#2+'��λ:����'+#2+#3;

  if ShowOptionForm('','����',Pchar(ss),Pchar(ChangeFileExt(AppPath,'.ini'))) then
	  ReadIni;
end;

end.
