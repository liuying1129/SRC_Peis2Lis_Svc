program SvcPeis2Lis;

uses
  SvcMgr,
  UPeis2Lis in 'UPeis2Lis.pas' {Peis2Lis: TService};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TPeis2Lis, Peis2Lis);
  Application.Run;
end.
