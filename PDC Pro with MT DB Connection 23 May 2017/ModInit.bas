Attribute VB_Name = "ModInit"
Public ConnectStringOnline As String
Public ConnectStringMirror As String
Public ConnectStringOnlineOLD As String

Public PDCUserName As String
Public PDCUserType As Integer
Public LocationCode As String
Public LocationName As String
Public MICRLocationCode As String
Public WaitFlag As Boolean

Public Sub Init()
'ConnectStringOnline = "Provider=SQLOLEDB.1;Password=MtEdum@te2013;Persist Security Info=True;User ID=pipl2013;Initial Catalog=MTeducare;Data Source=49.248.12.108,18553" '49.248.12.108
'ConnectStringOnline = "Provider=SQLOLEDB.1;Password=panaceea;Persist Security Info=True;User ID=sa;Initial Catalog=PDC_MT_DB;Data Source=vivek" '49.248.12.108
ConnectStringOnline = "Provider=SQLOLEDB.1;Password=Shvet@nk#2017#;Persist Security Info=True;User ID=ShvetankInfotech;Initial Catalog=PDC_Warehouse;Data Source=49.248.16.100" '49.248.12.108

'ConnectStringOnline = "Provider=SQLNCLI10.1;Persist Security Info=True;User ID=careerclasses_org_pdcsystem;Password=PdcM123@shvet;Initial Catalog=careerclasses_org_pdcsystem;Data Source=216.51.232.100"

'ConnectStringOnline = "Provider=SQLOLEDB.1;Password=PdcM123@shvet;Persist Security Info=True;User ID=careerclasses_org_pdcsystem;Initial Catalog=careerclasses_org_pdcsystem;Data Source=216.51.232.100" '49.248.12.108

Dim SourcePath As String
SourcePath = App.Path & "\ASPDC_Mirror.scd"
ConnectStringMirror = "Provider=MSDASQL.1;Persist Security Info=False;User ID=admin;Extended Properties=""DBQ=" & SourcePath & ";DefaultDir= " & App.Path & ";Driver={Microsoft Access Driver (*.mdb)};DriverId=281;FIL=MS Access;FILEDSN=" & SourcePath & ";MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UID=admin;UserCommitSync=Yes;"";Initial Catalog=" & SourcePath & """"


End Sub
