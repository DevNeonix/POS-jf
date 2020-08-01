<%
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.CommandTimeout = 900

'***********************************************************************************************
'   ESTE ES EL QUE ESTOY USANDO EN LA CASA y en EL MODELADOR
'***********************************************************************************************
dim txtConexion

'     MODELADOR
'    txtConexion= "Provider=SQLOLEDB.1;Data Source=192.168.3.4;Initial Catalog=JACINTA; User id=SA; Password=F3QY17;"
     

'    
'   txtConexion = 	"Provider=SQLOLEDB.1;Data Source=192.168.3.30,1433;Initial Catalog=JACINTA;User ID=SA;Password=c4r4m3l1t0;"
'    "Driver={SQL Server};Server=192.168.3.30;Database=JACINTA;User ID=sa;Pwd=c4r4m3lt0"
'    txtConexion = 	"Provider=SQLOLEDB;Data Source=PC-MMOLINA ;Initial Catalog=JACINTA;User ID=SA;Password=c4r4m3l1t0;"
'    txtConexion =  "DSN=desarrollo;UID=SA;PWD=c4r4m3l1t0"
'    txtConexion =  "Provider=SQLOLEDB.1;Data Source=PC-MMOLINA;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;Persist Security Info=True;"
'    txtConexion =   "Provider=SQLOLEDB.1;Data Source=192.168.3.30\MSSQLSERVER;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
'    txtConexion =  "DRIVER={SQL Server};SERVER=PC-MMOLINA;Database=JACINTA;Uid=sa;PWD=c4r4m3lt0"
'    txtConexion =  "Provider=SQLOLEDB.1;Data Source=192.168.3.30,1433;Initial Catalog=JACINTA;User ID=SA;Password=c4r4m3l1t0;"
'    txtConexion =  "Provider=SQLOLEDB.1;Data Source=192.168.3.30,1433;Network Library=DBMSSOCN;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
'    txtConexion =  "Provider=SQLOLEDB.1;Data Source=PC-MMOLINA;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
'    txtConexion =  "Provider=SQLNCLI;Data Source=PC-MMOLINA;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
'    txtConexion =  "Provider=SQLNCLI10;Data Source=PC-MMOLINA;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
'    txtConexion =  "Driver={SQL Native Client};Server=PC-MMOLINA;Database=JACINTA;Uid=sa;Pwd=c4r4m3lt0"
'    txtConexion =  "DSN=desarrollo;UID=SA;PWD=c4r4m3l1t0"
'    txtConexion =  "Provider=SQLNCLI;Server=PC-MMOLINA;Database=JACINTA; Uid=SA; Pwd=c4r4m3l1t0;"
'    txtConexion =  "Provider=SQLNCLI10;Server=PC-MMOLINA;Database=JACINTA; Uid=SA; Pwd=c4r4m3l1t0;"
'    txtConexion =  "Driver={SQL Native Client 10.0};Server=PC-MMOLINA;Database=JACINTA;Uid=sa;Pwd=c4r4m3lt0"
'    txtConexion = 	"Provider={SQL Server};Server=PC-MMOLINA;Database=JACINTA;User ID=sa;Pwd=c4r4m3lt0"
'    txtConexion =  "Provider=SQLXMLOLEDB.4.0;Data Provider=SQLNCLI;Data Source=PC-MMOLINA;Initial Catalog=JACINTA;User Id=sa;Password=c4r4m3l1t0;"



''''    MODELADOR - DESARROLLO 
''''  	txtConexion= "Provider=SQLOLEDB.1;Data Source=192.168.3.30;Initial Catalog=JACINTA; User id=SA; Password=c4r4m3l1t0;"
     
 txtConexion = "Provider=SQLOLEDB.1;Data Source=servsql\mssqlserver1;Initial Catalog=JACINTA;User ID=SA;Password=F3QY17;"
 



Cnn.Open txtConexion

Set RS = Server.CreateObject("ADODB.Recordset")
	RS.ActiveConnection = Cnn
	RS.CursorType       = 3 'CONST adOpenStatic = 3
	RS.LockType         = 1 'CONST adReadOnly = 1
	RS.CursorLocation   = 3 'CONST adUseClient = 3
' set txtConexion = nothing
%>
