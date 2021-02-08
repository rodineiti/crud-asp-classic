<% 

strDriver = "MySQL ODBC 5.1 Driver"
strServer = "localhost"
strUser   = "root"
strPass   = "root"
strDB     = "cms"

' Cria uma nova instancia da Classe formando o Objeto
Set conDB = Server.CreateObject ("ADODB.Connection")

strConexaoDB = "driver={"&strDriver&"};server="&strServer&";uid="&strUser&";pwd="&strPass&";database=" & strDB

Session("connectionstring")	=	strConexaoDB

' Abre a conexao com o Banco de dados
conDB.ConnectionString = Session("connectionstring")
conDB.Open

%>
