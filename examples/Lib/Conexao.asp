<%
dim conn
sub abreConexao
	'Criando a conex�o com o BD
	serverorigem = Request.ServerVariables("SERVER_NAME")




strcon =  "Provider=SQLNCLI11;Server=localhost;Database=Testando;Uid=sa;Pwd=123;"

Set conn = Server.CreateObject( "ADODB.Connection" )
	conn.open(strcon)	
end sub


sub fechaConexao
	'Fechando a conex�o com o BD
	conn.Close()
	Set conn = Nothing
end sub
%>