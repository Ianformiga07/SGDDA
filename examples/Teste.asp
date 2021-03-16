<!--#include file ="lib/Conexao.asp"-->
<%
Server.ScriptTimeout = 15000000
dim arrayModulo(4)
arrayModulo(0) = 2
arrayModulo(1) = 4
arrayModulo(2) = 4
arrayModulo(3) = 4
arrayModulo(4) = 2
call abreConexao
for i = 0 to 4 
 
  for j = 0 to arrayModulo(i) - 1  
	for k = 0 to 6
	  for p = 0 to 49
sql = "insert INTO  TB_Recipiente(CodigoModulo, CodigoFace, ID_Nivel, ID_Pasta) Values('"&i+1&"','"&j+1&"','"&k+1&"','"&p+1&"')"
	   'conn.execute(sql)
	  next  
	next
  Next
Next
call FechaConexao
response.write "finalizou"

%>