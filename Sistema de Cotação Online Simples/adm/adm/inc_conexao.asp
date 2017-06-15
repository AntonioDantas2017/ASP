<%
'A programaчуo abaixo щ utilizada para 
'evitar que o servidor armazene em cache o sistema

Session.LCID = 1046 
'Response.clear 'Limpa tudo o que estava armazenado no buffer 
'Response.AddHeader "cache-control", "private"
'Response.AddHeader "pragma", "no-cache"
'Response.CacheControl = "no-cache"
'Response.CacheControl = "Private"
'Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
'Response.Expires= 0
'Response.Buffer = False
server.ScriptTimeout = 300

' Abaixo sуo feitos dois procedimentos necessсrios para realizar e fechar a 
'conexуo com o banco de dados



sub abre_conexao01
		
	Set objConn01 = Server.CreateObject("ADODB.Connection") 
	objConn01.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:/inetpub/wwwroot/geral/adm/dados/bd.mdb;"
'	objConn01.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=e:/home/acquasolution/dados/acquasolution_ bd2.mdb;"
	Set objrs01 = Server.CreateObject("ADODB.Recordset") 
	objConn01.CursorLocation = 3 ' щ o mesmo que adUseClient
	
end sub

sub fecha_conexao01
	objConn01.close
	set objConn01=nothing
end sub

'CONEXУO 02
'----------
sub abre_conexao02
	
	Set objConn02 = Server.CreateObject("ADODB.Connection") 
	objConn02.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:/inetpub/wwwroot/geral/adm/dados/bd.mdb;"
'	objConn02.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=e:/home/acquasolution/dados/acquasolution_ bd2.mdb;"
	Set objrs02 = Server.CreateObject("ADODB.Recordset") 
	objConn02.CursorLocation = 3 ' щ o mesmo que adUseClient
	
end sub
sub fecha_conexao02
	objConn02.close
	set objConn02=nothing
end sub

'CONEXУO 03
'----------
sub abre_conexao03
	
	Set objConn03 = Server.CreateObject("ADODB.Connection") 
	objConn03.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=C:/inetpub/wwwroot/geral/adm/dados/bd.mdb;"
'	objConn03.open = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=e:/home/acquasolution/dados/acquasolution_ bd2.mdb;"
	Set objrs03 = Server.CreateObject("ADODB.Recordset") 
	objConn03.CursorLocation = 3 ' щ o mesmo que adUseClient
	
end sub
sub fecha_conexao03
	objConn03.close
	set objConn02=nothing
end sub
%>