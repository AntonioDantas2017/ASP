<!--#include file="adm/inc_conexao.asp"-->

<%
'RESGATA TODAS AS VARIÁVEIS
'--------------------------
Dim objConn01, objrs01
Dim cont,login,senha,resposta,data,inicio,ip,browser

login 		= Request("login")
senha 		= Request("senha")
resposta	= Request("resposta")

login 		= replace(login,"'","")
senha 		= replace(senha,"'","")
resposta 	= replace(resposta,"'","")

login 		= replace(login,"select","")
login 		= replace(login,"delete","")
login 		= replace(login,"insert","")
login 		= replace(login,"update","")

senha 		= replace(senha,"select","")
senha 		= replace(senha,"delete","")
senha 		= replace(senha,"insert","")
senha 		= replace(senha,"update","")

resposta 		= replace(resposta,"select","")
resposta 		= replace(resposta,"delete","")
resposta 		= replace(resposta,"insert","")
resposta 		= replace(resposta,"update","")




'response.write login & senha & resposta

data		= date
inicio		= time
ip			= Request.ServerVariables("REMOTE_ADDR")
browser		= Request.ServerVariables("HTTP_USER_AGENT")

'USUÁRIO E SENHA VAZIOS
'----------------------
if (login="" or senha="" or resposta="") then
    Response.Redirect ("default1.asp?acesso=vazio")
end if

call abre_conexao01

'DEFINE O Nº DO CONTADOR E A SESSÃO
'-----------------------------------------------------------------------------------------
set objrs01 = objConn01.execute("select max(CONT) as maxCONT from LOG")
if objrs01("maxcont") <> "" then
	maxcont= objrs01("maxcont") + 1 
else
	maxcont= "100000"
end if
Session("log") = maxcont

'VERIFICA SE O LOGIN EXISTE
'--------------------------
set objrs01 = objConn01.execute("SELECT * FROM funcionarios WHERE login='"&login&"' AND senha='"&senha&"'  AND resposta='"&resposta&"' ")
if  not objrs01.eof then
	bloqueado 	= objrs01("bloqueado")
	codfunc 	= objrs01("codfunc") 
	nacesso 	= objrs01("nacesso")

	'VERIFICA SE O IP É CADASTRADO
	'-----------------------------------------------------------------------------------------
	'	set objrs01 = objConn01.execute(" SELECT * FROM ip WHERE ip='"&ip&"'  ")
	'	if objrs01.eof then
	'		set objrs01 = objConn01.execute("INSERT INTO log (cont,data,inicio,login,senha,resposta,resultado,ip,browser) VALUES ('"&maxcont&"','"&data&"','"&inicio&"','"&login&"','"&senha&"','"&resposta&"','IP INVÁLIDO','"&ip&"','"&browser&"') ")
	'	    Response.Redirect ("default.asp?acesso=5")
	'	end if
	'-----------------------------------------------------------------------------------------

	'VERIFICA SE ESTÁ BLOQUEADO
	'--------------------------
	if bloqueado = "sim" then
			'REGISTRA O ACESSO NO LOG
			'------------------------
		    set objrs01 = objConn01.execute("INSERT INTO log (cont,data,inicio,login,senha,resposta,resultado,ip,browser) VALUES ('"&maxcont&"','"&data&"','"&inicio&"','"&login&"','"&senha&"','"&resposta&"','ACESSO BLOQUEADO','"&ip&"','"&browser&"') ")
		Response.Redirect("default1.asp?acesso=bloqueado")
	end if
	
	'VERIFICA SE EXISTEM MAIS DE 3 ACESSOS INVALIDOS
	'-----------------------------------------------
	if bloqueado = "não" then
		set objrs01 = objConn01.execute("SELECT cont FROM funcionarios_cont WHERE login='"&login&"' ")
		if  not objrs01.eof then
			cont = objrs01("cont") + 1
			set objrs01 = objConn01.execute("update funcionarios_cont set cont='"&cont&"' WHERE login='"&login&"' ")
			cont = 1
				if cont >= 4 then
						'REGISTRA O ACESSO NO LOG
						'------------------------
					   set objrs01 = objConn01.execute("INSERT INTO log (cont,data,inicio,login,senha,resposta,resultado,ip,browser) VALUES ('"&maxcont&"','"&data&"','"&inicio&"','"&login&"','"&senha&"','"&resposta&"','ACESSO BLOQUEADO','"&ip&"','"&browser&"') ")
				   Response.Redirect("default1.asp?acesso=bloqueado")
				end if
		else
			set objrs01 = objConn01.execute("INSERT INTO funcionarios_cont(login,cont) VALUES('"&login&"','1') ")
		end if		
	end if
	
	'USUÁRIO LIBERADO
	'----------------
	Session("mercado_login") = login
	Session("mercado_senha") = senha
    Session("Authenticated") = 1
	
	'ATUALIZA O CONTADOR DO Nº DE ACESSO DO USUÁRIO
	'----------------------------------------------
	if nacesso > 0 then
		nacesso = nacesso + 1
		set objrs01 = objConn01.execute("update funcionarios set nacesso='"&nacesso&"' WHERE codfunc='"&codfunc&"' ")
	else
		set objrs01 = objConn01.execute("update funcionarios set nacesso='1' WHERE codfunc='"&codfunc&"' ")
	end if

	'REGISTRA O ACESSO NO LOG
	'------------------------
	set objrs01 = objConn01.execute("INSERT INTO log (cont,data,inicio,login,senha,resposta,resultado,ip,browser) VALUES 	('"&maxcont&"','"&data&"','"&inicio&"','"&login&"','"&senha&"','"&resposta&"','LIBERADO','"&ip&"','"&browser&"') ")

	set objrs01 = objConn01.execute("delete from funcionarios_cont WHERE login='"&login&"' ")
	call fecha_conexao01

	'GRAVA O COOKIE DO LOGIN	
	'-----------------------
	if login <> "" then
		response.cookies("mercado")("login") = login
		response.cookies("mercado").expires = "31/12/2010"		
	end if
	
	'ABRE A JANELA DO SISTEMA E RETORNA À PÁGINA PRINCIPAL
	'-----------------------------------------------------

%>
	<script language="JavaScript" type="text/JavaScript">
	function MM_preloadImages() 
	{ 
		var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
		var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
		if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
	}
	function MM_openBrWindow(theURL,winName,features) 
	{ 
		window.open(theURL,winName,features);
	}
	</script>
	<script language="JavaScript">
		MM_openBrWindow('adm/default.asp','','status=yes','width=1024,height=740')
		history.back(-1)
	</script>

<%
'	response.write "login : " & login & "<BR>"
'	response.write "senha : " & senha & "<BR>"


Else
	'USUÁRIO NEGADO
	'--------------

	'CONTADOR PARA OS ACESSOS INVALIDOS DESTE LOGIN
	'----------------------------------------------
	set objrs01 = objConn01.execute("SELECT cont FROM funcionarios_cont WHERE login='"&login&"' ")
	if  not objrs01.eof then
		cont = objrs01("cont") + 1
		set objrs01 = objConn01.execute("update funcionarios_cont set cont='"&cont&"' WHERE login='"&login&"' ")
	else
		set objrs01 = objConn01.execute("INSERT INTO funcionarios_cont(login,cont) VALUES('"&login&"','1') ")
	end if		
	
	'REGISTRA O ACESSO NO LOG
	'------------------------
	set objrs01 = objConn01.execute("INSERT INTO log (cont,data,inicio,login,senha,resposta,resultado,ip,browser) VALUES 	('"&maxcont&"','"&data&"','"&inicio&"','"&login&"','"&senha&"','"&resposta&"','ACESSO NEGADO','"&ip&"','"&browser&"') ")
		
    Session("Authenticated") = 0
	call fecha_conexao01
	
   Response.Redirect ("default1.asp?acesso=negado")
End If


%>