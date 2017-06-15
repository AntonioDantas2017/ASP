<!--#include file="inc_conexao.asp"-->

<%
Dim objConn01,objrs01
Dim objConn02,objrs02
Dim objConn03,objrs03
%>
<!--#include file="inc_backup.asp"-->
<%
call abre_conexao01
set objrs01 = objConn01.execute("SELECT * FROM funcionarios WHERE login='"&Session("piemme_login")&"' AND senha='"&Session("piemme_senha")&"'  ")
if  objrs01.eof then
	call fecha_conexao01
    Response.Redirect ("default_acessonegado.asp")
else
	nomeresumido    = objrs01("nomeresumido")
	codfunc 		= objrs01("codfunc")
	nacesso 		= objrs01("nacesso")
	acesso  		= session("acesso")

	nivel_cotacao		= objrs01("nivel_cotacao")
	nivel_mercado		= objrs01("nivel_mercado")
	nivel_contratos		= objrs01("nivel_contratos")
	nivel_funcionarios	= objrs01("nivel_funcionarios")
	
	call fecha_conexao01	
end if
%>
