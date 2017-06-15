<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_cotacao
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
call fecha_conexao01
nivel_aux = nivel
'nivel_aux = 2
%><%
call abre_conexao01
call abre_conexao02
'nivel = nivel_bd
'if ((nivel="1")or(nivel="2")or(nivel="3")) then

'else
'    Response.Redirect ("default_acessonegado.asp")
'end if
'nivel_aux = nivel
%>
<!--#include file="inc_css.asp"-->


<html>
<style type="text/css">
<!--
.style3 {color: #000000; font-weight: bold; }
-->
</style>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


</head>
<script>
//select
</script>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">

<div align="left" class="linkA">
  <%
  CODCOT = REQUEST("CODCOT")
 set objrs01 = objconn01.execute("select * from produtos2 where CODCOT="&CODCOT&" order by codprc") 
 while not objrs01.eof
 produto = objrs01("produto")
 diff = 50 - len(produto)
 for i=1 to diff
 	produto = produto & "&nbsp;"
 next
 response.Write produto & "<BR>"
 objrs01.movenext
 wend
 %>

  </div>
    <%
call fecha_conexao01
call fecha_conexao02
%>

</body>
</html>