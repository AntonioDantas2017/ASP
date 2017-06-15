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


<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="javascript:print();">

<div align="center">
  <%
  CODCOT = REQUEST("CODCOT")
 set objrs01 = objconn01.execute("select * from produtos order by codpro") 
  if objrs01.eof then%>
  <br>
  <table width="750" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
    <tr> 
      <td width="228"><div align="center"> 
          <!--#include file="inc_logo.asp"-->
        </div></td>
      <td width="522"><div align="center"><strong>PRODUTOS</strong></div></td>
    </tr>
  </table>
  <br>
  <br>
  <br>
  <br>
  <table width="750" border="0" cellspacing="1" cellpadding="1" class=linka>
    <tr> 
      <td bgcolor="#FFFFFF"><div align="center"> <strong> NÃO FOI ENCONTRADO NENHUM 
          REGISTRO.</strong></div>
        <div align="center"><strong></strong></div></td>
    </tr>
  </table>
  <br>
  <br>
  <br>
  <a href="javascript:history.go(-1)"><img src="imagem/botoes_voltar.gif" width="75" height="19" border="0"></a><br>


<%
else
	Function Ceil( sNumero ) 
	mRet = sNumero 
	if (mRet - fix( mRet ) ) <> 0 then 
	mRet = fix( mRet )+1 
	end if 
	Ceil = mRet 
	End Function 

	n_reg_pag = 74
	totalregistros = objrs01.RecordCount / 2
	npaginas = Ceil(totalregistros / n_reg_pag)
%>

<%
For contpag = 1 to npaginas
%>

  <br>
  
  <div align="center"> <br>
    <div align="center">
      <table width="750" border="1" cellpadding="0" cellspacing="0" bordercolor="#7C7C7C" class="linkA">
        <tr>
          <td width="50" bgcolor="#CCCCCC"><div align="center"><strong>COD</strong></div></td>
          <td width="235" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
          <td width="46" bgcolor="#CCCCCC"><div align="center"><strong>ENTRA </strong></div></td>
          <td width="21" bordercolor="#FFFFFF" bgcolor="#FFFFFF"></td>
          <td width="50" bgcolor="#CCCCCC"><div align="center"><strong>COD</strong></div></td>
          <td width="236" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
          <td width="46" bgcolor="#CCCCCC"><div align="center"><strong>ENTRA </strong></div></td>
        </tr>
        <%

		'contador = 1
	For cont = 1 to n_reg_pag
	 		IF NOT OBJRS01.EOF THEN
			ELSE
				Exit For
			END IF
	'contador = contador + 1
			set objrs02 = objconn02.execute("select * from produtos_cot where codpro="&objrs01("codpro")&" and codcot="&codcot&"")		  
			if objrs02.eof then 
				continua = "sim"
			else
				continua = "NÃO"
			end if
		if continua = "sim" then
	%>
        <%if (contador mod 2)<>0 then%>
        <tr valign="middle" bgcolor="#FFFFFF">
          <%else%>
        <tr valign="middle" bgcolor="#E6E6E6">
          <%end if%>
          <td><div align="center"><%=objrs01("codpro")%></div></td>
          <td><%=left(objrs01("produto"),35)%></td>
          <td>&nbsp;</td>
          <td bgcolor="#FFFFFF"><%
	  	objrs01.movenext		
		  IF (NOT OBJRS01.EOF)THEN

			set objrs02 = objconn02.execute("select * from produtos_cot where codpro="&objrs01("codpro")&" and codcot="&codcot&"")		  
			if objrs02.eof then 
				continua = "sim"
			else
				continua = "NÃO"
			end if
		else
			continua = "NÃO"
		END IF
		%>
          </td>
          <td><div align="center">
            <%if continua = "sim" then%>
            <%=objrs01("codpro")%>
            <%end if%>
          </div></td>
          <td><%if continua = "sim" then%>
              <%=left(objrs01("produto"),35)%>
            <%end if%></td>
          <td><%if continua = "sim" then%>&nbsp; <%end if%></td>
        </tr>
        <%

			 IF NOT OBJRS01.EOF THEN
		  		objrs01.movenext
			ELSE
				Exit For
			END IF
		ELSE
			Exit For
		end if
	
		
	Next
	
	%>
      </table>
      <%
	Next
%>
    </div>
  </div>


  <div align="center">

     <%end if%>
    <br>
    <%
call fecha_conexao01
call fecha_conexao02
%>

</div></div>
</body>
</html>