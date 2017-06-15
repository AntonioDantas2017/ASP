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
%>
<!--#include file="inc_css.asp"-->


<html><head>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


</head>


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<div align="center"><br>
  <table width="650" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td>                         <div align="center"><strong><font color="#000000">OBSERVA&Ccedil;&Otilde;ES DA COTA&Ccedil;&Atilde;O </font></strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <%
auxapaga = request("auxapaga")
cont	 = request("cont")

auxapaga2 = request("auxapaga2")
codprc	  = request("codprc")

CODCOT = request("CODCOT")


if auxapaga = "sim" then
	sql=("update prod_cot_forn set preco='0', obs='' where cont="&cont&" ")
	'Response.write(sql)
	set objrs01 = objconn01.execute(sql)
	auxapaga = ""
	%>
  <table width="650" border="0" cellspacing="1" cellpadding="1" class=linka>
    <tr>
      <td height="20" bgcolor="#FFAAAA"><div align="center"><strong> O PRODUTO FOI APAGADO DA LISTA DO FORECEDOR COM SUCESSO! </strong></div></td>
    </tr>
  </table>
  <%END IF%>
  <%
if auxapaga2 = "sim" then
	sql=("delete from prod_cot_forn where codprc="&codprc&" ")
	'Response.write(sql)
	set objrs01 = objconn01.execute(sql)
	sql=("delete from produtos_cot where codprc="&codprc&" ")
	'Response.write(sql)
	set objrs01 = objconn01.execute(sql)	
	auxapaga = ""
	%>

  <table width="650" border="0" cellspacing="1" cellpadding="1" class=linka>
    <tr>
      <td height="20" bgcolor="#FFAAAA"><div align="center"><strong> O PRODUTO FOI APAGADO DA LISTA DA COTA&Ccedil;&Atilde;O COM SUCESSO! </strong></div></td>
    </tr>
  </table>
  <%END IF%>
  <br>
 
  <div align="center">
  <%
  
  
  SET OBJRS01 = OBJCONN01.EXECUTE("SELECT * FROM PROD_COT_FORN2 WHERE OBS<>'' AND CODCOT="&CODCOT)
  IF OBJRS01.EOF THEN
  %>
  
    <table width="650" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr>
        <td bgcolor="#FFFFFF"><div align="center"> <strong> N&Atilde;O FOI ENCONTRADO 
          NENHUMA OBSERVA&Ccedil;&Atilde;O.</strong></div>
            <div align="center"><strong></strong></div></td>
      </tr>
    </table>
    <%
RESPONSE.WRITE "<br><br><br><br><br><br>"
else
%>
    <table width="650" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
      <tr>
        <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
          <tr>
            <td width="11%" rowspan="2" bgcolor="#FFFFFF">LEGENDA:</td>
            <td width="89%" bgcolor="#FFFFFF"> AP FORN: APAGA O PRODUTO DA LISTA DO FORNECEDOR </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">AP PROD: APAGA O PRODUTO DA LISTA DA COTA&Ccedil;&Atilde;O </td>
          </tr>
        </table></td>
      </tr>
    </table>
    <div align="center"> <br>
        <div align="center">
          <table width="650" border="1" cellpadding="0" cellspacing="0" CLASS="linkA">
            <tr valign="top">
              <td width="200"><div align="center"><strong><font color="#000000">PRODUTO</font></strong></div></td>
              <td width="100"><div align="center"><font color="#000000"><strong>PRE&Ccedil;O<br>
                FORNECEDOR</strong></font></div></td>
              <td width="149"><div align="center"><strong><font color="#000000">OBSERVA&Ccedil;&Otilde;ES</font></strong></div></td>
              <td width="111"><div align="center"><strong><font color="#000000">MENOR<br>
              PRE&Ccedil;O</font></strong></div></td>
              <td width="38"><div align="center"><font color="#000000"><strong>
                <%if nivel_aux="2" then%>
                AP<br>
                FORN
  <%end if%>
              </strong></font></div></td>
              <td width="38"><div align="center"><font color="#000000"><strong> 
                <%if nivel_aux="2" then%>
                AP PROD
                <%end if%>
              </strong></font></div></td>
            </tr>
            <% 

'For cont = 1 to 50
contador = 0
while not objrs01.eof 
contador = contador + 1
cont  	= objrs01("cont")
codfor 	= objrs01("codfor")
codprc	= objrs01("codprc")
preco	= objrs01("preco")
obs		= objrs01("obs")
%>
            
           
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
       
			  
            <td width="200"><div align="left">
              <%
			
			set objrs02 = objconn02.execute("select * from produtos_cot where codprc="&codprc&"")
			if not objrs02.eof then
				codpro = objrs02("codpro")
				set objrs02 = objconn02.execute("select * from produtos where codpro="&codpro&"")
				if not objrs02.eof then
					response.Write objrs02("produto")
				else
					response.Write "Erro! Produto não Encontrado"
				end if
			else
				response.Write "Erro! Produto não Encontrado"
			end if
			
			
			
			%>
            </div></td>
            <td><div align="left">
              <%
			  	IF NOT ISNUMERIC(preco) THEN			  
					preco = 0
				END IF
			  	IF preco = 0 THEN
					RESPONSE.Write "Não Tem"&"<BR>"
				ELSE
					RESPONSE.Write FORMATCURRENCY(preco,2)&"<BR>"
				END IF
				set objrs02 = objconn02.execute("select empresa,vendedor from fornecedores where codfor="&codfor&"")
				if not objrs02.eof then
					response.Write ucase(objrs02("vendedor") & "-" & objrs02("empresa"))
				else
					response.Write "Erro! Fornecedor não Encontrado"
				end if
				%>
            
            </DIV></TD>
             
            <td><div align="left">
              <%
				response.Write obs
				%>
            </DIV></TD>
				
            <td><div align="left">
              <%
				sql = "SELECT min(preco) as menor,codfor FROM prod_cot_forn2 where codprc="&codprc&" and preco <> 0 group by preco,codfor"
				set objrs02 = objConn02.execute(sql)
				if (not isnumeric(objrs02("menor"))) or (isnull(objrs02("menor"))) then
			
				else
					codfor = objrs02("codfor")	
					preco = objrs02("menor")
					response.Write formatcurrency(preco,2)&"<BR>"
					set objrs02 = objconn02.execute("select empresa,vendedor from fornecedores where codfor="&codfor&"")
					if not objrs02.eof then
						response.Write ucase(objrs02("vendedor") & "-" & objrs02("empresa"))
					else
						response.Write "Erro! Fornecedor não Encontrado"
					end if
				end if
				%>
              <a href="javascript:MM_openBrWindow('produtos_precos.asp?codpro=<%=codpro%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_precos.gif" alt="Pre&ccedil;os desse produto" width="10" height="10" border="0"></a>
            </DIV></TD>
            <form name="form1" method="get" action="cotacao_obs.asp">
              <input type="hidden" value="<%=CONT%>" name="CONT">
              <input type="hidden" value="<%=codcot%>" name="codcot">			  
              <input type="hidden" value="sim" name="auxapaga">
            <td><div align="center">
              <script language="JavaScript">
					function confirmar()
					{
					   return (confirm('Deseja apagar este registro?'))
					}
					  </script>
              <%if nivel_aux="2" then%>
              <input NAME="delete2" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar Produto da Lista do Fornecedor" WIDTH="20" HEIGHT="20">
              <%end if%>
            </div></TD>
			</form>
			
			
			
			
            <form name="form1" method="get" action="cotacao_obs.asp">
              <input type="hidden" value="<%=codprc%>" name="codprc">
              <input type="hidden" value="sim" name="auxapaga2">
              <td width="38" valign="middle">
                <div align="center">
                  <script language="JavaScript">
					function confirmar()
					{
					   return (confirm('Deseja apagar este registro?'))
					}
					  </script>                  
                <%if nivel_aux="2" then%>    <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar Produto da Lista da Cota&ccedil;&atilde;o" WIDTH="20" HEIGHT="20">  
                <%end if%>
                </div></td>
            </form>
            </tr>
            <% 
	objrs01.movenext
	wend
	end if	
%>
          </table>
        </div>
    </div>
    <div align="center"> <br>
        <br>
        <table width="300" border="0" cellspacing="0" cellpadding="0" class=linka>
          <tr>
            <td width="50%"><div align="center">
                <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
            <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></div></td>
          </tr>
        </table>
    </div>
  </div>
</div>
<%
call fecha_conexao01
call fecha_conexao02
%>
</body>
</html>