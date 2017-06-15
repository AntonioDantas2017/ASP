<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_mercado
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
call fecha_conexao01
nivel_aux = nivel
'nivel_aux = 2
%>
<%
call abre_conexao01
%>
<!--#include file="inc_css.asp"-->


<html><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="MM_goToURL('parent.frames[\'topFrame\']','sistema2.asp?barra=NEWS');return document.MM_returnValue">
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codcli		= request("codcli")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")
codcli    = request("codcli")

set objrs01 = objConn01.execute("select * from clientes where codcli="&codcli&" ")
if not objrs01.eof then
	nome		=	objrs01("nome")
	obs 		=	objrs01("obs")
end if
call fecha_conexao01
%>
<div align="center"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="115"> <table width="115" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>CLIENTES</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="3">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="30"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"><strong></strong></div></td>
                          <td width="120"> <div align="center"></div></td>
                          <td width="130"> <div align="center"></div></td>
                          <td width="130"> <div align="right"></div></td>
                          <td width="420"><div align="right"></div>
                            <div align="center"></div>
                            <div align="center"><strong> </strong></div>
                            <div align="center"></div>
                            <div align="right"> </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <br>
  <table width="700" border="0" cellspacing="0" cellpadding="0" class=linka>
    <tr> 
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>
	    CADASTRO DE CLIENTES 
	  </strong></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td bgcolor="#FFFFFF"><div align="center">
<script language="JavaScript">
function ValidateOrder(form)
{
  if (form.nome.value == "")
  { 
	  alert("Por favor digite o Nome"); form.nome.focus(); return; 
  }  
     

  form.submit();

}
</script>
                <br>
                <font color="#FF0000" size="2"><strong>( n&atilde;o utilizar os 
                caracteres especiais : &quot; ' &amp; % ( ) &lt; &gt; )</strong></font><br>
				  <form action="clientes_editar2.asp" method="get" name="form" id="form" >
                    <input name="codcli" type="hidden" value="<%=codcli%>">
                   <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
					<input name="tipo" type="hidden" value="<%=tipo%>">
					<input name="filtro" type="hidden" value="<%=filtro%>">					
					<input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
					<input name="datainicio" type="hidden" value="<%=datainicio%>">
					<input name="datafim" type="hidden" value="<%=datafim%>">	
					<input name="filtro2" type="hidden" value="<%=filtro2%>">						
                  <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong><br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td width="100%"><div align="center"><table width="650" border="0" cellspacing="0" cellpadding="0">
                        <tr valign="top">
                          <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td bgcolor="#CCCCCC"><strong>NOME</strong></td>
                              </tr>
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td height="18"><input name="nome" type="text" id="nome" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=nome%>" size="60" maxlength="50"></td>
                              </tr>
                            
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                            
                              
                              <tr>
                                <td bgcolor="#CCCCCC"><strong>OBSERVA&Ccedil;&Otilde;ES</strong></td>
                              </tr>
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td bgcolor="#FFFFFF"><textarea name="obs" cols="30" rows="2" id="obs" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"><%=obs%></textarea></td>
                              </tr>
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                          </table></td>
                        </tr>
                      </table>
                      </div></td>
                    </tr>
				</table>
                  <br>
                  <br>
<input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">				  
                  <br>
                </form>                
                
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
</div>
</body>
</html>