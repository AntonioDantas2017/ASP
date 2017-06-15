<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
if nivel_clientes = "0" then
	session("pg") = request.ServerVariables("SCRIPT_NAME")
    Response.Redirect ("default_acessonegado.asp")
end if
call abre_conexao01
%>
<!--#include file="inc_css.asp"-->


<html><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0"  onLoad="javascript:form.nome.focus();">
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
	tel1		=	objrs01("tel1")
	tel2		=	objrs01("tel2")	
	celular		=	objrs01("cel")	
	endereco	=	objrs01("endereco")
	bairro		=	objrs01("bairro")
	cidade		=	objrs01("cidade")
	estado		=	objrs01("estado")
	contato		=	objrs01("contato")
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
                      <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#CCCCCC"><strong>NOME </strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18" colspan="3"><input name="nome" type="text" id="nome" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=nome%>" size="80" maxlength="100"></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td width="30%" bgcolor="#CCCCCC"><strong>TELEFONE 1 </strong></td>
                            <td width="29%" bgcolor="#CCCCCC"><div align="left"><strong>TELEFONE 2 </strong></div></td>
                            <td width="41%" bgcolor="#CCCCCC"><strong>CELULAR</strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td bgcolor="#FFFFFF"><input name="tel1" type="text" id="tel" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=tel%>" size="16" maxlength="14" onKeyPress="return txtBoxFormat('tel', '(99) 9999-9999', event);">
                              (00) 0000-0000 </td>
                            <td bgcolor="#FFFFFF"><input name="tel2" type="text" id="tel2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=tel2%>" size="16" maxlength="14"  onKeyPress="return txtBoxFormat('tel2', '(99) 9999-9999', event);"></td>
                            <td bgcolor="#FFFFFF"><input name="celular" type="text" id="celular" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=celular%>" size="16" maxlength="14" onKeyPress="return txtBoxFormat('celular', '(99) 9999-9999', event);"></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#CCCCCC"><strong>ENDERECO</strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#FFFFFF"><input name="endereco" type="text" id="endereco" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=endereco%>" size="60" maxlength="100"></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td bgcolor="#CCCCCC"><strong>BAIRRO</strong></td>
                            <td bgcolor="#CCCCCC"><strong>CIDADE</strong></td>
                            <td bgcolor="#CCCCCC"><strong>ESTADO</strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td><input name="bairro" type="text" id="bairro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=bairro%>" size="30" maxlength="50"></td>
                            <td><input name="cidade" type="text" id="cidade" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=cidade%>" size="35" maxlength="50"></td>
                            <td><select name="estado" id="estado" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                              <option value="AC" <%if estado="AC" then%> selected <%end if%>>Acre</option>
                              <option value="AL" <%if estado="AL" then%> selected <%end if%>>Alagoas</option>
                              <option value="AM" <%if estado="AM" then%> selected <%end if%>>Amaz&ocirc;nia</option>
                              <option value="AP" <%if estado="AP" then%> selected <%end if%>>Amap&aacute;</option>
                              <option value="BA" <%if estado="BA" then%> selected <%end if%>>Bahia</option>
                              <option value="CE" <%if estado="CE" then%> selected <%end if%>>Cear&aacute;</option>
                              <option value="DF" <%if estado="DF" then%> selected <%end if%>>Distrito Federal</option>
                              <option value="ES" <%if estado="ES" then%> selected <%end if%>>Esp&iacute;rito Santo</option>
                              <option value="GO" <%if estado="GO" then%> selected <%end if%>>Goiais</option>
                              <option value="MA" <%if estado="MA" then%> selected <%end if%>>Maranh&atilde;o</option>
                              <option value="MG" <%if estado="MG" then%> selected <%end if%>>Minas Gerais</option>
                              <option value="MS" <%if estado="MS" then%> selected <%end if%>>Mato Grosso do Sul</option>
                              <option value="MT" <%if estado="MT" then%> selected <%end if%>>Mato Grosso</option>
                              <option value="PA" <%if estado="PA" then%> selected <%end if%>>Par&aacute;</option>
                              <option value="PB" <%if estado="PB" then%> selected <%end if%>>Para&iacute;ba</option>
                              <option value="PE" <%if estado="PE" then%> selected <%end if%>>Pernambuco</option>
                              <option value="PI" <%if estado="PI" then%> selected <%end if%>>Piau&iacute;</option>
                              <option value="RJ" <%if estado="RJ" then%> selected <%end if%>>Rio de Janeiro</option>
                              <option value="RN" <%if estado="RN" then%> selected <%end if%>>Rio Grande do Norte</option>
                              <option value="RO" <%if estado="RO" then%> selected <%end if%>>Ror&acirc;ima</option>
                              <option value="RS" <%if estado="RS" then%> selected <%end if%>>Rio Grande do Sul</option>
                              <option value="PR" <%if estado="PR" then%> selected <%end if%>>Paran&aacute;</option>
                              <option value="SC" <%if estado="SC" then%> selected <%end if%>>Santa Catarina</option>
                              <option value="SE" <%if estado="SE" then%> selected <%end if%>>Sergipe</option>
                              <option value="SP" <%if estado="SP" then%> selected <%end if%>>S&atilde;o Paulo</option>
                              <option value="TO" <%if estado="TO" then%> selected <%end if%>>Tocantins</option>
                            </select></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#CCCCCC"><strong>CONTATO</strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3"><input name="contato" type="text" id="contato" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=contato%>" size="60" maxlength="100"></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#CCCCCC"><strong>OBSERVA&Ccedil;&Otilde;ES</strong></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3" bgcolor="#FFFFFF"><textarea name="obs" cols="50" rows="3" id="obs" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"><%=obs%></textarea></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td colspan="3"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                      </table></td>
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