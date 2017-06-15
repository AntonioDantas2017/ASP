
<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->

<%
call abre_conexao01
nivel = nivel_funcionarios
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

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="MM_goToURL('parent.frames[\'topFrame\']','sistema2.asp?barra=clientes');return document.MM_returnValue">
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa  = request("pesquisa")
ordem 	  = request("ordem")
listar	  = request("listar")
filtro	  = request("filtro")
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
                  <td bgcolor="#BBECB1"><div align="center"><strong>USU&Aacute;RIOS</strong></div></td>
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
                          <td width="30"><div align="center"></div></td>
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
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>CADASTRO 
          DE USU&Aacute;RIOS</strong></div></td>
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
 
  if (form.nomeresumido.value == "")
  { 
	  alert("Por favor digite o nome resumido"); form.nomeresumido.focus(); return; 
  }

   if (form.login.value == "")
  { 
	  alert("Por favor digite o login"); form.login.focus(); return; 
  }
  if(form.login.value.length < "3")
  {
  	alert("O campo login deve ter no mínimo 3 caract."); form.login.focus(); return;  
  }
   if (form.senha.value == "")
  { 
 	 alert("Por favor digite a senha"); form.senha.focus(); return; 
  }
  if(form.senha.value.length < "6")
  {
  	alert("O campo senha deve possuir 6 números."); form.senha.focus(); return;  
  }
   if (form.senha2.value == "")
  { 
	  alert("Por favor digite a senha de confirmação"); form.senha2.focus(); return; 
  }
  if (form.senha.value != form.senha2.value) 
  { 
    alert("A senha e a confirmação da senha não são iguais!");
    form.senha2.focus();
    return false;
  } 
 
  form.submit();
}
</script>
                  <form METHOD="get" ACTION="funcionarios_inserir2.asp">
                    <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                    <input name="maxcod" type="hidden" value="<%=maxcod%>">
                  <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong><br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
                    <tr bgcolor="#CCCCCC"> 
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td><strong>NOME</strong></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td width="20">&nbsp;</td>
                      <td width="98%"> <table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <tr> 
                                  <td><font color="#FF0000"><strong></strong></font>                                    <input name="nomeresumido" type="text" id="nomeresumido"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="12" maxlength="10"></td>
                                  <td width="24%">&nbsp;</td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    
                    <tr bgcolor="#CCCCCC"> 
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td width="33%"><strong>LOGIN</strong></td>
                            <td width="33%"><strong>SENHA</strong></td>
                            <td width="33%"><strong>CONFIRMA&Ccedil;&Atilde;O</strong></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td width="33%"><input name="login" type="text" id="login"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="20" maxlength="10">
                              ( m&iacute;n. 03 caract. )</td>
                            <td width="33%"><strong> 
                              <input name="senha" type="password" id="senha"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="8" maxlength="6">
                              </strong>( 6 n&uacute;meros )</td>
                            <td width="33%"><input name="senha2" type="password" id="senha2"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="8" maxlength="6">
                              ( 6 n&uacute;meros )</td>
                          </tr>
                        </table></td>
                    </tr>
                                      
                    
                    
                  </table>
                  <br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td bgcolor="#CCCCCC"><table width="650" border="0" cellspacing="1" cellpadding="1" CLASS=LINKA>
                          <tr bgcolor="#FFFFFF">
                            <td colspan="2"><div align="center">
                                <table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                  <tr>
                                    <td><div align="center"><strong>&Aacute;REAS 
                                      DE ACESSO</strong></div></td>
                                  </tr>
                                </table>
                            </div></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td width="28%">COTAC&Atilde;O:</td>
                            <td width="72%"><div align="center">
                                <input name="nivel_cotacao" type="radio" value="0" checked>
                              0-Sem Acesso
                              <input name="nivel_cotacao" type="radio" value="1">
                              1-Consulta
                              <input name="nivel_cotacao" type="radio" value="2">
                              2-Consulta+Altera&ccedil;&atilde;o </div></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>MERCADO :</td>
                            <td><div align="center">
                                <input name="nivel_mercado" type="radio" value="0" checked>
                              0-Sem Acesso
                              <input name="nivel_mercado" type="radio" value="1">
                              1-Consulta
                              <input name="nivel_mercado" type="radio" value="2">
                              2-Consulta+Altera&ccedil;&atilde;o </div></td>
                          </tr>
                          
                          <tr bgcolor="#FFFFFF">
                            <td>CONTRATOS  : </td>
                            <td><div align="center">
                                <input name="nivel_contratos" type="radio" value="0" checked>
                              0-Sem Acesso
                              <input name="nivel_contratos" type="radio" value="1">
                              1-Consulta
                              <input name="nivel_contratos" type="radio" value="2">
                              2-Consulta+Altera&ccedil;&atilde;o </div></td>
                          </tr>
                          <tr bgcolor="#FFFFFF">
                            <td>FUNCION&Aacute;RIOS :</td>
                            <td><div align="center">
                                <input name="nivel_funcionarios" type="radio" value="0" checked>
                              0-Sem Acesso
                              <input name="nivel_funcionarios" type="radio" value="1">
                              1-Consulta
                              <input name="nivel_funcionarios" type="radio" value="2">
                              2-Consulta+Altera&ccedil;&atilde;o </div></td>
                          </tr>
                          
                          
                          
                          
                          
                          
                          
                          
                          
                          
                          
                          
                      </table></td>
                    </tr>
                  </table>
                  <br>
                  <br>
                  <input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                  </form>                
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <br>
  <br>
</div>
</body>
</html>