
<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
if nivel_usuarios = 0 then
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

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" >
<div align="center"> 
  <%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa  = request("pesquisa")
ordem 	  = request("ordem")
listar	  = request("listar")
filtro	  = request("filtro")

codfunc = request("codfunc")

'BUSCA OS DADOS DO FUNCIONÁRIO
'-----------------------------
call abre_conexao01
set objrs01 = objConn01.execute("select * from funcionarios WHERE codfunc='"&codfunc&"'  ")
nome			= objrs01 ("nome")
nomeresumido	= objrs01 ("nomeresumido")
login 			= objrs01 ("login")
senha			= objrs01 ("senha")
'meta			= objrs01 ("meta")

permitebloqueio		=objrs01("permitebloqueio")
nivel_vendas		=objrs01("nivel_vendas")
nivel_clientes		=objrs01("nivel_clientes")
nivel_telefones		=objrs01("nivel_telefones")
nivel_avisos		=objrs01("nivel_avisos")
nivel_usuarios		=objrs01("nivel_usuarios")
nivel_log			=objrs01("nivel_log")
nivel_pagamentos	=objrs01("nivel_pagamentos")

call fecha_conexao01
%>
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
  
</div>
<div align="center"> </div>
<div align="center"> <br>
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
/*  if(form.login.value.length < "3")
  {
  	alert("O campo login deve ter no mínimo 3 caract."); form.login.focus(); return;  
  }*/
   if (form.senha.value == "")
  { 
 	 alert("Por favor digite a senha"); form.senha.focus(); return; 
  }
  /*if(form.senha.value.length < "6")
  {
  	alert("O campo senha deve possuir 6 números."); form.senha.focus(); return;  
  }*/
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
                <br>

 <form METHOD="get" ACTION="funcionarios_editar2.asp">
                    <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                    <input name="codfunc" type="hidden" value="<%=codfunc%>">
                  <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong> 
                  <br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
                    <tr bgcolor="#CCCCCC"> 
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td width="98"><strong>C&Oacute;DIGO</strong></td>
                            <td><strong>NOME RESUMIDO</strong></td>
                            <td width="167">&nbsp;</td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td width="20">&nbsp;</td>
                      <td width="98%"> <table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <tr> 
                                  <td width="72"><font color="#FF0000"><strong><%=codfunc%></strong></font></td>
                                  <td><input name="nomeresumido" type="text" id="nomeresumido"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" value="<%=nomeresumido%>" size="12" maxlength="10"></td>
                                  <td width="159">&nbsp;</td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr bgcolor="#CCCCCC"> 
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td width="42%"><strong>LOGIN</strong></td>
                            <td width="25%"><strong>SENHA</strong></td>
                            <td width="33%"><strong>CONFIRMA&Ccedil;&Atilde;O</strong></td>
                          </tr>
                        </table></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                          <tr> 
                            <td width="41%"><input name="login" type="text" id="login"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=login%>" size="20" maxlength="10"></td>
                            <td width="26%"><strong> 
                              <input name="senha" type="password" id="senha"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=senha%>" size="8" maxlength="8">
                            </strong></td>
                            <td width="33%"><input name="senha2" type="password" id="senha2"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=senha%>" size="8" maxlength="8"></td>
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
                            <td width="29%">VENDAS:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_vendas" type="radio" value="0" <%IF nivel_vendas = "0" or nivel_vendas = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_vendas" type="radio" value="1" <%IF nivel_vendas = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_vendas" type="radio" value="2" <%IF nivel_vendas = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						  
						  <tr bgcolor="#FFFFFF">
                            <td width="29%">PAGAMENTOS:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_pagamentos" type="radio" value="0" <%IF nivel_pagamentos = "0" or nivel_pagamentos = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_pagamentos" type="radio" value="1" <%IF nivel_pagamentos = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_pagamentos" type="radio" value="2" <%IF nivel_pagamentos = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						  
						  <tr bgcolor="#FFFFFF">
                            <td width="29%">PRODUTOS:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_produtos" type="radio" value="0" <%IF nivel_produtos = "0" or nivel_produtos = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_produtos" type="radio" value="1" <%IF nivel_produtos = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_produtos" type="radio" value="2" <%IF nivel_produtos = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						        <tr bgcolor="#FFFFFF">
                            <td width="29%">CLIENTES:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_clientes" type="radio" value="0" <%IF nivel_clientes = "0" or nivel_clientes = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_clientes" type="radio" value="1" <%IF nivel_clientes = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_clientes" type="radio" value="2" <%IF nivel_clientes = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  

                              </div></td>
                          </tr>
						        <tr bgcolor="#FFFFFF">
                            <td width="29%">TELEFONES:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_telefones" type="radio" value="0" <%IF nivel_telefones = "0" or nivel_telefones = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_telefones" type="radio" value="1" <%IF nivel_telefones = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_telefones" type="radio" value="2" <%IF nivel_telefones = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						        <tr bgcolor="#FFFFFF">
                            <td width="29%">AVISOS:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_avisos" type="radio" value="0" <%IF nivel_avisos = "0" or nivel_avisos = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_avisos" type="radio" value="1" <%IF nivel_avisos = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_avisos" type="radio" value="2" <%IF nivel_avisos = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						        <tr bgcolor="#FFFFFF">
                            <td width="29%">USU&Aacute;RIOS:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_usuarios" type="radio" value="0" <%IF nivel_usuarios = "0" or nivel_usuarios = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_usuarios" type="radio" value="1" <%IF nivel_usuarios = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_usuarios" type="radio" value="2" <%IF nivel_usuarios = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              </div></td>
                          </tr>
						        <tr bgcolor="#FFFFFF">
                            <td width="29%">LOG:</td>
                            <td width="71%"><div align="center">
                                <input name="nivel_log" type="radio" value="0" <%IF nivel_log = "0" or nivel_log = "" THEN%>checked<%END IF%>>
                              0-Sem Acesso
                              <input name="nivel_log" type="radio" value="1" <%IF nivel_log = "1" THEN%>checked<%END IF%>>
                              1-Consulta
                              <input name="nivel_log" type="radio" value="2" <%IF nivel_log = "2" THEN%>checked<%END IF%>>
                            2-Consulta+Altera&ccedil;&atilde;o  
                              
                              </div></td>
                          </tr>
                      </table></td>
                    </tr>
                  </table>
                  <br>
                  <input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma"style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                </form>
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
</div>
</body>
</html>