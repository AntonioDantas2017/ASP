<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<!--#include file="inc_css.asp"-->


<%
call abre_conexao01
nivel = nivel_contratos
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
call fecha_conexao01
nivel_aux = nivel
'nivel_aux = 2
%>


<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="MM_goToURL('parent.frames[\'topFrame\']','sistema2.asp?barra=news');return document.MM_returnValue">
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
                  <td bgcolor="#BBECB1"><div align="center"><strong> LOCAT&Aacute;RIOS</strong></div></td>
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
                          <td width="420"><div align="right"></div></td>
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
          DE LOCATARIO</strong></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td bgcolor="#FFFFFF"><div align="center"><br>
                <%
'RESGATA AS VARI�VEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codfor		= request("codfor")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")

nome		= TRIM(Request ("nome"))
qualificacao	= TRIM(Request ("qualificacao"))
cpf		= TRIM(Request ("cpf"))
rg	= TRIM(Request ("rg"))
contato		= TRIM(Request ("contato"))
call abre_conexao01
	
	set objRs01 = objConn01.execute("select * from locatarios where cpf = '"&cpf&"' ")
	
	if objrs01.eof then
	
	sql = "INSERT INTO locatarios (nome,qualificacao,cpf,rg,contato) values ('"&nome&"','"&qualificacao&"','"&cpf&"','"&rg&"','"&contato&"')"
	'response.write sql
	set objRs01 = objConn01.execute(sql)

call fecha_conexao01
	
%>
                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                O locat&aacute;rio foi realizado com sucesso!<br>
                </strong><strong><br>
                </strong><br>
                <br><%else%>
				
				                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                J&aacute; existe esse locat�rio cadastrado no sistema!<br>
                </strong><strong><br>
                </strong><br>
                <br>
				
				<%end if%>
                <table width="60%" border="0" cellspacing="0" cellpadding="0" class=linka>
                  <tr> 
                    <td width="33%"><div align="center"><a href="locatarios_inserir.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/botoes_inserir.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></div></td>
                    <td width="33%"><div align="center"><strong><a href="locatarios.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/botoes_principal.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></strong></div></td>
                  </tr>
                </table>
                <br>
                <br>
                <br>
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>