<!--#include file="inc_login.asp"-->

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
<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="MM_goToURL('parent.frames[\'topFrame\']','sistema2.asp?barra=clientes');return document.MM_returnValue">
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
            <td bgcolor="#FFFFFF"><div align="center"><br>
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa  	= request("pesquisa")
ordem 	  	= request("ordem")
listar	  	= request("listar")
filtro	  	= request("filtro")

rastreado = nomeresumido

nomeresumido= TRIM(Request("nomeresumido"))
nome 		= TRIM(Request("nome"))
login		= TRIM(Request("login"))
senha		= TRIM(Request("senha"))

nivel_cotacao		= TRIM(Request("nivel_cotacao"))
nivel_mercado		= TRIM(Request("nivel_mercado"))
nivel_contratos		= TRIM(Request("nivel_contratos"))
nivel_funcionarios	= TRIM(Request("nivel_funcionarios"))

'VERIFICA SE O LOGIN JÁ EXISTE
'-----------------------------
call abre_conexao01
	set objRs01 = objconn01.execute("select * from funcionarios where login='"&login&"' ")
	if not objRs01.eof then
%>
                <br>
                <div align="center"></div>
                <div align="center"></div>
                <div align="center"><strong><br>
                  <br>
                  <br>
                  <br>
                  O login j&aacute; existe!<br>
                  Favor digitar outro login.<br>
                  <br>
                  <br>
                  <br>
                  <br>
                  <br>
                  <a href="javascript:history.go(-1)"><img src="IMAGEM/botoes_voltar.gif" width="75" height="19" border="0"></a> 
                  </strong></div>
                      
                <br>
                <br>
                <br>
                <br>
<%
	else
	'CADASTRA O FUNCIONÁRIO NOVO
	'---------------------------
	set objRs01 = objconn01.execute("select max(codfunc) as maxcodfunc from funcionarios")
	if IsNull(objRs01("maxcodfunc")) then
	   maxcodfunc = "1000"
	else
	   aux = objRs01("maxcodfunc") + 1
	   maxcodfunc = right("0000" & aux,4)
	end if	


set objRs01 = objconn01.execute("INSERT INTO  Funcionarios(codfunc,nome,nomeresumido,login,senha) values ('"&maxcodfunc&"','"&nome&"','"&nomeresumido&"','"&login&"','"&senha&"') ")

set objRs01 = objconn01.execute("Update Funcionarios set nivel_cotacao='"&nivel_cotacao&"',nivel_mercado='"&nivel_mercado&"',nivel_contratos='"&nivel_contratos&"',nivel_funcionarios='"&nivel_funcionarios&"' where codfunc = '"&maxcodfunc&"' ")

%>
                <strong><br>
                <br>
                <br>
                <br>
                O cadastro do usu&aacute;rio foi realizado com sucesso!<br>
                <br>
                <br>
                <br>
                <br>
                </strong> 
                <table width="50%" border="0" cellspacing="0" cellpadding="0" class=linka>
                  <tr> 
                    <td width="50%"><div align="center"><a href="funcionarios_inserir.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"><img src="IMAGEM/botoes_inserir.gif" alt="Insere outro usu&aacute;rio" width="75" height="19" border="0"></a></div></td>
                    <td width="50%"><div align="center"><strong><a href="funcionarios.ASP?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"><img src="IMAGEM/botoes_principal.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></strong></div></td>
                  </tr>
                </table>
                <br>
                <br>
                <br>
                <%
	end if

'
call fecha_conexao01
%>
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>