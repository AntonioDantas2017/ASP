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
                  <td bgcolor="#BBECB1"><div align="center"><strong>TELEFONE</strong></div></td>
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
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>CADASTRO DE TELEFONE
      </strong></div></td>
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
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codtel		= request("codtel")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")


nome		=	request("nome")
telefone	=	request("telefone")
telefone2	=	request("telefone2")	
celular		=	request("celular")	
fax			=	request("fax")
obs 		=	request("obs")

achou = ""

call abre_conexao01

set objRs01 = objConn01.execute("select (codtel) from telefones where (telefone='"&telefone&"' or telefone2='"&telefone&"' or celular='"&telefone&"' or fax='"&telefone&"')")
if NOT objrs01.eof then
	achou = telefone
end if

set objRs01 = objConn01.execute("select (codtel) from telefones where (telefone='"&telefone2&"' or telefone2='"&telefone2&"' or celular='"&telefone2&"' or fax='"&telefone2&"') ")
if NOT objrs01.eof then
	achou = telefone2
end if

set objRs01 = objConn01.execute("select (codtel) from telefones where (telefone='"&celular&"' or telefone2='"&celular&"' or celular='"&celular&"' or fax='"&celular&"')")
if NOT objrs01.eof then
'	RESPONSE.Write "select (codtel) from telefones where (telefone='"&celular&"' or telefone2='"&celular&"' or celular='"&celular&"' or fax='"&celular&"')"
	achou = celular
end if

set objRs01 = objConn01.execute("select (codtel) from telefones where (telefone='"&fax&"' or telefone2='"&fax&"' or celular='"&fax&"' or fax='"&fax&"') ")
if NOT objrs01.eof then
	achou = fax
end if
'RESPONSE.Write ACHOU
if achou = "" then
	set objRs01 = objConn01.execute("select max(codtel) as maxcodtel from telefones")
	if IsNull(objRs01("maxcodtel")) then
	   codtel = 100000
	else
	   codtel = objRs01("maxcodtel") + 1
	end if
	
	sql = "INSERT INTO telefones (codtel,nome,telefone2,telefone,fax,celular,obs) values ("&codtel&",'"&nome&"','"&telefone2&"','"&telefone&"','"&fax&"','"&celular&"','"&obs&"') "
	'response.write sql	
	set objRs01 = objConn01.execute(sql)

call fecha_conexao01
	
%>
                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                O cadastro do telefone foi realizado com sucesso!<br>

                </strong>
<%else%>
                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                Este(s) Telefone(s) já existe(em) na lista!<br>
                </strong>
<%end if%>
				<br>			
                    <br>
                    <table width="50%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center"><a href="telefones.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=achou%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=telefones&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/botoes_principal.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></div></td>
                      </tr>
                    </table>
                    <br>
                    <br>
                  </center>
                </form>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>