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


valor		=	request("valor")
data_venc	=	request("data_venc")
codloc		=	request("codloc")	
codcon		=	request("codcon")	

if isdate(data_venc) then
	data_conv4 = split(data_venc,"/") 'funcao utilizada para criar um vetor delimitado po "/"
	data_venc2 = data_conv4(1)&"/"&data_conv4(0)&"/"&data_conv4(2) 'reorganiza data para o formato dd/mm/yy
end if

achou = ""

call abre_conexao01
valor2 = replace(valor,".","")
valor2 = replace(valor,",",".")
set objRs01 = objConn01.execute("select codcon from pagamentos where data_venc=#"&data_venc2&"# and valor="&valor2&" and codcon="&codcon&" ")
if not objrs01.eof then
	achou = "sim"
end if


if achou = "" then
	valor2 = replace(valor,".",",")
	sql = "INSERT INTO pagamentos (data_venc,valor,codcon,pago) values ('"&data_venc&"','"&valor2&"','"&codcon&"','não') "
	response.write sql	
	set objRs01 = objConn01.execute(sql)

call fecha_conexao01
	
%>
                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                O cadastro do pagamento foi realizado com sucesso!<br>
                </strong>
<%else%>
                <strong><br>
                </strong><font class="linkA"> </font> <strong><br>
                <br>
                Este pagamento já existe na lista!<br>
                </strong>
<%end if%>
				<br>			
                    <br>
                    <table width="50%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center"><a href="pagamentos.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=achou%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/botoes_principal.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></div></td>
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