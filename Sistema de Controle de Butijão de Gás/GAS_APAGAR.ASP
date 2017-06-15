<!--#include file="inc_conexao.asp"-->
 
<%
call abre_conexao01
call abre_conexao02
dim objrs01,objconn01,objrs02,objconn02
'nivel = nivel_site
'if ((nivel="1")or(nivel="2")or(nivel="3")) then

'else
'    Response.Redirect ("default_acessonegado.asp")
'end if
'nivel_aux = nivel
%>
<!--#include file="inc_css.asp"-->

<html><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
coditem  = REQUEST("coditem")
inserir = request("inserir")
set objrs01 = objconn01.execute("select * from gas_itens where coditem="&coditem&"")
if not objrs01.eof then
	hora 	= objrs01("hora")
	forma	= objrs01("forma")
	valor	= objrs01("valor")
	op		= objrs01("op")
	rua		= objrs01("rua")
	bairro	= objrs01("bairro")
	numero	= objrs01("numero")
	nome	= objrs01("nome")
	valor	= formatnumber(valor,2)
	pago	 = OBJRS01("pago")
else
	hora 	= "Erro!!!"
	forma	= "Erro!!!"
	valor	= "0,00"
	op		= "Erro!!!"
	rua		= "Erro!!!"
	bairro	= "Erro!!!"
	numero	= "Erro!!!"
	nome	= "Erro!!!"
	pago	= "Erro!!!"
end if
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form.senha.focus()">
<div align="center"> 

  <br>
  <table width="375" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td width="24"><div align="left"><strong>
						  <a href="javascript:history.go(-1)"></a>
                        </strong></div></td>
                        <td width="722"><div align="center"><strong>APAGAR A VENDA DO G&Aacute;S </strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <div align="center"> 
<%
coditem	= request("coditem")
inserir = request("inserir")
'op		= request("op")
if inserir <> "SIM" then
%>
  <br>
  <table width="375" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center"> 
                  <script language="JavaScript">
					function ValidateOrder(form)
					{
					  if (form.senha.value == "")
					  { 
						  alert("Por favor digite a Senha."); form.senha.focus(); return; 
					  }    
					 
					  form.submit();
					}
					</script>
                  <form ACTION="gas_apagar.asp" METHOD="post" name="form" id="form">
                    <input name="inserir" type="hidden" value="SIM">
                    <br>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						    <tr>
						      <td colspan="2"><div align="center"><strong>Informa&ccedil;&otilde;es</strong></div></td>
					        </tr>
						    <tr>
						      <td><div align="right">Hora:</div></td>
						      <td><%=hora%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Forma de Pagamento:&nbsp; </div></td>
						      <td><%=forma%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Pre&ccedil;o:&nbsp;</div></td>
						      <td><%=VALOR%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Operador:&nbsp;</div></td>
						      <td><%=op%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Rua:&nbsp;</div></td>
						      <td><%=rua%></td>
					        </tr>
						    <tr>
						      <td><div align="right">N&uacute;mero:&nbsp;</div></td>
						      <td><%=numero%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Bairro:&nbsp;</div></td>
						      <td><%=bairro%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Nome:&nbsp;</div></td>
						      <td><%=nome%></td>
					        </tr>
						    <tr>
						      <td><div align="right">Pago:&nbsp;</div></td>
						      <td><%=pago%></td>
					        </tr>
						    <tr>
						      <td><div align="right"></div></td>
						      <td>&nbsp;</td>
					        </tr>
						    <tr>
					          <td width="29%"><div align="right">SENHA :&nbsp;</div></td>
							  <td width="71%">
							  <%
							%>
							  <input name="senha" type="password" id="senha" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="6" ></td>
							</tr>
						  </table>                        </td>
                      </tr>
                    </table>
                    <input name="coditem"type="hidden" id="cont" value="<%=coditem%>">
                    <input name="op"type="hidden" id="coditem" value="<%=op%>">
                    <br>
                    <input name="button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Continua" style=" text-transform: uppercase; FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                    <br>
                </form>
                </div></td>
            </tr>
          </table></td>
      </tr>
</table>
<%
else
%>
<table width="375" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center">
                <%

senha  = REQUEST("senha")
coditem = REQUEST("coditem")

'set objrs01 = objconn01.execute("select senha from gas_configuracao where senha='"&senha&"'")
sql = "select * from funcionarios where senha='"&senha&"' and senha <> '1'"
set objrs01 = objconn01.execute(sql)
if not objrs01.eof then
   
sql = "delete from gas_itens where coditem = "&coditem&" "
'response.write sql	
set objRs01 = objConn01.execute(sql)	

set objRs01 = objConn01.execute("delete from estoque_gas where codgas = "&coditem&"")	
%>
					                  <SCRIPT LANGUAGE="JavaScript">

	opener.window.location.href = "gas_inserir.asp?inserir=não&op=<%=op%>"
//	opener.location.reload(true);
	self.window.close();
    </SCRIPT>
                  <br>
                  <table width="80%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"><strong>
					                  <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	//opener.location.reload(true);
	//self.window.close();
}
    </SCRIPT>
	<%else%>
						                  <SCRIPT LANGUAGE="JavaScript">

	alert("Senha Incorreta!");
//	alert("<%=sql%>");
	//opener.location.reload(true);
//	opener.form2.preco.focus();
	self.window.close();
    </SCRIPT>
                  <br>
                  <table width="80%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"><strong>
					                  <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	//opener.location.reload(true);
	//opener.form2.preco.focus();
	self.window.close();
}
    </SCRIPT>	
	<%end if%>
					 <a href="#" onClick="fecha_refresh()"><img src="IMAGEM/botoes_fechar.gif" width="75" height="19" border="0"></a>
					  </strong></td>
                    </tr>
                  </table>
                  <br>
                </div></td>
            </tr>
          </table></td>
      </tr>
    </table>
    <%

end if
call fecha_conexao01
%>
    <br>
  </div>
</div>
</body>
</html>