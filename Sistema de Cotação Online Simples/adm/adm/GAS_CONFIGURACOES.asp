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

<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
inserir = request("inserir")
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.form.estoque.focus()">
<div align="center"> 

  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
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
                        <td width="722"><div align="center"><strong>CONFIGURA&Ccedil;&Otilde;ES</strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
  <%
inserir = request("inserir")
if inserir <> "SIM" then
set objrs01 = objconn01.execute("select * from gas_configuracao")
valor 	= objrs01("valor")
estoque = objrs01("estoque")
senha 	= objrs01("senha")
valor 	= formatnumber(valor,2)
%>
  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center"> 
                  <script language="JavaScript">
					function ValidateOrder(form)
					{
					  if (form.valor.value == "")
					  { 
						  alert("Por favor digite o preco."); form.valor.focus(); return; 
					  }    
					  if (form.estoque.value == "")
					  { 
						  alert("Por favor digite a quantidade no estoque."); form.estoque.focus(); return; 
					  } 
					  if (form.senha.value == "")
					  { 
						  alert("Por favor digite a senha."); form.senha.focus(); return; 
					  } 					  					  
					 
					  form.submit();
					}
					</script>
                  <form ACTION="gas_configuracoes.asp" METHOD="post" name="form" id="form">
                    <input name="inserir" type="hidden" value="SIM">
                    <br>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						    <tr>
					          <td width="50%"><div align="right">PRE&Ccedil;O :&nbsp;</div></td>
							  <td width="50%">
							  <input name="valor" type="text" id="valor" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=valor%>" size="10" maxlength="5" ></td>
							</tr>
						    <tr>
						      <td><div align="right">ADICIONAR AO ESTOQUE :&nbsp;</div></td>
						      <td><input name="estoque" type="text" id="estoque" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="4" ></td>
					        </tr>
						    <tr>
						      <td><div align="right">SENHA :&nbsp;</div></td>
						      <td><input name="senha" type="text" id="senha" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=senha%>" size="10" maxlength="6" ></td>
					        </tr>
						  </table>                        </td>
                      </tr>
                    </table>
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
<table width="500" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center">
                <%

valor   = REQUEST("valor")
senha   = REQUEST("senha")
estoque = REQUEST("estoque")
valor = REPLACE(valor,".",",")
if not isnumeric(valor) then
	valor = 0
end if
valor = REPLACE(valor,",",".")
	   
	sql = "UPDATE GAS_CONFIGURACAO SET VALOR="&VALOR&",SENHA='"&SENHA&"' WHERE SENHA <> '' "
'	response.Write sql
	set objrs01 = objConn01.execute(sql)
	
	set objRs01 = objConn01.execute("insert into estoque_gas (quantidade,codgas) values("&estoque&",0)")	
	
	if acabou="sim" then
%>
					                  <SCRIPT LANGUAGE="JavaScript">

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
	self.window.close();
}
    </SCRIPT>
	<%else%>
						                  <SCRIPT LANGUAGE="JavaScript">

	//opener.location.reload(true);
	//opener.form2.preco.focus();
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
                  <br></td>
            </tr>
          </table></td>
      </tr>
    </table>
    <%

end if
call fecha_conexao01
%>
    <br>
</body>
</html>