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
acabou  = REQUEST("acabou")
inserir = request("inserir")
codforn	= request("codforn")
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form.preco.focus()">
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
                        <td width="722"><div align="center"><strong>ALTERAR PRECO DO PRODUTO 
                              <%
				cont = request("cont")
				sql = "SELECT * FROM prod_cot_forn where cont = "&cont&""
				set objrs01 = objConn01.execute(sql)
				codprc = objrs01("codprc")
				preco = objrs01("preco")
				obs = objrs01("obs")				

				sql = "SELECT * FROM produtos_cot where codprc = "&codprc&""
				set objrs01 = objConn01.execute(sql)
				codpro = objrs01("codpro")							  
				sql = "SELECT produto FROM produtos where codpro = "&codpro&""
				set objrs01 = objConn01.execute(sql)
				if not objrs01.eof then
					response.Write objrs01("produto")
				else
					response.Write "-"
				end if
					%>
</strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
  <div align="center"> 
<%
inserir = request("inserir")
if inserir <> "SIM" then
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
					  if (form.preco.value == "")
					  { 
						  alert("Por favor digite o preco."); form.preco.focus(); return; 
					  }    
					 
					  form.submit();
					}
					</script>
                  <form ACTION="alterar_preco.asp" METHOD="post" name="form" id="form">
                    <input name="inserir" type="hidden" value="SIM">
                    <input name="acabou" type="hidden" value="<%=acabou%>">
                    <input name="codforn" type="hidden" value="<%=codforn%>">					
					<br>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						    <tr>
					          <td width="41%"><div align="right">PRE&Ccedil;O :&nbsp;</div></td>
							  <td width="59%">
							  <%
							  if not isnumeric(preco) then
							  	preco = 0
							end if
							set objrs01 = objconn01.execute("select fora from fornecedores where codfor="&codforn&"")
						if objrs01("fora") = "sim" then
							preco = preco - (preco*0.06)
						end if

							%>
							  <input name="preco" type="text" id="preco" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="5" ></td>
							</tr>
						    <tr>
						      <td><div align="right">OBSERVA&Ccedil;&Atilde;O :&nbsp;</div></td>
						      <td><input name="obs" type="text" id="obs"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#EAEAEA" value="<%=obs%>" size="40" maxlength="50"></td>
					        </tr>
						  </table>                        </td>
                      </tr>
                    </table>
                    <input name="cont"type="hidden" id="cont" value="<%=cont%>">
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

cont  = REQUEST("cont")
preco = REQUEST("preco")
OBS   = REQUEST("OBS")
codforn       = REQUEST("codforn")
preco = REPLACE(preco,".",",")
if not isnumeric(preco) then
	preco = 0
end if

'set objrs01 = objconn01.execute("select fora from fornecedores where codfor="&codforn&"")
'if objrs01("fora") = "sim" then
'	preco = preco + (preco*0.06)
'end if

	   
	sql = "UPDATE PROD_COT_FORN SET PRECO='"&PRECO&"',OBS='"&OBS&"' WHERE CONT = "&CONT&""
	response.Write sql
	set objrs01 = objConn01.execute(sql)
	
	if acabou="sim" then
%>
					                  <SCRIPT LANGUAGE="JavaScript">

	opener.location.reload(true);
	self.window.close();
    </SCRIPT>
                  <br>
                  <table width="80%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"><strong>
					                  <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
	<%else%>
						                  <SCRIPT LANGUAGE="JavaScript">

	//opener.location.reload(true);
	opener.form2.preco.focus();
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
	opener.form2.preco.focus();
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