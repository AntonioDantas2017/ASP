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
codgas  = REQUEST("codgas")
set objrs01 = objconn01.execute("select * from gas where aberta='sim'")
if objrs01.eof and codgas = "" then

ELSE
	if codgas = "" then
		codgas = objrs01("codgas")
	end if
END IF
if not isnumeric(codgas) then
	codgas =0
end if
inserir = request("inserir")
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.form.senha.focus()">
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
                        <td width="722"><div align="center"><strong>FECHAMENTO</strong></div></td>
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
'CODGAS	= request("CODGAS")
inserir = request("inserir")
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
                  <form ACTION="gas_fechamento.asp" METHOD="post" name="form" id="form">
                    <input name="inserir" type="hidden" value="SIM">
                    <br>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						    <tr>
						      <td><div align="right">TOTAL DE G&Aacute;S VENDIDO:&nbsp;</div></td>
						      <td><% 
if not isnumeric(codgas) then
codgas = 0
end if
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"")
if not objrs02.eof then
	response.Write right("00"&objrs02.recordcount,3)
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "00"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL DINHEIRO :&nbsp; </div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Dinheiro'")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%>
					          <% 
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Dinheiro'")
if not objrs02.eof then
	response.Write "("&right("00"&objrs02.recordcount,2)&")"
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "(00)"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL CART&Atilde;O :&nbsp;</div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Cartão'")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%>
					          <% 
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Cartão'")
if not objrs02.eof then
	response.Write "("&right("00"&objrs02.recordcount,2)&")"
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "(00)"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL CHEQUE :&nbsp;</div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Cheque'")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%>
					          <% 
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Cheque'")
if not objrs02.eof then
	response.Write "("&right("00"&objrs02.recordcount,2)&")"
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "(00)"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL CREDI&Aacute;RIO :&nbsp;</div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Crediário'")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%>
					          <% 
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Crediário'")
if not objrs02.eof then
	response.Write "("&right("00"&objrs02.recordcount,2)&")"
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "(00)"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL VALE :&nbsp;</div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Vale'")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%>
					          <% 
set objrs02 =objconn02.execute("select coditem from gas_itens where codgas="&codgas&"  and (contado='' or contado='não' or contado=null) and forma='Vale'")
if not objrs02.eof then
	response.Write "("&right("00"&objrs02.recordcount,2)&")"
	quantidade_vendida = quantidade_vendida + objrs02.recordcount
else
	response.Write "(00)"
end if
%></td>
					        </tr>
						    <tr>
						      <td><div align="right">TOTAL:&nbsp; </div></td>
						      <td><% 
set objrs02 =objconn02.execute("select sum(valor) as total from gas_itens where (contado='' or contado='não' or contado=null) and codgas="&codgas&"")
if isnumeric(objrs02("total")) then
	response.Write formatnumber(objrs02("total"),2)
else
	response.Write formatnumber(0,2)
end if
%></td>
					        </tr>
						    <tr>
						      <td>&nbsp;</td>
						      <td>&nbsp;</td>
					        </tr>
						    <tr>
						      <td><div align="right">ESTOQUE ATUAL :&nbsp;</div></td>
						      <td><%
							  set objrs02 =objconn02.execute("select sum(quantidade) as estoque from estoque_gas")
if isnumeric(objrs02("estoque")) then
	response.Write objrs02("estoque")
else
	response.Write "0"
end if
							  %></td>
					        </tr>
						    <tr>
						      <td><div align="right"></div></td>
						      <td>&nbsp;</td>
					        </tr>
						    <tr>
					          <td width="38%"><div align="right">SENHA :&nbsp;</div></td>
							  <td width="62%">
							  <%

							%>
							  <input name="senha" type="password" id="senha" style=" FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="6" ></td>
							</tr>
						  </table>                        </td>
                      </tr>
                    </table>
                    <input name="codgas"type="hidden" id="cont" value="<%=codgas%>">
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
codgas = REQUEST("codgas")

set objrs01 = objconn01.execute("select senha from gas_configuracao where senha='"&senha&"'")
if not objrs01.eof then
   
sql = "update gas_itens set contado='sim'"
'response.write sql	
set objRs01 = objConn01.execute(sql)	
sql = "update gas set aberta='não'"
'response.write sql	
set objRs01 = objConn01.execute(sql)
	
set objRs01 = objConn01.execute("select max(codgas) as maxcodgas from gas")
if IsNull(objRs01("maxcodgas")) then
   codgas = 100000
else
   codgas = objRs01("maxcodgas") + 1
end if

set objRs01 = objConn01.execute("select estoque from gas_configuracao")	
estoque = objrs01("estoque")

sql = "INSERT INTO gas (codgas,data,aberta,estoque) values ("&codgas&",'"&date&"','sim',"&estoque&") "
'response.write sql	
set objRs01 = objConn01.execute(sql)

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

	alert("Senha Incorreta!");
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
        
      </div>
</div>
</body>
</html>