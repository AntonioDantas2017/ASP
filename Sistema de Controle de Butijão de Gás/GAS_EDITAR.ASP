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
	telefone= objrs01("telefone")
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
	telefone= "Erro!!!"
end if
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form2.forma.focus()">
<div align="center"> 

  
  <table width="475" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td><div align="center"><strong>EDITAR A VENDA DO G&Aacute;S </strong></div></td>
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
<table width="475" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center"> 
                  <script language="JavaScript">
					function ValidateOrder(form)
					{
						  if (form.forma.value == "")
	  { 
		  alert("Por favor escolha a Forma de Pagamento."); form.forma.focus(); return; 
	  }    
	  if (form.valor.value == "")
	  { 
		  alert("Por favor digite o preço."); form.valor.focus(); return; 
	  }    
	  if (form.op.value == "")
	  { 
		  alert("Por favor escolha o Operador."); form.op.focus(); return; 
	  }
					  if (form.senha.value == "")
					  { 
						  alert("Por favor digite a Senha."); form.senha.focus(); return; 
					  }    
					 
					  form.submit();
					}
					</script>
                  <input name="inserir" type="hidden" value="SIM">
             
                <form ACTION="gas_editar.asp" METHOD="post" name="form2" id="form2">
                      <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						   
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
						      <td><div align="right">Quantidade:&nbsp;</div></td>
						      <td><%=telefone%></td>
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
						      <td><div align="right"><strong>Forma de Pagamento:&nbsp; </strong></div></td>
						      <td><select name="forma" id="forma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  >
                                <option value="" <%if forma="" then%>selected<%end if%>>Forma Pagamento</option>
                                <option value="Dinheiro" <%if forma="Dinheiro" then%>selected<%end if%>>Dinheiro</option>
                                <option value="Cartão" <%if forma="Cartão" then%>selected<%end if%>>Cart&atilde;o</option>
                                <option value="Cheque" <%if forma="Cheque" then%>selected<%end if%>>Cheque</option>
                                <option value="Crediário" <%if forma="Crediário" then%>selected<%end if%>>Credi&aacute;rio</option>
                                <option value="Vale" <%if forma="Vale" then%>selected<%end if%>>Vale</option>
                              </select></td>
					        </tr>
							 <tr>
						      <td><div align="right"><strong>Pre&ccedil;o:&nbsp;</strong></div></td>
						      <td>          <input name="valor" type="text" id="valor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=valor%>" size="10" maxlength="6" ></td>
					        </tr> <tr>
						      <td><div align="right"><strong>Operador:&nbsp;</strong></div></td>
						      <td><select name="select" id="op" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  >
                                <option value="" <%IF op = "" then%> selected <%end if%>>OP</option>
                                <option value="1" <%IF op = "1" then%> selected <%end if%>>1</option>
                                <option value="5" <%IF op = "5" then%> selected <%end if%>>5</option>
                                <option value="6" <%IF op = "6" then%> selected <%end if%>>6</option>
                                <option value="7" <%IF op = "7" then%> selected <%end if%>>7</option>
                                <option value="9" <%IF op = "9" then%> selected <%end if%>>9</option>
                                <option value="10" <%IF op = "10" then%> selected <%end if%>>10</option>
                                <option value="11" <%IF op = "11" then%> selected <%end if%>>11</option>
                              </select></td>
					        </tr> <tr>
						      <td><div align="right"><strong>Rua:&nbsp;</strong></div></td>
						      <td><input name="rua" type="text" id="rua"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=rua%>" size="30" maxlength="60" ></td>
					        </tr> <tr>
						      <td><div align="right"><strong>N&uacute;mero:&nbsp;</strong></div></td>
						      <td><input name="numero" type="text" id="numero"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=numero%>" size="10" maxlength="6" ></td>
					        </tr> <tr>
						      <td><div align="right"><strong>Bairro:&nbsp;</strong></div></td>
						      <td><select name="bairro" id="bairro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"  >
                                <option value="" <%if bairro="" then%> selected <%end if%>>Bairro</option>
                                <option value="Campos de São José" <%if bairro="Campos de São José" then%> selected <%end if%>>Campos de S&atilde;o Jos&eacute;</option>
                                <option value="Mariana 1" <%if bairro="Mariana 1" then%> selected <%end if%>>Mariana 1</option>
                                <option value="Mariana 2" <%if bairro="Mariana 2" then%> selected <%end if%>>Mariana 2</option>
                                <option value="Santa Cecília 1" <%if bairro="Santa Cecília 1" then%> selected <%end if%>>Santa Cec&iacute;lia 1</option>
                                <option value="Santa Cecília 2" <%if bairro="Santa Cecília 2" then%> selected <%end if%>>Santa Cec&iacute;lia 2</option>
                                <option value="Pouzada" <%if bairro="Pouzada" then%> selected <%end if%>>Pouzada</option>
                                  <option value="Pouzada">Pouzada</option>
                      <option value="Cajuru" <%if bairro="Cajuru" then%> selected <%end if%>>Cajuru</option>
                      <option value="Serrote" <%if bairro="Serrote" then%> selected <%end if%>>Serrote</option>
                      <option value="Jardim Helena" <%if bairro=">Jardim Helena" then%> selected <%end if%>>Jardim Helena</option>  
                                <option value="Outros" <%if bairro="Outros" then%> selected <%end if%>>Outros</option>
                              </select></td>
					        </tr> <tr>
						      <td><div align="right"><strong>Quantidade</strong></div></td>
						      <td><input name="telefone" type="text" id="telefone"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"  onKeypress="tecla_enter2(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=telefone%>" size="20" maxlength="20" ></td>
					        </tr>
							 <tr>
						      <td><div align="right"><strong>Nome Entregador </strong></div></td>
						      <td>
					<select name="nome" id="nome" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" tabindex="7" >
                      <option value=""  <%if nome="" then%> selected <%end if%>>Entregador</option>
                      <option value="Claudio"  <%if nome="Claudio" then%> selected <%end if%>>Claudio</option>
                      <option value="Omary" <%if nome="Omary" then%> selected <%end if%>>Omary</option>
                      <option value="Julio" <%if nome="Julio" then%> selected <%end if%>>Julio</option>
                      <option value="Jurandir" <%if nome="Jurandir" then%> selected <%end if%>>Jurandir</option>                                                                  
                      <option value="Outro" <%if nome="Outro" then%> selected <%end if%>>Outro</option>                                                                                        
                      
                  </select></td>
					        </tr>  <tr>
						      <td><div align="right"><strong> Pago </strong></div></td>
						      <td><select name="pago" id="pago" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  >
                                <option value="N" <%IF pago = "N" then%> selected <%end if%>>N&atilde;o</option>
                                <option value="S" <%IF pago = "S" then%> selected <%end if%>>Sim</option>
                              </select></td>
					        </tr> 
							  <script language="JavaScript">
<!--


function tecla_enter2(tecla,form)
{ 
	if(tecla == 13){ 
		MM_openBrWindow('busca_telefone.ASP?telefone='+form.telefone.value,'','status=yes,scrollbars=yes,width=750,height=572');
	}
}

//-->
                   </script>
							<tr>
						      <td><div align="right"></div></td>
						      <td>&nbsp;</td>
					        </tr>
						    <tr>
					          <td width="36%"><div align="right">SENHA :&nbsp;</div></td>
							  <td width="64%">
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
<table width="475" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center">
                <%

senha  = REQUEST("senha")
coditem = REQUEST("coditem")

rua		   	= request("rua")
numero   	= request("numero")
bairro	   	= request("bairro")
nome	   	= request("nome")
telefone	= request("telefone")
valor		= replace(request("valor"),".",",")
forma   	= request("forma")
op		   	= request("op")
pago	   	= request("pago")
coditem	  	= Request ("coditem")
if not isnumeric(valor) then
	valor = 0
end if
'set objrs01 = objconn01.execute("select senha from gas_configuracao where senha='"&senha&"'")
sql = "select * from funcionarios where senha='"&senha&"' and senha <> '1'"
set objrs01 = objconn01.execute(sql)
if not objrs01.eof then
sql = "update gas_itens set rua='"&rua&"',numero='"&numero&"',bairro='"&bairro&"',nome='"&nome&"',telefone='"&telefone&"',valor='"&valor&"',forma='"&forma&"',op='"&op&"',pago='"&pago&"'  where coditem = "&coditem&" "
'response.write sql	
%>
	<SCRIPT LANGUAGE="JavaScript">
	//	alert("<%=sql%>");
    </SCRIPT>
<%

set objRs01 = objConn01.execute(sql)	

%>
					                  <SCRIPT LANGUAGE="JavaScript">
//	alert("<%=sql%>");
	opener.window.location.href = "gas_inserir.asp?inserir=não&op=<%=op%>"
//	opener.location.reload(true);
	self.window.close();
    </SCRIPT>
                  <br>
                  <table width="80%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td align="center"><strong>
					      
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