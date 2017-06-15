<!--#include file="inc_conexao.asp"-->
 
<%
call abre_conexao01
call abre_conexao02
dim objrs01,objconn01,objrs02,objconn02
%>
<!--#include file="inc_css.asp"-->

<html><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form2.forma.focus()">
<div align="center"> 
  <table width="100" border="0" height="7">
    <tr>
      <td height="7"></td>
    </tr>
  </table>
  <table width="475" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td><div align="center"><strong>INSERIR A VENDA DO G&Aacute;S </strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
 
  <table width="100" border="0" height="7">
    <tr>
      <td height="7"></td>
    </tr>
  </table>
  <%
inserir	  	= request("inserir")
forma   	= request("forma")
valor   	= request("valor")
op		   	= request("op")
pago	   	= request("pago")
codgas	  	= Request ("codgas")

rua		   	= request("rua")
numero   	= request("numero")
bairro	   	= request("bairro")
nome	   	= request("nome")
telefone	= request("telefone")

auxpago 	= request("auxpago")	
auxcoditem	= request("auxcoditem")	

valor		= replace(valor,".",",")
if not isnumeric(valor) then
	valor2 = request("valor2")
	valor2 = replace(valor2,".",",")
	if isnumeric(valor2) then
		valor = valor2
	else
		valor = 0
	end if	
end if


if inserir <> "sim" then
%>
 
  <table width="450" border="0">
    <tr>
      <td><form METHOD="get" ACTION="gas_insert.asp" name="form2">
        <table width="475" border="0" bgcolor="#CCCCCC">
          <tr>
            <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
                <tr bgcolor="#FFFFFF">
                  <td width="39%" bgcolor="#FFFFFF"><div align="left"><strong>Forma de Pagamento </strong></div></td>
                  <td bgcolor="#FFFFFF"><div align="center"><strong>Pre&ccedil;o</strong></div></td>
                  <td width="14%" bgcolor="#FFFFFF"><div align="center"><strong>QTDE</strong></div></td>
                  <td width="8%" rowspan="6" bgcolor="#FFFFFF"><div align="center">
                      <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" tabindex="9">
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td bgcolor="#FFFFFF"><div align="left">
                      <script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a fun&ccedil;&atilde;o 
		ValidateOrder(form);
	}
}

function tecla_enter2(tecla,form)
{ 
	if(tecla == 13){ 
		MM_openBrWindow('busca_telefone.ASP?telefone='+form.telefone.value,'','status=yes,scrollbars=yes,width=750,height=572');
	}
}
function ValidateOrder(form)
{
	  if (form.forma.value == "")
	  { 
		  alert("Por favor escolha a Forma de Pagamento."); form.forma.focus(); return; 
	  }    
	  if (form.valor.value == "")
	  { 
		  alert("Por favor digite o pre&ccedil;o."); form.valor.focus(); return; 
	  }    
	  if (form.op.value == "")
	  { 
		  alert("Por favor escolha o Operador."); form.op.focus(); return; 
	  }    	  	  
  form.submit();
}
//-->
                   </script>
                      <select name="forma" id="forma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" >
                        <option value="" selected>Forma Pagamento</option>
                        <option value="Dinheiro">Dinheiro</option>
                        <option value="Cart&atilde;o">Cart&atilde;o</option>
                        <option value="Cheque">Cheque</option>
                        <option value="Credi&aacute;rio">Credi&aacute;rio</option>
                        <option value="Vale">Vale</option>
                      </select>
                  </div></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <%
			set objrs01 = objconn01.execute("select valor from gas_configuracao")
			valor = formatnumber(objrs01("valor"),2)
			%>
                      <input name="valor" type="text" id="valor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=valor%>" size="10" maxlength="6" tabindex="1">
                      <input type="hidden" name="inserir" value="sim">
                      <input type="hidden" name="codgas" value="<%=request("codgas")%>">
                      <input type="hidden" name="valor2" value="<%=valor%>">
                  </div></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <input name="telefone" type="text" id="telefone"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" tabindex="2"  onKeyPress="tecla_enter2(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="1" size="2" maxlength="2">
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td bgcolor="#FFFFFF"><div align="left"><strong>Rua</strong></div></td>
                  <td bgcolor="#FFFFFF"><div align="center"><strong>Operador</strong></div></td>
                  <td bgcolor="#FFFFFF"><div align="center"><strong>N&uacute;mero</strong></div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td bgcolor="#FFFFFF"><div align="left">
                      <input name="rua" type="text" id="rua"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="30" maxlength="60"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" tabindex="3" >
                  </div></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <select name="op" id="op" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" tabindex="4"  >
                        <option value="" <%IF request("op") = "" then%> selected <%end if%>>OP</option>
                        <option value="1" <%IF request("op") = "1" then%> selected <%end if%>>01  - Mossor&oacute;</option>
                        <option value="5" <%IF request("op") = "5" then%> selected <%end if%>>05  - Auxiliadora</option>
                        <option value="6" <%IF request("op") = "6" then%> selected <%end if%>>06  - Fabiola</option>
                        <option value="7" <%IF request("op") = "7" then%> selected <%end if%>>07  - Nen&ecirc;</option>
                        <option value="9" <%IF request("op") = "9" then%> selected <%end if%>>09  - Deusa</option>
                        <option value="10" <%IF request("op") = "10" then%> selected <%end if%>>10 - Fagner</option>
                        <option value="11" <%IF request("op") = "11" then%> selected <%end if%>>11 - Rosana</option>
                        <option value="12" <%IF request("op") = "12" then%> selected <%end if%>>12 - Socorro</option>
                      </select>
                  </div></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <input name="numero" type="text" id="numero"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="10" maxlength="6"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" tabindex="5" >
                  </div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td bgcolor="#FFFFFF"><strong>Bairro</strong></td>
                  <td bgcolor="#FFFFFF" align="center"><strong>Nome Entregador </strong></td>
                  <td bgcolor="#FFFFFF"><div align="center"><strong>Pago</strong></div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td bgcolor="#FFFFFF"><select name="bairro" id="bairro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" tabindex="6" >
                      <option value="" selected>Bairro</option>
                      <option value="Campos de S&atilde;o Jos&eacute;">Campos de S&atilde;o Jos&eacute;</option>
                      <option value="Mariana 1">Mariana 1</option>
                      <option value="Mariana 2">Mariana 2</option>
                      <option value="Santa Cec&iacute;lia 1">Santa Cec&iacute;lia 1</option>
                      <option value="Santa Cec&iacute;lia 2">Santa Cec&iacute;lia 2</option>
                      <option value="Pouzada">Pouzada</option>
                      <option value="Cajuru">Cajuru</option>
                      <option value="Serrote">Serrote</option>
                      <option value="Jardim Helena">Jardim Helena</option>                                                                  
                      <option value="Outros">Outros</option>
                  </select></td>
                  <td bgcolor="#FFFFFF" align="center">
					<select name="nome" id="nome" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" tabindex="7" >
                      <option value="" selected>Entregador</option>
                      <option value="Claudio">Claudio</option>
                      <option value="Omary">Omary</option>
                      <option value="Julio">Julio</option>
                      <option value="Jurandir">Jurandir</option>                                                                  
                      <option value="Outro">Outro</option>                                                                                        
                      
                  </select></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <select name="pago" id="pago" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" tabindex="8"  >
                        <option value="N" selected >N&atilde;o</option>
                        <option value="S">Sim</option>
                      </select>
                  </div></td>
                </tr>
            </table></td>
          </tr>
        </table>
      </form></td>
    </tr>
  </table>
  <%
else
	function preparaPalavra(str)
	   preparaPalavra = replace(str,"á","a")
	   preparaPalavra = replace(preparaPalavra,"à","a")
	   preparaPalavra = replace(preparaPalavra,"ã","a")
	   preparaPalavra = replace(preparaPalavra,"â","a")
	   preparaPalavra = replace(preparaPalavra,"ä","a")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"é","e")	
	   preparaPalavra = replace(preparaPalavra,"è","e")	
	   preparaPalavra = replace(preparaPalavra,"ê","e")	
	   preparaPalavra = replace(preparaPalavra,"ë","e")		   	   	   
	   preparaPalavra = replace(preparaPalavra,"í","i")
	   preparaPalavra = replace(preparaPalavra,"ì","i")
	   preparaPalavra = replace(preparaPalavra,"î","i")
	   preparaPalavra = replace(preparaPalavra,"ï","i")	   	   	   
	   preparaPalavra = replace(preparaPalavra,"ó","o")
	   preparaPalavra = replace(preparaPalavra,"ò","o")
	   preparaPalavra = replace(preparaPalavra,"õ","o")
	   preparaPalavra = replace(preparaPalavra,"ô","o")
	   preparaPalavra = replace(preparaPalavra,"ö","o")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"ú","u")
	   preparaPalavra = replace(preparaPalavra,"ù","u")
	   preparaPalavra = replace(preparaPalavra,"û","u")
	   preparaPalavra = replace(preparaPalavra,"ü","u")	   	   
	   preparaPalavra = preparaPalavra
	end function
	nome = ucase(preparaPalavra(lcase(nome)))
	rua  = ucase(preparaPalavra(lcase(rua)))
	
		set objRs01 = objConn01.execute("select max(coditem) as maxcoditem from gas_itens")
		if IsNull(objRs01("maxcoditem")) then
		   coditem = 100000
		else
		   coditem = objRs01("maxcoditem") + 1
		end if
		
		sql = "INSERT INTO gas_itens (coditem,codgas,forma,valor,op,rua,telefone,numero,bairro,nome,pago,hora,contado) values ("&coditem&",'"&codgas&"','"&forma&"','"&valor&"','"&op&"','"&rua&"','"&telefone&"','"&numero&"','"&bairro&"','"&nome&"','"&pago&"','"&time&"','não') "		
		'response.write sql	
		set objRs01 = objConn01.execute(sql)	

		%>
		<script>
		//	MM_openBrWindow('gas_imp.asp?coditem=<%=coditem%>','','status=yes,scrollbars=yes,width=100,height=100');
		</script>
		<%
	
		
		set objRs01 = objConn01.execute("insert into estoque_gas (quantidade,codgas) values(-"&telefone&","&coditem&")")	

%>
					                  <SCRIPT LANGUAGE="JavaScript">
//	alert("<%=sql%>");
	opener.window.location.href = "gas_inserir.asp?inserir=não&op=<%=op%>"
//	opener.location.reload(true);
	self.window.close();
    </SCRIPT>

    <%

end if
call fecha_conexao01
%>
   
</div>
</body>
</html>