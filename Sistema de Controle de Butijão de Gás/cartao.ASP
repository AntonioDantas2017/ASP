<!--#include file="inc_conexao.asp"-->
 
<%
call abre_conexao01
call abre_conexao02
dim objrs01,objconn01,objrs02,objconn02

if request("nome") = "" then
%>
<!--#include file="inc_css.asp"-->

<html><head>
<title>Cartão Inserir</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form2.nome.focus()">
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
                        <td><div align="center"><strong>INSERIR CARTÃO</strong></div></td>
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
 
  <table width="450" border="0">
    <tr>
      <td><form METHOD="get" ACTION="cartao.asp" name="form2">
        <table width="475" border="0" bgcolor="#CCCCCC">
          <tr>
            <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
                <tr bgcolor="#FFFFFF">
                  <td width="34%" bgcolor="#FFFFFF" align="center"><strong>Nome</strong></td>
                  <td bgcolor="#FFFFFF" align="center"><div align="center"><strong>Valor</strong></div></td>
                  <td width="33%" bgcolor="#FFFFFF" align="center"><strong>Operador</strong></td>
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

function ValidateOrder(form)
{
	  if (form.nome.value == "")
	  { 
		  alert("Por favor escolha o nome."); form.nome.focus(); return; 
	  }    
	  if (form.valor.value == "")
	  { 
		  alert("Por favor digite o valor."); form.valor.focus(); return; 
	  }    
	  if (form.op.value == "")
	  { 
		  alert("Por favor escolha o Operador."); form.op.focus(); return; 
	  }    	  	  
  form.submit();
}
//-->
                   </script>
                      <input name="nome" type="text" id="nome0"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="" size="30" maxlength="50" tabindex="1" >
                  </div></td>
                  <td bgcolor="#FFFFFF"><div align="center">
                      <input name="valor" type="text" id="valor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="" size="10" maxlength="6" tabindex="2">
                      <input type="hidden" name="inserir" value="sim">
                                       <input type="hidden" name="valor2" value="<%=valor%>">
                  </div></td>
                  <td bgcolor="#FFFFFF">
                      <p align="center">
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
                      </select></td>
                </tr>
                </table></td>
          </tr>
        </table>
      </form></td>
    </tr>
  </table>
  <%
else
	op = request("op")
	nome = request("nome")
	valor = request("valor")
	
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
	
	sql = "INSERT INTO cartao (nome,valor,op,data,hora,rec) values ('"&nome&"','"&valor&"','"&op&"','"&date&"','"&time&"','N') "		
		'response.write sql	
		set objRs01 = objConn01.execute(sql)	

		%>
		<script>
		//	MM_openBrWindow('gas_imp.asp?coditem=<%=coditem%>','','status=yes,scrollbars=yes,width=100,height=100');
		</script>
		<%
	
%>
					                  <SCRIPT LANGUAGE="JavaScript">
//	alert("<%=sql%>");
	window.location.href = "cartao.asp?op=<%=op%>"
	opener.location.reload(true);
//	self.window.close();
    </SCRIPT>

    <%

end if
call fecha_conexao01
%>
   
</div>
</body>
</html>