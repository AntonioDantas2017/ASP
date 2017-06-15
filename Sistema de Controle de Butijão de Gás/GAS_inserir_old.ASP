<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_css.asp"-->

<html><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

</head>


<%
call abre_conexao01
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02

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

'auxapaga  	= Request ("auxapaga")
'auxcoditem	= Request ("auxcoditem")

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

if auxpago = "S" then
	set objrs01 = objconn01.execute("update gas_itens set pago='S' where coditem="&auxcoditem&"")
end if

if auxpago = "N" then
	set objrs01 = objconn01.execute("update gas_itens set pago='N' where coditem="&auxcoditem&"")
end if
%>

<body onLoad="form2.forma.focus()">	

<div align="center">
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center">&nbsp;&nbsp; <strong>COTA&Ccedil;&Atilde;O</strong></div></td>
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
                      <table width="100%" height="27" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="27">
						                               <div align="center"><a href="javascript:MM_openBrWindow('GAS_FECHAMENTO.asp?codgas=<%=REQUEST("codgas")%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_cfop.gif" alt="Fechamento" width="20" height="20" border="0"></a></div>
                       
					      </td>
					      <td width="20"><div align="center"><strong><a href="JavaScript:window.location.href='gas_inserir.asp?inserir=não&op=<%=op%>'"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="295"></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%
  set objrs01 = objconn01.execute("select * from gas where aberta='sim'")
  if objrs01.eof and codgas = "" then
  %>
  <br>
  <br>
  <table width="100%" border="0" class="linkA">
    <tr>
      <td><div align="center">O SISTEMA DO G&Aacute;S N&Atilde;O EST&Aacute;  ABERTO</div></td>
    </tr>
  </table>
  <br>
  <%
else
	if codgas = "" then
		codgas = objrs01("codgas")
	end if

if inserir="sim" then

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
end if

%>
  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

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
  form.submit();
}
</script>
               
			   
			    <br>
			    <div align="center">
                  <strong> </strong><br>
<form METHOD="get" ACTION="gas_inserir.asp" name="form2">
  <table width="95%" border="0" bgcolor="#CCCCCC">
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
            <td bgcolor="#FFFFFF">
              
                    <div align="left">
                      <script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a função 
		ValidateOrder(form);
	}
}

function tecla_enter2(tecla,form)
{ 
	if(tecla == 13){ 
		MM_openBrWindow('busca_telefone.ASP?telefone='+form.telefone.value,'','status=yes,scrollbars=yes,width=750,height=572');
	}
}

//-->
                   </script>
                      <select name="forma" id="forma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" onBlur="javascript:document.form2.valor.focus()" >
                        
                        <option value="" selected>Forma Pagamento</option>
                        <option value="Dinheiro">Dinheiro</option>
                        <option value="Cartão">Cartão</option>
                        <option value="Cheque">Cheque</option>
                        <option value="Crediário">Crediário</option>				  				  				  
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
              <input type="hidden" name="codgas" value="<%=codgas%>">
              <input type="hidden" name="valor2" value="<%=valor%>">			  
            </div>              </td>
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
            <td bgcolor="#FFFFFF">
              <div align="left">
                <input name="rua" type="text" id="rua"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="30" maxlength="60"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" tabindex="3" >
                </div></td>
            <td bgcolor="#FFFFFF"><div align="center">
              <select name="op" id="op" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" tabindex="4"  >
                <option value="" <%IF request("op") = "" then%> selected <%end if%>>OP</option>
                <option value="1" <%IF request("op") = "1" then%> selected <%end if%>>01  - Mossoró</option>
                <option value="5" <%IF request("op") = "5" then%> selected <%end if%>>05  - Auxiliadora</option>
                <option value="6" <%IF request("op") = "6" then%> selected <%end if%>>06  - Fabiola</option>
                <option value="7" <%IF request("op") = "7" then%> selected <%end if%>>07  - Nenê</option>
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
            <td bgcolor="#FFFFFF"><strong>Nome Entregador </strong></td>
            <td bgcolor="#FFFFFF"><div align="center"><strong>Pago</strong></div></td>
            </tr>
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#FFFFFF"><select name="bairro" id="bairro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" tabindex="6" >
              <option value="" selected>Bairro</option>
              <option value="Campos de São José">Campos de São José</option>
              <option value="Mariana 1">Mariana 1</option>
              <option value="Mariana 2">Mariana 2</option>
              <option value="Santa Cecília 1">Santa Cecília 1</option>
              <option value="Santa Cecília 2">Santa Cecília 2</option>
              <option value="Pouzada">Pouzada</option>
              <option value="Outros">Outros</option>			  			  			  			  			  
            </select></td>
            <td bgcolor="#FFFFFF"><input name="nome" type="text" id="nome"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" value="<%=request("nome")%>" size="30" maxlength="50" tabindex="7" ></td>
            <td bgcolor="#FFFFFF"><div align="center">
              <select name="pago" id="pago" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" tabindex="8"  >
                <option value="N" selected >Não</option>
                <option value="S">Sim</option>			  
              </select>
            </div></td>
            </tr>
      </table></td>
    </tr>
  </table>
  </form>
                  
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><img src="imagem/img_transp.gif" width="1" height="1"></td>
                    </tr>
                  </table>
                  <%
set objrs01 = objConn01.execute("select * from gas_itens WHERE codgas="&codgas&" and (contado='' or contado='não' or contado=null) ORDER BY coditem DESC ")
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#FFDDCC"> 
                            <td width="9%" height="20" bgcolor="#CCCCCC"><div align="center"><strong>HORA</strong></div>
                              <div align="center"></div>
                            <div align="center"></div>                              <div align="center"></div></td>
                            <td width="15%" bgcolor="#CCCCCC"><div align="center"><strong>FORMA  </strong></div></td>
                            <td width="13%" bgcolor="#CCCCCC"><div align="center"><strong>PRE&Ccedil;O</strong></div></td>
                            <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong>OP</strong></div></td>
                            <td width="43%" bgcolor="#CCCCCC"><div align="center"><strong>ENDERE&Ccedil;O/NOME</strong></div></td>
                            <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong>PG</strong></div></td>
                            <td width="5%" bgcolor="#CCCCCC"><div align="center"><strong>ED</strong></div></td>
                            <td width="5%" height="20" bgcolor="#CCCCCC"><div align="center"><strong>AP</strong></div></td>
                          </tr>
<%
if objrs01.eof then 
%>

<tr bgcolor="#FFFFFF"> 
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td  >&nbsp; </td>
</tr>

<%else%>
<%
while not objrs01.eof 
contador 		= contador+1
coditem 		= objrs01("coditem")
hora	 		= objrs01("hora")
forma	 		= objrs01("forma")
valor	 		= objrs01("valor")
op		 		= objrs01("op")
pago	 		= objrs01("pago")

rua		 		= objrs01("rua")
bairro	 		= objrs01("bairro")
numero	 		= objrs01("numero")
nome	 		= objrs01("nome")
telefone 		= objrs01("telefone")
aux_dados		= ""
if (rua<>"") or (numero<>"") then
	aux_dados = "Rua: "&rua&" n°: "&numero&"<BR>"
end if
if (bairro<>"") or (telefone<>"") then
	aux_dados = aux_dados & "Bairro: "&bairro&" Telefone: "&telefone&"<BR>"
end if
if (nome<>"") then
	aux_dados = aux_dados & "Nome: "&nome
end if

if aux_dados = "" then
	aux_dados = "LEVOU"
end if
%>
                          		<%if PAGO = "N" then%>
			<tr valign="middle" bgcolor="#FFC8C8" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFC8C8')"> 
	<%else%>
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
	<%end if%>
                            <td height="20"><div align="left">
                              <div align="center"><%=hora%></div>
                            </div>
                            <div align="center"></div> <div align="left"></div>                              <div></div> </td>
                            <td height="20"><%=forma%></td>
                            <td height="20"><div align="right"><%=formatnumber(valor,2)%></div></td>
                            <td height="20"><div align="center"><%=right("00"&op,2)%></div></td>
                            <td height="20"><%=aux_dados%></td>
                            <td height="20"><table width="90%" border="0" cellspacing="0" cellpadding="0" class=linka>
                              <tr>
                                <td><div align="right">
                                      <%if pago="S" then%>
                                    <a href="gas_inserir.asp?auxpago=N&auxcoditem=<%=coditem%>&codgas=<%=codgas%>&op=<%=request("op")%>" tabindex="10"><img src="imagem/icones_confirma.gif" alt="Gás Pago" width="20" height="20" border="0"></a>
                                    <%else%>
                                    <a href="gas_inserir.asp?auxpago=S&codgas=<%=codgas%>&auxcoditem=<%=coditem%>" tabindex="10"><img src="IMAGEM/icones_cancela.gif" alt="Gás Não Pago" width="20" height="20" border="0"></a>
                                    <%end if%>
                                </div></td>
                                </tr>
                            </table></td>
                            <td height="20"><div align="center"><a href="javascript:MM_openBrWindow('GAS_EDITAR.ASP?codgas=<%=codgas%>&CODITEM=<%=CODITEM%>&op=<%=request("op")%>','','status=yes,scrollbars=yes,width=550,height=550')"  class="linkA5" tabindex="11"><img src="imagem/icones_edita.gif" alt="Editar G&aacute;s" width="20" height="20" border="0"></a></div></td>
                            <td width="5%" height="20">
                                <div align="center"><a href="javascript:MM_openBrWindow('GAS_APAGAR.ASP?codgas=<%=codgas%>&CODITEM=<%=CODITEM%>&op=<%=request("op")%>','','status=yes,scrollbars=yes,width=550,height=550')"  class="linkA5" tabindex="12"><img src="imagem/icones_apaga.gif" alt="Apagar Gás" width="20" height="20" border="0"></a></div>                            </td>
                          </tr>
                          <%
	objrs01.movenext
wend
if isnumeric(totalprodutos) then
	totalprodutos = replace(formatnumber(totalprodutos,2),".","")
end if

%>
<%end if%>                        </table></td>
                    </tr>
                  </table>
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%end if%>
  <br>
    <br>
    <%
		  call fecha_conexao02
          call fecha_conexao01
%>
</div>
</body>
</html>