<!--#include file="inc_login.asp"-->
<!--#include file="inc_css.asp"-->

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
              <script language="JavaScript">
/***
* Descrição.: formata um campo do formulário de
* acordo com a máscara informada...
* Parâmetros: - objForm (o Objeto Form)
* - strField (string contendo o nome
* do textbox)
* - sMask (mascara que define o
* formato que o dado será apresentado,
* usando o algarismo "9" para
* definir números e o símbolo "!" para
* qualquer caracter...
* - evtKeyPress (evento)
* Uso.......: <input type="textbox"
* name="xxx".....
* onkeypress="return txtBoxFormat(document.rcfDownload, 'str_cep', '99999-999', event);">
* Observação: As máscaras podem ser representadas como os exemplos abaixo:
* CEP -> 99.999-999
* CPF -> 999.999.999-99
* CNPJ -> 99.999.999/9999-99
* Data -> 99/99/9999
* Tel Resid -> (99) 999-9999
* Tel Cel -> (99) 9999-9999
* Processo -> 99.999999999/999-99
* C/C -> 999999-!
* E por aí vai...
***/

function txtBoxFormat(strField, sMask, evtKeyPress) {
    var i, nCount, sValue, fldLen, mskLen,bolMask, sCod, nTecla;

    if(window.event) { // Internet Explorer
      nTecla = evtKeyPress.keyCode; }
    else if(evtKeyPress.which) { // Nestcape / firefox
      nTecla = evtKeyPress.which;
    }
//se for backspace não faz nada
if (nTecla != 8){
    sValue = document.getElementById(strField).value;
// alert(sValue);

    // Limpa todos os caracteres de formatação que
    // já estiverem no campo.
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( "-", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ".", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( ":", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "/", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( "(", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( ")", "" );
    sValue = sValue.toString().replace( " ", "" );
    sValue = sValue.toString().replace( " ", "" );
    fldLen = sValue.length;
    mskLen = sMask.length;

    i = 0;
    nCount = 0;
    sCod = "";
    mskLen = fldLen;

    while (i <= mskLen) {
      bolMask = ((sMask.charAt(i) == "-") || (sMask.charAt(i) == ".")  || (sMask.charAt(i) == ":") || (sMask.charAt(i) 

== "/"))
      bolMask = bolMask || ((sMask.charAt(i) == "(") || (sMask.charAt(i) == ")") || (sMask.charAt(i) == " "))

      if (bolMask) {
        sCod += sMask.charAt(i);
        mskLen++; }
      else {
        sCod += sValue.charAt(nCount);
        nCount++;
      }

      i++;
    }

    document.getElementById(strField).value = sCod;

    if (nTecla != 8) { // backspace
      if (sMask.charAt(i-1) == "9") { // apenas números...
        return ((nTecla > 47) && (nTecla < 58)); } // números de 0 a 9
      else { // qualquer caracter...
        return true;
      } }
    else {
      return true;
    }
}//fim do if que verifica se é backspace
}
//Fim da Função Máscaras Gerais
</script>

<script language="JavaScript" type="text/JavaScript">
<!--

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<script>


function Limpar(valor, validos) { 
// retira caracteres invalidos da string 
var result = ""; 
var aux; 
for (var i=0; i < valor.length; i++) { 
aux = validos.indexOf(valor.substring(i, i+1)); 
if (aux>=0) { 
result += aux; 
} 
} 
return result; 
} 
//Formata número tipo moeda usando o evento onKeyDown 

function Formata(campo,tammax,teclapres,decimal) { 
var tecla = teclapres.keyCode; 
vr = Limpar(campo.value,"0123456789"); 
tam = vr.length; 
dec=decimal 

if (tam < tammax && tecla != 8){ tam = vr.length + 1 ; } 

if (tecla == 8 ) 
{ tam = tam - 1 ; } 

if ( tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105 ) 
{ 

if ( tam <= dec ) 
{ campo.value = vr ; } 

if ( (tam > dec) && (tam <= 5) ){ 
campo.value = vr.substr( 0, tam - 2 ) + "," + vr.substr( tam - dec, tam ) ; } 
if ( (tam >= 6) && (tam <= 8) ){ 
campo.value = vr.substr( 0, tam - 5 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; 
} 
if ( (tam >= 9) && (tam <= 11) ){ 
campo.value = vr.substr( 0, tam - 8 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; } 
if ( (tam >= 12) && (tam <= 14) ){ 
campo.value = vr.substr( 0, tam - 11 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - dec, tam ) ; } 
if ( (tam >= 15) && (tam <= 17) ){ 
campo.value = vr.substr( 0, tam - 14 ) + "." + vr.substr( tam - 14, 3 ) + "." + vr.substr( tam - 11, 3 ) + "." + vr.substr( tam - 8, 3 ) + "." + vr.substr( tam - 5, 3 ) + "," + vr.substr( tam - 2, tam ) ;} 
} 

} 

 </script>
</head>


<%
call abre_conexao01
call abre_conexao02

inserir	  	= request("inserir")
codche	  	= Request ("codche")
fornecedor = request("fornecedor")

cliente   	= request("cliente")
codcli 		= request("codcli")	
valor 		= request("valor")	
data 		= request("data")	

auxapaga  	= Request ("auxapaga")
auxcont	= Request ("auxcont")

pesquisa = request("pesquisa")
pesquisa2 = request("pesquisa2")
%>

<body onLoad="form2.codcli.focus()">	

<div align="center"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center">&nbsp;&nbsp;<strong>CHEQUE</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="3">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="28" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" height="28" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td><strong>FORNECEDOR:</strong>
                          <%
						  set objrs02= objConn02.execute("select vendedor from fornecedores WHERE codfor="&fornecedor&"")

						  if not objrs02.eof then
							response.Write ucase(objrs02("vendedor"))
						else
							response.Write "Erro! O fornecedor não Existe"
						end if
						  %></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  
  <table width="10%" height="10" border="0">
    <tr>
      <td></td>
    </tr>
  </table>
  <%
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
	   preparaPalavra = replace(preparaPalavra,"ç","c")	   	   	   
	   preparaPalavra = preparaPalavra
	end function

if inserir = "sim" then 
	sql = "update cheques set codfor="&fornecedor&" where codche = "&codche&" "
	'response.write sql	
	set objRs01 = objConn01.execute(sql)
end if

if inserir="não" then

if cliente <> "" then
	cliente = ucase(preparaPalavra(lcase(cliente)))
	set objrs01 = objconn01.execute("select * from clientes where nome = '"&cliente&"'")
	codcli = request("codcli")		
	if codcli <> "" then
		if objrs01.eof then		
			sql = "INSERT INTO clientes (codcli,nome) values ("&codcli&",'"&cliente&"') "
			'response.write sql	
			set objRs01 = objConn01.execute(sql)
		else
			codcli = objrs01("codcli")
		end if
	end if
end if		

if codcli <> "" then
	valor = replace(valor,".","")
	valor = replace(valor,",",".")
	set objrs01 = objconn01.execute("select * from aux_cheques where codche = "&codche&" and codcli="&codcli&" and  valor="&valor&" and  data=#"&data&"#")
	valor = replace(valor,".",",")	
	if not objrs01.eof then
%>
  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
    <tr> 
      <td bgcolor="#FF9999"><div align="center"><strong><font size="2">ESSE CHEQUE J&Aacute; FOI INSERIDO </font></strong></div></td>
    </tr>
  </table>
  <%
	else
		set objrs01 = objconn01.execute("select * from clientes where codcli = "&codcli&"")
	 	if not objrs01.eof then
			set objRs01 = objConn01.execute("select max(cont) as maxcont from aux_cheques")
			if IsNull(objRs01("maxcont")) then
			   cont = 100000
			else
			   cont = objRs01("maxcont") + 1
			end if
			
			sql = "INSERT INTO aux_cheques (cont,codche,codcli,valor,data) values ("&cont&",'"&codche&"','"&codcli&"','"&valor&"','"&data&"') "
			'response.write sql	
			set objRs01 = objConn01.execute(sql)	
		else
		AXCODCLI = codcli
		AXVALOR = valor
		AXDATA = data
%>
  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
    <tr> 
      <td bgcolor="#FF9999"><div align="center"><strong><font size="2">ESSE CLIENTE NÃO EXISTE FAVOR DIGITAR O CÓDIGO E O NOME DO CLIENTE</font></strong></div></td>
    </tr>
  </table>
  <%
		end if
	END IF
END IF

if auxapaga = "sim" then
sql = "delete from aux_cheques where cont = "&auxcont&" "
'response.write sql	
set objRs01 = objConn01.execute(sql)
%>
	  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
		<tr> 
		  <td bgcolor="#FF9999"><div align="center"><strong><font size="2">O CHEQUE FOI APAGADO </font></strong></div></td>
		</tr>
  </table>
	  <%
end if
end if
%>
	  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

<script language="JavaScript">

function ValidateOrder(form)
{
   if (form.codcli.value == "")
  { 
	  alert("Por favor digite o Código do cliente"); form.codcli.focus(); return; 
  } 
     if (form.data.value == "")
  { 
	  alert("Por favor digite a data do Cheque"); form.data.focus(); return; 
  } 
     if (form.valor.value == "")
  { 
	  alert("Por favor digite o Valor do Cheque"); form.valor.focus(); return; 
  } 
  form.submit();
}
</script>
               
			   
			    <br>
			    <div align="center">
                  <strong> </strong><br>
<form METHOD="get" ACTION="cheques_inserir.asp" name="form2">
  <table width="95%" border="0" bgcolor="#CCCCCC">
    <tr>
      <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
          <tr bgcolor="#FFFFFF">
            <td width="13%"><div align="center"><strong>Codcli</strong></div></td>
            <td width="50%"><div align="center"><strong>Cliente</strong></div></td>
            <td width="18%"><div align="center"><strong>Valor</strong></div></td>
            <td width="12%"><div align="center"><strong>Data</strong></div></td>
            <td width="7%" rowspan="2"><div align="center">
                <input name="x" type="button" id="x" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onBlur="javascript:document.form2.codcli.focus();">
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td height="21"><div align="center">
              <input name="codcli" type="text" id="codcli"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR: #FFC4C4"  onKeyPress="return txtBoxFormat('data', '99999', event);" value="<%=AXCODCLI%>" size="10" maxlength="8" >
            </div></td>
            <td height="21"><div align="center">
              <input name="cliente" type="text" id="cliente"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR: #FFC4C4" size="50" maxlength="50"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" >
            </div></td>
            <td><div align="center">
			<script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a função 
		ValidateOrder(form);
	}
}

//-->
</script>	
			R$
			<input type="hidden" name="inserir" value="não">
			<input type="hidden" name="fornecedor" value="<%=fornecedor%>">
			<input type="hidden" name="codche" value="<%=codche%>">			
			                <input name="valor" type="text" id="valor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4"  onKeypress="if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;"   onKeyDown="Formata(this,20,event,2)" value="<%=AXVALOR%>" size="10" maxlength="10">
            </div></td>
            <td><div align="center">
              <input name="data" type="text" id="data"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" onBlur="javascript:document.form2.x.focus();"  onKeyPress="return txtBoxFormat('data', '99/99/9999', event);" value="<%=AXDATA%>" size="10" maxlength="10">
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
set objrs01 = objConn01.execute("select * from aux_cheques WHERE codche="&codche&" ORDER BY cont ")
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#CCCCCC"> 
                            <td width="10%" height="20"><div align="center"><strong>CODCLI 
                                </strong></div>
                              <div align="center"></div></td>
                            <td width="54%"><div align="center"><strong>CLIENTE</strong></div></td>
                            <td width="15%"><div align="center"><strong>VALOR </strong></div></td>
                            <td width="16%"><div align="center"><strong>DATA</strong></div></td>
                            <td width="5%" height="20"><div align="center"><strong>AP</strong></div></td>
                          </tr>
<%
if objrs01.eof then 
%>

<tr bgcolor="#FFFFFF"> 
<td >&nbsp;</td>
<td >&nbsp;</td>
<td >&nbsp;</td>
<td >&nbsp;</td>
<td  >&nbsp; </td>
</tr>

<%else%>
<%
while not objrs01.eof 
contador 		= contador+1
cont 		= objrs01("cont")
codcli 		= objrs01("codcli")
valor   	= formatnumber(objrs01("valor"),2)
data		= (objrs01("data"))
%>
<script>
//	alert("<%=data%>");
</script>
<%
data2		= split(data,"/")
data		= right("0"&data2(0),2)&"/"&right("0"&data2(1),2)&"/"&data2(2)

set objrs02= objConn02.execute("select * from clientes WHERE codcli="&codcli&"")
if not objrs02.eof then
	cliente  = ucase(objrs02("nome"))
else
	cliente = "Erro! Favor Apagar Registro"
end if

achou = ""
if pesquisa2 <> "" and pesquisa <> "" then

	if pesquisa2  = "nome" then
		pesquisa = ucase(preparapalavra(pesquisa))
		set objrs02= objConn02.execute("select * from clientes WHERE nome Like '%"&(pesquisa)&"%' and codcli="&codcli&"")
		if not objrs02.eof then
			achou = "cliente"
		end if
	else
		if pesquisa2  = "aux_cheques.data" then
			if isdate(pesquisa) then
				data2		= split(pesquisa,"/")
				pesquisa	= right("0"&data2(0),2)&"/"&right("0"&data2(1),2)&"/"&data2(2)
				if pesquisa = data then
					achou = "data"
				end if
			end if
		else
			if pesquisa2  = "valor" then
				if isnumeric(pesquisa) then
					pesquisa = formatnumber(pesquisa,2)
		
					if valor = pesquisa then
						achou = "valor"
					end if
				end if
			else
				if pesquisa2  = "codcli" then
					if isnumeric(pesquisa) then
						if right("000000"&codcli,6) = right("000000"&pesquisa,6) then
							achou = "codcli"
						end if
					end if
				end if
			end if
		end if
	end if
end if
	
if achou <> "" then
%>
					
                          <tr bgcolor="#FFFFBB"> 
<%else%>
                <tr bgcolor="#FFFFFF"> 
			<%end if%>					
							
                            <td height="20" <%if achou="codcli" then%> bgcolor="#FFC8C8"<%end if%>><div align="left">
                              <div align="center"><%=right("00000"&codcli,5)%></div>
                            </div>
                            <div align="center"></div> <div align="left"></div>                              <div></div> </td>
                            <td height="20"  <%if achou="cliente" then%> bgcolor="#FFC8C8"<%end if%>><div align="left"><%
															response.Write ucase(cliente)
														
							%></div>
                            <div align="center"></div></td>
                            <td height="20"  <%if achou="valor" then%> bgcolor="#FFC8C8"<%end if%>>
                              <div align="right"><%=valor%></div></td>
                            <td height="20"  <%if achou="data" then%> bgcolor="#FFC8C8"<%end if%>>
                              <div align="center"><%
							  data2 = split(data,"/")
							  response.Write right("0"&data2(0),2) & "/" & right("0"&data2(1),2) & "/" & data2(2)
							  %></div>                            </td>
                            <td width="5%" height="20"><div align="center"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <form name="form4" method="get" action="cheques_inserir.asp">
<input name="auxcont" type="hidden" value="<%=cont%>">
<input name="auxapaga" type="hidden" value="sim">
<input name="fornecedor" type="hidden" value="<%=fornecedor%>">

<input name="codche" type="hidden" value="<%=codche%>">
<input name="inserir" type="hidden" value="não">


                                    <tr> 
                                      <td><div align="center"> 
                                          <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja apagar este cheque?'))
	}
	</script>
                                          <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga este item" WIDTH="20" HEIGHT="20">
                                        </div></td>
                                    </tr>
                                  </form>
                                </table>
                                
                              </div></td>
                          </tr>
                          <%
	objrs01.movenext
wend
if isnumeric(totalclientes) then
	totalclientes = replace(formatnumber(totalclientes,2),".","")
end if

%>
<%end if%>                        </table></td>
                    </tr>
                  </table>
 
                  <br>
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <p><br>
    <br>
    <%
		  call fecha_conexao02
          call fecha_conexao01
%>
  </p>
 
</div>
</body>
</html>