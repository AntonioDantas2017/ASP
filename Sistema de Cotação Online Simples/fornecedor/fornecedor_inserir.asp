<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_css.asp"-->

<html><head>
<script>

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

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

function setPointer(theRow, thePointerColor)
{
    if (thePointerColor == '' || typeof(theRow.style) == 'undefined') {
        return false;
    }
    if (typeof(document.getElementsByTagName) != 'undefined') {
        var theCells = theRow.getElementsByTagName('td');
    }
    else if (typeof(theRow.cells) != 'undefined') {
        var theCells = theRow.cells;
    }
    else {
        return false;
    }

    var rowCellsCnt  = theCells.length;
    for (var c = 0; c < rowCellsCnt; c++) {
        theCells[c].style.backgroundColor = thePointerColor;
    }

    return true;
} // end of the 'setPointer()' function

</script>

</head>


<%
if session("codforn") = "" then
%>
<script language="javascript" type="text/javascript">
	alert("Você não está logado favor pedir para algum administrador digitar a senha")
	window.location.href = "sair.asp"
</script>
<%
end if
call abre_conexao01
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02

inserir	  	= request("inserir")
codforn   	= session("codforn")
codprc 		= request("codprc")	
cont		= request("cont")	
preco		= TRIM(request("preco"))
obs			= TRIM(request("obs"))

%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form2.preco.focus()">	

<div align="center"> 
 
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>FORNECEDORES</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="3">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="27" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" height="27" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="92%"><strong>FORNECEDOR:</strong>
                            <%
						  set objrs02= objConn02.execute("select vendedor from fornecedores WHERE codfor="&codforn&"")

						  if not objrs02.eof then
							response.Write ucase(objrs02("vendedor"))
						else
							response.Write "Erro! O fornecedor n&atilde;o Existe"
						end if
						  %></td>
                          <td width="8%"><div align="center"><a href="javascript:MM_openBrWindow('fornecedores_help.asp','','status=yes,scrollbars=yes,width=565,height=465')"  class="linkA5"><img src="imagem/icones_base.gif" alt="Ajuda" width="20" height="20" border="0"></a></div></td>
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
  <%
  set objrs01 = objconn01.execute("select * from cotacao where aberta='sim'")
  if objrs01.eof then
  %>
  <br>
  <br>
  <table width="100%" border="0" class="linkA">
    <tr>
      <td><div align="center">N&Atilde;O EXISTE COTA&Ccedil;&Atilde;O ABERTA </div></td>
    </tr>
  </table>
  <br>
  <%
else
codcot = objrs01("codcot")
	

if inserir = "sim" then

preco	= replace(preco,".",",")
if not isnumeric(preco) then
	preco = 0
end if
preco	= formatnumber(preco,2)

'set objrs01 = objconn01.execute("select fora from fornecedores where codfor="&codforn&"")
'if objrs01("fora") = "sim" then
'	preco = preco + (preco*0.06)
'end if

set objrs01 = objconn01.execute("update prod_cot_forn set preco = '"&preco&"',obs='"&obs&"' where cont = "&cont&"")
end if
codprc = ""
%>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

<script language="JavaScript">

function ValidateOrder(form)
{

  form.submit();
}
</script>
               
			   
			    <br>
			    <div align="center">
                  
				  <br>
				  <%
'				  set objrs01 = objconn01.execute("select * from produtos_cot where not exist (select codfor from prod_cot_forn where codfor = "&codforn&" and prod_cot_forn.codprc=codprc)")
				  'set objrs01 = objconn01.execute("select * from produtos_cot where not exist (SELECT prod_cot_forn.codfor FROM prod_cot_forn INNER JOIN produtos_cot ON prod_cot_forn.codprc = produtos_cot.codprc where codfor = "&codforn&")")
				  achou = "não"
				  set objrs01 = objconn01.execute("select * from produtos_cot WHERE CODCOT="&codcot&"")
				  while not objrs01.eof and achou = "não"
						set objrs02 = objConn02.execute("select codprc from prod_cot_forn WHERE codprc="&objrs01("codprc")&" and codfor="&codforn&"")
						if not objrs02.eof then
							objrs01.movenext
						else
							achou = "sim"
						end if
				  wend		
										  
if not objrs01.eof then
	set objrs02 = objConn02.execute("select max(cont) as maxcont from prod_cot_forn")
	if IsNull(objrs02("maxcont")) then
	   cont = "000001"
	else
	   cont = Right( "000000" & (Cint(objrs02("maxcont")) + 1) ,6)
	end if
	set objrs02 = objconn02.execute("insert into prod_cot_forn (codfor,codprc,preco,cont) values("&codforn&","&objrs01("codprc")&",'0,00','"&cont&"') ")	
	set objrs02 = objconn02.execute("select * from produtos_cot where codprc="&objrs01("codprc"))		
	if not objrs02.eof then
		set objrs02 = objconn02.execute("select * from produtos where codpro="&objrs01("codpro"))	
		if not objrs02.eof then		
			aux_prod = objrs01("codpro")
			produto = objrs02("produto")
		else
			produto = "-"
		end if
	end if

aux_cont_ap = cont

				  %>
				  <form METHOD="get" ACTION="fornecedor_inserir.asp" name="form2">
				    <table width="95%" border="0" bgcolor="#CCCCCC">
    <tr>
      <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
          <tr bgcolor="#FFFFFF">
            <td colspan="2"><div align="center"><strong>Produto</strong></div></td>
            <td width="18%"><div align="center"><strong>PRE&Ccedil;O</strong></div></td>
            <td width="8%" rowspan="2"><div align="center">
                <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
            </div></td>
          </tr>
          
          <tr bgcolor="#FFFFFF">
            <td colspan="2"><div align="left"><%=produto%></div>      </td>
            <td><div align="center">
			<input type="hidden" name="inserir" value="sim">
			<input type="hidden" name="cont" value="<%=cont%>">			
			               <!-- <input name="preco" type="text" id="preco"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" size="10" maxlength="7"  onKeyDown="Formata(this,20,event,2)">-->
						    <input name="preco" type="text" id="preco"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" size="10" maxlength="7"  onKeypress="tecla_enter(+event.keyCode,this.form);">
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="19%"><div align="center"><strong>Observa&ccedil;&atilde;o:</strong></div></td>
            <td colspan="2">&nbsp;
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
              <input name="obs" type="text" id="obs"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#EAEAEA" size="50" maxlength="50"  onKeypress="tecla_enter(+event.keyCode,this.form);"></td>
            <td width="8%"><div align="center"><a href="javascript:MM_openBrWindow('cotacao_busca.asp?CODPROD1=<%=cont%>&codforn=<%=codforn%>&codprc=<%=codprc%>&codcot=<%=codcot%>','','status=yes,scrollbars=yes,width=400,height=400');"  class="linkA5"><img src="imagem/icones_ver.gif" alt="Buscar Produto" width="15" height="15" border="0"></a></div></td>
          </tr>
      </table></td>
    </tr>
  </table>
  </form>
                
				<%else
				acabou = "sim"
				%>
				
				  <table width="95%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                    <tr> 
                      <td><div align="center"><strong>ACABOU OS PRODUTOS DA COTA&Ccedil;&Atilde;O </strong></div></td>
                    </tr>
                  </table>
				  <br>
				  <%end if%>
				  
               
                  <%
if acabou <> "sim" then
	if aux_prod <> "" then
		sql = "select top 15 * from prod_cot_forn2 WHERE codfor="&codforn&" and preco >= 0 and codcot = "&codcot&" and codpro <> "&aux_prod&" ORDER BY cont desc "
	else
		sql = "select top 15 * from prod_cot_forn2 WHERE codfor="&codforn&" and preco >= 0 and codcot = "&codcot&" ORDER BY cont desc "
	end if
else
	sql = "select * from prod_cot_forn2 WHERE codfor="&codforn&" and preco >= 0 and codcot = "&codcot&" ORDER BY cont "

end if
'response.Write sql
set objrs01 = objConn01.execute(sql)
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#CCCCCC"> 
                            <td width="80%" height="20"><div align="right"><strong>PRODUTO&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%IF acabou = "sim" then%><a href="javascript:MM_openBrWindow('cotacao_busca.asp?codforn=<%=codforn%>&codprc=<%=codprc%>&codcot=<%=codcot%>','','status=yes,scrollbars=yes,width=400,height=400');"  class="linkA5"><img src="imagem/icones_ver.gif" alt="Buscar Produto" width="15" height="15" border="0"></a><%else%>&nbsp;<%end if%></strong></div>
                              <div align="center"></div>
                            <div align="center"></div>                              <div align="center"></div></td>
                            <td width="14%"><div align="center">
                              <div align="center"><strong>PREÇO</strong></div>
                            </div></td>
                            <td width="6%" height="20"><div align="center"><strong>ALT</strong></div></td>
                          </tr>
<%
if objrs01.eof then 
%>

<tr bgcolor="#FFFFFF"> 
<td >&nbsp;</td>
<td >&nbsp;</td>
<td  >&nbsp; </td>
</tr>

<%else%>
<%
contador = 0
while not objrs01.eof 
contador 		= contador+1
cont 		= objrs01("cont")
codprc   	= objrs01("codprc")
preco   	= objrs01("preco")

set objrs02= objConn02.execute("select * from produtos_cot WHERE codprc="&codprc&"")
if not objrs02.eof then
	codpro  = objrs02("codpro")
else
	produto = "Erro! Produto Não Encontrado. Favor colocar valor R$ 0,00"
end if

set objrs02= objConn02.execute("select * from produtos WHERE codpro="&codpro&"")
if not objrs02.eof then
	produto  = objrs02("produto")
else
	produto = "Erro! Produto Não Encontrado. Favor colocar valor R$ 0,00"
end if

if not isnumeric(preco) then
	preco = 0
end if

%>

                           <%if preco = 0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
                            <td height="20"><div align="left">
                              <div align="left"><%=produto%></div>
                            </div>
                                           <div></div> </td>
                            <td height="20"><div align="left">
                              <div align="right"><%=formatnumber(preco)%></div>
                            </div></td>
                            <td width="6%" height="20"><div align="center"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr> 
                                      <td><div align="center"> 
                                    <%'if acabou="sim" then%>   
									
					<!--				  <a href="javascript:MM_openBrWindow('alterar_preco.asp?cont=<%=cont%>&codforn=<%=codforn%>&acabou=<%=acabou%>','','status=yes,scrollbars=yes,width=800,height=572')">  <img SRC="imagem/icones_edita.gif" alt="Editar produto" WIDTH="20" HEIGHT="20" border="0"></a>
						-->
						
									  <a href="javascript:MM_openBrWindow('cotacao_busca.asp?CODPROD1=<%=aux_cont_ap%>&CODPROD2=<%=cont%>&PRODUTO=<%=produto%>&AUX=1','','status=yes,scrollbars=yes,width=800,height=572')">  <img SRC="imagem/icones_edita.gif" alt="Editar produto" WIDTH="20" HEIGHT="20" border="0"></a>
									  
									  
									  <% 'else response.Write "-" 'end if%>
                                        </div></td>
                                    </tr>                        
                                </table>
                                
                            </div></td>
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
                  <br>  
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%end if%>
  <%
		  call fecha_conexao02
          call fecha_conexao01
%>
  
 
    <br>
   <a href="fornecedores_imp.asp?codcot=<%=codcot%>"><img src="imagem/botoes_fechar.gif" width="75" height="19" border="0"></a></div>
</body>
</html>