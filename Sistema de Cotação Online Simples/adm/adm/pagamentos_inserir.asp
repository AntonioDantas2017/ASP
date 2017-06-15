<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
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
<%
call abre_conexao01
%>
<!--#include file="inc_css.asp"--><html><head>
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
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>

</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" >
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codtel		= request("codtel")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")

%>
<div align="center"> 

  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="115"> <table width="115" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>TELEFONE</strong></div></td>
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
                      <table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="30"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"><strong></strong></div></td>
                          <td width="120"> <div align="center"></div></td>
                          <td width="130"> <div align="center"></div></td>
                          <td width="130"> <div align="right"></div></td>
                          <td width="420"><div align="right"></div>
                            <div align="center"></div>
                            <div align="center"><strong> </strong></div>
                            <div align="center"></div>
                            <div align="right"> </div></td>
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
  <table width="700" border="0" cellspacing="0" cellpadding="0" class=linka>
    <tr> 
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>
	    CADASTRO DE TELEFONE
	  </strong></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td bgcolor="#FFFFFF"><div align="center">
                <script language="JavaScript">
function ValidateOrder(form)
{
  if (form.valor.value == "")
  { 
	  alert("Por favor digite o Valor"); form.valor.focus(); return; 
  }  
    if (form.data_venc.value == "")
  { 
	  alert("Por favor digite a Data de Vencimento"); form.data_venc.focus(); return; 
  }  
    if (form.codcon.value == "")
  { 
	  alert("Por favor escolha o Contrato"); form.codcon.focus(); return; 
  }  

  form.submit();

}
</script>
                <br>
                  <form method="GET" ACTION="pagamentos_inserir2.asp" name="form" id="form">
                    <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                    <input name="filtro" type="hidden" value="<%=filtro%>">					
					<input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
					<input name="datainicio" type="hidden" value="<%=datainicio%>">
					<input name="datafim" type="hidden" value="<%=datafim%>">	
					<input name="filtro2" type="hidden" value="<%=filtro2%>">																				
					
                  <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong><br>
                  <br>
                  <font color="#FF0000" size="2"><strong>( n&atilde;o utilizar 
                  os caracteres especiais : &quot; ' &amp;  ( )  )</strong></font><br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td width="100%">
					    
						<table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						  

                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>							  
						  <tr> 
                            <td bgcolor="#CCCCCC"><strong>DATA VENCIMENTO </strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><input name="data_venc" type="text" id="data_venc" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="12" maxlength="10"  onKeyPress="return txtBoxFormat('data_venc', '99/99/9999', event);"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          
						  
						  <tr>
						    <td bgcolor="#CCCCCC"><strong>VALOR</strong></td>
					      </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>							  
						  <tr>
						    <td bgcolor="#FFFFFF"><input name="valor" type="text" id="valor" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="15" maxlength="10"  onKeypress="if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;"   onKeyDown="Formata(this,20,event,2)"></td>
					      </tr>

                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>							  
						  <tr>
						    <td bgcolor="#CCCCCC"><strong>CONTRATO</strong></td>
					      </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>	
						  <tr>
						    <td><select name="codcon" id="codcon"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" >
                              <option value="">Contrato</option>
                              <%
set objrs01 = objconn01.execute("select codcon from contratos order by impressao ")
while not objrs01.eof
%>
                              <option value="<%response.write objrs01("codcon")%>">
                              <%response.write objrs01("codcon")%>
                              </option>
                              <%
	objrs01.movenext
wend
%>
                            </select></td>
					      </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>	
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>							   
                        </table></td>
                    </tr>
				</table>
                  <br>
                  <br>
<input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                  <br>
                </form>                
                
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
</div>
</body>
</html>
 <%
'call fecha_conexao01
%>