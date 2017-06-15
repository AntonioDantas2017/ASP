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
<!--#include file="inc_css.asp"-->


<html><head>
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
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="document.form.data.focus()">
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codrec		= request("codrec")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")

codrec    	= request("codrec")
inserir		= request("inserir")

if inserir = "sim" then
	set objRs01 = objConn01.execute("select max(codrec) as maxcodrec from recibos")
	if IsNull(objRs01("maxcodrec")) then
	   codrec = 1
	else
	   codrec = objRs01("maxcodrec") + 1
	end if
	sql = "INSERT INTO recibos (codrec) values ('"&codrec&"') "
	set objRs01 = objConn01.execute(sql)
end if

set objrs01 = objConn01.execute("select * from recibos where codrec="&codrec&" ")
if not objrs01.eof then
	valor		=	objrs01("valor")
	data		=	objrs01("data")
	nome_rec	=	objrs01("nome_rec")
	cpf_rec		=	objrs01("cpf_rec")
	rg_rec		=	objrs01("rg_rec")	
	nome_ent	=	objrs01("nome_ent")
	cpf_ent		=	objrs01("cpf_ent")
	rg_ent		=	objrs01("rg_ent")			
	referente	=	objrs01("referente")			
end if

if inserir = "sim" then
	data		=	date
	nome_rec	=	"Antonio Josivaldo Dantas"
	cpf_rec		=	"480.939.124-87"
	rg_rec		=	"959462"
	nome_ent	=	"Antonio Josivaldo Dantas"
	cpf_ent		=	"480.939.124-87"
	rg_ent		=	"959462"	
end if
call fecha_conexao01
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
                  <td bgcolor="#BBECB1"><div align="center"><strong>RECIBOS</strong></div></td>
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
	    CADASTRO DE RECIBOS 
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
  if (form.data.value == "")
  { 
	  alert("Por favor digite a Data"); form.data.focus(); return; 
  }  
  if (form.valor.value == "")
  { 
	  alert("Por favor digite o valor"); form.valor.focus(); return; 
  }  
 
  if (form.nome_rec.value == "")
  { 
	  alert("Por favor digite o Nome de quem irá recebe"); form.nome_rec.focus(); return; 
  }  
  if (form.cpf_rec.value == "")
  { 
	  alert("Por favor digite  o CPF de quem irá recebe"); form.cpf_rec.focus(); return; 
  }  
  if (form.rg_rec.value == "")
  { 
	  alert("Por favor digite  o RG de quem irá recebe"); form.rg_rec.focus(); return; 
  }  
  
    if (form.nome_ent.value == "")
  { 
	  alert("Por favor digite o Nome de quem irá fornecer"); form.nome_ent.focus(); return; 
  }  
  if (form.cpf_ent.value == "")
  { 
	  alert("Por favor digite  o CPF de quem irá fornecer"); form.cpf_ent.focus(); return; 
  }  
  if (form.rg_ent.value == "")
  { 
	  alert("Por favor digite  o RG de quem irá fornecer"); form.rg_ent.focus(); return; 
  }  
    	    
 if (form.referente.value == "")
  { 
	  alert("Por favor a que é Referente o recibo"); form.referente.focus(); return; 
  }  
  form.submit();

}
</script>
                <br>
                <font color="#FF0000" size="2"><strong>( n&atilde;o utilizar os 
                caracteres especiais : &quot; ' &amp; % ( ) &lt; &gt; )</strong></font><br>
				  <form action="recibos_editar2.asp" method="get" name="form" id="form" >
                    <input name="codrec" type="hidden" value="<%=codrec%>">
                   <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
					<input name="tipo" type="hidden" value="<%=tipo%>">
					<input name="filtro" type="hidden" value="<%=filtro%>">					
					<input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
					<input name="datainicio" type="hidden" value="<%=datainicio%>">
					<input name="datafim" type="hidden" value="<%=datafim%>">	
					<input name="filtro2" type="hidden" value="<%=filtro2%>">						
                  <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong><br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td width="100%"><div align="center"><table width="650" border="0" cellspacing="0" cellpadding="0">
                        <tr valign="top">
                          <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td colspan="2" bgcolor="#CCCCCC"><strong>DATA</strong></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td height="18" colspan="2"><input name="data" type="text" id="data" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=data%>" size="15" maxlength="10"   onKeyPress="return txtBoxFormat('data', '99/99/9999', event);"></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
							 
                              <tr>
                                <td colspan="2" bgcolor="#CCCCCC"><strong>VALOR</strong> </td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td colspan="2" bgcolor="#FFFFFF"><input name="valor" type="text" id="valor" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" value="<%=valor%>" size="15" maxlength="10"></td>
                              </tr>
                            
                               <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr> <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr> <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr> <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td width="50%" bgcolor="#CCCCCC"><div align="center"><strong>DADOS DE QUEM RECEBE </strong></div></td>
                                <td width="50%" bgcolor="#CCCCCC"><div align="center"><strong>DADOS DE QUEM FORNECE </strong></div></td>
                              </tr>
							   <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
							    <tr>
                                <td width="50%" bgcolor="#CCCCCC"><strong>NOME</strong></td>
                                <td width="50%" bgcolor="#CCCCCC"><strong>NOME</strong></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td bgcolor="#FFFFFF"><input name="nome_rec" type="text" id="nome_rec" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=nome_rec%>" size="50" maxlength="60"></td>
                                <td bgcolor="#FFFFFF"><input name="nome_ent" type="text" id="nome_ent" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=nome_ent%>" size="50" maxlength="60"></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
							   <tr>
                                <td width="50%" bgcolor="#CCCCCC"><strong>RG</strong></td>
                                <td width="50%" bgcolor="#CCCCCC"><strong>RG </strong></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td bgcolor="#FFFFFF"><input name="rg_rec" type="text" id="rg_rec" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=rg_rec%>" size="35" maxlength="30"></td>
                                <td bgcolor="#FFFFFF"><input name="rg_ent" type="text" id="rg_ent" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=rg_ent%>" size="35" maxlength="30"></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
							   <tr>
                                <td width="50%" bgcolor="#CCCCCC"><strong>CPF</strong></td>
                                <td width="50%" bgcolor="#CCCCCC"><strong>CPF</strong></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td bgcolor="#FFFFFF"><input name="cpf_rec" type="text" id="cpf_rec" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=cpf_rec%>" size="17" maxlength="14"  onKeyPress="return txtBoxFormat('cpf_rec', '999.999.999-99', event);"></td>
                                <td bgcolor="#FFFFFF"><input name="cpf_ent" type="text" id="cpf_ent" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=cpf_ent%>" size="17" maxlength="14"  onKeyPress="return txtBoxFormat('cpf_ent', '999.999.999-99', event);"></td>
                              </tr>
							   <tr>
                                <td colspan="2" bgcolor="#CCCCCC"><strong>REFERENTE A </strong></td>
                                </tr>
                              <tr>
                                <td colspan="2"><input name="referente" type="text" id="referente" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=referente%>" size="90"></td>
                              </tr>
                              
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                              <tr>
                                <td colspan="2"><img src="imagem/img_transp.gif" width="1" height="2"></td>
                              </tr>
                          </table></td>
                        </tr>
                      </table>
                      </div></td>
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