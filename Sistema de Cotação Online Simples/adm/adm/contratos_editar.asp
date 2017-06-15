<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
%>
<!--#include file="inc_css.asp"-->


<html>
<head>
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

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0"  onLoad="form.locatario.focus()">
<%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codcli		= request("codcli")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")
codcon    	= request("codcon")
acao		= request("acao")

if acao = "renovar" or acao = "inserir" then
	set objRs01 = objConn01.execute("select max(codcon) as maxcodcon from contratos")
	if IsNull(objRs01("maxcodcon")) then
	   codcon = 100000
	else
	   codcon = objRs01("maxcodcon") + 1
	end if
	sql = "INSERT INTO contratos (codcon) values ('"&codcon&"') "
	set objRs01 = objConn01.execute(sql)
end if

set objrs01 = objConn01.execute("select * from contratos where codcon="&codcon&" ")
if not objrs01.eof then
	auxcodloc 	= objrs01 ("codloc")
	auxcodimo 	= objrs01 ("codimo")
	periodo 	= objrs01 ("periodo")
	dataini 	= objrs01 ("dataini")
	datafim 	= objrs01 ("datafim")
	aluguel 	= objrs01 ("aluguel")
	vencimento 	= objrs01 ("vencimento")
	multa 		= objrs01 ("multa")
	destinado 	= objrs01 ("destinado")
	reajuste 	= objrs01 ("reajuste")
	impressao 	= objrs01 ("impressao")
	testemunha1 = objrs01 ("testemunha1")
	testemunha2 = objrs01 ("testemunha2")
end if

if acao = "renovar" or acao = "inserir" then
	multa 		= "de um aluguel"
	impressao 	= date
end if


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
                  <td bgcolor="#BBECB1"><div align="center"><strong> CONTRATOS</strong></div></td>
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
          <% if acao="editar" then %>
          ALTERA&Ccedil;&Atilde;O DE
          <% else %>
			  <% if acao="renovar" then %>
          RENOVA&Ccedil;&Atilde;O  DE<% else %> INSERIR <% END IF %><% END IF %> CONTRATO </strong></div></td>
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
   if (form.locatario.value == "")
  { 
	  alert("Por favor escolha o Locatário"); form.locatario.focus(); return; 
  }
  if (form.imovel.value == "")
  { 
	  alert("Por favor escolha o Imóvel"); form.imovel.focus(); return; 
  }
  if (form.periodo.value == "")
  { 
	  alert("Por favor digite o Período"); form.periodo.focus(); return; 
  }

  if (form.dataini.value == "")
  { 
	  alert("Por favor digite a Data Início"); form.dataini.focus(); return; 
  }
  
   if (form.datafim.value == "")
  { 
	  alert("Por favor digite a Data Final"); form.datafim.focus(); return; 
  }
  
   if (form.aluguel.value == "")
  { 
	  alert("Por favor digite o valor do aluguel"); form.aluguel.focus(); return; 
  }
  
     if (form.vencimento.value == "")
  { 
	  alert("Por favor digite o Vencimento"); form.vencimento.focus(); return; 
  }
  
     if (form.multa.value == "")
  { 
	  alert("Por favor digite o valor da multa"); form.multa.focus(); return; 
  }
   
     if (form.reajuste.value == "")
  { 
	  alert("Por favor digite o dia do reajuste"); form.reajuste.focus(); return; 
  }
  
     if (form.impressao.value == "")
  { 
	  alert("Por favor digite o dia da impressao"); form.impressao.focus(); return; 
  }
  
  form.submit();
}
</script>
                <br>
                <font color="#FF0000" size="2"><strong>( n&atilde;o utilizar os 
                caracteres especiais : &quot; ' &amp; % ( ) &lt; &gt; )</strong></font><br>
				  <form action="contratos_editar2.asp" method="get" name="form" id="form" >
				    <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="ordem" type="hidden" value="<%=ordem%>">
                    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                    <input name="filtro" type="hidden" value="<%=filtro%>">
                    <input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
                    <input name="datainicio" type="hidden" value="<%=datainicio%>">
                    <input name="datafim2" type="hidden" value="<%=datafim%>">
                    <input name="filtro2" type="hidden" value="<%=filtro2%>">
					<input name="codcon" type="hidden" value="<%=codcon%>">
					<input name="acao" type="hidden" value="<%=acao%>">
					
                    <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO OBRIGAT&Oacute;RIO</strong><br>
                  <br>
                  <table width="650" border="0" cellspacing="0" cellpadding="0">
                    <tr valign="top"> 
                      <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0" CLASS="linkA">
                          
						 <tr> 
                            <td bgcolor="#CCCCCC"><strong>LOCAT&Aacute;RIO*</strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><select name="locatario" id="locatario"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" >
                                <option value="">Nome</option>
                                <%
set objrs01 = objconn01.execute("select * from locatarios order by nome ")
while not objrs01.eof
%>
                                <option value="<%response.write objrs01("codloc")%>" <% if INT(auxcodcon) = INT(objrs01("codloc")) then%> selected <% end if %>> 
                                <%response.write objrs01("nome")%>
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
                            <td bgcolor="#CCCCCC"><strong>IM&Oacute;VEL*</strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><select name="imovel" id="imovel"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" >
                                <option value="">Imóvel</option>
                                <%
set objrs01 = objconn01.execute("select * from imoveis order by rua,numero ")
while not objrs01.eof
%>
                                <option value="<%response.write objrs01("codimo")%>"  <% if int(auxccodimo) = int(objrs01("codimo")) then%> selected <% end if %>> 
                                <%response.write "Rua: "&objrs01("rua") & " N°: "&objrs01("numero")%>
                                </option>
                                <%
	objrs01.movenext
wend
%>
                              </select></td>
                          </tr>
						  <% if renovar<>"sim" then %>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  <tr> 
                            <td bgcolor="#CCCCCC"><strong>PER&Iacute;ODO</strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><input name="periodo" type="text" id="periodo" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="5" maxlength="2" value="<%=periodo%>"  onKeyPress="return txtBoxFormat('periodo', '99', event);">
                              meses </td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr> 
                            <td bgcolor="#CCCCCC"><strong>DATA IN&Iacute;CIO</strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><input name="dataini" type="text" id="dataini" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="12" maxlength="10" value="<%=dataini%>"  onKeyPress="return txtBoxFormat('dataini', '99/99/9999', event);"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr> 
                            <td bgcolor="#CCCCCC"><strong>DATA FIM</strong></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr> 
                            <td height="18"><input name="datafim" type="text" id="datafim" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="12" maxlength="10" value="<%=datafim%>"  onKeyPress="return txtBoxFormat('datafim', '99/99/9999', event);"></td>
                          </tr>
                          <tr>
                            <td bgcolor="#CCCCCC"><strong>ALUGUEL MENSAL</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="aluguel" type="text" id="aluguel" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="15" maxlength="10" value="<%=aluguel%>"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>DIA VENCIMENTO DO ALUGUEL</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="vencimento" type="text" id="vencimento" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="5" maxlength="2" value="<%=vencimento%>"  onKeyPress="return txtBoxFormat('vencimento', '99', event);"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>MULTA</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="multa" type="text" id="multa" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="30" maxlength="20" value="<%=multa%>">
                              ( max. 20 caract. ) </td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  <% end if %>
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>IMOVEL EXCLUSIVAMENTE 
                              DESTINADO A</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="destinado" type="text" id="destinado" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="60" maxlength="50" value="<%=destinado%>">
                              ( max. 50 caract. ) </td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>DATA DO REAJUSTE</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="reajuste" type="text" id="reajuste" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="12" maxlength="10" value="<%=reajuste%>"  onKeyPress="return txtBoxFormat('reajuste', '99/99/9999', event);"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>DATA DA IMPRESSAO</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="impressao" type="text" id="impressao" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="12" maxlength="10" value="<%=impressao%>"   onKeyPress="return txtBoxFormat('impressao', '99/99/9999', event);"></td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>TESTEMUNHA 1</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="testemunha1" type="text" id="testemunha1" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="60" maxlength="50" value="<%=testemunha1%>">
                              ( max. 50 caract. ) </td>
                          </tr>
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
						  
						  <tr>
                            <td bgcolor="#CCCCCC"><strong>TESTEMUNHA 2</strong></td>
                          </tr>
                          <tr>
                            <td><img src="imagem/img_transp.gif" width="1" height="2"></td>
                          </tr>
                          <tr>
                            <td height="18"><input name="testemunha2" type="text" id="testemunha2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" size="60" maxlength="50" value="<%=testemunha2%>">
                              ( max. 50 caract. ) </td>
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
  <%
call fecha_conexao01
%>  
</div>
</body>
</html>