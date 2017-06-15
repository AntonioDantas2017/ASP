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
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="form.preco.focus()">
<div align="center"> 

  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                                             <td width="722"><div align="center"><strong>AJUDA COTA&Ccedil;&Atilde;O </strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
            <tr> 
              <td bgcolor="#FFFFFF"><div align="center"> 
                 
                 
                  <br>
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>
						 
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
						    <tr>
						      <td><p>PARA FAZER A COTA&Ccedil;&Atilde;O:<br>
				                <br>
				                1 - O PRODUTO APARECER&Aacute; AO LADO DA CAIXA PARA SE DIGITAR O PRE&Ccedil;O.<br>
				                <br>
				                2 - CASO VOC&Ecirc; TENHA O PRODUTO DIGITE O PRE&Ccedil;O E APERTE &quot;ENTER&quot;, SE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;N&Atilde;O TIVER DEIXE A CAIXA EM BRANCO OU DIGITE &quot;0,00&quot; E APERTE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&quot;ENTER&quot;.<br>
				                <br>
				                3 - O PRE&Ccedil;O DEVE SER DIGITADO COM V&Iacute;RGULA E DUAS CASAS DECIMAIS<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;EXEMPLO: &quot;1,32&quot;. CASO CONTR&Aacute;RIO O PRODUTO SER&Aacute; INSERIDO COM O<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;VALOR &quot;0,00&quot;, ASSIM AFIRMANDO QUE VOC&Ecirc; N&Atilde;O TEM O MESMO.<br>
					          <br>
					          4 - O CAMPO OBSERVA&Ccedil;&Atilde;O SE REFERE AS OBSERVA&Ccedil;&Otilde;ES QUE VOC&Ecirc; PODE &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;FAZER DO PRODUTO. EX: &quot;N&Atilde;O TEM O SABOR CHOCOLATE&quot; OU &quot;MARCA X&quot;.<br>
					          <br>
					          5 - CASO PRECISE ALTERAR O PRE&Ccedil;O OU OBSERVA&Ccedil;&Atilde;O DE UM PRODUTO, &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CLIQUE NO &Iacute;CONE AMARELO, DIGITE O PRE&Ccedil;O E/OU OBSERVA&Ccedil;&Atilde;O &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;NORMALMENTE</p>
						        <p>6 - PARA BUSCAR UM PRODUTO CLIQUE NO S&Iacute;MBOLO DA LUPA ABAIXO DO &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;BOTAO OK APERTE	ENTER	DIGITE	O	PRODUTO, APERTE ENTER SELECIONE<br>
&nbsp;&nbsp;&nbsp;						          &nbsp;O&nbsp;PRODUTO E CLIQUE NO S&Iacute;MBOLO DE EDITAR, DIGITE O PRE&Ccedil;O E/OU &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;OBSERVA&Ccedil;&Atilde;O E APERTE ENTER.<br>
						          <br>
						          7 - AO FINALIZAR A COTA&Ccedil;&Atilde;O APARECER&Aacute; &quot;ACABOU OS PRODUTOS DA <br>
  &nbsp;&nbsp;&nbsp;&nbsp;					          COTA&Ccedil;&Atilde;O&quot;, CONFIRA OS PRE&Ccedil;OS DOS PRODUTOS E CLIQUE EM FECHAR.</p></td>
					        </tr>
						  </table>                        </td>
                      </tr>
                    </table>
                    <br>
                <a href="#" onClick="javascript:self.window.close();"><img src="IMAGEM/botoes_fechar.gif" width="75" height="19" border="0"></a>
                <br>
                <br>
                <br>
              </div></td>
            </tr>
          </table></td>
      </tr>
</table>
</body>
</html>