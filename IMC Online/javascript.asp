<%@ LANGUAGE="VBScript"%>
<!--#include file="inc_css.asp"-->
			
<html>
<head>
<style type="text/css">
 .botao{
        font-size:10px;
        font-family:Verdana,Helvetica;
        font-weight:bold;
        color:blue;
        background:#638cb5;
        border:0px;
        width:100px;
        height:19px;
       }
</style>
<TITLE>Índice de Massa Corpórea On-Line</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body background="img/bgfull.jpg" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="21%" height="583" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="98%" height="583" align="center" valign="top" background="img/bgsquare.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" valign="top"><!--#include file="toplogo.asp"-->
          </td>
        </tr>
        <tr> 
          <td align="center" valign="top"><img src="img/1x6v.gif" width="1" height="6"></td>
        </tr>
        <tr> 
          <td align="left" valign="top">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1%" rowspan="9" align="right" valign="top"><img src="img/6x1h.gif" width="7" height="1"></td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><img src="img/tpcadastro.jpg" width="752" height="200"></td>
                <td width="1%" rowspan="9" align="left" valign="top"><img src="img/6x1h.gif" width="6" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#FFFFFF">
<div align="left"><font color="#727292"><br>
                    </font></div></td>
              </tr>

              <tr> 
                <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#FF0000"><img src="img/blank25x8h.gif" width="25" height="8"></font></td>
              </tr>
              <tr>
                <td align="left" valign="top" bgcolor="#FFFFFF"><div align="center"> 
                    <form name="form1" method="post" action="">
                      <table width="75%" border="1" class="linkA">
                        <tr> 
                          <td colspan="2">Nome: 
                            <input name="nome" type="text" id="nome" size="60" maxlength="50"></td>
                        </tr>
                        <tr> 
                          <td colspan="2"><div align="center">Calculadora personalizada</div></td>
                        </tr>
                        <tr> 
                          <td  colspan="2">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td colspan="2">Escolha como ser&aacute; sua calculadora:</td>
                        </tr>
                        <tr> 
                          <td width="32%">Cor do fundo: </td>
                        <td width="68%">    <select name="fundo" size="1" id="fundo">
                              <option value="#ffffff">Branco</option>
                              <option value="#000000">Preto</option>
                              <option value="#ff0000">Vermelho</option>
                              <option value="#0000ff">Azul</option>
                              <option value="#00ff00">Verde</option>
                            </select></td>
                        </tr>
                        <tr> 
                          <td> Tamanho dos N&uacute;meros:</td>
                            <td><input name="tamanho" type="text" id="tamanho" size="10" maxlength="5"></td>
                        </tr>
                        <tr> 
                          <td>Cor dos Bot&otilde;es:</td>
<td>                            <select name="botoes" size="1" id="botoes">
                              <option value="#ffffff">Branco</option>
                              <option value="#000000">Preto</option>
                              <option value="#ff0000">Vermelho</option>
                              <option value="#0000ff">Azul</option>
                              <option value="#00ff00">Verde</option>
                            </select></td>
                        </tr>

                      </table>
                      <br>
                      <input type="button" name="Enviar" value="Enviar" class="botao" onClick="calculadora(this)">
                      <br>
                      <br>
                    </form>
                    <p>&nbsp;</p>
                    <br>
                  </div>
                </td>
              </tr>

           </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
