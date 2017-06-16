<%@ LANGUAGE="VBScript"%>
<!--#include file="inc_css.asp"-->
			
<html>
<head>
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
                <td align="left" valign="top" bgcolor="#FFFFFF" class="linkA">
<div align="right"><font color="#727292"> <%response.Write("Data: "&date&" Hora: "&Time)%><br>
                    </font></div></td>
              </tr>

              <tr> 
                <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#FF0000"><img src="img/blank25x8h.gif" width="25" height="8"></font></td>
              </tr>
              <tr>
                <td align="left" valign="top" bgcolor="#FFFFFF" class="linkA"><div align="center"><br>
                    <%
					if time >= 12 and time < 18 then
						response.Write("Boa Tarde ")
					else
						if time >= 1 and time < 12 then
							response.Write("Boa Dia ")
						else
							response.Write("Bom Noite ")
						end if
					end if
					
					%>
                    d igite suas informa&ccedil;&otilde;es:<br>
                  </div>
                  <form name="form1" method="post" action="resultado.asp">
                    <div align="center">
                      <table width="75%" border="0" class="linkA">
                        <tr> 
                          <td width="10%">Nome:</td>
                          <td> <input name="nome" type="text" id="nome" value="" size="60" maxlength="50"></td>
                        </tr>
                        <tr> 
                          <td width="10%">Sexo:</td>
                          <td> <input name="sexo" type="radio" value="Masculino" checked>
                            Masculino 
                            <input type="radio" name="sexo" value="Feminino">
                            Feminino </td>
                        </tr>
                        <tr> 
                          <td>Peso:</td>
                          <td> <input name="peso" type="text" id="txtpeso2" size="15" maxlength="10">
                            Kg (ex. 78,400)</td>
                        </tr>
                        <tr> 
                          <td>Altura:</td>
                          <td> <input name="altura" type="text" id="altura" size="15" maxlength="10">
                            Mt (ex. 1,78)</td>
                        </tr>
                      </table>
                    </div>
                    <p align="center">
                      <input type="hidden" name="cadastro" value="sim">
                      <br>
                      <input name="cmdenviar" type="submit" id="cmdenviar" value="Resultado">
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input name="cmdlimpar" type="reset" id="cmdlimpar" value="Limpar">
                    </p>
                    <p align="center">&nbsp;</p>
                  </form> 
                  
                </td>
              </tr>

           </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
