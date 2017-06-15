<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->

<!--#include file="inc_css.asp"-->


<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">
<%
'RESGATA AS VARIÁVEIS
'--------------------
listar	  = request("listar")
cod       = request("cod")

'response.write cod & "<BR>"
call abre_conexao01
call abre_conexao02
	set objrs01 = objConn01.execute(" select * from comunicador where cod = '"&cod&"' ")
	if not objrs01.eof then
		mensagem 	= objrs01("mensagem")
		data	 	= objrs01("data")
		hora		= objrs01("hora")
		codfunc		= objrs01("codfunc")
		set objrs01 = objConn01.execute("select * from  Funcionarios where codfunc='"&codfunc&"' ")	
		if not objrs01.eof then
			inic = objrs01 ("nomeresumido")
		end if
'		set objrs01 = objConn01.execute("select * from  comunicador_func where cod='"&cod&"' ")	
'		if not objrs01.eof then
'			codfunci = objrs01 ("codfunc")
'		end if
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
      <td width="155" valign="top"> <table width="150" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>AVISOS</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="795"><table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                     <table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="6%"> <div align="center"><strong> <a class="linkB2" href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"><br>
                          </a></strong></div></td>
                          <td width="6%"><div align="center"><strong> <a class="linkB2"  href="javascript:location.reload()"> 
                              </a></strong><strong><a class="linkB2" href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"><br>
                              </a></strong></div></td>
                          <td width="6%"><div align="center">
                              <div align="center">
                                <div align="center"><strong> </strong> </div>
                              </div>
                              </div></td>
                          <td width="6%"> <div align="center"> </div></td>
                          <td width="6%"> <div align="center"><a class="linkB2" href="#" onClick="MM_openBrWindow('sp_grupo.asp','','status=yes,scrollbars=yes,width=600,height=572')"></a></div></td>
                          <td width="6%"> <div align="center"><a class="linkB2" href="#" onClick="MM_openBrWindow('ft_cc_emp.asp','','status=yes,scrollbars=yes,width=600,height=572')"></a></div></td>
                          <td width="6%"> <div align="center"><a class="linkB2" href="#" onClick="MM_openBrWindow('sp_prod_cat.asp','','status=yes,scrollbars=yes,width=600,height=572')"></a> 
                            </div></td>
                          <td width="6%"> <div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="6%"><div align="center"></div></td>
                          <td width="16%"><div align="center"></div></td>
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
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>EDITA O REGISTRO</strong></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td><div align="center">
                <script language="JavaScript">
function ValidateOrder(form)
{
  if (form.data2 == "")
  { 
  alert("Por favor digite a Data."); form.data2.focus(); return; 
  }  

  if (form.mensagem.value == "")
  { 
  alert("Por favor digite a Mensagem."); form.mensagem.focus(); return; 
  }  

  form.submit();
}
</script>
                <br>
                <form METHOD="post" ACTION="fr_mensageiro_editar2.asp">
                    
                  <p> 
                    <input name="listar" type="hidden" value="<%=listar%>">
                    <input name="axcod" type="hidden" value="<%=cod%>">
					<input name="codfunc" type="hidden" value="<%=codfunc%>">
                    <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO 
                    OBRIGAT&Oacute;RIO</strong></p>
                  <table width="700" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td>
						  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
							  <tr>
								<td width="20%"><div align="right">REMETENTE : &nbsp;</div></td>
								
                            <td width="80%"><strong><%=inic%></strong></td>
							  </tr>
						 </table>
						 <table width="100%" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td><img src="imagem/img_transp.gif" width="1" height="7"></td>
						  </tr> 
						</table>
						 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                          <tr> 
                            <td width="20%"><div align="right">DATA : &nbsp;</div></td>
                            <td width="80%"> <strong><%=data%> às <%=hora%>h</strong></td>
                          </tr>
                        </table>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td><img src="imagem/img_transp.gif" width="1" height="7"></td>
                          </tr>
                        </table>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                          <tr> 
                            <td width="20%" valign="top"><div align="right">DESTINAT&Aacute;RIO 
                                : &nbsp;</div></td>
                            <td width="80%"> <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                                <tr> 
                                 <% 
								
							
                            
							set objrs01 = objConn01.execute("select * FROM  funcionarios WHERE nome <> 'd4w' ORDER BY NOME ")
	cont = 0
	While Not objrs01.EOF 
 		codfunc2 = objrs01("codfunc")
 		NOME = objrs01("nomeresumido")
		
		If cont Mod 4 = 0 AND cont <> 0 Then
			Response.Write "</tr>" & VbCrLf & "<tr>" 
		End If
		%>
                                  <td width="20%"  align="left" valign="top"> <div align="left"> 
                                      <table width="109" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" class="LINKA">
                                        <tr> 
                                          <td width="109"  align="center"   valign="bottom"> 
                                            <div align="left">
											
											<%
											 set objrs02 = objconn02.execute("select * FROM  comunicador_func WHERE cod = '"&cod&"' and codfunc = '"&codfunc2&"'")
											if Not objRs02.EOF then
											 											    
											%>
											 
                                              <input name="checkfunc<%=codfunc2%>"  type="checkbox"  value="<%=codfunc2%>" checked>
											  
											<%else%>
											
										        <input name="checkfunc<%=codfunc2%>"  type="checkbox"  value="<%=codfunc2%>" >
											<%end if
											=UCASE(LEFT(NOME,20))%></div></td>
                                        </tr>
                                      </table>
                                    </div></td>
                                  <%
	cont = cont + 1
	objrs01.MoveNext	
	WEND
	'End if
'end if


%>
                                </tr>
                              </table></td>
                            <BR>
                          </tr>
                        </table>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td><img src="imagem/img_transp.gif" width="1" height="7"></td>
						  </tr> 
						</table>
						 <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
							  <tr>
								<td width="20%"><div align="right">MENSAGEM: &nbsp;</div></td>
								
                            <td width="80%"><textarea onKeypress="if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" name="mensagem" cols="100" rows="15" id="mensagem" style="text-transform: uppercase; FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4"><%=mensagem%></textarea></td>
							  </tr>
						 </table>
						 <table width="100%" border="0" cellspacing="0" cellpadding="0">
						  <tr>
							<td><img src="imagem/img_transp.gif" width="1" height="7"></td>
						  </tr> 
						</table>
						
					</td>
                    </tr>
                  </table>
                  <p><strong></strong></p>
                  <p> 
                    <input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                    <br>
                  </p>
                </form>                
                
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
<%
	call fecha_conexao01
	call fecha_conexao02

%>  
</div>
</body>
</html>