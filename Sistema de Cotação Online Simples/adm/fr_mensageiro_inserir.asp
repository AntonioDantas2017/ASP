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
'dim objrs01
'dim objConn01
'RESGATA AS VARIÁVEIS
'--------------------
listar		 = request("listar")
ordem		 = request("ordem")
pesquisa 	 = request("pesquisa")
datainicio   = request("datainicio")
datafim 	 = request("datafim")
filtro 	 	 = request("filtro")


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
      <td height="13" bgcolor="#CCCCCC"><div align="center"> <strong>CADASTRO 
          DE UMA NOVA MENSAGEM</strong></div></td>
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
  
  if (form.mensagem.value == "")
  { 
  alert("Por favor digite a Mensagem."); form.mensagem.focus(); return; 
  }  

  form.submit();
}
</script>
                <br>
                <form METHOD="post" ACTION="fr_mensageiro_inserir2.asp">
                     <input name="listar" type="hidden" value="<%=listar%>">
					 <input name="ordem" type="hidden" value="<%=ordem%>">
					 <input name="filtro" type="hidden" value="<%=filtro%>">
					 <input name="datainicio" type="hidden" value="<%=datainicio%>">
					 <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
					 <input name="datafim" type="hidden" value="<%=datafim%>">
					 <input name="nomeresumido" type="hidden" value="<%=nomeresumido%>">
                    <strong>OS CAMPOS EM VERMELHO S&Atilde;O DE PREENCHIMENTO 
                    OBRIGAT&Oacute;RIO</strong><br>
                  </p>
                  <table width="600" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                          <tr>
                            <td width="20%"><div align="right">REMETENTE : &nbsp;</div></td>
                            <td width="80%">
							
							<%=nomeresumido%></td>
                          </tr>
                        </table>
							

							
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
                          <tr> 
                            <td width="20%" valign="top"><div align="right">DESTINAT&Aacute;RIO 
                                : &nbsp;</div></td>
                            <td width="80%"> 
                              <table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
   
    <tr> 
      <% Call abre_conexao01%>
	 
	   <input name="checktodos" type="checkbox" value="T">TODOS&nbsp;&nbsp; <br>
	 <% set objrs01 = objConn01.execute("select * FROM  funcionarios ORDER BY NOME  ")
	 cont = 0
	   While Not objrs01.EOF 
 		 codfunc = objrs01("codfunc")
		  NOME    = objrs01("NOMERESUMIDO")
		
		If cont Mod 4 = 0 AND cont <> 0 Then
			Response.Write "</tr>" & VbCrLf & "<tr>" 
		End If
		%>
      <td  align="left" valign="top"> <div align="left"> 
          <table width="80" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" class="LINKA">
            <tr> 
                                          <td width="80"  align="center"   valign="bottom"> 
                                            <div align="left">     
                                              <input name="checkfunc<%=codfunc%>"  type="checkbox"  value="<%=codfunc%>">
                                          <%=LEFT(NOME,20)%></div></td>
            </tr>
          </table>
          
        </div></td>
      <%
		cont = cont + 1
	objrs01.MoveNext	
	WEND
	'End if
'end if
Call fecha_conexao01
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
						    <td width="20%"><div align="right">DATA : &nbsp;</div></td>
						    <td width="80%"><strong> 
                              <%RESPONSE.Write(DATE)%>
                              </strong></td>
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
					        <td width="80%"> <div align="left">
                                <textarea onKeypress="if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" name="mensagem" cols="100" rows="15" id="mensagem" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4"></textarea>
                              </div></td>
				  </tr>
				</table>

					</td>
                    </tr>
                  </table>
                  <p>&nbsp;</p>
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
  
</div>
</body>
</html>