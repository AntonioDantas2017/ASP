<!--#include file="inc_login.asp"-->
<%topo = "sim"%>
<!--#include file="inc_css.asp"-->
<%
primeira = request("primeira")
if primeira <> "sim" then
%>
	<Script Language="JavaScript">
		window.open("default_centro_branco.asp" , "mainFrame");
	</script>
<%end if%>

<html><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>


<body bgcolor="<!--#include file="inc_background.asp"-->" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="center">
<%
menu1 = request("menu1")

'RESPONSE.Write codfunc


%>

<table width="200" height="5" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td></td>
  </tr>
</table>

<table width="710" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="22">&nbsp;</td>
    <td width="999" valign="top"><div align="center"></div>
        <div align="left"> 
          <table width="700" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td valign="bottom"> 
			  
                
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000"><table width="700" height="60" border="0" cellpadding="1" cellspacing="1">
                  <tr> 
                    <td width="4" height="20" bgcolor="#bbecb1"> 
					  
                      <table width="55" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="55"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="vendas.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/fat_cc_ent.gif" alt="Cheques" width="30" height="30" border="0"></a> 
                                  </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="vendas.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>VENDAS</strong></a></div></td>
                        </tr>
                      </table>
					  
					               <table width="86" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="86"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="pagamentos.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/icones_portes.gif" alt="Pagamentos" width="30" height="30" border="0"></a> 
                                  </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="pagamentos.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>PAGAMENTOS</strong></a></div></td>
                        </tr>
                      </table>
					  
					  
					  
					  
					  
					  
					  
					  
					   <table width="69" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="69"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="produtos.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/icones_reservas.gif" alt="Produtos" width="32" height="32" border="0"></a> 
                                  </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="produtos.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>PRODUTOS</strong></a></div></td>
                        </tr>
                      </table>
					  
					  
					  
					  
					  
					  
                      <table width="65" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="65"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="clientes.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/icones_clientes.gif" alt="Clientes" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="clientes.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>&nbsp;CLIENTES&nbsp;</strong></a></div></td>
                        </tr>
                      </table>

                      
					  
					  
					  
					   
					  
                     
                      <table width="65" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="65"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="telefones.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_usu2.gif" alt="Telefones" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="telefones.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>TELEFONES</strong></a></div></td>
                        </tr>
                      </table>
                    
					  
					   
                           
					  
					
						
						
						
						
						
					
						
						
						
						
						
						
						
						
						
						
						
						
						
						
					<table width="64" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="64"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="fr_mensageiro.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/cupons.gif" alt="Avisos" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="fr_mensageiro.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>AVISOS</strong></a></div></td>
                        </tr>
                      </table> 

					  
                      <table width="35" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="35"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="funcionarios.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_func.gif" alt="Usuários" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="funcionarios.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>USU</strong></a></div></td>
                        </tr>
                      </table>
					  

                      <table width="36" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="36"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="usuarios_log.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_log.gif" alt="Log" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="usuarios_log.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>LOG</strong></a></div></td>
                        </tr>
                      </table>
					  
					
					
					
					 <table width="172" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="172"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35">&nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="15">&nbsp;</td>
                        </tr>
                      </table>
					  
					  
					   <table width="36" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="36"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="sair.asp"  target="mainFrame"  class="linkB2"><img src="imagem/icones_apagar.gif" alt="Log" width="32" height="32" border="0"></a> 
                                  </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="sair.asp"  target="mainFrame"  class="linkB2"><strong>SAIR</strong></a></div></td>
                        </tr>
                      </table>
					
					
					
                   </td>
                  </tr>
                </table></td>
            </tr>
          </table>
              </td>
            </tr>
          </table>
               </div></td>
  </tr>
</table>
</div>
</body>
</html>
