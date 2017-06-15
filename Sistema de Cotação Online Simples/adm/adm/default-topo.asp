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

<html><style type="text/css">
<!--
.style1 {COLOR: #000099; FONT-FAMILY: verdana; FONT-SIZE: 11px; TEXT-DECORATION: none; font-weight: bold; }
-->
</style><head>
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
			  
                <%if nivel_cotacao <> "0"  then %>
                <table width="120" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td bgcolor="#000000"><div align="center"> 
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">
                          <tr> 
                            <td   <%if menu1="1" then%>bgcolor="#bbecb1"<%else%>bgcolor="#F2F2F2"<%end if%>  ><div align="center"><a href="default-topo.asp?menu1=1"  class="linkB"><strong>COTA&Ccedil;&Atilde;O</strong></a></div></td>
                          </tr>
                        </table>
                      </div></td>
                  </tr>
                </table> 
                <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td>&nbsp;</td>
                  </tr>
                </table>
                <%end if%>
				
				
				<%if nivel_mercado <> "0"  THEN%>
                <table width="120" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td bgcolor="#000000"><div align="center"> 
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">
                          <tr> 
                            <td   <%if menu1="2" then%>bgcolor="#bbecb1"<%else%>bgcolor="#F2F2F2"<%end if%>  ><div align="center"><a href="default-topo.asp?menu1=2"  class="style1">MERCADO</a></div></td>
                          </tr>
                        </table>
                      </div></td>
                  </tr>
                </table> 
                <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td>&nbsp;</td>
                  </tr>
                </table>
                <%end if%>
				
				<%if nivel_contratos <> "0"  THEN%>
                <table width="120" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td bgcolor="#000000"><div align="center"> 
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">
                          <tr> 
                            <td   <%if menu1="3" then%>bgcolor="#bbecb1"<%else%>bgcolor="#F2F2F2"<%end if%>  ><div align="center"><a href="default-topo.asp?menu1=3"  class="linkB"><strong>CONTRATOS</strong></a></div></td>
                          </tr>
                        </table>
                      </div></td>
                  </tr>
                </table> 
                <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td>&nbsp;</td>
                  </tr>
                </table>
                <%end if%>
				
				<%if nivel_funcionarios <> "0"  then %>
                <table width="120" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr>
                    <td bgcolor="#000000"><div align="center">
                        <table width="100%" border="0" cellspacing="1" cellpadding="1">
                          <tr>
                            <td   <%if menu1="4" then%>bgcolor="#bbecb1"<%else%>bgcolor="#F2F2F2"<%end if%>  ><div align="center"><a href="default-topo.asp?menu1=4"  class="linkB"><strong>FUNCION&Aacute;RIOS</strong></a></div></td>
                          </tr>
                        </table>
                    </div></td>
                  </tr>
                </table></td>
            </tr>
          </table>
		  <%end if%>
		  
		  
          <%if menu1="1" then%>
		  
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000">
			    
				<table width="700" height="60" border="0" cellpadding="1" cellspacing="1" >
                  <tr> 
                    <td width="700" height="20" bgcolor="#bbecb1"> 
                		
						<table width="100%" border="0" cellpadding="0" cellspacing="0">
						  <tr>
							<td>
							
								<table width="74" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
									<tr> 
										<td width="82"><img src="imagem/img_transp.gif" height="5"></td>
									</tr>
									<tr> 
										<td height="35"><div align="center"><a href="cotacao.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_cotacao.gif" alt="Cota&ccedil;&atilde;o" width="32" height="32" border="0"></a> 
									</div></td>
									</tr>
									<tr> 
										<td height="15"><div align="center"> <a href="cotacao.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>COTA&Ccedil;&Atilde;O</strong></a></div></td>
									</tr>
							  </table>

							  							  
								<table width="108" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td width="108"><img src="imagem/img_transp.gif" height="5"></td>
                                  </tr>
                                  <tr>
                                    <td height="35"><div align="center"><a href="fornecedores.asp?primeira=sim&filtro=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_usu.gif" alt="Fornecedores" width="32" height="32" border="0"></a> </div></td>
                                  </tr>
                                  <tr>
                                    <td height="15"><div align="center"> <a href="fornecedores.asp?primeira=sim&filtro=sim"  target="mainFrame"  class="linkB2"><strong>FORNECEDORES</strong></a></div></td>
                                  </tr>
                              </table>
								<table width="82" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td width="82"><img src="imagem/img_transp.gif" height="5"></td>
                                  </tr>
                                  <tr>
                                    <td height="35"><div align="center"><a href="produtos.asp"  target="mainFrame"  class="linkB2"><img src="imagem/icones_plano.gif" alt="Produtos da Cota&ccedil;&atilde;o" width="32" height="32" border="0"></a> </div></td>
                                  </tr>
                                  <tr>
                                    <td height="15"><div align="center"> <a href="produtos.asp"  target="mainFrame"  class="linkB2"><strong>PRODUTOS</strong></a></div></td>
                                  </tr>
                              </table>
								</td>
						  </tr>
						</table>
						
				    </td>
                  </tr>
                </table>
				
			  </td>
            </tr>
          </table>
		  
          <%else%>
		  
          <%if menu1="2" then%>
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000"><table width="700" height="60" border="0" cellpadding="1" cellspacing="1">
                  <tr> 
                    <td width="4" height="20" bgcolor="#bbecb1"> 
					  
                      <table width="63" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="63"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="cheques.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/fat_chq.gif" alt="Cheques" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="cheques.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>CHEQUES</strong></a></div></td>
                        </tr>
                      </table>
                      <table width="85" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td width="69"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="clientes.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><img src="imagem/icones_clientes.gif" alt="Clientes" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="clientes.asp?primeira=sim&tipo=C"  target="mainFrame"  class="linkB2"><strong>&nbsp;CLIENTES&nbsp;</strong></a></div></td>
                        </tr>
                      </table>

                      <table width="108" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="108"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="fornecedores.asp?primeira=sim&filtro=não"  target="mainFrame"  class="linkB2"><img src="imagem/icones_usu.gif" alt="Fornecedores" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="fornecedores.asp?primeira=sim&filtro=não"  target="mainFrame"  class="linkB2"><strong>FORNECEDORES</strong></a></div></td>
                        </tr>
                      </table>
                      <table width="30" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="38"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"></div></td>
                        </tr>
                      </table>
                      <table width="74" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="74"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="telefones.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_usu2.gif" alt="Telefones" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="telefones.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>TELEFONES</strong></a></div></td>
                        </tr>
                      </table>
                      <table width="30" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="38"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"></div></td>
                        </tr>
                      </table>
                      <table width="74" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="74"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="15">&nbsp;</td>
                        </tr>
                      </table>
                      <table width="90" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="15">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <%else%>
		  
		  <%if menu1="3" then%>
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000"><table width="700" height="60" border="0" cellpadding="1" cellspacing="1">
                  <tr> 
                    <td width="4" height="20" bgcolor="#bbecb1"><table width="86" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="86"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="locatarios.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><img src="imagem/icones_locador.gif" alt="Locadores" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="locatarios.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><strong>LOCAT&Aacute;RIOS</strong></a></div></td>
                        </tr>
                      </table>
					  
					  
                      <table width="89" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="89"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="IMOVEIS.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><img src="imagem/HOME.gif" alt="Pagamentos" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="IMOVEIS.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><strong>IM&Oacute;VEIS</strong></a></div></td>
                        </tr>
                      </table>
                      <table width="90" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="contratos.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><img src="imagem/icones_contrato.gif" alt="Contratos" width="32" height="32" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="contratos.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><strong>CONTRATOS</strong></a></div></td>
                        </tr>
                      </table>
                      <table width="89" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="89"><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr>
                          <td height="35"><div align="center"><a href="pagamentos.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><img src="imagem/fat_cc_hist.gif" alt="Pagamentos" width="30" height="30" border="0"></a> </div></td>
                        </tr>
                        <tr>
                          <td height="15"><div align="center"> <a href="pagamentos.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"><strong>PAGAMENTOS</strong></a></div></td>
                        </tr>
                      </table>

						

                        <table width="33" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="80"><img src="imagem/img_transp.gif" height="5"></td>
                          </tr>
                          <tr>
                            <td height="35">&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="15"><div align="center"> <a href="contato.asp?primeira=sim&tipo=P"  target="mainFrame"  class="linkB2"></a></div></td>
                          </tr>
                        </table>
                        <table width="100" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                          <tr>
                            <td><img src="imagem/img_transp.gif" height="5"></td>
                          </tr>
                          <tr>
                            <td height="35"><div align="center"><a href="recibos.asp"  target="mainFrame"  class="linkB2"><img src="imagem/botoes_conv.gif" alt="Recibos" width="32" height="32" border="0"></a> </div></td>
                          </tr>
                          <tr>
                            <td height="15"><div align="center"> <a href="recibos.asp"  target="mainFrame"  class="linkB2"><strong>RECIBOS</strong></a></div></td>
                          </tr>
                        </table>
                    </td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <%else%>
		  
		  <%if menu1="4" then%>
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000">
			    <table width="700" height="60" border="0" cellpadding="1" cellspacing="1">
                  <tr> 
                    <td width="4" height="20" bgcolor="#bbecb1"> 
					  
                      <table width="70" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="funcionarios.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_func.gif" alt="Usuários" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="funcionarios.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>USU&Aacute;RIOS</strong></a></div></td>
                        </tr>
                      </table>

                      <table width="50" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td><img src="imagem/img_transp.gif" height="5"></td>
                        </tr>
                        <tr> 
                          <td height="35"><div align="center"><a href="usuarios_log.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_log.gif" alt="Log" width="32" height="32" border="0"></a> 
                            </div></td>
                        </tr>
                        <tr> 
                          <td height="15"><div align="center"> <a href="usuarios_log.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>LOG</strong></a></div></td>
                        </tr>
                      </table>
					  
                    </td>
                  </tr>
                </table>
			  </td>
            </tr>
          </table>
          <%else%>
          <table width="700" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000">
			  	<table width="700" height="60" border="0" cellpadding="1" cellspacing="1">
                  <tr> 
					<td height="20" bgcolor="#F2F2F2">&nbsp;</td>
                  </tr>
                </table>
			  </td>
            </tr>
          </table>
		  
         
		  <%end if%>
		  <%end if%>
		  <%end if%>
		  <%end if%>
        </div></td>
  </tr>
</table>
</div>
</body>
</html>
