<!--#include file="inc_conexao.asp"-->
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
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>


<body bgcolor="#FFFFFF">
<div align="center">
  <table width="500" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="22">&nbsp;</td>
    <td width="999" valign="top">
        <div align="left">
          <table width="500" height="60" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td bgcolor="#000000">
			    
				<table width="500" height="60" border="0" cellpadding="1" cellspacing="1" >
                  <tr> 
                    <td width="500" height="20" bgcolor="#bbecb1"> 
                		
						<table width="100%" border="0" cellpadding="0" cellspacing="0">
						  <tr>
							<td>
							
								
								
								<table width="74" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
									<tr> 
										<td width="82"><img src="imagem/img_transp.gif" height="5"></td>
									</tr>
									<tr> 
										<td height="35"><div align="center"><a href="cotacao_inserir.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_cotacao.gif" alt="Cota&ccedil;&atilde;o" width="32" height="32" border="0"></a><a href="cotacao_inserir.asp?primeira=sim"  target="mainFrame"  class="linkB2"></a> 
									</div></td>
									</tr>
									<tr> 
										<td height="15"><div align="center"> <a href="cotacao_inserir.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>COTA&Ccedil;&Atilde;O</strong></a></div></td>
									</tr>
							  </table>

							  							  
								<table width="108" height="55" border="0" align="left" cellpadding="0" cellspacing="0">
                                  <tr>
                                    <td width="108"><img src="imagem/img_transp.gif" height="5"></td>
                                  </tr>
                                  <tr>
                                    <td height="35"><div align="center"><a href="fornecedores.asp?primeira=sim"  target="mainFrame"  class="linkB2"><img src="imagem/icones_chat.gif" alt="Fornecedores" width="32" height="32" border="0"></a> </div></td>
                                  </tr>
                                  <tr>
                                    <td height="15"><div align="center"> <a href="fornecedores.asp?primeira=sim"  target="mainFrame"  class="linkB2"><strong>FORNECEDORES</strong></a></div></td>
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
		  
        </div></td>
  </tr>
</table>
</div>
</body>
</html>