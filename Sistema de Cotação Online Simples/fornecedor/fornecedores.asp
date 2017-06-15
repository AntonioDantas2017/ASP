<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_css.asp"-->
<%
call abre_conexao01
call abre_conexao02

%>
<html><style type="text/css">
<!--
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

<script language="JavaScript" type="text/JavaScript">
<!--

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>

</head>

<%
dim objrs01,objrs02,objconn01,objconn02

senha = request("senha")
fornecedor = request("fornecedor")
if senha <> "" then
	if senha = "123" then
		session("codforn")	= fornecedor
		set objrs01 = objconn01.execute("select * from prod_cot_forn where codfor="&fornecedor&"")
		if objrs01.eof then
		%>
			<script language="JavaScript">
			//	alert('É a primeira vez que você acessa a cotação\nPara ajuda clique no ícone "?"');
			</script>
		<%	
		end if
		
		set objrs01 = objconn01.execute("select max(prioridade) as maior from fornecedores")
		maior = objrs01("maior")
		
		if not isnumeric(maior) then
			maior = 0
		end if
				
		set objrs01 = objconn01.execute("update fornecedores set prioridade = "&(maior+1)&", datahora = '"&date&" "&time&"' where codfor="&fornecedor&"")
		
	%>
		<script language="JavaScript">
			window.location.href = "fornecedor_inserir.asp"
		</script>
	<%
	else
	%>
		<script language="JavaScript">
		{
		alert('Senha errada!!!\ntente novamente')
		}
		</script>
	<%
	end if
	%>
<%
end if
%>

<body onLoad="form.fornecedor.focus()">	

<div align="center"> 
    <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> <strong>FORNECEDORES</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="3">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="25" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" height="27" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td></td>
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
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

<script language="JavaScript">

function ValidateOrder(form)
{

	  if (form.fornecedor.value == "")
  { 
	  alert("Por favor escolha o fornecedor"); form.fornecedor.focus(); return; 
  }  

  form.submit();
}
</script>
               
			   
			    <br>
			    <div align="center">
			      <form action="fornecedores.asp" method="post" name="form" id="form">
			        <table width="100%" border="0" bgcolor="#CCCCCC">
                      <tr bgcolor="#FFFFFF">
                        <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
                            <tr>
                              <td width="20%"><div align="right">Fornecedor:&nbsp;&nbsp;</div></td>
                              <td width="42%"><select name="fornecedor" id="fornecedor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA" >
                                  <option value="">FORNECEDOR</option>
                                
                                  <%
set objrs02 = objconn02.execute("select codfor,vendedor,empresa from fornecedores where fora='sim' order by vendedor ")
while not objrs02.eof
%>
                                  <option value="<%response.write objrs02("codfor")%>">
                                  <%response.write LEFT(ucase(objrs02("vendedor")),15)& " - " &LEFT(ucase(objrs02("empresa")),15)%>
                                  </option>
                                  <%
	objrs02.movenext
wend
%>
                              </select></td>
                              <td width="13%"><div align="right">Senha:&nbsp;&nbsp;</div></td>
                              <td width="13%"><input name="senha" type="password" id="senha"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR: #EAEAEA" size="7" maxlength="6"></td>
                              <td width="12%"><div align="center">
                                <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                              </div></td>
                            </tr>
                        </table></td>
                      </tr>
                    </table>
                    <br>
			      </form>
			      </div>
            </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%
		  call fecha_conexao02
          call fecha_conexao01
%>
  
 
    <br>
    <br>
    <a href="sair.asp"><img src="imagem/botoes_fechar.gif" width="75" height="19" border="0"></a></div>
</body>
</html>