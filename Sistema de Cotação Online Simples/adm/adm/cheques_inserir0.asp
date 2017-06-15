<!--#include file="inc_login.asp"-->
<!--#include file="inc_css.asp"-->

<%
call abre_conexao01
nivel = nivel_mercado
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
call fecha_conexao01
nivel_aux = nivel
'nivel_aux = 2
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
call abre_conexao01
call abre_conexao02

inserir	  	= request("inserir")
cliente   	= request("cliente")
codcli 		= request("codcli")	
valor 		= request("valor")	
data 		= request("data")	
codche	  	= Request ("codche")
auxapaga  	= Request ("auxapaga")
auxcont	= Request ("auxcont")
fornecedor = request("fornecedor")
%>

<body onLoad="form2.codcli.focus()">	

<div align="center"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center">&nbsp;&nbsp;<strong>CHEQUE</strong></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td width="3">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="28" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" height="28" border="0" cellpadding="0" cellspacing="0" class=linka>
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
  
  
  <table width="10%" height="10" border="0">
    <tr>
      <td></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

<script language="JavaScript">

function ValidateOrder(form)
{
   if (form.fornecedor.value == "")
  { 
	  alert("Por favor escolha o Fornecedor"); form.fornecedor.focus(); return; 
  } 
  form.submit();
}
</script>
               
			   
			   
			    <div align="center">
                  <strong>
                  </strong>
                  <table width="100%" border="0" class="linkA">
                    <tr>
                      <td><div align="center"><strong>Escolha o Fornecedor</strong>:</div></td>
                    </tr>
                  </table>
                  <strong>                   </strong><br>
<form METHOD="get" ACTION="cheques_inserir.asp" name="form2">
  <table width="95%" border="0" bgcolor="#CCCCCC">
    <tr>
      <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
          <tr bgcolor="#FFFFFF">
            <td width="50%"><div align="right">
             <%
			 	set objRs01 = objConn01.execute("select max(codche) as maxcodche from cheques")
	if IsNull(objRs01("maxcodche")) then
	   codche = 100000
	else
	   codche = objRs01("maxcodche") + 1
	end if

	sql = "INSERT INTO cheques (codche,data) values ("&codche&",'"&date&"') "
	'response.write sql	
	set objRs01 = objConn01.execute(sql)

			 
			 %>
			  <select name="fornecedor" id="fornecedor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" >
                <option value="">FORNECEDOR</option>
                <%
set objrs02 = objconn02.execute("select codfor,vendedor,empresa from fornecedores where fora='não' order by vendedor ")
while not objrs02.eof
%>
                <option value="<%response.write objrs02("codfor")%>">
                  <%response.write LEFT(ucase(objrs02("vendedor")),15)& " - " &LEFT(ucase(objrs02("empresa")),15)%>
                  </option>
                <%
	objrs02.movenext
wend
%>
              </select>
&nbsp;&nbsp; </div></td>
            <td width="50%">
                <div align="left">
                  &nbsp;&nbsp;
                  <input name="inserir" type="hidden" id="button3" value="sim">
				  <input name="codche" type="hidden" id="button3" value="<%=codche%>">
				  <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                    </div></td>
            </tr>
          
      </table></td>
    </tr>
  </table>
  </form>
                  
                                   <br>
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <p><%
		  call fecha_conexao02
          call fecha_conexao01
%>
  </p>
 
</div>
</body>
</html>