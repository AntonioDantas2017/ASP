<!--#include file="inc_login.asp"-->
<!--#include file="inc_css.asp"-->
<%
call abre_conexao01
nivel = nivel_cotacao
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
call fecha_conexao01
nivel_aux = nivel
'nivel_aux = 2
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

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
produto   	= request("produto")
codpro 		= request("codpro")	
codcot	  	= Request ("codcot")
auxapaga  	= Request ("auxapaga")
auxcodprc	= Request ("auxcodprc")
%>
<body onLoad="form2.produto.focus()">


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
                  <td bgcolor="#BBECB1"><div align="center">&nbsp;&nbsp; <strong>COTA&Ccedil;&Atilde;O</strong></div></td>
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
  <%
if inserir = "sim" then 
	
	set objRs01 = objConn01.execute("select max(codcot) as maxcodcot from cotacao")
	if IsNull(objRs01("maxcodcot")) then
	   codcot = 100000
	else
	   codcot = objRs01("maxcodcot") + 1
	end if
	
	sql = "INSERT INTO cotacao (codcot,data,aberta) values ("&codcot&",'"&date&"','sim') "
	'response.write sql	
	set objRs01 = objConn01.execute(sql)
	
		
	sql = "update fornecedores set prioridade = 99, datahora='1/1/1900 00:00:00'"
	'response.write sql	
	set objRs01 = objConn01.execute(sql)	
	
	
	Dim fso, f, dir 
	dir = "\inetpub\wwwroot\geral\caixa\fornecedores\"&codcot 
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set f = fso.CreateFolder(dir) 
	CreateFolderDemo = f.Path 

end if

if inserir="n�o" then

if produto <> "" then
	function preparaPalavra(str)
	   preparaPalavra = replace(str,"�","a")
	   preparaPalavra = replace(preparaPalavra,"�","a")
	   preparaPalavra = replace(preparaPalavra,"�","a")
	   preparaPalavra = replace(preparaPalavra,"�","a")
	   preparaPalavra = replace(preparaPalavra,"�","a")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"�","e")	
	   preparaPalavra = replace(preparaPalavra,"�","e")	
	   preparaPalavra = replace(preparaPalavra,"�","e")	
	   preparaPalavra = replace(preparaPalavra,"�","e")		   	   	   
	   preparaPalavra = replace(preparaPalavra,"�","i")
	   preparaPalavra = replace(preparaPalavra,"�","i")
	   preparaPalavra = replace(preparaPalavra,"�","i")
	   preparaPalavra = replace(preparaPalavra,"�","i")	   	   	   
	   preparaPalavra = replace(preparaPalavra,"�","o")
	   preparaPalavra = replace(preparaPalavra,"�","o")
	   preparaPalavra = replace(preparaPalavra,"�","o")
	   preparaPalavra = replace(preparaPalavra,"�","o")
	   preparaPalavra = replace(preparaPalavra,"�","o")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"�","u")
	   preparaPalavra = replace(preparaPalavra,"�","u")
	   preparaPalavra = replace(preparaPalavra,"�","u")
	   preparaPalavra = replace(preparaPalavra,"�","u")	   	   
	   preparaPalavra = preparaPalavra
	end function
	produto = ucase(preparaPalavra(lcase(produto)))
	set objrs01 = objconn01.execute("select * from produtos where produto = '"&produto&"'")
	if objrs01.eof then
		set objRs01 = objConn01.execute("select max(codpro) as maxcodpro from produtos")
		if IsNull(objRs01("maxcodpro")) then
		   codpro = 100000
		else
		   codpro = objRs01("maxcodpro") + 1
		end if
		
		sql = "INSERT INTO produtos (codpro,produto) values ("&codpro&",'"&produto&"') "
		'response.write sql	
		set objRs01 = objConn01.execute(sql)
	else
		codpro = objrs01("codpro")
	end if
end if		

if codpro <> "" then
	set objrs01 = objconn01.execute("select * from produtos_cot where codcot = "&codcot&" and codpro="&codpro&"")
	if not objrs01.eof then
%>
  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
    <tr> 
      <td bgcolor="#FF9999"><div align="center"><strong><font size="2">ESSE PRODUTO J&Aacute; FOI INSERIDO </font></strong></div></td>
    </tr>
  </table>
  <%
	else
		set objrs01 = objconn01.execute("select * from produtos where codpro = "&codpro&"")
	 	if not objrs01.eof then
			set objRs01 = objConn01.execute("select max(codprc) as maxcodprc from produtos_cot")
			if IsNull(objRs01("maxcodprc")) then
			   codprc = 100000
			else
			   codprc = objRs01("maxcodprc") + 1
			end if
			
			sql = "INSERT INTO produtos_cot (codprc,codcot,codpro) values ("&codprc&",'"&codcot&"','"&codpro&"') "
			'response.write sql	
			set objRs01 = objConn01.execute(sql)	
		else
%>
  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
    <tr> 
      <td bgcolor="#FF9999"><div align="center"><strong><font size="2">ESSE PRODUTO N�O EXISTE </font></strong></div></td>
    </tr>
  </table>
  <%
		end if
	END IF
END IF

if auxapaga = "sim" then
sql = "delete from produtos_cot where codprc = "&auxcodprc&" "
'response.write sql	
set objRs01 = objConn01.execute(sql)
%>
	  <table width="700" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
		<tr> 
		  <td bgcolor="#FF9999"><div align="center"><strong><font size="2">O PRODUTO FOI APAGADO </font></strong></div></td>
		</tr>
  </table>
	  <%
end if
end if
%>
	  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 

<script language="JavaScript">

function ValidateOrder(form)
{

  form.submit();
}
</script>
               
			   
			    <br>
			    <div align="center">
                  <strong> </strong><br>
<form METHOD="get" ACTION="cotacao_inserir.asp" name="form2">
  <table width="95%" border="0" bgcolor="#CCCCCC">
    <tr>
      <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
          <tr bgcolor="#FFFFFF">
            <td width="28%"><div align="center"><strong>Cod</strong></div></td>
            <td width="65%"><div align="center"><strong>Produto</strong></div></td>
            <td width="7%" rowspan="2"><div align="center">
                <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><div align="center">
              <input name="codpro" type="text" id="codpro"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="6"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" >
            </div></td>
            <td><div align="center">
			<script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a fun��o 
		ValidateOrder(form);
	}
}

//-->
</script>	
			<input type="hidden" name="inserir" value="n�o">
			<input type="hidden" name="codcot" value="<%=codcot%>">			
			                <input name="produto" type="text" id="produto"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" size="60" maxlength="50"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" >
							<script>
							function busca(codcot)
							{
								MM_openBrWindow('cotacao_busca.asp?codcot='+codcot+'&produto='+document.form2.produto.value,'','status=yes,scrollbars=yes,width=400,height=400');	
							}
							</script>
                 <a href="javascript:busca(<%=codcot%>);"  class="linkA5"><img src="imagem/icones_ver.gif" alt="Pre&ccedil;os desse produto" width="15" height="15" border="0"></a></div></td>
          </tr>
      </table></td>
    </tr>
  </table>
  </form>
                  
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><img src="imagem/img_transp.gif" width="1" height="1"></td>
                    </tr>
                  </table>
                  <%
set objrs01 = objConn01.execute("select * from produtos_cot WHERE codcot="&codcot&" ORDER BY codprc DESC ")
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#CCCCCC"> 
                            <td height="20"><div align="center"><strong>PRODUTO 
                                </strong></div>
                            <div align="center"></div></td>
                            <td width="5%" height="20"><div align="center"><strong>AP</strong></div></td>
                          </tr>
<%
if objrs01.eof then 
%>

<tr bgcolor="#FFFFFF"> 
<td >&nbsp;</td>
<td  >&nbsp; </td>
</tr>

<%else%>
<%
while not objrs01.eof 
contador 		= contador+1
codprc 		= objrs01("codprc")
codpro   	= objrs01("codpro")


set objrs02= objConn02.execute("select * from produtos WHERE codpro="&codpro&"")
if not objrs02.eof then
	produto  = objrs02("produto")
else
	produto = "Erro! Favor Apagar Registro"
end if
%>
                          <tr bgcolor="#FFFFFF"> 
                            <td height="20"><div align="left">
                              <div align="left"><%=produto%></div>
                            </div>
                            <div align="center"></div> <div align="left"></div>                              <div></div> </td>
                            <td width="5%" height="20"><div align="center"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <form name="form<%=codprc%>" method="get" action="cotacao_inserir.asp">
<input name="auxcodprc" type="hidden" value="<%=codprc%>">
<input name="auxapaga" type="hidden" value="sim">
<input name="codcot" type="hidden" value="<%=codcot%>">
<input name="inserir" type="hidden" value="n�o">


                                    <tr> 
                                      <td><div align="center"> 
                                          <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja apagar este produto da cota��o?'));
	}
	</script>
                                         <input NAME="auxdelete<%=codprc%>" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga este item" WIDTH="20" HEIGHT="20" id="auxdelete<%=codprc%>">
										                                          </div></td>
                                    </tr>
                                  </form>
                                </table>
                                
                              </div></td>
                          </tr>
                          <%
	objrs01.movenext
wend
%>
<%end if%>                        </table></td>
                    </tr>
                  </table>
 
                  <br>
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <p><br>
    <br>
    <%
		  call fecha_conexao02
          call fecha_conexao01
%>
  </p>
 
</div>
</body>
</html>