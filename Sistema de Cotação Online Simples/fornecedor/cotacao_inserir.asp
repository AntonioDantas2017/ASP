<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_css.asp"-->

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
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02

inserir	  	= request("inserir")
produto   	= request("produto")
codpro 		= request("codpro")	
codcot	  	= Request ("codcot")
auxapaga  	= Request ("auxapaga")
auxcodprc	= Request ("auxcodprc")
%>

<body onLoad="form2.produto.focus()">	

<div align="center">
  <table width="500" border="0" cellspacing="0" cellpadding="0">
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
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      <table width="100%" height="27" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td> </td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%
  set objrs01 = objconn01.execute("select * from cotacao where aberta='sim'")
  if objrs01.eof then
  %>
  <br>
  <br>
  <table width="100%" border="0" class="linkA">
    <tr>
      <td><div align="center">N&Atilde;O EXISTE COTA&Ccedil;&Atilde;O ABERTA </div></td>
    </tr>
  </table>
  <br>
  <%
else
	codcot = objrs01("codcot")
	inserir = "não"

if inserir="não" then

if produto <> "" then
	function preparaPalavra(str)
	   preparaPalavra = replace(str,"á","a")
	   preparaPalavra = replace(preparaPalavra,"à","a")
	   preparaPalavra = replace(preparaPalavra,"ã","a")
	   preparaPalavra = replace(preparaPalavra,"â","a")
	   preparaPalavra = replace(preparaPalavra,"ä","a")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"é","e")	
	   preparaPalavra = replace(preparaPalavra,"è","e")	
	   preparaPalavra = replace(preparaPalavra,"ê","e")	
	   preparaPalavra = replace(preparaPalavra,"ë","e")		   	   	   
	   preparaPalavra = replace(preparaPalavra,"í","i")
	   preparaPalavra = replace(preparaPalavra,"ì","i")
	   preparaPalavra = replace(preparaPalavra,"î","i")
	   preparaPalavra = replace(preparaPalavra,"ï","i")	   	   	   
	   preparaPalavra = replace(preparaPalavra,"ó","o")
	   preparaPalavra = replace(preparaPalavra,"ò","o")
	   preparaPalavra = replace(preparaPalavra,"õ","o")
	   preparaPalavra = replace(preparaPalavra,"ô","o")
	   preparaPalavra = replace(preparaPalavra,"ö","o")	   	   	   	   
	   preparaPalavra = replace(preparaPalavra,"ú","u")
	   preparaPalavra = replace(preparaPalavra,"ù","u")
	   preparaPalavra = replace(preparaPalavra,"û","u")
	   preparaPalavra = replace(preparaPalavra,"ü","u")	   	   
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
  <table width="500" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
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
  <table width="500" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
    <tr> 
      <td bgcolor="#FF9999"><div align="center"><strong><font size="2">ESSE PRODUTO NÃO EXISTE </font></strong></div></td>
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
	  <table width="500" border="0" cellspacing="0" cellpadding="0" align="center" class="linkA">
		<tr> 
		  <td bgcolor="#FF9999"><div align="center"><strong><font size="2">O PRODUTO FOI APAGADO </font></strong></div></td>
		</tr>
	  </table>
	  <%
end if
end if
%>
  <br>
  <table width="500" border="0" cellspacing="0" cellpadding="0">
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
            <td width="33%"><div align="center"><strong>Cod</strong></div></td>
            <td width="60%"><div align="center"><strong>Produto</strong></div></td>
            <td width="7%" rowspan="2"><div align="center">
                <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
            </div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><div align="center"><script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a função 
		ValidateOrder(form);
	}
}

//-->
</script>	
              <input name="codpro" type="text" id="codpro"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR: #FFC4C4" size="10" maxlength="6"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;">
            </div></td>
            <td><div align="center">
			<input type="hidden" name="inserir" value="não">
			<input type="hidden" name="codcot" value="<%=codcot%>">			
			                <input name="produto" type="text" id="produto"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" size="50" maxlength="50"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;">
            </div></td>
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
set objrs01 = objConn01.execute("select * from produtos_cot WHERE codcot="&codcot&" ORDER BY codprc ")
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#FFDDCC"> 
                            <td height="20" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO 
                                </strong></div>
                              <div align="center"></div>
                            <div align="center"></div>                              <div align="center"></div></td>
                            <td width="5%" height="20" bgcolor="#CCCCCC"><div align="center"><strong>AP</strong></div></td>
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
            <form name="form4" method="get" action="cotacao_inserir.asp">
<input name="auxcodprc" type="hidden" value="<%=codprc%>">
<input name="auxapaga" type="hidden" value="sim">
<input name="codcot" type="hidden" value="<%=codcot%>">
<input name="inserir" type="hidden" value="não">


                                    <tr> 
                                      <td><div align="center"> 
                                          <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja apagar este prodto da cotação?'))
	}
	</script>
                                          <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga este item" WIDTH="20" HEIGHT="20">
                                        </div></td>
                                    </tr>
                                  </form>
                                </table>
                                
                              </div></td>
                          </tr>
                          <%
	objrs01.movenext
wend
if isnumeric(totalprodutos) then
	totalprodutos = replace(formatnumber(totalprodutos,2),".","")
end if

%>
<%end if%>                        </table></td>
                    </tr>
                  </table>
                  <br>
			    </div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <%end if%>
  <br>
    <br>
    <%
		  call fecha_conexao02
          call fecha_conexao01
%>
</div>
</body>
</html>