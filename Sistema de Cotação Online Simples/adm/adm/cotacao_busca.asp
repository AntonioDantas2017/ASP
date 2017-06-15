<!--#include file="inc_conexao.asp"-->
<%
dim objrs01,objconn01,objrs02,objconn02
call abre_conexao01
call abre_conexao02
%>
<!--#include file="inc_css.asp"-->


<html><head>
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

<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


</head>


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<div align="center"><br>
  <table width="350" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td>                         <div align="center"><strong><font color="#000000">BUSCA DA COTA&Ccedil;&Atilde;O </font></strong></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <br>
 
  <div align="center">
  <%
  CODCOT = REQUEST("CODCOT")
  PRODUTO = REQUEST("PRODUTO")
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
	   preparaPalavra = replace(preparaPalavra,"'","`")	   	   
	   preparaPalavra = replace(preparaPalavra,""&chr(39)&"","´")	   	   	   	   
	   preparaPalavra = preparaPalavra
	end function
	produto = ucase(preparaPalavra(lcase(produto)))
	
  SET OBJRS01 = OBJCONN01.EXECUTE("SELECT * FROM PRODUTOS2 WHERE PRODUTO LIKE '%"&PRODUTO&"%' AND CODCOT="&CODCOT)
  IF OBJRS01.EOF THEN
  %>
  
    <table width="350" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr>
        <td bgcolor="#FFFFFF"><div align="center"> <strong>ESSE PRODUTO N&Atilde;O FOI ENCONTRADO .</strong></div>
            <div align="center"><strong></strong></div></td>
      </tr>
    </table>
    <br>
    <br>
    <div align="center">
      <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	//opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
    <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></div>
    <%
else
%>
    <div align="center"> <br>
        <div align="center">
          <table width="350" border="1" cellpadding="0" cellspacing="0" CLASS="linkA">
            <tr valign="top">
              <td width="317"><div align="center"><strong><font color="#000000">PRODUTOS ENCONTRADOS </font></strong></div></td>
           <!--   <td width="27"><div align="center"><strong>AP</strong></div></td>-->
            </tr>
            <% 

'For cont = 1 to 50
contador = 0
while not objrs01.eof 
contador = contador + 1
codprc = objrs01("codprc")
%>
            
           
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
            <%end if%>
       
			  
            <td><div align="left">
              <%
			
			
				'codpro = objrs02("codpro")
				'set objrs02 = objconn02.execute("select * from produtos where codpro="&codpro&"")
				'if not objrs02.eof then
					response.Write ucase(objrs01("produto"))
				'else
				'	response.Write "Erro! Produto não Encontrado"
				'end if
			
			
			
			
			%>
            </div></td>
<!--			   <SCRIPT LANGUAGE="JavaScript">
function apagar<%=codprc%>() 
{
	opener.document.form<%=codprc%>.auxdelete<%=codprc%>.focus();
// Só funciona se for botao

	self.window.close();
}
    </SCRIPT>
            <td> <a href="#"  onClick="apagar<%=codprc%>()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></td>-->
          </tr>
            <% 
	objrs01.movenext
	wend
	end if	
%>
          </table>
        </div>
    </div>
    <div align="center"><br>
        <table width="300" border="0" cellspacing="0" cellpadding="0" class=linka>
          <tr>
            <td width="50%"><div align="center">
                <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	//opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
            <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></div></td>
          </tr>
        </table>
    </div>
  </div>
</div>
<%
call fecha_conexao01
call fecha_conexao02
%>
</body>
</html>