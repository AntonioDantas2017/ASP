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
      
	 
function fecha_refresh() 
{
	opener.location.reload(true);
	self.window.close();
}

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


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  onLoad="form5.produto.focus()">

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
  CODPRC = REQUEST("CODPRC")
  CODFORN = REQUEST("CODFORN")
  CODCOT = REQUEST("CODCOT")
  CODPROD1 = REQUEST("CODPROD1")
  CODPROD2 = REQUEST("CODPROD2")  
      
  PRODUTO = REQUEST("PRODUTO")
  if CODPROD1 <> "" then
		  SET OBJRS01 = OBJCONN01.EXECUTE("delete from prod_cot_forn where cont = "&CODPROD1&"")
		 end if
  if PRODUTO = "" then
  		
  	
  %>
    <script language="JavaScript">
	function ValidateOrder(form)
	{
		form.submit();
	}
	</script>
  <form METHOD="get" ACTION="cotacao_BUSCA.asp" name="form5">
    <table width="350" border="0" bgcolor="#CCCCCC">
      <tr>
        <td><table width="100%" border="0" cellpadding="1" cellspacing="1" class="linkA">
            <tr bgcolor="#FFFFFF">
              <td><div align="center"><strong>Produto</strong></div></td>
              <td width="7%" rowspan="2"><div align="center">
                  <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
              </div></td>
            </tr>
            <tr bgcolor="#FFFFFF">
              <td><div align="center">
                    <script language="JavaScript">
<!--

function tecla_enter(tecla,form)
{ 
	if(tecla == 13){ 
		//chama a fun&ccedil;&atilde;o 
		ValidateOrder(form);
	}
}

//-->
  </script>
                    <input type="hidden" name="CODPRC" value="<%=CODPRC%>">
                    <input type="hidden" name="CODFORN" value="<%=CODFORN%>">
                    <input type="hidden" name="CODPROD1" value="<%=CODPROD1%>">
                    <input type="hidden" name="CODCOT" value="<%=CODCOT%>">
                    <input name="produto" type="text" id="produto"  style="FONT-FAMILY: Verdana; FONT-SIZE: 11px; BACKGROUND-COLOR:#FFC4C4" size="45" maxlength="50"  onKeypress="tecla_enter(+event.keyCode,this.form);if (event.keyCode == 34 || event.keyCode == 35 || event.keyCode == 37 || event.keyCode == 38 || event.keyCode == 39 || event.keyCode == 96) event.returnValue = false;" >
                  
          </div></td>
              </tr>
        </table></td>
      </tr>
    </table>
  </form>
  
  <%
  else
  	if CODPROD2 <> "" OR Request("AUX") = 1 then
		if CODPROD2 <> "" then
			SET OBJRS01 = OBJCONN01.EXECUTE("delete from prod_cot_forn where cont = "&CODPROD2&"")
		end if
	%>
		<SCRIPT LANGUAGE="JavaScript">
			opener.location.reload(true);
			self.window.close();
    	</SCRIPT>
	<%
	else
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

'  SET OBJRS01 = OBJCONN01.EXECUTE("select * from prod_cot_forn2 WHERE codfor="&codforn&" and codcot = "&codcot&" and produto like '%"&produto&"%' ORDER BY cont desc ")
  'SET OBJRS01 = OBJCONN01.EXECUTE("select * from prod_cot_forn2 WHERE codfor="&codforn&" and codcot = "&codcot&" ORDER BY cont desc ")
  SET OBJRS01 = OBJCONN01.EXECUTE("SELECT * FROM PRODUTOS2 WHERE PRODUTO LIKE '%"&PRODUTO&"%' AND CODCOT="&CODCOT&"")
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

    <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></div>
    <%
else
%>
    <div align="center"> <br>
        <div align="center">
          <table width="350" border="1" cellpadding="0" cellspacing="0" CLASS="linkA">
            <tr valign="top">
              <td width="271"><div align="center"><strong><font color="#000000">PRODUTOS ENCONTRADOS </font></strong></div></td>
              <td width="50"><div align="center"><strong>PRE&Ccedil;O</strong></div></td>
              <td width="21"><div align="center"><strong>ED</strong></div></td>
              <!--   <td width="27"><div align="center"><strong>AP</strong></div></td>-->
            </tr>
            <% 

'For cont = 1 to 50
contador = 0
while not objrs01.eof 
contador 		= contador+1
'cont 		= objrs01("cont")
codprc   	= objrs01("codprc")
'preco   	= objrs01("preco")

'response.Write "select * from prod_cot_forn2 WHERE codfor="&codforn&" and codcot = "&codcot&" and codprc = "&codprc&" "

SET OBJRS02 = OBJCONN02.EXECUTE("select * from prod_cot_forn2 WHERE codfor="&codforn&" and codcot = "&codcot&" and codprc = "&codprc&" ")
if not objrs02.eof then
	cont 		= objrs02("cont")
	preco   	= objrs02("preco")


set objrs02= objConn02.execute("select * from produtos_cot WHERE codprc="&codprc&"")
if not objrs02.eof then
	codpro  = objrs02("codpro")
else
	produto = "Erro! Produto Não Encontrado. Favor colocar valor R$ 0,00"
end if

set objrs02= objConn02.execute("select * from produtos WHERE codpro="&codpro&"")
if not objrs02.eof then
	produto  = objrs02("produto")
else
	produto = "Erro! Produto Não Encontrado. Favor colocar valor R$ 0,00"
end if

if not isnumeric(preco) then
	preco = 0
end if

%>
            
           
            <%'if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
            <%'end if%>
       
			  
            <td><div align="left">
              <%
					response.Write ucase(produto)

			%>
            </div></td>
            <td>
              <div align="right">
                <%
					response.Write formatnumber(preco,2)
			%>
              </div></td>
            <td> <a href="cotacao_busca.asp?CODPROD2=<%=cont%>&PRODUTO=<%=produto%>&AUX=1">  <img SRC="imagem/icones_edita.gif" alt="Editar produto" WIDTH="20" HEIGHT="20" border="0"></a></td>
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
end if			
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
               
              <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a><a href="#"  onClick="fecha_refresh()" class=linkb>
              <%end if%>
              </a></div></td>
          </tr>
        </table>
    </div>
  </div>
</div>
<%
	end if
call fecha_conexao01
call fecha_conexao02
%>
</body>
</html>