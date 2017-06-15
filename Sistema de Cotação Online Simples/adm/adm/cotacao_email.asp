<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
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
%><%
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
  <table width="650" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td>                         <div align="center"><strong><font color="#000000">ENVIAR  COTA&Ccedil;&Atilde;O </font></strong></div></td>
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
  
  codcot = request("codcot")
  enviar = request("enviar")
  
  if enviar <> "sim" then
  %>
  
  
  <% 
  SET OBJRS01 = OBJCONN01.EXECUTE("SELECT msg_cot from config")
  msg_cot = objrs01("msg_cot")
  
  SET OBJRS01 = OBJCONN01.EXECUTE("SELECT * from email_cot2 where codcot = "&codcot)
  IF OBJRS01.EOF THEN
  %>
  
    <table width="650" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr>
        <td bgcolor="#FFFFFF"><div align="center"> <strong> N&Atilde;O FOI ENCONTRADO 
          NENHUM FORNECEDOR.</strong></div>
            <div align="center"><strong></strong></div></td>
      </tr>
    </table>
    <%
RESPONSE.WRITE "<br><br><br><br><br><br>"
else
%>
    <div align="center"> <br>
        <div align="center">
		<form name="form1" action="cotacao_email.asp" method="get">
		<input type="hidden" value="<%=codcot%>" name="codcot">
          <table width="650" border="1" cellpadding="0" cellspacing="0" CLASS="linkA">
            <tr valign="top">
              <td width="180"><div align="center"><strong><font color="#000000">FORNECEDOR</font></strong></div></td>
              <td width="180"><div align="center"><font color="#000000"><strong>EMPRESA</strong></font></div></td>
              <td width="250"><div align="center"><strong><font color="#000000">EMAIL</font></strong></div></td>
              <td width="30"><div align="center">
                <input type="checkbox" name="td" value="false">
              </div></td>
            </tr>
            <% 

'For cont = 1 to 50
contador = 0
while not objrs01.eof 
contador = contador + 1
vendedor	= objrs01("vendedor")
empresa		= objrs01("empresa")
email		= objrs01("email")
recebido	= objrs01("recebido")
enviado		= objrs01("enviado")
codfor		= objrs01("codfor")

msg = ""

if isdate(enviado) then
	msg = "Enviado em "&enviado
end if
%>
            
           
           
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
       
            <td width="180"><div align="left">
              <%=vendedor%>
            </div></td>
            <td><div align="left">
              <%=empresa%>
            
            </DIV></TD>
             
            <td><div align="left">
              <%=lcase(email)%>
            </DIV></TD>
				
            <td><div align="center">
              <input type="checkbox" name="<%=codfor%>" value="<%=codfor%>" alt="<%=msg%>" checked="<%if msg <> "" then%>true<%else%>false<%end if%>">
            </div></TD>
            
			
            </tr>
            <% 
	objrs01.movenext
	wend
	end if	
%>
          </table>
		  <br>
		  <table width="184" border="1" cellpadding="0" cellspacing="0" CLASS="linkA">
            <tr valign="top">
              <td width="180"><div align="center"><strong><font color="#000000">MENSAGEM</font></strong></div></td>
            </tr>
            <tr valign="middle" bgcolor="#E6E6E6">
            
              <td width="180"><div align="left">
                <textarea name="msg" cols="50" rows="5" class="linkA" id="msg">
				<%=msg_cot%>
				</textarea>
              </div></td>
              </tr>
          
          </table>
		  <br>
		  <input name="Button" TYPE="button" onClick="ValidateOrder(this.form)" VALUE="Confirma" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
		</form>
		<script language="JavaScript">
function ValidateOrder(form)
{     

  form.submit();

}
</script>

        </div>
    </div>
    <div align="center"> <br>
	
	<%else%>
	
	<%
	
	
	
	%>
	
        <br>
        <table width="300" border="0" cellspacing="0" cellpadding="0" class=linka>
          <tr>
            <td width="50%"><div align="center">
                <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
	opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
            <a href="#"  onClick="fecha_refresh()" class=linkb><img src="IMAGEM/botoes_fechar.gif" alt="Fechar" width="75" height="19" border="0"></a></div></td>
          </tr>
        </table>
    </div>
  </div>
</div>
<%end if%>
<%
call fecha_conexao01
call fecha_conexao02
%>
</body>
</html>