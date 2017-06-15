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
<title>Cartão - Ver</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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

</head>


<%
call abre_conexao01
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02



%>
 <%
                          datainicio = request("datainicio")
                          if datainicio = "" then
                          datainicio  = date
                          else
                          datainicio  = request("datainicio")
                          end if
                          

op = request("filtro2")                  

if op <> "" then
	sql3 = " and op = '"&op&"' "
end if


  if request("apagar") = "sim" then
     set objrs01 = objConn01.execute("delete from cartao WHERE cod = "&request("cod")&" ")
                          end if
                          
if request("aux")= "S" then
	set objrs01 = objconn01.execute("update cartao  set REC='S' where cod="&REQUEST("cod")&"")
end if

if request("aux")= "N" then
	set objrs01 = objconn01.execute("update cartao  set REC='N' where cod="&REQUEST("cod")&"")
end if    

sql2 = "select * from cartao WHERE data=#"&datainicio&"# "&sql3&" and rec = 'N'"
'response.write sql2
set objrs01 = objConn01.execute(sql2)
if  objrs01.eof then                      
                          %>

<body  onLoad="document.form3.datainicio.focus()">
<%else%>

<body  onLoad="ap1.focus()">	
<%end if%>
<div align="center">
  
  <table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="150"> <table width="150" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>CARTAO</strong></div></td>
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
                     	
					    						  <form action="cartao_ver.asp" method="get" name="form3">
                          
                            <input name="filtro2" type="hidden" value="<%=request("filtro2")%>">
                               <td width="29"><div align="center"><strong><a href="javascript:MM_openBrWindow('cartao.asp?inserir=não&filtro2=<%=request("filtro2")%>','','status=yes,scrollbars=yes,width=500,height=130')"><img src="imagem/icones_extra.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
					      <td width="39"><div align="center"><strong><a href="JavaScript:window.location.href='cartao_ver.asp?inserir=não&filtro2=<%=request("filtro2")%>&titulo=CARTAO'"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="139"> <div align="center">
                         
                          
                          &nbsp;<input name="datainicio" type="text" id="datainicio"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=datainicio%>" size="11" maxlength="10">&nbsp;
<input name="imageField32" type="image" id="imageField3" src="imagem/icones_ok.gif" alt="Procura Usu&aacute;rio (Nome, Login)" width="20" height="13" border="0">
                          </div></td>
						  </form>
                          <td width="139"> <div align="center">
                          
                           <select name="select2" id="filtro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                                            
                              <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=" <%IF request("filtro2") = "" then%> selected <%end if%>>OP</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=1" <%IF request("filtro2") = "1" then%> selected <%end if%>>01  - Mossor&oacute;</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=5" <%IF request("filtro2") = "5" then%> selected <%end if%>>05  - Auxiliadora</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=6" <%IF request("filtro2") = "6" then%> selected <%end if%>>06  - Fabiola</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=7" <%IF request("filtro2") = "7" then%> selected <%end if%>>07  - Nen&ecirc;</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=9" <%IF request("filtro2") = "9" then%> selected <%end if%>>09  - Deusa</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=10" <%IF request("filtro2") = "10" then%> selected <%end if%>>10 - Fagner</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=11" <%IF request("filtro2") = "11" then%> selected <%end if%>>11 - Rosana</option>
                        <option value="cartao_ver.asp?datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=12" <%IF request("filtro2") = "12" then%> selected <%end if%>>12 - Socorro</option>
                               </select>  </div></td>

                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <script>
		//MM_openBrWindow('gas_imp.asp?coditem=<%=coditem%>','','status=yes,scrollbars=yes,width=100,height=100');
		</script>
  <br>
  &nbsp;<table width="500" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#BBECB1"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td valign="top" bgcolor="#FFFFFF"><div align="center"> 


			   
			    <br>
			    <div align="center">
			    
                  <%
op = request("filtro2")                  

if op <> "" then
	sql = " and op = '"&op&"' "
end if

                      

set objrs01 = objConn01.execute("select * from cartao WHERE data=#"&datainicio&"# "&sql&"  ORDER BY hora desc")
%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#FFDDCC"> 
                            <td width="9%" height="20" bgcolor="#CCCCCC"><div align="center"><strong>HORA</strong></div>
                              <div align="center"></div>
                            <div align="center"></div>                              <div align="center"></div></td>
                            <td width="35%" bgcolor="#CCCCCC"><div align="center">
								<strong>NOME </strong></div></td>
                            <td width="6%" bgcolor="#CCCCCC"><div align="center"><strong>OP</strong></div></td>
                            <td width="21%" bgcolor="#CCCCCC"><div align="center">
								<strong>VALOR</strong></div></td>
                          <td width="19%" bgcolor="#CCCCCC"><div align="center">
							<strong>ENTREGUE</strong></div></td>
                          <td width="5%" bgcolor="#CCCCCC"><div align="center">
							<strong>AP</strong></div></td>
                          </tr>
<%
if objrs01.eof then 
%>

<tr bgcolor="#FFFFFF"> 
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
<td bgcolor="#FFFFFF" >&nbsp;</td>
</tr>

<%else%>
<%
tote = 0
totn = 0
contn = 0
conte = 0

while not objrs01.eof 
contador 		= contador+1
cod 		= objrs01("cod")
hora	 		= objrs01("hora")
valor	 		= objrs01("valor")
if not isnumeric(valor) then
	valor = 0
end if
op		 		= objrs01("op")
nome	 		= objrs01("nome")
RECEBIDO 				= objrs01("REC")
%>
      
            <%if RECEBIDO = "N" then
            totn = totn + valor
                        contn = contn + 1
            %>
			<tr valign="middle" bgcolor="#FFC8C8"> 
	<%else
	        tote = tote + valor
	        conte = conte + 1
	%>
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6"> 
              <%end if%>
	<%end if%>

                            <td height="20"><div align="left">
                              <div align="center"><%=hora%></div>
                            </div>
                                               </td>
                            <td height="20"><%=nome%></td>
                            <td height="20"><div align="center"><%=right("00"&op,2)%></div></td>
                            <td height="20"><%=formatnumber(valor,2)%></td>
                                                        <td height="20"><table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka align=left>
                              <tr>
                                <td><div align="center">
                                      <%if RECEBIDO="S" then%>
                                    <a href="cartao_ver.asp?aux=N&cod=<%=cod%>&datainicio=<%=datainicio%>&filtro2=<%=request("filtro2")%>" class="LinkA"><img src="imagem/icones_confirma.gif" alt="Gás Pago" width="20" height="20" border="0"> SIM</a>
                                    <%else%>
                                    <a href="cartao_ver.asp?aux=S&cod=<%=cod%>&datainicio=<%=datainicio%>&filtro2=<%=request("filtro2")%>" class="LinkA"  name="ap<%=contador%>"><img src="IMAGEM/icones_cancela.gif" alt="Gás Não Pago" width="20" height="20" border="0"> NÃO</a>
                                    <%end if%>
                                </div></td>
                                </tr>
                            </table></td>

                            <td width="5%" height="20">
                                <div align="center" ><a href="CARTAO_VER.ASP?filtro2=<%=request("filtro2")%>&datainicio=<%=datainicio%>&cod=<%=cod%>&apagar=sim"  class="linkA5"><img src="imagem/icones_apaga.gif" alt="Apagar Gás" width="20" height="20" border="0"></a></div>                            </td>

                           
                          </tr>
                          <%
	objrs01.movenext
wend


%>
 <%end if%>   
                      </table></td>
                    </tr>
                  </table>
                  <br>
					<br>

                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class="linkA">
                          <tr bgcolor="#FFDDCC"> 
                            <td width="25%" height="20" bgcolor="#CCCCCC" align="center"><div align="center">
								<b>TOTAL GERAL</b></div>
                              <div align="center"></div>
                            <div align="center"></div>                              <div align="center"></div></td>
                          <td width="35%" bgcolor="#CCCCCC" align="center"><b>TOTAL ENTREGUE</b></td>
                          <td width="32%" bgcolor="#CCCCCC" align="center"><b>TOTAL NÃO 
							ENTREGUE</b></td>
                          </tr>


<tr bgcolor="#FFFFFF"> 
<td bgcolor="#FFFFFF" align="center" ><%=formatnumber((tote+totn),2)%>(<%=contador%>)</td>
<td bgcolor="#FFFFFF" align="center" ><%=formatnumber(tote,2)%>(<%=conte%>)</td>
<td bgcolor="#FFFFFF" align="center" ><%=formatnumber(totn,2)%>(<%=contn%>)</td>
</tr>

                      </table></td>
                    </tr>
                  </table>
                

                  &nbsp;</div>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
 
  <br>
    <br>
    <%
		  call fecha_conexao02
          call fecha_conexao01
%>
</div>
</body>
</html>