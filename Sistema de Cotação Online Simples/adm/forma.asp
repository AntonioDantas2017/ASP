<!--#include file="inc_conexao.asp"-->
<%
dim objrs01,objconn01,objrs02,objconn02
call abre_conexao01
call abre_conexao02
%>
<!--#include file="inc_css.asp"-->


<html><head>
			                  <SCRIPT LANGUAGE="JavaScript">
function fecha_refresh() 
{
//	opener.location.reload(true);
	self.window.close();
}
    </SCRIPT>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<%
codven	 	= request("codven")

if axcodpag <> "" then
	'valor		= request("valor")
	set objrs01 = objconn01.execute("delete from pagamentos where codpag="&axcodpag&"")
	'restam = restam + valor
end if

set objrs01 = objconn01.execute("select sum(valor*qtd) as tot from itens where codven="&codven)
tot			= objrs01("tot")

set objrs02 = objconn02.execute("select sum(valor) as tot from pagamentos where codven="&codven)
restam		= objrs02("tot")

if not isnumeric(tot) then
	tot = 0
end if

if not isnumeric(restam) then
	restam = 0
end if



if tot <> restam then
	restam = tot - restam
else
	restam = 0
end if

axcodpag 	= request("axcodpag")
forma		= request("forma")

if forma <> "" then
	valor		= request("valor")
	data		= request("data")
	obs			= request("obs")
	pago		= request("pago")	
	if isnumeric(valor) and isdate(data) and isnumeric(restam) then
		'if formatnumber(restam,2) > formatnumber(valor,2)  then
			valor	= replace(valor,".","")
			valor	= replace(valor,",",".")
			set objrs01 = objconn01.execute("insert into pagamentos (valor,data,codmod,codven,obs,pago) values ("&valor&",'"&data&"',"&forma&","&codven&",'"&obs&"','"&pago&"')")
			restam = restam - replace(valor,".",",")
			
		'else
			%> <!--<script>alert("Valor Superior ao Restante!<%=valor%>-<%=restam%>");</script> --><%
		'end if
	else
		%> <script>alert("Alguns Dados Informados estão Incorretos!");</script> <%	
	end if
end if



%>


</head>


<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="f1.modalidade.focus()">


<div align="center"><br>
  <table width="350" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr>
                  <td bgcolor="#BBECB1" ><table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                      <tr>
                        <td>                         <div align="center"><strong><font color="#000000">FORMAS DE PAGAMENTO </font></strong></div></td>
						<td width="6%"><div align="center"><a href="#"  onClick="fecha_refresh()" class=linkb ><img  border="0" src="imagem/icones_sair.gif" width="20" height="20"></a></div></td>
                      </tr>
                  </table></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
    </tr>
  </table>

 
<form name="f1" action="forma.asp">
<input type="hidden" value="<%=codven%>" name="codven"%>
  <table width="350" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr>
        <td bgcolor="#FFFFFF"><div align="center">
          <table width="100%" border="0" class="linkA">
		  		
            <tr bgcolor="#CCCCCC">
			  
              <td colspan="4" ><div align="center"><strong>RESTAM R$ <%=formatnumber(restam,2)%></strong></div></td>
			  <script>


function ValidateOrder(form)
{
	if(form.valor.value == "")
	{
		alert("Digite o Valor!");form.valor.focus();
	}
	
	if(form.data.value == "")
	{
		alert("Digite a data!");form.data.focus();
	}

  form.submit();
}

			  </script>
              </tr>
            <tr>
              <td width="14%">VALOR:</td>
              <td width="28%"><input name="valor" type="text" id="modalidades"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" size="8"  maxlength="7">
                0,00</td>
              <td width="12%">DATA:</td>
              <td width="46%"><input name="data" type="text" id="data"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFC4C4" onKeyPress="return txtBoxFormat('data', '99/99/9999', event);" value="<%=fdata(date)%>" size="12"  maxlength="10">
                00/00/0000</td>
            </tr>
			<tr>
              <td colspan="2">FORMA:
                <select name="forma" id="forma"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#FFC4C4" >
                  <%
set objrs01 = objconn01.execute("select * from modalidades order by modalidades")
while not objrs01.eof
%>
                  <option value="<%response.write objrs01("codmod")%>">
                    <%response.write (ucase(objrs01("modalidades")))%>
                    </option>
                  <%
	objrs01.movenext
wend
%>
                  </select></td>
              <td colspan="2" valign="middle">PAGO: 
                <input name="pago" type="radio" value="s">
                Sim&nbsp;
                <input name="pago" type="radio" value="n"  checked="checked">
                Não</td>
              </tr>
			  <tr>
              <td colspan="4">OBS: 
                <input name="obs" type="text" id="obs"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #cccccc" size="50"  maxlength="100"></td>
              </tr>
            <tr>
              <td colspan="4"><div align="center">
                <div align="center">
                 <%if restam > 0 then%> <input name="Button3" type="button" id="button3" onClick="ValidateOrder(this.form)" value="OK" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA"> <%end if%>
                </div>
              </div></td>
              </tr>
          </table>
          <br>
        </div>
            <div align="center"><strong></strong></div></td>
	  </tr>
	  	
    </table>
</form>
<table width="350" border="0" cellspacing="1" cellpadding="1" class=linka>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">
      <table width="100%" border="0" class="linkA">
        <tr bgcolor="#FFFFFF">
          <td width="22%" ><div align="center"><strong>VALOR</strong></div></td>
          <td width="21%"><div align="center"><strong>DATA</strong></div></td>
          <td width="50%"><div align="center"><strong>FORMA</strong></div></td>
          <td width="7%"><div align="center"><strong>AP</strong></div></td>
        </tr>
        <%			
			set objrs01= objConn01.execute("select * from pagamentos2 where codven = "&codven)
			contador = 0
			while not objrs01.eof
			contador = contador + 1
			
%>
        <%if (contador mod 2)<>0 then%>
        <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')">	
          <%else%>
        <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')">

          <%end if%>
          <td ><div align="center"><%=formatnumber(objrs01("valor"),2)%></div></td>
          <td><div align="center"><%=fdata(objrs01("data"))%></div></td>
          <td><div align="center"><%=ucase(objrs01("modalidades"))%></div></td>
          <td><div align="center"><a href="forma.asp?axcodpag=<%=objrs01("codpag")%>&valor=<%=objrs01("valor")%>&codven=<%=codven%>" class="linkA5"><img src="imagem/icones_apaga.gif" width="20" height="20" border="0" class="linkA"></a></div></td>
        </tr>
        <%
			objrs01.movenext
			wend
			%>
      </table>
      <br>
      <br>
    </div>
        <div align="center"><strong></strong></div></td>
  </tr>
</table>
<br>
<br>
<br>
<br>

</div>
</body>
</html>