<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
'call abre_conexao01
'nivel = nivel_cotacao
'if ((nivel="1")or(nivel="2")) then

'else
'    Response.Redirect ("default_acessonegado.asp")
'end if
'call fecha_conexao01
'nivel_aux = nivel
'nivel_aux = 2
%>
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


</head>
<%
fornecedor = request("fornecedor")
%>

<body>

<div align="center"> 
<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
codcot = request("codcot")

call abre_conexao01
dim objrs01,objconn01
call abre_conexao02
dim objrs02,objconn02
call abre_conexao03
dim objrs03,objconn03


if fornecedor = "" then
%>
  <table width="750" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
    <tr> 
      <td width="228"><div align="center">
          <!--#include file="inc_logo.asp"-->
        </div></td>
      <td width="522"><div align="center"><strong>IMPRIMIR COTA&Ccedil;&Atilde;O </strong></div></td>
    </tr>
  </table>
  <br>
  <div align="center"></div>
  <div align="center">
<table width="750" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
      <form action="COTACAO_imp.asp" method="post" name="form1">
        <input name="tipo2" type="hidden" value="<%=tipo2%>">
		
		
		<tr> 
          <td valign="top"><div align="center"><strong>
            
              </strong><br>
              <br>
              <table width="500" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
                <tr>
                  <td height="2"></td>
                </tr>
                <tr>
                  <td bgcolor="#CCCCCC"><strong>FORNECEDOR</strong> </td>
                </tr>
                <tr>
                  <td height="2"></td>
                </tr>
                <tr>
                  <td><select name="fornecedor" id="fornecedor"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA" >
                      <option value="">Fornecedor</option>
				      <option value="todos">Todos</option>
                      <%
set objrs02 = objconn02.execute("select codfor,vendedor from fornecedores where fora='sim' order by vendedor ")
while not objrs02.eof
%>
                      <option value="<%response.write objrs02("codfor")%>">
                      <%response.write objrs02("vendedor")%>
                      </option>
                      <%
	objrs02.movenext
wend
%>
                    </select>                  </td>
                </tr>
                <tr>
                  <td height="2"></td>
                </tr>
              </table>
              <strong><input type="hidden" name="codcot" value="<%=codcot%>" ><br>
              </strong>
              <table width="400" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="33%"><div align="center"> 
                      <input name="Submit" TYPE="submit" VALUE="Confirma"style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA">
                    </div></td>
                  <td width="33%"><div align="center"><a href="javascript:self.close()" class=linkb><img src="IMAGEM/botoes_fechar.gif" width="75" height="19" border="0"></a></div></td>
                </tr>
              </table>
            </div></td>
        </tr>
      </form>
    </table>
    
  </div>
  <%
else
fornecedor		= request("fornecedor")
codcot		= request("codcot")


set objrs01 = objConn01.execute("select quantidade from temp_cotacao where quantidade <> ''")
if objrs01.eof then
	set objrs01 = objConn01.execute("delete from prod_cot_forn where preco <= 0.00 ")

		if fornecedor = "todos" then
		'COTACAO DO MENOR VALOR DOS PRODUTOS
		set objrs01 = objConn01.execute("delete from temp_cotacao")
		sql = "SELECT * FROM produtos_cot where  codcot = "&codcot&""
		set objrs01 = objConn01.execute(sql)
		while not objrs01.eof 
			codprc = objrs01("codprc")
			sql = "SELECT min(preco) as menor,codfor FROM prod_cot_forn2 where codprc="&codprc&" group by preco,codfor,prioridade order by preco,prioridade"
			'sql = "SELECT min(preco) as menor,codfor FROM prod_cot_forn2 where codprc="&codprc&" group by preco,codfor order by preco"
			'response.Write sql&"<BR>"
			set objrs02 = objConn02.execute(sql)
			if (not isnumeric(objrs02("menor"))) or (isnull(objrs02("menor"))) then
			
			else
				preco = objrs02("menor")
	'			preco= replace(preco,".","")
	'			preco= replace(preco,",",".")			
	'			sql = "SELECT * FROM prod_cot_forn where preco="&preco&" order by cont"
	'			set objrs02 = objConn02.execute(sql)
				codfor = objrs02("codfor")
				codpro = objrs01("codpro")
				preco= replace(preco,".",",")			
				sql = "insert into temp_cotacao (data,codfor,codpro,preco,tipo) values(#"&date&"#,"&codfor&","&codpro&",'"&preco&"','-')"
				'response.Write sql&"<BR><BR>"			
				objConn02.execute(sql)
			end if
			objrs01.movenext
		wend
	
		'COTACAO DOS MAIORES VALOR DOS PRODUTOS
		sql = "SELECT * FROM produtos_cot where  codcot = "&codcot&""
		set objrs01 = objConn01.execute(sql)
		while not objrs01.eof 
			codprc = objrs01("codprc")
			sql = "SELECT max(preco) as maior,codfor FROM prod_cot_forn2 where codprc="&codprc&" group by preco,codfor,prioridade order by preco desc,prioridade desc"
			'sql = "SELECT max(preco) as maior,codfor FROM prod_cot_forn2 where codprc="&codprc&" group by preco,codfor order by preco desc"
			set objrs02 = objConn02.execute(sql)
			if (not isnumeric(objrs02("maior"))) or (isnull(objrs02("maior"))) then
				
			else
				preco = objrs02("maior")
				'preco = replace(preco,".","")
				'preco = replace(preco,",",".")			
				'sql = "SELECT * FROM prod_cot_forn where preco="&preco&" order by cont"
				'set objrs02 = objConn02.execute(sql)
				codfor = objrs02("codfor")
				codpro = objrs01("codpro")
		
				preco = replace(preco,".",",")			
		
				sql = "insert into temp_cotacao (data,codfor,codpro,preco,tipo) values(#"&date&"#,"&codfor&","&codpro&",'"&preco&"','+')"
				'response.Write sql & "<BR>"
				objConn02.execute(sql)
			end if
			objrs01.movenext
		wend
	end if
end if
%>
  <div align="center">
  <%
  fornecedor2 = fornecedor
  if fornecedor2 = "todos" then
	sql = "SELECT codfor FROM fornecedores where fora='sim'"
	set objrs01 = objConn01.execute(sql)
  	while not objrs01.eof 
		fornecedor = objrs01("codfor")

	sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
	set objrs02 = objConn02.execute(sql)
	if not objrs02.eof then
  %> 
<table width="100%" border="0">
  <tr>
    <td>  
	<form action="cotacao_imp2.asp" method="post">
      <div align="center">
      <table width="600" border="0" class="linkA">
<%
		sql = "SELECT * FROM fornecedores where codfor = "&fornecedor&""
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
			nome = objrs02("vendedor")
			empresa =  objrs02("empresa")
			fax =  objrs02("fax")
			telefone=  objrs02("telefone")
		end if
%>			
      <tr>
        <td><table width="600" border="1" cellpadding="0" cellspacing="0" bordercolor="#7C7C7C" class="linkA">

		   <tr bordercolor="#FFFFFF" bgcolor="#7C7C7C">
            <td width="740"><table width="100%" border="0" cellpadding="2" cellspacing="2" class="linkA">
                <tr bgcolor="#FFFFFF">
                  <td width="16%"><strong>FORNECEDOR:</strong></td>
                  <td width="36%"><%=nome%></td>
                  <td width="16%"><strong>EMPRESA:</strong></td>
                  <td width="32%"><%=empresa%></td>
                  </tr>
              </table></td>
            </tr>
		  <% if fax<>"" or telefone <> "" then%>
		   <tr bordercolor="#FFFFFF" bgcolor="#7C7C7C">
		     <td><table width="100%" border="0" cellpadding="2" cellspacing="2" class="linkA">
               <tr bgcolor="#FFFFFF">
                 <td width="15%"><strong>FAX:</strong></td>
                 <td width="18%"><%=fax%></td>
                 <td width="15%"><strong>TELEFONE:</strong></td>
                 <td width="18%"><%=telefone%></td>
                 <td width="34%">
				   <div align="center">
				     <%
				 data_aux2 = split(date,"/")
				 data_aux = right("0"&data_aux2(0),2) &"/"&right("0"&data_aux2(1),2) &"/"&data_aux2(2)
				 response.Write data_aux &"&nbsp;-&nbsp;"&time
				 %>				 
			        </div></td>
                 </tr>
             </table></td>
            </tr>
		  <% end if%>
          <tr>
            <td height="995" valign="top">
              <div align="center">
                <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&" and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
	%>			
                  <table width="100%" border="0" bgcolor="#FFFFFF">
                    <tr>
                      <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#3C3C3C" class="linkA">
                        <tr bgcolor="#CCCCCC">
                          <td width="61%" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
                          <td width="19%" bgcolor="#CCCCCC"><div align="center"><strong>QUANTIDADE</strong></div></td>
                          <td width="16%" bgcolor="#CCCCCC"><div align="center"><strong>VALOR</strong></div></td>
                          <td width="4%" bgcolor="#CCCCCC"><div align="center"><strong>AP</strong></div></td>
                        </tr>
                        <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			preco = objrs02("preco")			
			cont = objrs02("cont")
			quantidade = objrs02("quantidade")
	%>
                        <tr id="tr<%=cont%>">
                          <td bgcolor="#FFFFFF"><%
				sql = "SELECT produto FROM produtos where codpro = "&objrs02("codpro")&""
				set objrs03 = objConn03.execute(sql)
				produto = ""
				if not objrs03.eof then
					produto		=  objrs03("produto")
					response.Write objrs03("produto")
				else
					response.Write "-"
				end if
			
					%>                          </td>
                          <td bgcolor="#FFFFFF">
                            <div align="center">
                              <input name="quant<%=cont%>" type="text" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA" size="10" maxlength="10" value="<%=quantidade%>"  onfocus="javascript:setPointer(tr<%=cont%>, '#CCFFCC');" onBlur="javascript:setPointer(tr<%=cont%>, '#FFFFFF');">
                              
                            </div></td>
                          <td bgcolor="#FFFFFF"><div align="right">
                            <%
				response.Write formatnumber(preco)
				%>
                          </div></td>
                          <td bgcolor="#FFFFFF">
                            <input type="checkbox" name="ap<%=cont%>" value="sim" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px" onClick="javascript:if(this.checked){setPointer(tr<%=cont%>, '#FF8888');}else{setPointer(tr<%=cont%>, '#FFFFFF');}" title="Apagar Produto: <%=produto%>"><!---->
							                          </td>
                        </tr>
                        <%	
	objrs02.movenext		
  wend
  
  %>
                      </table></td>
                    </tr>
                  </table>

                  <%
						  
		else
		
		'response.Write "Nenhum produto nesse fornecedor!"
		end if%>
              </div></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
	<%
	end if	
	objrs01.movenext
	wend
	%>
	<br>
	<label>
	  <input type="hidden" name="fornecedor2" value="<%=fornecedor2%>">
      <input type="submit" name="Submit2" value="Imprimir" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA">
      </label>
      </div>
	<br>
</form>
</td>
    </tr>
</table>
<br>
<br>

	<%
else

  %>
  <table border="1" width="100%">
  <TR>
  <TD>
  <form action="cotacao_imp2.asp" method="post">
    <div align="center">
      <table width="600" border="0">
<%
		sql = "SELECT * FROM fornecedores where codfor = "&fornecedor&""
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
			nome = objrs02("vendedor")
			empresa =  objrs02("empresa")
			fax =  objrs02("fax")
			telefone=  objrs02("telefone")
		end if
%>			

      <tr>
        <td><table width="600" border="1" cellpadding="0" cellspacing="0" bordercolor="#7C7C7C" class="linkA">
          <tr bordercolor="#FFFFFF" bgcolor="#7C7C7C">
            <td><table width="100%" border="0" cellpadding="2" cellspacing="2" class="linkA">
              <tr bgcolor="#FFFFFF">
                <td width="16%"><strong>FORNECEDOR:</strong></td>
                <td width="36%"><%=nome%></td>
                <td width="16%"><strong>EMPRESA:</strong></td>
                <td width="32%"><%=empresa%></td>
              </tr>
            </table></td>
            </tr>
		  <% if fax<>"" or telefone <> "" then%>
		   <tr bordercolor="#FFFFFF" bgcolor="#7C7C7C">
            <td><table width="100%" border="0" cellpadding="2" cellspacing="2" class="linkA">
              <tr bgcolor="#FFFFFF">
                <td width="15%"><strong>FAX:</strong></td>
                <td width="18%"><%=fax%></td>
                <td width="15%"><strong>TELEFONE:</strong></td>
                <td width="18%"><%=telefone%></td>
                <td width="34%"><div align="center">
                    <%
				 data_aux2 = split(date,"/")
				 data_aux = right("0"&data_aux2(0),2) &"/"&right("0"&data_aux2(1),2) &"/"&data_aux2(2)
				 response.Write data_aux &"&nbsp;-&nbsp;"&time
				 %>
                </div></td>
              </tr>
            </table></td>
            </tr>
		  <% end if%>
          <tr>
            <td height="850" valign="top">
              <div align="center">
                <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"  and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
	%>
                <table width="100%" border="0" bgcolor="#FFFFFF">
                  <tr>
                    <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#3C3C3C" class="linkA">
                        <tr bgcolor="#CCCCCC">
                          <td width="61%" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
                          <td width="20%" bgcolor="#CCCCCC"><div align="center"><strong>QUANTIDADE</strong></div></td>
                          <td width="15%" bgcolor="#CCCCCC"><div align="center"><strong>VALOR</strong></div></td>
                          <td width="4%" bgcolor="#CCCCCC"><div align="center"><strong>AP</strong></div></td>
                        </tr>
                        <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"  and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			preco = objrs02("preco")			
			cont	= objrs02("cont")	
			quantidade = objrs02("quantidade")		
	%>
                        <tr id="tr<%=cont%>">
                          <td bgcolor="#FFFFFF"><%
				sql = "SELECT produto FROM produtos where codpro = "&objrs02("codpro")&""
				set objrs03 = objConn03.execute(sql)
				if not objrs03.eof then
					response.Write objrs03("produto")
				else
					response.Write "-"
				end if
					%>                          </td>
                          <td bgcolor="#FFFFFF"><div align="center">
                            <input name="quant<%=cont%>" type="text" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA" size="10" maxlength="10"  value="<%=quantidade%>"  onfocus="javascript:setPointer(tr<%=cont%>, '#CCFFCC');" onBlur="javascript:setPointer(tr<%=cont%>, '#FFFFFF');" >
                          </div></td>
                          <td bgcolor="#FFFFFF"><div align="right">
                              <%
				response.Write formatnumber(preco)
				%>
                          </div></td>
                          <td bgcolor="#FFFFFF">                            <input type="checkbox" name="ap<%=cont%>" value="sim" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px" onClick="javascript:if(this.checked){setPointer(tr<%=cont%>, '#FF8888');}else{setPointer(tr<%=cont%>, '#FFFFFF');}" title="Apagar Produto: <%=produto%>"><!----></td>
                        </tr>
                        <%	
	objrs02.movenext		
  wend
  
  %>
                    </table></td>
                  </tr>
                </table>
                <%
			else
					response.Write "Nenhum produto nesse fornecedor!"
			end if%>
              </div></td>
          </tr>
        </table></td>
      </tr>
    </table>
	  <br>
	  <input type="hidden" name="fornecedor22" value="<%=fornecedor2%>">
	  <input type="submit" name="Submit22" value="Imprimir" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR:#EAEAEA">
	  <br>
	  <br>
	</div>
	</form></TD>
	</TR>
	</TABLE>
	<%
end if
%>

    <%
call fecha_conexao01
call fecha_conexao02
call fecha_conexao03
%>
    <%end if%>
</div></div>
</body>
</html>