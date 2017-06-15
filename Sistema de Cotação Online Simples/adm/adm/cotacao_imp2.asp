<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_TEMPO.asp"-->
<%
call abre_conexao01
call abre_conexao02
call abre_conexao03
DIM OBJRS01,OBJRS02,OBJRS03,OBJCONN01,OBJCONN02,OBJCONN03
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
<!--#include file="inc_css2.asp"-->


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
fornecedor = request("fornecedor2")
%>

<body <%if fornecedor <> "" then%>onLoad="javascript:if(confirm('Atualizar Tela?')){location.reload();}else{print();}"<%end if%>>

<div align="center">
  <div align="center">
  <%
  fornecedor2 = request("fornecedor2")
  if fornecedor2 = "todos" then
	sql = "SELECT codfor FROM fornecedores where fora='sim'"
	set objrs01 = objConn01.execute(sql)
  	while not objrs01.eof 
		fornecedor = objrs01("codfor")

	sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
	set objrs02 = objConn02.execute(sql)
	if not objrs02.eof then
  %> 
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
				                                
                <b>MERC. MOSSORÓ</b>&nbsp; -&nbsp; (&nbsp; ) OK</div></td>
                 </tr>
             </table></td>
            </tr>
		  <% end if%>
          <tr>
            <td height="925" valign="top">
              <div align="center">
                <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			cont = objrs02("cont")			
			if request("ap"&cont) = "sim" then
				set objrs03 = objConn03.execute("delete from temp_cotacao where cont="&cont&"")
			END IF
			objrs02.MOVENEXT
		WEND
				
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&" and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
	%>			
                  <table width="100%" border="0" bgcolor="#FFFFFF">
                    <tr>
                      <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#3C3C3C" class="linkA">
                        <tr bgcolor="#CCCCCC">
                          <td width="62%" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
                          <td width="18%" bgcolor="#CCCCCC"><div align="center"><strong>QUANTIDADE</strong></div></td>
                          <td width="20%" bgcolor="#CCCCCC"><div align="center"><strong>VALOR</strong></div></td>
                        </tr>
                        <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			preco = objrs02("preco")			
			cont = objrs02("cont")			
	%>
                        <tr>
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
                          <%				
                          aux_qtde = request("quant"&cont)
                          set objrs03 = objConn03.execute("update temp_cotacao set quantidade='"&aux_qtde&"' where cont="&cont)%>
                          <%=request("quant"&cont)%>&nbsp;</div></td>
                          <td bgcolor="#FFFFFF"><div align="right">
                            <%
				response.Write formatnumber(preco)
				%>
                          </div></td>
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
        <br>
        <br></td>
      </tr>
    </table>
	<%
	end if	
	objrs01.movenext
	wend
	
else

  %> 
    <table width="600" border="0">
<%
FORNECEDOR = REQUEST("FORNECEDOR22")
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
                    
                MOSSORÓ<input type="checkbox" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px">OK</div></td>
              </tr>
            </table></td>
            </tr>
		  <% end if%>
          <tr>
            <td heighT="925" valign="top">
              <div align="center">
                <%
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"   and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			cont = objrs02("cont")			
			if request("ap"&cont) = "sim" then
				set objrs03 = objConn03.execute("delete from temp_cotacao where cont="&cont&"")
			END IF
			objrs02.MOVENEXT
		WEND			
				
		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"  and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		if not objrs02.eof then
	%>
                <table width="100%" border="0" bgcolor="#FFFFFF">
                  <tr>
                    <td><table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#3C3C3C" class="linkA">
                        <tr bgcolor="#CCCCCC">
                          <td width="62%" bgcolor="#CCCCCC"><div align="center"><strong>PRODUTO</strong></div></td>
                          <td width="20%" bgcolor="#CCCCCC"><div align="center"><strong>QUANTIDADE</strong></div></td>
                          <td width="18%" bgcolor="#CCCCCC"><div align="center"><strong>VALOR</strong></div></td>
                        </tr>
                        <%

		sql = "SELECT * FROM temp_cotacao where codfor="&fornecedor&"  and tipo='-'"
		set objrs02 = objConn02.execute(sql)
		while not objrs02.eof 	
			preco = objrs02("preco")			
			cont	= objrs02("cont")	
            if request("ap"&cont) = "sim" then
			set objrs03 = objConn03.execute("delete from temp_cotacao where cont="&cont&"")
			else
		
	%>
                        <tr>
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
                          <%
                           aux_qtde = request("quant"&cont)
                          set objrs03 = objConn03.execute("update temp_cotacao set quantidade='"&aux_qtde&"' where cont="&cont)%>
                          <%=request("quant"&cont)%>&nbsp;</div></td>
                          <td bgcolor="#FFFFFF"><div align="right">
                              <%
				response.Write formatnumber(preco)
				%>
                          </div></td>
                        </tr>
						
                        <%	
	end if
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
	<%
end if
%>

    <%
call fecha_conexao01
call fecha_conexao02
call fecha_conexao03
%>

</div></div>
</body>
</html>