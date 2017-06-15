<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
%>
<!--#include file="inc_css.asp"-->

<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" onLoad="MM_goToURL('parent.frames[\'topFrame\']','sistema2.asp?barra=news');return document.MM_returnValue">
<div align="center"> 
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="115"> <table width="115" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong> CONTRATOS</strong></div></td>
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
                      <table width="100%" border="0" cellpadding="0" cellspacing="0" class=linka>
                        <tr> 
                          <td width="30"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"></div></td>
                          <td width="30"><div align="center"><strong></strong></div></td>
                          <td width="120"> <div align="center"></div></td>
                          <td width="130"> <div align="center"></div></td>
                          <td width="130"> <div align="right"></div></td>
                          <td width="420"><div align="right"></div>
                            <div align="center"></div>
                            <div align="center"><strong> </strong></div>
                            <div align="center"></div>
                            <div align="right"> </div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <br>
  <table width="700" border="0" cellspacing="0" cellpadding="0" class=linka>
    <tr> 
      <td height="13" bgcolor="#CCCCCC"><div align="center"><strong> CONTRATO </strong></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="2" cellpadding="2" CLASS=LINKA>
          <tr> 
            <td bgcolor="#FFFFFF"><div align="center"><br>
                <%
'RESGATA AS VARIÁVEIS
'--------------------
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
listar      = request("listar")
codcli		= request("codcli")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")
codcon    	= request("codcon")
acao		= request("acao")


codcon		= TRIM(Request ("codcon"))
locatario 	= TRIM(Request ("locatario"))
imovel 		= TRIM(Request ("imovel"))
auxcodloc 	= TRIM(Request ("locatario"))
auxcodimo 	= TRIM(Request ("imovel"))
periodo 	= TRIM(Request ("periodo"))
dataini 	= TRIM(Request ("dataini"))
datafim 	= TRIM(Request ("datafim"))
aluguel 	= TRIM(Request ("aluguel"))
vencimento 	= TRIM(Request ("vencimento"))
multa 		= TRIM(Request ("multa"))
destinado 	= TRIM(Request ("destinado"))
reajuste 	= TRIM(Request ("reajuste"))
impressao 	= TRIM(Request ("impressao"))
testemunha1 = TRIM(Request ("testemunha1"))
testemunha2	= TRIM(Request ("testemunha2"))
reajuste	= TRIM(Request ("reajuste"))

sql = "update contratos set  multa='"&multa&"',destinado='"&destinado&"',testemunha1='"&testemunha1&"',testemunha2='"&testemunha2&"',via=1 where codcon = "&codcon&""
set objRs01 = objConn01.execute(sql)

sql = "update contratos set renovado='n' where codcon = "&codcon&""

if isdate(impressao) then
	sql = "update contratos set impressao='"&impressao&"' where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if


if isnumeric(periodo) then
	sql = "update contratos set periodo="&periodo&" where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isnumeric(auxcodloc) then
	sql = "update contratos set codloc="&auxcodloc&" where codcon="&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isnumeric(auxcodimo) then
	sql = "update contratos set codimo="&auxcodimo&" where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isnumeric(vencimento) then
	sql = "update contratos set vencimento="&vencimento&" where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isnumeric(aluguel) then
	sql = "update contratos set aluguel='"&aluguel&"' where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isdate(dataini) then
	sql = "update contratos set dataini='"&dataini&"' where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isdate(datafim) then
	sql = "update contratos set datafim='"&datafim&"' where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if isdate(reajuste) then
	sql = "update contratos set reajuste='"&reajuste&"' where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if


if (acao  <> "renovar") and (acao  <> "inserir") then
	sql = "update contratos set via=via+1 where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
	
	sql = "delete from  pagamentos where codcon = "&codcon&""
	set objRs01 = objConn01.execute(sql)
end if

if vencimento > 31 then
	vencimento = 31
end if

if isdate(dataini) and isnumeric(periodo) and isnumeric(vencimento) then
	d = split(dataini,"/")
	mes = d(1)
	ano = d(2)
	cont = 1
	
	cont = 0
	venc_old = vencimento	
	mes = int(mes)
	passou = "sim"
	while int(cont) <= int(periodo)
		if vencimento >= 28 then
			if (mes = 2) and (vencimento>=29) and (passou="sim") then
				if (ano / 4 = 0) and vencimento >= 29 then
					vencimento = 29
				else
					vencimento = 28	
				end if	
				'response.Write mes
			else
				if ((mes = 4) or (mes = 6) or (mes = 9) or (mes = 11)) and (vencimento=31) and (passou="sim") then
					vencimento = 30
				else
					vencimento = venc_old	
				end if		
			end if	
		end if

		
		data_venc = vencimento & "/" & mes & "/" & ano
		'passou = "não"
		'while not isdate(data_venc) 
			'vencimento = vencimento + 1 
			'if vencimento >= 31 then
			'	vencimento = 1
			'	mes = mes + 1 
			'end if
			
		'	data_venc = vencimento & "/" & mes & "/" & ano
			
			'passou = "sim"
		'wend
		sql = "INSERT into pagamentos (codcon,data_venc,valor,pago) values ("&codcon&",'"&data_venc&"','"&aluguel&"','não') "
		'response.Write sql&"<BR>"
		set objRs01 = objConn01.execute(sql)
		'response.Write data_venc&"<BR>"
		cont = cont + 1
		mes = mes + 1		
		if int(mes) >= 13 then
			mes = 1
			ano = ano + 1
		end if

	wend
else
	response.Write "DATA(S) INFORMADA(S) ERRADA(S)"
end if


%>
                <br>				
                <div align="center"><strong> <br>
                  <br>
                  O cadastro do contrato foi 
                  <% if renovar<>"sim" then
				  
				    sql = "update contratos set renovado='s' where codcon = "&codcon&""
					set objRs01 = objConn01.execute(sql)
				  
				   %>
                  alterado 
                  <% else %>
                  renovado 
                  <% END IF %>
                  com sucesso!<br>
                  <br>
                  <br>
                  <br>
                  <br>
                  <br>
                  <a href="contratos.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/botoes_principal.gif" alt="volta para a p&aacute;gina principal" width="75" height="19" border="0"></a></strong></div>
                <br>
                <br>
              </div></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <%call fecha_conexao01%>
</div>
</body>
</html>