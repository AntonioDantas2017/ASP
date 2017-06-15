<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_mercado
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

<html>

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
<style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0"  onLoad="form1.pesquisa.focus()">
<div align="center"> 

<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodpag, auxapaga, listar, contador,auxverifica1,auxprincipal,filtro


'RESGATA AS VARIÁVEIS
'--------------------
auxcodpag   = request("auxcodpag")
auxapaga    = request("auxapaga")
auxapaga2    = request("auxapaga2")

auxpago		= request("auxpago")
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
ordem1 	 	= request("ordem1")
listar      = request("listar")
codpag		= request("codpag")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")

if auxpago="não" then
	set objrs01 = objConn01.execute("update pagamentos set pago='não' WHERE codpag="&auxcodpag&" ")
end if
if auxpago="sim" then
	set objrs01 = objConn01.execute("update pagamentos set pago='sim' WHERE codpag="&auxcodpag&" ")
end if


if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "codpag"
end if
if ordem1="" then
   ordem1 = 0
end if
if listar="" then
   listar = "30"
end if

if datainicio = "" then
	datainicio = date - 60
end if
if datafim = "" then
	datafim = date + 60
end if
%>
<%
'FUNÇÃO PARA PESQUISA
'--------------------
function preparaPalavra(str)
   preparaPalavra = replace(str,"a","[a,á,à,ã,â,ä]")
   preparaPalavra = replace(preparaPalavra,"e","[e,é,è,ê,ë]")	
   preparaPalavra = replace(preparaPalavra,"i","[i,í,ì,î,ï]")
   preparaPalavra = replace(preparaPalavra,"o","[o,ó,ò,õ,ô,ö]")
   preparaPalavra = replace(preparaPalavra,"u","[u,ú,ù,û,ü]")
   preparaPalavra = preparaPalavra
end function
%>
<%
'ROTINA DE SQL
'-------------
sub rotinasql

	sql = "SELECT * FROM pagamentos2 where codpag  > 0 "	

	if pesquisa <> "" then
		nometratado =  preparapalavra(pesquisa)
		if pesquisa2 = "" then
			sql = sql & " and (nome Like '%"&(nometratado)&"%')"
		else
			sql = sql & " and ("&pesquisa2&" Like '%"&(nometratado)&"%') "
		end if
	end if

	if datainicio <> "" then
		data_conv3 = split(datainicio,"/") 'funcao utilizada para criar um vetor delimitado po "/"
		datainicio = data_conv3(1)&"/"&data_conv3(0)&"/"&data_conv3(2) 'reorganiza data para o formato dd/mm/yy
	end if
	if datafim <> "" then
		data_conv4 = split(datafim,"/") 'funcao utilizada para criar um vetor delimitado po "/"
		datafim = data_conv4(1)&"/"&data_conv4(0)&"/"&data_conv4(2) 'reorganiza data para o formato dd/mm/yy
	end if
	
	if (datainicio <> "") and (datafim<>"") then
		sql = sql & " and data_venc between #"&datainicio&"# and #"&datafim&"# "
	end if
	
'	if codpag <> "" then
'	  sql = sql & " and (codpag = '"&codpag&"') "
'	end if
	
	if filtro <> "" then
	  sql = sql & " and (pago = '"&filtro&"') "
	end if	
	
	'if filtro2 <> "" then
	'  sql = sql & " and (tipo = "&filtro2&") "
	'end if		
				
	sql = sql & " order by " & ordem
	
'	response.Write sql
	objrs01.Open Sql, objConn01
	
end sub
%>
  <%



'AVISO PARA RESULTADO DA PESQUISA EM BRANCO
'------------------------------------------
sub vazio
	response.write "<br><br><br><br><br><br><br>"		
	response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class=linka><tr><td><div align='center'><b><font face='Verdana, Arial, Helvetica, sans-serif'>Não foi encontrado nenhum registro.</font> </b></div></td></tr></table>"
end sub
%>
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="115">
<table width="115" height="40" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="43" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>PAGAMENTOS</strong></div></td>
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
                          <td width="20"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="20"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                         <%if nivel_aux = "2" then%> <td width="20"><div align="center"><strong> 
                              
                          </strong><strong><a href="pagamentos_inserir.asp?tipo=<%=tipo%>"><img src="imagem/icones_insere.gif" alt="Insere novo cliente" width="20" height="20" border="0"></a></strong></div></td><%end if%>
<td width="4"><div align="center">
						<!--  <a href="#" onClick="MM_openBrWindow('tipos.asp','','status=yes,menubar=yes,scrollbars=yes,width=790,height=572')"><img src="imagem/icones_usu.gif" alt="Categorias de downloads" width="20" height="20" border="0"></a>-->
						  </div></td>    
						  <%'end if%>                      
						  <td width="103"> <div align="center">
                            <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="pagamentos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" selected>LISTAR:30</option>
                              <option value="pagamentos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="pagamentos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="pagamentos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="pagamentos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>>" >LISTAR:30</option>
                              <option value="pagamentos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:100</option>
                              <option value="pagamentos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="pagamentos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="pagamentos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="pagamentos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="pagamentos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:200</option>
                              <option value="pagamentos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="pagamentos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="pagamentos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="pagamentos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="pagamentos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
                          <form action="pagamentos.asp" method="get" name="form3">
                            <input name="listar" type="hidden" value="<%=listar%>">
                            <input name="ordem" type="hidden" value="<%=ordem%>">
                            <input name="tipo" type="hidden" value="<%=tipo%>">
                            <input name="filtro" type="hidden" value="<%=filtro%>">
                            <input name="filtro2" type="hidden" value="<%=filtro2%>">
                            <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                            <input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
                          <td width="307"> 
                            <div align="left">
                              <!-- <select name="ordem" id="ordem" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <option value="cotacao.asp?ordem=codcot&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="codcot" then%> selected <%end if%>>ORDEM:Cod</option>
                              <option value="cotacao.asp?ordem=tamanho&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="tamanho desc" then%> selected <%end if%>>ORDEM:Tamanho</option>
                              <option value="cotacao.asp?ordem=nome&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="nome" then%> selected <%end if%>>ORDEM:Nome</option>							
                            </select>-->
                              &nbsp;De:
                              <input name="datainicio" type="text" id="datainicio"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=datainicio%>" size="8" maxlength="10">
                              at&eacute;:
                              <input name="datafim" type="text" id="datafim"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=datafim%>" size="8" maxlength="10">
  <input name="imageField32" type="image" id="imageField3" src="imagem/icones_ok.gif" alt="Procura Usu&aacute;rio (Nome, Login)" width="20" height="13" border="0">
                              </div></td>
						  </form>
							<td width="102"><div align="center">
							  <select name="select2" id="filtro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
							    <option value="pagamentos.asp?ordem=<%=ordem%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if filtro="" then%> selected <%end if%>>Todos</option>
							    <option value="pagamentos.asp?ordem=<%=ordem%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=não&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if filtro="não" then%> selected <%end if%>>Não Pagos</option>
							    <option value="pagamentos.asp?ordem=<%=ordem%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=sim&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if filtro="sim" then%> selected <%end if%>>Pagos</option>
						      </select>
						    </div></td>
                        </tr>
                        <tr>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td width="307">   
						  
                            <div align="left">
                              <table width="303" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="pagamentos.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="tipo" type="hidden" value="<%=tipo%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">
                                  <input name="filtro2" type="hidden" value="<%=filtro2%>">								  	
                                  <input name="datainicio" type="hidden" value="<%=datainicio%>">
                                  <input name="datafim" type="hidden" value="<%=datafim%>">									  	
                                    
                                  <tr> 
                                    <td width="303">&nbsp;Pesq:
                                      <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="select" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codloc" <%if pesquisa2 = "codloc" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="nome" <%if pesquisa2 = "nome" then%> selected <% end if %>>Nome</option>
                                        <option value="data_venc" <%if pesquisa2 = "data_venc" then%> selected <% end if %>>Data Vencimento</option>
                                        <option value="valor" <%if pesquisa2 = "valor" then%> selected <% end if %>>Valor</option>
                                       
                                      </select>
                                    <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" alt="Procura" width="20" height="13" border="0"></td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>

                          <td width="102">&nbsp;</td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  
  <div align="center">
    <%
if auxapaga="sim" then
	set objrs01 = objConn01.execute("delete from pagamentos WHERE codpag="&auxcodpag&"  ")
%>
	<table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong><font color="#FF0000"> 
            </font>O REGISTRO DE COD <%=auxcodpag%> FOI APAGADO COM SUCESSO!</strong></div></td>
      </tr>
    </table>
    <%	
end if
%>
    <%
'SE É A PRIMEIRA VEZ QUE A PÁGINA É CARREGADA
'--------------------------------------------
If Session("primeiravez_downloads") <> "Nao" then 
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	call rotinasql
	if objrs01.RecordCount<>"0" then
		session("pagina_downloads") = 1 
		CALL MostraDados 
		Session("primeiravez_downloads") = "Nao"
	else
		CALL Vazio
	end if

'SE A PÁGINA JÁ FOI CARREGADA
'----------------------------
Else
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	call rotinasql
	if objrs01.RecordCount<>"0" then
		if Request("Navegacao") = "" then
			Session("pagina_downloads") = 1 
		end if
		if Request("Navegacao") = "Primeira" then
			Session("pagina_downloads") = 1 
		end if
		if Request("Navegacao") = "Proxima" then
			Session("pagina_downloads") = Session("pagina_downloads") + 1 
		end if
		If Request("Navegacao") = "Anterior" then
			Session("pagina_downloads") = Session("pagina_downloads") - 1
		End If
		if Request("Navegacao") = "Ultima" then
			Session("pagina_downloads") = objrs01.PageCount
		end if
		CALL MostraDados
	else
		CALL Vazio
	end if
End If 

'MENU DE PAGINAÇÃO
'-----------------
sub menu
End sub
%>
    <%
Sub MostraDados()
	objrs01.AbsolutePage = Session("pagina_downloads") ' vai para o número da página que está armazenado em session("pagina")
%>
	
    <%
if (Session("pagina_downloads") <> 1 )or(Session("pagina_downloads") <> 1)or(Session("pagina_downloads") <> objrs01.PageCount )or(Session("pagina_downloads") <> objrs01.PageCount) then
%>
    <table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class=linka>
                        <tr> 
                          <td width="50%" bgcolor="#F4F4F4"><div align="center"><strong> 
							<%
	Response.Write "Existe(m) " & objrs01.RecordCount

	  response.write " pagamento(s) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 

							%>	
						  </strong></div></td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
			  </td>
            </tr>
          </table></td>
        <td width="110"><div align="right">
            <%
if (Session("pagina_downloads") <> 1 )or(Session("pagina_downloads") <> 1)or(Session("pagina_downloads") <> objrs01.PageCount )or(Session("pagina_downloads") <> objrs01.PageCount) then

	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'>"
	End If
	
End if
%>
          </div></td>
      </tr>
    </table>
    <%
else
%>
	<table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><table width="700" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class=linka>
                  <tr> 
                    <td width="50%" bgcolor="#F4F4F4"><div align="center"><strong> 
                        <%
	Response.Write "Existe(m) " & objrs01.RecordCount
	
	  response.write " pagamento(s) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 
	%>
                        </strong></div>
					</td>
                  </tr>
                </table></td>
            </tr>
          </table>
          <div align="right"> </div></td>
      </tr>
    </table>
    <%
end if
%>
    <table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0">
      <tr>
        <td><table width="700" border="0" cellpadding="1" cellspacing="1" class="linkA">
            <tr bgcolor="#FFFFFF">
              <td width="23">&nbsp;</td>
              <td width="21"  bgcolor="#FFFEBF"></td>
              <td width="196">PAGAMENTO APROXIMADO </td>
              <td width="20" bgcolor="#FFB7B7">&nbsp;</td>
              <td width="202">PAGAMENTO ATRASADO </td>
              <td width="20">&nbsp;</td>
              <td width="196">&nbsp;</td>
            </tr>
        </table></td>
      </tr>
    </table>
    <table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
  </div>
  <div align="center"> 
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#999999"><table width="700" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
            <tr valign="top" bgcolor="#FFFFFF"> 
              <td width="109" ><div align="center" class="linkA5"><strong><a href="pagamentos.asp?ordem=<% if ordem1 = 1 then %>data_venc&ordem1=0<%else%>data_venc desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Data Vencimento </a></strong>
                    <%if ordem="data_venc desc" then%>
                <img src="imagem/cima.gif">
                <%else if ordem="data_venc" then%>
                <img src="imagem/baixo.gif">
                <%end if%>
                    <%end if%>
              </div></td>
              <td width="116"><div align="center" class="linkA5"><strong><a href="pagamentos.asp?ordem=<% if ordem1 = 1 then %>pago&ordem1=0<%else%>pago desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Pago </a></strong>
                    <%if ordem="pago desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="pago" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="89"><div align="center" class="linkA5"><strong> <a href="pagamentos.asp?ordem=<% if ordem1 = 1 then %>valor&ordem1=0<%else%>valor desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"  class="linkA5">Valor</a></strong>
                    <%if ordem="valor" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="valor desc" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
			  <td width="317"><div align="center" class="linkA5"><strong> <a href="pagamentos.asp?ordem=<% if ordem1 = 1 then %>nome&ordem1=0<%else%>nome desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"  class="linkA5">Locat&aacute;rio</a></strong>
                    <%if ordem="nome" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="nome desc" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>			  
			  <%'if nivel_aux = "2" then%>
			  <td width="27"><div align="center" class="linkA5"><strong>Con</strong></div></td>
			  <td width="23"><div align="center" class="linkA5"><strong>Ap</strong></div></td><%'end if%>
            </tr>
<%	
	valor_tot = 0	
	For contador = 1 to listar
		codpag		= objrs01("codpag")
		data_venc	=	objrs01("data_venc")
		if isdate(data_venc) then
			data2	= split(data_venc,"/")
			data_venc 	= cdate(right("0"&data2(0),2) & "/" & right("0"&data2(1),2) & "/" & data2(2))
		end if
		valor		=	objrs01("valor")		
		if isnumeric(valor) then
			valor = formatnumber(valor,2)
			valor_tot = valor_tot + valor
		end if
		pago = objrs01("pago")
		nome = objrs01("nome")
		codcon		= objrs01("codcon")


%>
                    <%
					  if (data_venc < date) AND (pago = "não") then
					%>
           
		    <tr valign="middle" bgcolor="#FFB7B7" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFB7B7')"> 
              <%else
if ((data_venc - 15) < date) AND (pago = "não") then
			  %>
            <tr valign="middle" bgcolor="#FFFEBF"  onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFEBF')"> 
			<%else%>
			 <%'if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%'else%>
           <!-- <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> -->
              <%'end if%>
              <%end if
			    end if%>
 
			  <td width="109" valign="middle"><div align="center"><strong>
                <% 
	if data_venc <> "" then
		data2	= split(data_venc,"/")
		data_venc 	= (right("0"&data2(0),2) & "/" & right("0"&data2(1),2) & "/" & data2(2))
		response.write data_venc
	else
		response.write "-"
	end if				  
	
%>
              </strong></div></td>
              
			  <td width="116" valign="middle"><div align="center">
			    <table width="90%" border="0" cellspacing="0" cellpadding="0" class=linka>
                  <tr>
                    <td width="31%"><div align="right">
                        <%if nivel_aux = "2" then%>
                        <%if pago="sim" then%>
                        <a href="pagamentos.asp?auxpago=não&auxcodpag=<%=codpag%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="imagem/icones_confirma.gif" alt="Não Pago" width="20" height="20" border="0"></a>
                        <%else%>
                        <a href="pagamentos.asp?auxpago=sim&auxcodpag=<%=codpag%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/icones_cancela.gif" alt="Pago" width="20" height="20" border="0"></a>
                        <%end if%>
                        <%end if%>
                    </div></td>
                    <td width="69%"><%
IF pago = "sim" then
	response.write "&nbsp;"
	response.write "PAGO."
else
	response.write "&nbsp;"
	response.write "NÃO PAGO."
end if
%>
                        <div align="left"></div></td>
                  </tr>
                </table>
			  </div></td>
			  
			  <td valign="middle">
			    
		        <div align="right">
		          <% 
response.write (valor)
%>			  
                </div></td>
			  <td valign="middle">
                <div align="left">
                  <% 
response.write ucase(nome)
%>
                  </div>
                  <div align="center"> 
                    <a href="downloads_ver.asp?codpag=<%=codpag%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"></a> 
                  </div></td><td width="27" valign="middle"><div align="center"><a href="contratos.asp?codcon=<%=codcon%>"><img src="imagem/icones_edita.gif" alt="Edita esta Cota&ccedil;&atilde;o" width="20" height="20" border="0"></a></div></td>
				
				<td width="23" valign="middle"><div align="center"> 
                  
				  <table width="100%" border="0" cellspacing="0" cellpadding="0">
					<!-- ABAIXO ESTÁ A PROGRAMAÇÃO NECESSÁRIA PARA MOSTRAR O ÍCONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMAÇÃO-->
					<form name="form4" method="get" action="pagamentos.asp">
					  <input name="auxapaga" type="hidden" value="sim">
					  <input name="auxcodpag" type="hidden" value="<%=codpag%>">
					  <input name="listar" type="hidden" value="<%=listar%>">
					  <input name="ordem" type="hidden" value="<%=ordem%>">
					  <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
					  <input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">					  
					  <input name="tipo" type="hidden" value="<%=tipo%>">
					  <input name="filtro" type="hidden" value="<%=filtro%>">					  
					  <input name="filtro2" type="hidden" value="<%=filtro2%>">	
 					  <input name="datainicio" type="hidden" value="<%=datainicio%>">
					  <input name="datafim" type="hidden" value="<%=datafim%>">							  				  					  
					  
					  <tr> 
						<td><div align="center"> 
						  <script language="JavaScript">
							function confirmar()
							{
							   return (confirm('Deseja apagar este registro?'))
							}
						  </script>
						  <%if nivel_aux = "2" then%>
						  <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar este cliente" WIDTH="20" HEIGHT="20">
						  <%end if%>
						</div></td>
					  </tr>
					</form>
				  </table>
				  
              </div></td>
			  <%'end if%>
            </tr>
            <%
	objrs01.MoveNext
	If objrs01.Eof then Exit For
		Next
		
%>
          </table></td>
      </tr>
    </table>
    <table width="50%" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0">
      <tr>
        <td><div align="right">
          <table width="200" border="0" bgcolor="#DFDFDF">
            <tr>
              <td><table width="200" border="0" cellpadding="1" cellspacing="1" class="linkA">
                <tr bgcolor="#784E4B">
                  <td bgcolor="#FF9999"><div align="center"><strong>VALOR TOTAL </strong></div></td>
                </tr>
                <tr bgcolor="#FFFFFF">
                  <td><div align="center"><%=FORMATNUMBER(valor_tot,2)%></div></td>
                </tr>
              </table></td>
            </tr>
          </table>
        </div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td></td>
        <td width="110"><div align="right"> 
            <%
if (Session("pagina_downloads") <> 1 )or(Session("pagina_downloads") <> 1)or(Session("pagina_downloads") <> objrs01.PageCount )or(Session("pagina_downloads") <> objrs01.PageCount) then

	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""pagamentos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'>"
	End If
	
End if
%>
          </div></td>
      </tr>
    </table>
    <%
		CALL fecha_conexao01
		CALL fecha_conexao02
	%>
    <%
End Sub
%>
  </div>
</div>
</body>
</html>