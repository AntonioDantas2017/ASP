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
%>
<%
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
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">
<div align="center"> 

<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodpro, auxapaga, listar, contador,auxverifica1,auxprincipal,filtro


'RESGATA AS VARIÁVEIS
'--------------------
auxcodpro   = request("auxcodpro")
auxapaga    = request("auxapaga")
auxapaga2    = request("auxapaga2")

pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
ordem1 	 	= request("ordem1")
listar      = request("listar")
codpro		= request("codpro")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")

if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "codpro"
end if
if ordem1="" then
   ordem1 = 0
end if
if listar="" then
   listar = "30"
end if
if filtro="" then
   filtro = "não"
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
	
	if filtro = "não" then
		sql = "SELECT * FROM produtos_aux where codpro  > 0 and tipo='-'"	
	else
		sql = "SELECT * FROM produtos where codpro  > 0 "	
	end if
	if pesquisa <> "" then
		nometratado =  preparapalavra(pesquisa)
		if pesquisa2 = "" then
			sql = sql & " and (produto Like '%"&(nometratado)&"%') "
		else
			sql = sql & " and ("&pesquisa2&" Like '%"&(nometratado)&"%') "
		end if
	end if

'	if datainicio <> "" then
'		data_conv3 = split(datainicio,"/") 'funcao utilizada para criar um vetor delimitado po "/"
'		datainicio = data_conv3(1)&"/"&data_conv3(0)&"/"&data_conv3(2) 'reorganiza data para o formato dd/mm/yy
'	end if
'	if datafim <> "" then
'		data_conv4 = split(datafim,"/") 'funcao utilizada para criar um vetor delimitado po "/"
'		datafim = data_conv4(1)&"/"&data_conv4(0)&"/"&data_conv4(2) 'reorganiza data para o formato dd/mm/yy
'	end if
	
	'if (datainicio <> "") and (datafim<>"") then
	'	sql = sql & " and data between #"&datainicio&"# and #"&datafim&"# "
	'end if
	
'	if codpro <> "" then
'	  sql = sql & " and (codpro = '"&codpro&"') "
'	end if
	
	'if filtro <> "" then
	 ' sql = sql & " and (fora = '"&filtro&"') "
	'end if	
	
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
<table width="115" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"><strong>PRODUTO</strong></div></td>
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
                          <td width="26"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="26"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                         <%'if nivel_aux = "2" then%> <td width="25"><div align="center"><strong> 
                              
                          </strong><strong><a href="produtos_inserir.asp?tipo=<%=tipo%>"></a></strong></div></td>
<td width="24"><div align="center">
						<!--  <a href="#" onClick="MM_openBrWindow('tipos.asp','','status=yes,menubar=yes,scrollbars=yes,width=790,height=572')"><img src="imagem/icones_usu.gif" alt="Categorias de downloads" width="20" height="20" border="0"></a>-->
						  </div></td>    
						  <%'end if%>                      
						  <td width="118"> <div align="center">
                            <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="produtos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" selected>LISTAR:30</option>
                              <option value="produtos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="produtos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="produtos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="produtos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>>" >LISTAR:30</option>
                              <option value="produtos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:100</option>
                              <option value="produtos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="produtos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="produtos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="produtos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="produtos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:200</option>
                              <option value="produtos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="produtos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="produtos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="produtos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="produtos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
                          <td width="208"> <div align="center">
                           <!-- <select name="ordem" id="ordem" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <option value="produtos.asp?ordem=codpro&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="codpro" then%> selected <%end if%>>ORDEM:Cod</option>
                              <option value="produtos.asp?ordem=tamanho&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="tamanho desc" then%> selected <%end if%>>ORDEM:Tamanho</option>
                              <option value="produtos.asp?ordem=nome&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="nome" then%> selected <%end if%>>ORDEM:Nome</option>							
                            </select>-->
                          </div></td>
							
<td width="200">   
						  
                            <div align="left">
                              <table width="236" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="produtos.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="tipo" type="hidden" value="<%=tipo%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">
                                  <input name="filtro2" type="hidden" value="<%=filtro2%>">								  	
                                  <input name="datainicio" type="hidden" value="<%=datainicio%>">
                                  <input name="datafim" type="hidden" value="<%=datafim%>">									  	
                                    
                                  <tr> 
                                    <td width="236">Pesq:
                                      <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="pesquisa2" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codpro" <%if pesquisa2 = "codpro" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="produto" <%if pesquisa2 = "produto" then%> selected <% end if %>>Descrição</option>
                                      </select>
                                    <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" alt="Procura" width="20" height="13" border="0"></td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                        <td width="200"><div align="right">
                          <select name="select2" id="filtro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                            <option value="produtos.asp?ordem=<%=ordem%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=sim&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if filtro="sim" then%> selected <%end if%>>Todos:Sim</option>
                            <option value="produtos.asp?ordem=<%=ordem%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=não&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if filtro="não" then%> selected <%end if%>>Todos:N&atilde;o</option>
                          </select>
                        </div></td>
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
	set objrs01 = objConn01.execute("delete from produtos WHERE codpro="&auxcodpro&"  ")
%>
	<table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong><font color="#FF0000"> 
            </font>O PRODUTO DE COD <%=auxcodpro%> FOI APAGADO COM SUCESSO!</strong></div></td>
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

	  response.write " produto(s) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 

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
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
	
	  response.write " produto(s) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 
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
  </div>
  <div align="center"> 
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#999999"><table width="700" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
            <tr valign="top" bgcolor="#FFFFFF"> 
              <td width="36" rowspan="3" ><div align="center" class="linkA5"><strong><a href="produtos.asp?ordem=<% if ordem1 = 1 then %>codpro&ordem1=0<%else%>codpro desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Cod</a></strong><%if ordem="codpro" then%><img src="imagem/cima.gif"><%else if ordem="codpro desc" then%><img src="imagem/baixo.gif">
                <%end if%>
                <%end if%>
              </div></td>
              <td width="180" rowspan="3"><div align="center" class="linkA5"><strong><a href="produtos.asp?ordem=<% if ordem1 = 1 then %>produto&ordem1=0<%else%>produto desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Descri&ccedil;&atilde;o</a></strong>
                  <%if ordem="produto desc" then%><img src="imagem/cima.gif"><%else if ordem="produto" then%><img src="imagem/baixo.gif"><%end if%>
                <%end if%>
              </div></td>
			  <td colspan="4"><div align="center" class="linkA5"><strong>&Uacute;ltima compra data :

	              <%
			  set objrs02 =  objconn02.execute("select * from temp_cotacao")
			  if not objrs02.eof then
			  		response.Write objrs02("data")
			else
					response.Write "-"
				end if
			  
			%>
			  </strong></div>			 </td>
			  <%'if nivel_aux = "2" then%><td width="30" rowspan="3"><div align="center" class="linkA5"><strong>Pre</strong></div></td>
			  <td width="30" rowspan="3"><div align="center" class="linkA5"><strong>Ap</strong></div></td>
			  <%'end if%>
            </tr>
            <tr valign="top" bgcolor="#FFFFFF">
              <td colspan="2"><div align="center" class="linkA5"><strong>Menor Pre&ccedil;o </strong></div></td>
              <td colspan="2"><div align="center" class="linkA5"><strong>Maior Pre&ccedil;o </strong></div></td>
            </tr>
            <tr valign="top" bgcolor="#FFFFFF">
              <td width="62"><div align="center" class="linkA5"><strong>Valor</strong></div></td>
              <td width="137"><div align="center" class="linkA5"><strong>Fornecedor</strong></div></td>
              <td width="63"><div align="center" class="linkA5"><strong>Valor </strong></div></td>
              <td width="137"><div align="center" class="linkA5"><strong>Fornecedor</strong></div></td>
            </tr>
<%		
	For contador = 1 to listar

		codpro		= objrs01("codpro")
		produto		=	objrs01("produto")
	
	%>
                <%if (contador mod 2)<>0 then%>
  <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')">
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
	<%end if%>
              
              
			  <td width="36" valign="middle"><div align="center"> 
                  <%response.write codpro%>
              </div></td>
              
			  <td width="180" valign="middle"> <div align="left"><strong>
			    <% 
	if produto <> "" then
		response.write produto
	else
		response.write "-"
	end if				  
	
%>
			  </strong></div>			   </td>
			  
			  <td valign="middle"><div align="center"><%
			  set objrs02 =  objconn02.execute("select * from temp_cotacao where codpro = "&codpro&" and tipo='-'")
			  if not objrs02.eof then
			  		response.Write formatnumber(objrs02("preco"))
			else
					response.Write "-"
				end if
			  
			%>
			  </div> <div align="center"></div> 	    <div align="center"> 
		                    </div></td>
			  <td valign="middle"><div align="center">
			    <% 
						codfor		= ""
			  if not objrs02.eof then
					codfor = objrs02("codfor")
					sql= "select empresa from fornecedores where codfor = "&codfor&""
			  		set objrs02 =  objconn02.execute(sql)
					if not objrs02.eof then
						response.Write objrs02("empresa")
					else
						response.Write "-"
					end if
			else
					response.Write "-"
				end if
			%>
              </div></td>
			  <%'if nivel_aux = "2" then%><td valign="middle"><div align="center">
			    <div align="center">
			      <%
			  set objrs02 =  objconn02.execute("select * from temp_cotacao where codpro = "&codpro&" and tipo='+'")
			  if not objrs02.eof then
			  		response.Write formatnumber(objrs02("preco"))
			else
					response.Write "-"
				end if
			  
			%>
			    </div> 
                   
              </div></td>
				
				<td valign="middle"><div align="center">
				  <%
			   		codfor		= ""
				  if not objrs02.eof then
					codfor = objrs02("codfor")
					sql= "select empresa from fornecedores where codfor = "&codfor&""
			  		set objrs02 =  objconn02.execute(sql)
					if not objrs02.eof then
						response.Write objrs02("empresa")
					else
						response.Write "-"
					end if
			else
					response.Write "-"
				end if
			%>
				</div></td>
				<td valign="middle"><div align="center"><a href="javascript:MM_openBrWindow('produtos_precos.asp?codpro=<%=codpro%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_precos.gif" alt="Pre&ccedil;os desse produto" width="20" height="20" border="0"></a></div></td>
			    <td valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <!-- ABAIXO EST&Aacute; A PROGRAMA&Ccedil;&Atilde;O NECESS&Aacute;RIA PARA MOSTRAR O &Iacute;CONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMA&Ccedil;&Atilde;O-->
                    <form name="form4" method="get" action="produtos.asp">
                      <input name="auxapaga" type="hidden" value="sim">
                      <input name="auxcodpro" type="hidden" value="<%=codpro%>">
                      <input name="listar2" type="hidden" value="<%=listar%>">
                      <input name="ordem2" type="hidden" value="<%=ordem%>">
                      <input name="pesquisa3" type="hidden" value="<%=pesquisa%>">
                      <input name="pesquisa22" type="hidden" value="<%=pesquisa2%>">
                      <input name="tipo2" type="hidden" value="<%=tipo%>">
                      <input name="filtro3" type="hidden" value="<%=filtro%>">
                      <input name="filtro22" type="hidden" value="<%=filtro2%>">
                      <input name="datainicio2" type="hidden" value="<%=datainicio%>">
                      <input name="datafim2" type="hidden" value="<%=datafim%>">
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
	            </table></td>
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
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td></td>
        <td width="110"><div align="right"> 
            <%
if (Session("pagina_downloads") <> 1 )or(Session("pagina_downloads") <> 1)or(Session("pagina_downloads") <> objrs01.PageCount )or(Session("pagina_downloads") <> objrs01.PageCount) then

	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""produtos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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