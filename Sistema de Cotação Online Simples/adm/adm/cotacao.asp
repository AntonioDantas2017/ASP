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
.style1 {COLOR: #0000FF; FONT-FAMILY: verdana; FONT-SIZE: 11px; TEXT-DECORATION: none; font-weight: bold; }
-->
</style>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">
<div align="center"> 

<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodcot, auxapaga, listar, contador,auxverifica1,auxprincipal,filtro


'RESGATA AS VARIÁVEIS
'--------------------
auxcodcot   = request("auxcodcot")
auxapaga    = request("auxapaga")
auxapaga2    = request("auxapaga2")

pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
ordem1 	 	= request("ordem1")
listar      = request("listar")
codcot		= request("codcot")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")
auxbloqueado= request("auxbloqueado")


if auxbloqueado="não" then
	set objrs01 = objConn01.execute("update cotacao set aberta='não' ")
end if
if auxbloqueado="sim" then
	set objrs01 = objConn01.execute("update cotacao set aberta='sim' WHERE codcot="&auxcodcot&" ")
end if



if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "codcot"
end if
if ordem1="" then
   ordem1 = 0
end if
if listar="" then
   listar = "30"
end if

if datainicio = "" then
	datainicio = DATE - 360
end if

if datafim = "" then
	datafim = Date
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

	sql = "SELECT * FROM cotacao where codcot  > 0 "	

	datainicio2 = split(datainicio,"/")
	datainicio = datainicio2(1) & "/" &  datainicio2(0) & "/" &  datainicio2(2)
	datafim2 = split(datafim,"/")
	datafim = datafim2(1) & "/" &  datafim2(0) & "/" &  datafim2(2)	

	if (datainicio <> "") and (datafim<>"") then
		sql = sql & " and data between #"&datainicio&"# and #"&datafim&"# "
	end if
	
'	if codcot <> "" then
'	  sql = sql & " and (codcot = '"&codcot&"') "
'	end if
	
'	if filtro <> "" then
'	  sql = sql & " and (fora = '"&filtro&"') "
'	end if	
	
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
                  <td bgcolor="#BBECB1"><div align="center"><strong>
				    COTA&Ccedil;&Atilde;O
				  </strong></div></td>
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
                          <td width="23"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="23"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                         <%if nivel_aux = "2" then%> <td width="22"><div align="center"><strong> 
                              
                          </strong><strong><a href="cotacao_inserir.asp?inserir=sim"><img src="imagem/icones_insere.gif" alt="Insere nova Cota&ccedil;&atilde;o" width="20" height="20" border="0"></a></strong></div></td><%end if%>
<td width="14"><div align="center">
						<!--  <a href="#" onClick="MM_openBrWindow('tipos.asp','','status=yes,menubar=yes,scrollbars=yes,width=790,height=572')"><img src="imagem/icones_usu.gif" alt="Categorias de downloads" width="20" height="20" border="0"></a>-->
						  </div></td>    
						  <%'end if%> 
						                       
						  <td width="111"> <div align="center">
                            <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="cotacao.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" selected>LISTAR:30</option>
                              <option value="cotacao.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="cotacao.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="cotacao.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="cotacao.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>>" >LISTAR:30</option>
                              <option value="cotacao.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:100</option>
                              <option value="cotacao.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="cotacao.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="cotacao.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="cotacao.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="cotacao.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:200</option>
                              <option value="cotacao.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="cotacao.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="cotacao.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="cotacao.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="cotacao.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
						  <form action="cotacao.asp" method="get" name="form3">
                            <input name="listar" type="hidden" value="<%=listar%>">
                            <input name="ordem" type="hidden" value="<%=ordem%>">
                            <input name="tipo" type="hidden" value="<%=tipo%>">
                            <input name="filtro" type="hidden" value="<%=filtro%>">
                            <input name="filtro2" type="hidden" value="<%=filtro2%>">
                            <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                            <input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
                          <td width="242"> <div align="center">
                           <!-- <select name="ordem" id="ordem" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <option value="cotacao.asp?ordem=codcot&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="codcot" then%> selected <%end if%>>ORDEM:Cod</option>
                              <option value="cotacao.asp?ordem=tamanho&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="tamanho desc" then%> selected <%end if%>>ORDEM:Tamanho</option>
                              <option value="cotacao.asp?ordem=nome&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" <%if ordem="nome" then%> selected <%end if%>>ORDEM:Nome</option>							
                            </select>-->
                          De:
                            <input name="datainicio" type="text" id="datainicio"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=datainicio%>" size="8" maxlength="10">
At&eacute;:
<input name="datafim" type="text" id="datafim"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=datafim%>" size="8" maxlength="10">
<input name="imageField32" type="image" id="imageField3" src="imagem/icones_ok.gif" alt="Procura Usu&aacute;rio (Nome, Login)" width="20" height="13" border="0">
                          </div></td>
						  </form>
<td width="142" valign="middle">   
						  
                            <div align="left">
                              <!--<table width="291" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="cotacao.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="tipo" type="hidden" value="<%=tipo%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">
                                  <input name="filtro2" type="hidden" value="<%=filtro2%>">								  	
                                  <input name="datainicio" type="hidden" value="<%=datainicio%>">
                                  <input name="datafim" type="hidden" value="<%=datafim%>">									  	
                                    
                                  <tr> 
                                    <td>Pesq:
                                      <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="pesquisa2" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codcot" <%if pesquisa2 = "codcot" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="vendedor" <%if pesquisa2 = "nome" then%> selected <% end if %>>Nome</option>
                                        <option value="empresa" <%if pesquisa2 = "descricao" then%> selected <% end if %>>Empresa</option>										
                                        <option value="obs" <%if pesquisa2 = "obs" then%> selected <% end if %>>OBS</option>																				
                                      </select>
                                      <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" alt="Procura" width="20" height="13" border="0"></td>
                                  </tr>
                                </form>
                              </table>-->
                      <a href="cotacao.asp?auxbloqueado=não&auxcodcot=<%=codcot%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="IMAGEM/icones_cancela.gif" alt="Abrir cota&ccedil;&atilde;o" width="20" height="20" border="0">BLOQUEAR TUDO</a>                         </div></td>
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
	set objrs01 = objConn01.execute("delete from cotacao WHERE codcot="&auxcodcot&"  ")
	set objrs01 = objConn01.execute("select * from produtos_cot WHERE codcot="&auxcodcot&"  ")	
	while not objrs01.eof 
		set objrs02 = objConn02.execute("delete from prod_cot_forn WHERE codprc="&objrs01("codprc")&"  ")	
		objrs01.movenext
	wend
	set objrs01 = objConn01.execute("delete from produtos_cot WHERE codcot="&auxcodcot&"  ")		
	
	set obj=server.createobject("Scripting.FileSystemObject")
	dir = "\inetpub\wwwroot\geral\caixa\fornecedores\"&auxcodcot
	obj.DeleteFolder dir, FALSE



%>
	<table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong><font color="#FF0000"> 
            </font>O COTACAO DE COD <%=auxcodcot%> FOI APAGADO COM SUCESSO!</strong></div></td>
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

	  response.write " cotacao(ões) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 

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
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
	
	  response.write " cotacao(ões) cadastrado(s) - Mostrando página " & Session("pagina_downloads") & " de " & objrs01.PageCount 
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
              <td width="54" ><div align="center" class="linkA5"><strong><a href="cotacao.asp?ordem=<% if ordem1 = 1 then %>codcot&ordem1=0<%else%>codcot desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Cod</a></strong><%if ordem="codcot" then%><img src="imagem/cima.gif"><%else if ordem="codcot desc" then%><img src="imagem/baixo.gif">
                <%end if%>
                <%end if%>
              </div></td>
              <td width="70"><div align="center" class="linkA5"><strong><a href="cotacao.asp?ordem=<% if ordem1 = 1 then %>data&ordem1=0<%else%>data desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Data</a></strong>
                  <%if ordem="data desc" then%><img src="imagem/cima.gif"><%else if ordem="data" then%><img src="imagem/baixo.gif"><%end if%>
                <%end if%>
              </div></td>
			  <td width="108"><div align="center" class="style1">Quantidade de Produtos </div>
		      <div align="center" class="linkA5"></div>			    <div align="center" class="linkA5"></div></td>
			  <td width="167"><div align="center" class="style1">Quantidade de Fornecedores </div></td>
			  <td width="183"><div align="center" class="linkA5"><strong><a href="cotacao.asp?ordem=<% if ordem1 = 1 then %>aberta&ordem1=0<%else%>aberta desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Aberta</a></strong>
			      <%if ordem="aberta" then%><img src="imagem/cima.gif"><%else if ordem="aberta desc" then%><img src="imagem/baixo.gif"><%end if%>
			    <%end if%>
			  </div>			  </td>			  
			  <td width="26"><div align="center" class="linkA5"><strong>Imp</strong></div></td>
			  <td width="24"><div align="center" class="linkA5"><strong>Obs</strong></div></td>
			  <%'if nivel_aux = "2" then%><td width="20"><div align="center" class="linkA5"><strong>Ed</strong></div></td>
			  <td width="20"><div align="center" class="linkA5"><strong>Ap</strong></div></td><%'end if%>
            </tr>
<%		
	For contador = 1 to listar
		codcot		= 	objrs01("codcot")
		data		=	objrs01("data")
		aberta		=	objrs01("aberta")
%>

			<%if aberta = "não" then%>
			<tr valign="middle" bgcolor="#FFC8C8" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFC8C8')"> 
	<%else%>
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
	<%end if%>
              
			  <td width="54" valign="middle"><div align="center"> 
                <a href="javascript:MM_openBrWindow('cotacao_produtos_imp.asp?codcot=<%=codcot%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA">
			 <%response.write codcot%></a>
              </div></td>
              
			  <td width="70" valign="middle"> <div align="left">
			    <div align="center"><strong>
			      <% 
	if data <> "" then
		data_conv4 = split(data,"/") 'funcao utilizada para criar um vetor delimitado po "/"
		data = right("0"&data_conv4(0),2)&"/"&right("0"&data_conv4(1),2)&"/"&data_conv4(2) 'reorganiza data para o formato dd/mm/yy
		response.write data
	else
		response.write "-"
	end if				  
	
%>
			      </strong>			      </div>
			  </div></td>
			  
			  <td valign="middle"><div align="center">
                <% 
set objrs02 =objconn02.execute("select codprc from produtos_cot where codcot="&codcot&"")
if not objrs02.eof then
	response.Write right("00"&objrs02.recordcount,3)
else
	response.Write "00"
end if
%>
              </div>
			    <div align="center"></div></td>
			  <td valign="middle"><div align="center">
                <% 
set objrs02 =objconn02.execute("select codfor from prod_cot_forn2 where codcot="&codcot&" group by codfor")
if not objrs02.eof then
	response.Write right("0"&objrs02.recordcount,2)
else
	response.Write "00"
end if
%>
              </div></td>
			  <td valign="middle"> <div align="center">
			                   <table width="90%" border="0" cellspacing="0" cellpadding="0" class=linka>
                  <tr>
                    <td width="31%"><div align="right">
                        <%if nivel_aux = "2" then%>
						<%
						'datainicio2 = split(datainicio,"/")
						'datainicio = datainicio2(1) & "/" & datainicio2(0) & "/" & datainicio2(2)
						'datafim2 = split(datafim,"/")
						'datafim = datafim2(1) & "/" & datafim2(0) & "/" & datafim2(2)						
					    if aberta="sim" then%>
                                                <%else%>
                                                <a href="cotacao.asp?auxbloqueado=sim&auxcodcot=<%=codcot%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA"><img src="imagem/icones_confirma.gif" alt="Fechar Cota&ccedil;&atilde;o" width="20" height="20" border="0"></a>                         </div></td> 
						<%end if%>
                        <%end if%>
                    </div></td>
                    <td width="69%"><%
IF aberta = "sim" then
	response.write "&nbsp;"
	response.write "LIBERADA."
else
	response.write "&nbsp;"
	response.write "BLOQUEADA."
end if
%>
                      <div align="left"></div></td>
                  </tr>
                </table>
 			  </div>			    <div align="center"> 
                  <a href="downloads_ver.asp?codcot=<%=codcot%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"></a> 
                </div></td>			  
			  
			  <td valign="middle"><div align="center"><a href="javascript:MM_openBrWindow('cotacao_imp.asp?codcot=<%=codcot%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_precos.gif" alt="Imprimir a Cota&ccedil;&atilde;o" width="20" height="20" border="0"></a></div></td>
			  <td valign="middle"><div align="center"><a href="javascript:MM_openBrWindow('cotacao_obs.asp?codcot=<%=codcot%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_precos.gif" alt="Visualizar as Observa&ccedil;&otilde;es" width="20" height="20" border="0"></a></div></td>
			  <td width="20" valign="middle"><div align="center"> 
                 <%if nivel_aux = "2" then%> <a href="cotacao_inserir.asp?codcot=<%=codcot%>&inserir=não&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="imagem/icones_edita.gif" alt="Edita esta Cota&ccedil;&atilde;o" width="20" height="20" border="0"></a> <%end if%>
              </div></td>
				
				<td width="20" valign="middle"><div align="center"> 
                  
				  <table width="100%" border="0" cellspacing="0" cellpadding="0">
					<!-- ABAIXO ESTÁ A PROGRAMAÇÃO NECESSÁRIA PARA MOSTRAR O ÍCONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMAÇÃO-->
					<form name="form4" method="get" action="cotacao.asp">
					  <input name="auxapaga" type="hidden" value="sim">
					  <input name="auxcodcot" type="hidden" value="<%=codcot%>">
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
						 <%if nivel_aux = "2" then%> <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar esta Cotação" WIDTH="20" HEIGHT="20">
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
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td></td>
        <td width="110"><div align="right"> 
            <%
if (Session("pagina_downloads") <> 1 )or(Session("pagina_downloads") <> 1)or(Session("pagina_downloads") <> objrs01.PageCount )or(Session("pagina_downloads") <> objrs01.PageCount) then

	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> 1 then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_downloads") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""cotacao.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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