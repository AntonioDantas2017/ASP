<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
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
.style1 {font-weight: bold}
.style2 {font-weight: bold}
-->
</style>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0"  onLoad="form1.pesquisa.focus()">
<div align="center"> 

<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodcon, auxapaga, listar, contador,auxverifica1,auxprincipal


'RESGATA AS VARIÁVEIS
'--------------------
auxcodcon   = request("auxcodcon")
auxapaga    = request("auxapaga")
pesquisa    = request("pesquisa")
pesquisa2    = request("pesquisa2")
ordem 	 	= request("ordem")
ordem1 	 	= request("ordem1")
listar      = request("listar")
codfor		= request("codfor")
filtro		= request("filtro")
filtro2		= request("filtro2")
datainicio	= request("datainicio")
datafim		= request("datafim")
codcon		= request("codcon")

if datainicio = "" then
	datainicio = DateAdd("yyyy",-2,date)
	'response.Write datainicio
end if

if datafim = "" then
	datafim = DateAdd("d",30,date)
end if

if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "nome"
end if
if listar="" then
   listar = "30"
end if


'sql = "delete from contratos where aluguel=''"
'set objRs01 = objConn01.execute(sql)
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
	
	sql = "SELECT * FROM contratos2 where codcon <> 0 "
	
	if pesquisa <> "" then
	nometratado =  preparapalavra(pesquisa)
		if pesquisa2 = "" then
			sql = sql & " and (nome  Like '%"&(nometratado)&"%')"
		else
			
			sql = sql & " and ("&pesquisa2&" Like '%"&(nometratado)&"%') "
		end if
	end if
	
	datainicio2 = split(datainicio,"/")
	datainicio = datainicio2(1) & "/" &  datainicio2(0) & "/" &  datainicio2(2)
	datafim2 = split(datafim,"/")
	datafim = datafim2(1) & "/" &  datafim2(0) & "/" &  datafim2(2)	

	
	if (datainicio <> "") and (datafim<>"") then
		sql = sql & " and ((datafim between #"&datainicio&"# and #"&datafim&"#) or (dataini between #"&datainicio&"# and #"&datafim&"#)) "
	end if
	
	if codcon <> "" then
	  sql = sql & " and (codcon = "&codcon&") "
	end if
		
	sql = sql & " order by " & ordem
'	response.Write(sql)
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
                          <td width="20"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="24"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="22"><div align="center"><strong> 
                              
                          </strong><strong></strong>
                          <div align="center"> <a href="contratos_editar.asp?acao=inserir&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="imagem/icones_insere.gif" alt="Edita este Registro" width="20" height="20" border="0"></a> </div>
                          </div></td>
                          
						
						  <td width="11"><div align="left"><strong></strong></div></td>
                          <td width="100"> <div align="center">
                            <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="contratos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" selected>LISTAR:30</option>
                              <option value="contratos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="contratos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="contratos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="contratos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="contratos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:100</option>
                              <option value="contratos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="contratos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="contratos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="contratos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="contratos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:200</option>
                              <option value="contratos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="contratos.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="contratos.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="contratos.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="contratos.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
                          <td width="4"> <div align="center"></div></td>
                          <td width="200">  
						  
                            <div align="left">
                              <table width="196" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="contratos.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">								  
                                  <tr> 
                                    <td width="196"> <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="pesquisa2" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codcon" <%if pesquisa2 = "codcon" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="nome" <%if pesquisa2 = "nome" then%> selected <% end if %>>Nome</option>
                                        <option value="rua" <%if pesquisa2 = "rua" then%> selected <% end if %>>Rua</option>
                                        <option value="numero" <%if pesquisa2 = "numero" then%> selected <% end if %>>Número</option>
                                        <option value="bairro" <%if pesquisa2 = "bairro" then%> selected <% end if %>>Bairro</option>
										 <option value="cidade" <%if pesquisa2 = "cidade" then%> selected <% end if %>>Cidade</option>
										 
                                      </select> 
                                      <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" width="20" height="13" border="0">                                    </td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                         						  <form action="contratos.asp" method="get" name="form3">
                            <input name="listar" type="hidden" value="<%=listar%>">
                            <input name="ordem" type="hidden" value="<%=ordem%>">
                            <input name="tipo" type="hidden" value="<%=tipo%>">
                            <input name="filtro" type="hidden" value="<%=filtro%>">
                            <input name="filtro2" type="hidden" value="<%=filtro2%>">
                            <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                            <input name="pesquisa2" type="hidden" value="<%=pesquisa2%>">
                            <td width="196">
                               <div align="left">
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
                          </div></td></form>
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
	set objrs01 = objConn01.execute("delete from contratos WHERE codcon="&auxcodcon&"  ")
	set objrs01 = objConn01.execute("delete from pagamentos WHERE codcon="&auxcodcon&"  ")
%>
	<table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong><font color="#FF0000"> 
            </font>O 
			CADASTRO
			DE COD <%=auxcodcon%> FOI APAGADO COM SUCESSO!</strong></div></td>
      </tr>
    </table>
    <%	
end if
%>
    <%
'SE É A PRIMEIRA VEZ QUE A PÁGINA É CARREGADA
'--------------------------------------------
If Session("primeiravez_locatarios") <> "Nao" then 
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	call rotinasql
	if objrs01.RecordCount<>"0" then
		session("pagina_locatarios") = 1 
		CALL MostraDados 
		Session("primeiravez_locatarios") = "Nao"
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
			Session("pagina_locatarios") = 1 
		end if
		if Request("Navegacao") = "Primeira" then
			Session("pagina_locatarios") = 1 
		end if
		if Request("Navegacao") = "Proxima" then
			Session("pagina_locatarios") = Session("pagina_locatarios") + 1 
		end if
		If Request("Navegacao") = "Anterior" then
			Session("pagina_locatarios") = Session("pagina_locatarios") - 1
		End If
		if Request("Navegacao") = "Ultima" then
			Session("pagina_locatarios") = objrs01.PageCount
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
	objrs01.AbsolutePage = Session("pagina_locatarios") ' vai para o número da página que está armazenado em session("pagina")
%>
	
    <%
if (Session("pagina_locatarios") <> 1 )or(Session("pagina_locatarios") <> 1)or(Session("pagina_locatarios") <> objrs01.PageCount )or(Session("pagina_locatarios") <> objrs01.PageCount) then
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


	  response.write " Registro(s) cadastrado(s) - Mostrando página " & Session("pagina_locatarios") & " de " & objrs01.PageCount 



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
if (Session("pagina_locatarios") <> 1 )or(Session("pagina_locatarios") <> 1)or(Session("pagina_locatarios") <> objrs01.PageCount )or(Session("pagina_locatarios") <> objrs01.PageCount) then

	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
	
	  response.write " Registro(s) cadastrado(s) - Mostrando página " & Session("pagina_locatarios") & " de " & objrs01.PageCount 

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
            <td width="22"  bgcolor="#FFFC82"></td>
            <td width="195">RENOVA&Ccedil;&Atilde;O DE CONTRATO </td>
            <td width="20" bgcolor="#FE6F6B">&nbsp;</td>
            <td width="202">CONTRATO VENCIDO </td>
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
  <div align="right">
  
  </div>
  <div align="center"> 
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#999999"><table width="700" border="0" cellpadding="1" cellspacing="1" class="linkA5">
            <tr valign="top" bgcolor="#FFFFFF"> 
              <td width="97" ><div align="center" class="style1"><a href="contratos.asp?ordem=<% if ordem1 = 1 then %>dataini&ordem1=0<%else%>dataini desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Data Início</a>
                <%if ordem="dataini" then%><img src="imagem/cima.gif"><%else if ordem="dataini desc" then%><img src="imagem/baixo.gif">
                <%end if%>
                <%end if%>
				
				<BR>
				
				<a href="contratos.asp?ordem=<% if ordem1 = 1 then %>datafim&ordem1=0<%else%>datafim desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Data Fim</a></strong><%if ordem="datafim" then%><img src="imagem/cima.gif"><%else if ordem="datafim desc" then%><img src="imagem/baixo.gif">
                <%end if%>
              <%end if%></strong></div></td>
              <td width="129"><div align="center"><strong><a href="contratos.asp?ordem=<% if ordem1 = 1 then %>vencimento&ordem1=0<%else%>vencimento desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Dia Vencimento</a>
                  <%if ordem="vencimento" then%>
                  <img src="imagem/cima.gif">
                  <%else if ordem="vencimento desc" then%>
                  <img src="imagem/baixo.gif">
                  <%end if%>
                  <%end if%>
                  <BR>
                  <a href="contratos.asp?ordem=<% if ordem1 = 1 then %>aluguel&ordem1=0<%else%>aluguel desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Valor Aluguel</a></strong>
                  <%if ordem="aluguel" then%>
                  <img src="imagem/cima.gif">
                  <%else if ordem="aluguel desc" then%>
                  <img src="imagem/baixo.gif">
                  <%end if%>
                  <%end if%>
              </strong></strong></div></td>
              <td width="352"><div align="center" class="style2"><a href="contratos.asp?ordem=<% if ordem1 = 1 then %>nome&ordem1=0<%else%>nome desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Nome Locatário</a>
                <%if ordem="nome" then%><img src="imagem/cima.gif"><%else if ordem="nome desc" then%><img src="imagem/baixo.gif">
                <%end if%>
                <%end if%><BR><a href="contratos.asp?ordem=<% if ordem1 = 1 then %>rua,numero,bairro&ordem1=0<%else%>rua,numero,bairro desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Imóvel</a></strong><%if ordem="rua,numero,bairro" then%><img src="imagem/cima.gif"><%else if ordem="rua,numero,bairro desc" then%><img src="imagem/baixo.gif">
                <%end if%>
              <%end if%></strong></div></td>
              <td width="26"><div align="center"><strong>Imp</strong></div></td>
              <td width="28"><div align="center"><strong>Ren</strong></div></td>
              <td width="20"><div align="center"><strong>Ed</strong></div></td>
              <td width="26"><div align="center"><strong>Ap</strong></div></td>
            </tr>
            <%		
	For contador = 1 to listar
	codcon 	= objrs01("codcon")
	nome 	= objrs01("nome")
	rua 	= objrs01("rua")
	numero 	= objrs01("numero")		
	periodo = objrs01("periodo")
	dataini = objrs01("dataini")
	renovado= objrs01("renovado")
	if isdate(dataini) then
		dataini2 = split(dataini,"/")
		dataini = right("0"&dataini2(0),2) &"/" & right("0"&dataini2(1),2) &"/" & dataini2(2)
	end if
	datafim = objrs01 ("datafim")
	if isdate(datafim) then
		datafim2 = split(datafim,"/")
		datafim = cdate(right("0"&datafim2(0),2) &"/" & right("0"&datafim2(1),2) &"/" & datafim2(2))
	end if
	aluguel = objrs01 ("aluguel")
	if isnumeric(aluguel) then
		aluguel = formatnumber(aluguel,2)
	end if
	vencimento = objrs01 ("vencimento")
	reajuste = objrs01 ("reajuste")
	if isdate(reajuste) then
		reajuste2 = split(reajuste,"/")
		reajuste = right("0"&reajuste2(0),2) &"/" & right("0"&reajuste2(1),2) &"/" & reajuste2(2)
	end if
	
	if renovado = "" then
		renovado = "n"
	end if
	
'	set objrs02 = objConn02.execute("select * from locatarios where codcon="&codcon&" ")
'	nome = objrs02("nome")
	
'	set objrs02 = objConn02.execute("select * from imoveis where codimo="&codimo&" ")
'	rua = objrs02("rua")
'	numero = objrs02("numero")	

%>
                  <%
				  if ((datafim - 15) < date) and (renovado = "n") then
				  
				  %>
            <tr valign="middle" bgcolor="#FFFC82"  onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFC82')" class="linkA"> 
              <%else
			  if datafim < date and (renovado = "n") then
			  %>
            <tr valign="middle" bgcolor="#FE6F6B" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FE6F6B')" class="linkA"> 
			<%else%> 
			 <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')" class="linkA"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')" class="linkA"> 
              <%end if%>
              <%end if
			    end if%>

			  
			  
			  
              <td width="97" valign="middle"><div align="center"> 
                  <%
				dataini2 = split(dataini,"/")
				dataini = right("0"&dataini2(0),2) &"/" & right("0"&dataini2(1),2) &"/" & dataini2(2) 
				datafim2 = split(datafim,"/")
				datafim = right("0"&datafim2(0),2) &"/" & right("0"&datafim2(1),2) &"/" & datafim2(2) 
				  response.write dataini&"<BR>"&datafim
				  
				  %>
              </div></td>
              <td width="129" valign="middle">  
                <div align="right">
                  <%
				response.write  "Dia: "&vencimento&"<BR>R$ "&aluguel
						%>
              </div></td>
              <td valign="middle"> <div align="left">
                <%
				imovel = "Rua: "&rua&" n°: "&numero
				response.write UCASE(nome)&"<BR>"&imovel
				%>
              </div></td>
              <td valign="middle"><div align="center"><a href="javascript:MM_openBrWindow('contratos_imp.asp?codcon=<%=codcon%>','','status=yes,scrollbars=yes,width=750,height=572')"  class="linkA5"><img src="imagem/icones_precos.gif" alt="Imprimir o Contrato" width="20" height="20" border="0"></a></div></td>
              <td valign="middle"><div align="center"> <a href="contratos_editar.asp?codcon=<%=codcon%>&acao=renovar&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="imagem/icones_cfop.gif" alt="Edita este Registro" width="20" height="20" border="0"></a> </div></td>
              <td width="20" valign="middle"><div align="center"> <a href="contratos_editar.asp?codcon=<%=codcon%>&acao=editar&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"><img src="imagem/icones_edita.gif" alt="Edita este Registro" width="20" height="20" border="0"></a> 
                </div></td>
              <td width="26" valign="middle"><div align="center"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <!-- ABAIXO ESTÁ A PROGRAMAÇÃO NECESSÁRIA PARA MOSTRAR O ÍCONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMAÇÃO-->
                    <form name="form4" method="post" action="CONTRATOS.asp"> 
                      <input name="auxapaga" type="hidden" value="sim">
                      <input name="auxcodcon" type="hidden" value="<%=codcon%>">
                      <input name="listar" type="hidden" value="<%=listar%>">
                      <input name="ordem" type="hidden" value="<%=ordem%>">
                      <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                      <input name="tipo" type="hidden" value="<%=tipo%>">
                      <input name="filtro" type="hidden" value="<%=filtro%>">
                      <tr> 
                        <td><div align="center"> 
                            <script language="JavaScript">
							function confirmar()
							{
							   return (confirm('Deseja apagar este registro?'))
							}
						  </script>
                            <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar este registro" WIDTH="20" HEIGHT="20">
                          </div></td>
                      </tr>
                    </form>
                  </table>
                </div></td>
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
    <table width="950" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td>&nbsp;</td>
        <td width="110"><div align="right"> 
            <%
if (Session("pagina_locatarios") <> 1 )or(Session("pagina_locatarios") <> 1)or(Session("pagina_locatarios") <> objrs01.PageCount )or(Session("pagina_locatarios") <> objrs01.PageCount) then

	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""contratos.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'>"
	End If
	
End if
%>
          </div></td>
      </tr>
    </table>
    <%
End Sub
%>
    <%
		CALL fecha_conexao01
	%>
  </div>
</div>
</body>
</html>