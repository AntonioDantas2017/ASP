<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_contratos
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
%>
<!--#include file="inc_css.asp"-->

<html>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">
<div align="center"> 

<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodloc, auxapaga, listar, contador,auxverifica1,auxprincipal


'RESGATA AS VARIÁVEIS
'--------------------
auxcodloc   = request("auxcodloc")
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

if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "nome"
end if
if listar="" then
   listar = "30"
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
	
	sql = "SELECT * FROM locatarios where codloc <> 0 "
	
	if pesquisa <> "" then
		nometratado =  preparapalavra(pesquisa)
		if pesquisa2 = "" then
			sql = sql & " and (contato Like '%"&(nometratado)&"%'  or nome Like '%"&(nometratado)&"%') "
		else
			sql = sql & " and ("&pesquisa2&" Like '%"&(nometratado)&"%') "
		end if
	end if
	
	if codloc <> "" then
	  sql = sql & " and (codloc = '"&codloc&"') "
	end if
		
	sql = sql & " order by " & ordem
	'response.Write(sql)
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
                  <td bgcolor="#BBECB1"><div align="center"><strong> LOCAT&Aacute;RIOS</strong></div></td>
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
                          <td width="29"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="29"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="27"><div align="center"><strong> 
                              
                          </strong><strong>
						  <%if nivel_aux = "2" then%> 
						  <a href="locatarios_inserir.asp?tipo=<%=tipo%>"><img src="imagem/icones_insere.gif" alt="Insere novo registro" width="20" height="20" border="0"></a>
						  <%end if%>
						  </strong></div></td>
                          <td width="26"><div align="center"></div></td>
						  <td width="26"><div align="center">
					      </div></td>
						  <td width="29"><div align="center"><strong></strong></div></td>
                          <td width="121"> <div align="center">
                            <select name="select" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="locatarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" selected>LISTAR:30</option>
                              <option value="locatarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="locatarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="locatarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="locatarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>>" >LISTAR:30</option>
                              <option value="locatarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:100</option>
                              <option value="locatarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="locatarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="locatarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="locatarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="locatarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:200</option>
                              <option value="locatarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="locatarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" >LISTAR:30</option>
                              <option value="locatarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:100</option>
                              <option value="locatarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>">LISTAR:200</option>
                              <option value="locatarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>&filtro=<%=filtro%>&pesquisa2=<%=pesquisa2%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
                         
                          <td width="227">  
						  
                            <div align="left">
                              <table width="193" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="locatarios.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">								  
                                  <tr> 
                                    <td width="193"> <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="pesquisa2" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codloc" <%if pesquisa2 = "codloc" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="nome" <%if pesquisa2 = "nome" then%> selected <% end if %>>Nome</option>
                                        <option value="contato" <%if pesquisa2 = "contato" then%> selected <% end if %>>Contato</option>
                                        <option value="cpf" <%if pesquisa2 = "obs" then%> selected <% end if %>>CPF</option>
										 <option value="rg" <%if pesquisa2 = "rg" then%> selected <% end if %>>RG</option>
                                        </select>
                                      <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" width="20" height="13" border="0">                                    </td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                          <td width="63"><div align="center"></div></td>
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
	set objrs01 = objConn01.execute("delete from locatarios WHERE codloc="&auxcodloc&"  ")
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
			DE COD <%=auxcodloc%> FOI APAGADO COM SUCESSO!</strong></div></td>
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
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
  </div>
  <div align="center"> 
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#999999"><table width="700" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
            <tr valign="top" bgcolor="#FFFFFF" class="linkA5"> 
              <td width="58" ><div align="center" class="linkA5"><strong><a href="locatarios.asp?ordem=<% if ordem1 = 1 then %>codloc&ordem1=0<%else%>codloc desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Cód</a></strong>
                    <%if ordem="codloc desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="codloc" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="201"><div align="center" class="linkA5"><strong><a href="locatarios.asp?ordem=<% if ordem1 = 1 then %>nome&ordem1=0<%else%>nome desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Nome</a></strong>
                    <%if ordem="nome desc" then%>
                <img src="imagem/cima.gif">
                <%else if ordem="nome" then%>
                <img src="imagem/baixo.gif">
                <%end if%>
                <%end if%>
              </div></td>
              <td width="95"><div align="center"><strong>CPF</strong></div></td>
              <td width="94"><div align="center"><strong>RG</strong></div></td>
              <td width="180"><div align="center"><strong>CONTATO</strong></div></td>
              <td width="25"><div align="center"><strong>Ed</strong></div></td>
              <td width="25"><div align="center"><strong>Ap</strong></div></td>
            </tr>
            <%		
	For contador = 1 to listar
	
		codloc		= objrs01("codloc")
	nome		= objrs01 ("nome")
	qualificacao	= objrs01 ("qualificacao")
	cpf		= objrs01 ("cpf")
	rg	= objrs01 ("rg")
	contato		= objrs01 ("contato")		


%>
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
              <td width="58" valign="middle"><div align="center"> 
                  <%response.write codloc%>
                </div></td>
              <td width="201" valign="middle"> <div align="left"> 
                  <%
				if nome <> "" then
					response.write nome
				else
					response.write "-"
				end if
				%>
                </div></td>
              <td width="95" valign="middle"> 
                <div align="center"> <%
				if cpf <> "" then
					response.write cpf
				else
					response.write "-"
				end if
				%>
               </div></td>
              <td width="94" valign="middle"> <div align="center"> 
                  <%
				if rg <> "" then
					response.write rg
				else
					response.write "-"
				end if
				%>
                </div></td>
              <td width="180" valign="middle"> <div align="left"> 
                  <%
				if contato <> "" then
					response.write contato
				else
					response.write "-"
				end if
				%>
                </div></td>
              <td width="25" valign="middle"><div align="center"> <a href="locatarios_editar.asp?codloc=<%=codloc%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>"><%if nivel_aux = "2" then%> <img src="imagem/icones_edita.gif" alt="Edita este Registro" width="20" height="20" border="0"><%end if%></a> 
                </div></td>
              <td width="25" valign="middle"><div align="center"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <!-- ABAIXO ESTÁ A PROGRAMAÇÃO NECESSÁRIA PARA MOSTRAR O ÍCONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMAÇÃO-->
                    <form name="form4" method="post" action="locatarios.asp"> 
                      <input name="auxapaga" type="hidden" value="sim">
                      <input name="auxcodloc" type="hidden" value="<%=codloc%>">
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
                           <%if nivel_aux = "2" then%>  <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar este registro" WIDTH="20" HEIGHT="20"><%end if%>
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
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""locatarios.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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