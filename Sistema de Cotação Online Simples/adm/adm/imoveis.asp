<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_contratos
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
nivel_aux = nivel
'nivel_aux = 2
%><%
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
'DECLARA TODAS AS VARI�VEIS
'--------------------------
Dim sql, ordem, pesquisa, auxcodimo, auxapaga, listar, contador,auxverifica1,auxprincipal


'RESGATA AS VARI�VEIS
'--------------------
auxcodimo   = request("auxcodimo")
auxapaga    = request("auxapaga")
pesquisa    = request("pesquisa")
ordem 	 	= request("ordem")
listar      = request("listar")
codimo		= request("codimo")
tipo		= request("tipo")
filtro		= request("filtro")

if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "codimo"
end if
if listar="" then
   listar = "30"
end if
%>
<%
'FUN��O PARA PESQUISA
'--------------------
function preparaPalavra(str)
   preparaPalavra = replace(str,"a","[a,�,�,�,�,�]")
   preparaPalavra = replace(preparaPalavra,"e","[e,�,�,�,�]")	
   preparaPalavra = replace(preparaPalavra,"i","[i,�,�,�,�]")
   preparaPalavra = replace(preparaPalavra,"o","[o,�,�,�,�,�]")
   preparaPalavra = replace(preparaPalavra,"u","[u,�,�,�,�]")
   preparaPalavra = preparaPalavra
end function
%>
<%
'ROTINA DE SQL
'-------------
sub rotinasql
	
	sql = "SELECT * FROM imoveis where codimo <> 0 "
	
	if pesquisa <> "" then
	  nometratado =  preparapalavra(pesquisa)
	  sql = sql & " and (rua Like '%"&(nometratado)&"%') "
	end if
	
	if codimo <> "" then
	  sql = sql & " and (codimo = '"&codimo&"') "
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
	response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class=linka><tr><td><div align='center'><b><font face='Verdana, Arial, Helvetica, sans-serif'>N�o foi encontrado nenhum registro.</font> </b></div></td></tr></table>"
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
                  <td bgcolor="#BBECB1"><div align="center"><strong> IM&Oacute;VEIS</strong></div></td>
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
                          <td width="25"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="25"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="24"><div align="center"><strong> 
                              
                          </strong><strong>
                          <%if nivel_aux = "2" then%>
                          <a href="imoveis_inserir.asp?tipo=<%=tipo%>"><img src="imagem/icones_insere.gif" alt="Insere novo registro" width="20" height="20" border="0"></a>
                          <%end if%>
                          </strong></div></td>
                          <td width="17"><div align="center"></div></td>
						  
                          <td width="110"> <div align="center">
                            <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                              <%if listar="30" then%>
                              <option value="imoveis.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" selected>LISTAR:30</option>
                              <option value="imoveis.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="imoveis.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="imoveis.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="100" then%>
                              <option value="imoveis.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="imoveis.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:100</option>
                              <option value="imoveis.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="imoveis.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="200" then%>
                              <option value="imoveis.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="imoveis.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="imoveis.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:200</option>
                              <option value="imoveis.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                              <%end if%>
                              <%if listar="500" then%>
                              <option value="imoveis.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                              <option value="imoveis.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                              <option value="imoveis.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                              <option value="imoveis.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:500</option>
                              <%end if%>
                            </select>
                          </div>
                          <div align="right"></div></td>
                          <td width="23"> <div align="center"></div></td>
                          <td width="271">  
						  
                            <div align="left">
                              <table width="261" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="imoveis.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="filtro" type="hidden" value="<%=filtro%>">								  
                                  <tr> 
                                    <td width="261"> <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15">
                                      <select name="pesquisa2" id="pesquisa2" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" >
                                        <option value="codimo" <%if pesquisa2 = "codimo" then%> selected <% end if %>>C&oacute;d </option>
                                        <option value="rua" <%if pesquisa2 = "rua" then%> selected <% end if %>>Rua</option>
                                        <option value="numero" <%if pesquisa2 = "numero" then%> selected <% end if %>>N�mero</option>
                                        <option value="cidade" <%if pesquisa2 = "cidade" then%> selected <% end if %>>Cidade</option>
                                        <option value="bairro" <%if pesquisa2 = "bairro" then%> selected <% end if %>>Bairro</option>
                                        </select> 
                                      <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" width="20" height="13" border="0">                                    </td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                          <td width="45"><div align="center"></div></td>
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
	set objrs01 = objConn01.execute("delete from imoveis WHERE codimo="&auxcodimo&"  ")
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
			DE COD <%=auxcodimo%> FOI APAGADO COM SUCESSO!</strong></div></td>
      </tr>
    </table>
    <%	
end if
%>
    <%
'SE � A PRIMEIRA VEZ QUE A P�GINA � CARREGADA
'--------------------------------------------
If Session("primeiravez_locatarios") <> "Nao" then 
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da p�gina de registros
	call rotinasql
	if objrs01.RecordCount<>"0" then
		session("pagina_locatarios") = 1 
		CALL MostraDados 
		Session("primeiravez_locatarios") = "Nao"
	else
		CALL Vazio
	end if

'SE A P�GINA J� FOI CARREGADA
'----------------------------
Else
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da p�gina de registros
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

'MENU DE PAGINA��O
'-----------------
sub menu
End sub
%>
    <%
Sub MostraDados()
	objrs01.AbsolutePage = Session("pagina_locatarios") ' vai para o n�mero da p�gina que est� armazenado em session("pagina")
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


	  response.write " Registro(s) cadastrado(s) - Mostrando p�gina " & Session("pagina_locatarios") & " de " & objrs01.PageCount 



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
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Pr�xima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Pr�xima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='�ltima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='�ltima tela' width='20' height='20' border='0'>"
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
	
	  response.write " Registro(s) cadastrado(s) - Mostrando p�gina " & Session("pagina_locatarios") & " de " & objrs01.PageCount 

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
              <td width="49" ><div align="center" class="linkA5"><strong><a href="imoveis.asp?ordem=<% if ordem1 = 1 then %>codimo&ordem1=0<%else%>codimo desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">C&oacute;d</a></strong>
                    <%if ordem="codimo desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="codimo" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="300"><div align="center" class="linkA5"><strong><a href="imoveis.asp?ordem=<% if ordem1 = 1 then %>rua&ordem1=0<%else%>rua desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Rua</a></strong>
                    <%if ordem="rua desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="rua" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="63"><div align="center" class="linkA5"><strong><a href="imoveis.asp?ordem=<% if ordem1 = 1 then %>numero&ordem1=0<%else%>numero desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">N&uacute;mero</a></strong>
                    <%if ordem="numero desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="numero" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="199"><div align="center" class="linkA5"><strong><a href="imoveis.asp?ordem=<% if ordem1 = 1 then %>bairro&ordem1=0<%else%>bairro desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Bairro</a></strong>
                    <%if ordem="bairro desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="bairro" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="218"><div align="center" class="linkA5"><strong><a href="imoveis.asp?ordem=<% if ordem1 = 1 then %>cidade&ordem1=0<%else%>cidade desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Cidade</a></strong>
                    <%if ordem="cidade desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="cidade" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </div></td>
              <td width="48"><div align="center"><strong>Ed</strong></div></td>
              <td width="51"><div align="center"><strong>Ap</strong></div></td>
            </tr>
            <%		
	For contador = 1 to listar
	
		codimo		= objrs01("codimo")
	rua		= objrs01 ("rua")
	numero	= objrs01 ("numero")
	cidade		= objrs01 ("cidade")
	bairro	= objrs01 ("bairro")
%> <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
              <td width="49" valign="middle"><div align="center"> 
                  <%response.write codimo%>
                </div></td>
              <td width="300" valign="middle"> <div align="left"> 
                  <%
				if rua <> "" then
					response.write rua
				else
					response.write "-"
				end if
				%>
                </div></td>
              <td width="63" valign="middle"> 
                <div align="center"> <%
				if numero <> "" then
					response.write numero
				else
					response.write "-"
				end if
				%>
               </div></td>
              <td width="199" valign="middle"> 
                <%
				if bairro <> "" then
					response.write bairro
				else
					response.write "-"
				end if
				%>
              </td>
              <td width="218" valign="middle"> <div align="left"> 
                  <%
				if cidade <> "" then
					response.write cidade
				else
					response.write "-"
				end if
				%>
                </div></td>
              <td width="48" valign="middle"><div align="center"> <a href="imoveis_editar.asp?codimo=<%=codimo%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&tipo=<%=tipo%>"><strong>
                <%if nivel_aux = "2" then%>
              </strong><img src="imagem/icones_edita.gif" alt="Edita este Registro" width="20" height="20" border="0"><strong>
              <%end if%>
              </strong></a> 
                </div></td>
              <td width="51" valign="middle"><div align="center"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <!-- ABAIXO EST� A PROGRAMA��O NECESS�RIA PARA MOSTRAR O �CONE DE APAGAR REGISTRO DE FORMA A TER UAM MENSSAGEM DE CONFIRMA��O-->
                    <form name="form4" method="post" action="imoveis.asp"> 
                      <input name="auxapaga" type="hidden" value="sim">
                      <input name="auxcodimo" type="hidden" value="<%=codimo%>">
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
                            <strong>
                            <%if nivel_aux = "2" then%>
                            </strong>
                            <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apagar este registro" WIDTH="20" HEIGHT="20">
                            <strong>
                            <%end if%>
                            </strong>                          </div></td>
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
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Primeira&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> 1 then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Anterior&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Proxima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Pr�xima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Pr�xima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_locatarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""imoveis.asp?Navegacao=Ultima&ordem="&ordem&"&filtro="&filtro&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='�ltima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='�ltima tela' width='20' height='20' border='0'>"
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