<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_funcionarios
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


<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0" >
<div align="center"> 
<%
'RESGATA AS VARIÁVEIS
'--------------------
Dim sql, ordem, listar, Contador, pesquisa, auxapaga, auxapaga2, auxcont
Dim data, inicio, termino, login, senha, resultado, tempo, ip, browser, cont, resposta

auxapaga    = request ("auxapaga")
auxapaga2   = request ("auxapaga2")
auxcont     = request ("auxcont")

listar		= request ("listar")
pesquisa    = request ("pesquisa")
ordem 		= request ("ordem")

call abre_conexao01
call abre_conexao02

if pesquisa = "" then
	pesquisa = "Pesquise_aqui"
end if
if ordem="" then
   ordem = "data"
end if
listar      = request("listar")
if listar="" then
   listar = "30"
end if

'LIMPA TABELA DE LOG
'-------------------
if auxapaga="sim" then
	set objrs01 = objConn01.execute("select max(CONT) as maxCONT from LOG")
	maxcont = objrs01("maxcont")
	set objrs01 = objConn01.execute("delete from log  where cont<>'"&maxcont&"' ")
end if

if auxapaga2="sim" then
	set objrs01 = objConn01.execute("delete from log  where cont='"&auxcont&"' ")
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
	if pesquisa = "Pesquise_aqui" then
		if ordem = "data" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by data desc,inicio desc"
		end if
		if ordem = "inicio" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by inicio,data desc"
		end if
		if ordem = "login" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by login,data desc"
		end if
		if ordem = "senha" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by senha,data desc"
		end if
		if ordem = "resposta" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by resposta,data desc"
		end if
		if ordem = "resultado" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by resultado,data desc"
		end if
		if ordem = "ip" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by ip,data desc"
		end if
		if ordem = "browser" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where login <>'d4w' order by browser,data desc"
		end if
		objrs01.Open Sql, objConn01
	else
		nometratado =  preparapalavra(pesquisa)
		if ordem = "data" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by data desc,inicio desc "
		end if
		if ordem = "inicio" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%' ) and login <> 'd4w' order by inicio,data desc"
		end if
		if ordem = "login" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by login,data desc"
		end if
		if ordem = "senha" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by senha,data desc"
		end if
		if ordem = "resposta" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by resposta,data desc"
		end if
		if ordem = "resultado" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where(  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by resultado,data desc"
		end if
		if ordem = "ip" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by ip,data desc"
		end if
		if ordem = "browser" then
			sql= "SELECT data,inicio,termino,login,senha,resultado,ip,browser,cont,resposta FROM log where (  senha Like '%"&(nometratado)&"%')  and login <> 'd4w' order by browser,data desc"
		end if
		objrs01.Open Sql, objConn01
	end if
end sub
%>
<%
'GRAVA VARIÁVEIS DO BD
'---------------------
sub variaveis
	data	 = objrs01 ("data")
	inicio 	 = objrs01 ("inicio")
	termino  = objrs01 ("termino")
	login	 = objrs01 ("login")
	senha 	 = objrs01 ("senha")
	resultado= objrs01 ("resultado")
	ip	     = objrs01 ("ip")
	browser  = objrs01 ("browser")
	cont	 = objrs01 ("cont")
	resposta = objrs01 ("resposta")
end sub
%>
<%
'AVISO PARA RESULTADO DA PESQUISA EM BRANCO
'------------------------------------------
sub vazio
	response.write "<br><br><br><br><br><br><br>"		
	response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class=linka><tr><td><div align='center'><b><font face='Verdana, Arial, Helvetica, sans-serif'>N&atilde;o existe nenhum registro cadastrado.</font> </b></div></td></tr></table>"
end sub
%>
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="115"> <table width="115" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> <strong>LOG</strong></div></td>
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
                          <td width="28"> <div align="center"><strong><a href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="28"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
						  <%if ((nivel_aux="2")) then%><td width="28"><div align="center"><strong> </strong>
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <form name="form4" method="post" action="usuarios_log.asp">
                                  <input name="auxapaga" type="hidden" value="sim">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                                  <tr> 
                                    <td><div align="center"> 
                                        <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja limpar a tabela de LOG?'))
	}
	</script>
                                        <input NAME="delete2" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Limpa a tabela de LOG" WIDTH="20" HEIGHT="20">
                                      </div></td>
                                  </tr>
                                </form>
                              </table>
</div></td><%end if%>
                          <td width="25">&nbsp;</td>
                          <td width="126"> <div align="center"> 
                              <table width="100" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="usuarios_log.asp" method="post" name="form2">
                                  <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <tr> 
                                    <td height="18"> <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                                        <%if listar="30" then%>
                                        <option value="usuarios_log.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" selected>LISTAR:30</option>
                                        <option value="usuarios_log.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                        <option value="usuarios_log.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                        <option value="usuarios_log.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                        <%end if%>
                                        <%if listar="100" then%>
                                        <option value="usuarios_log.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                        <option value="usuarios_log.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:100</option>
                                        <option value="usuarios_log.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                        <option value="usuarios_log.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                        <%end if%>
                                        <%if listar="200" then%>
                                        <option value="usuarios_log.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                        <option value="usuarios_log.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                        <option value="usuarios_log.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:200</option>
                                        <option value="usuarios_log.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                        <%end if%>
                                        <%if listar="500" then%>
                                        <option value="usuarios_log.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                        <option value="usuarios_log.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                        <option value="usuarios_log.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                        <option value="usuarios_log.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:500</option>
                                        <%end if%>
                                      </select> </td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                          <td width="210">  
                            <div align="left">
                              <table width="120" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="usuarios_log.asp" method="post" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <tr> 
                                    <td> <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15"> 
                                      <input name="imageField32" type="image" id="imageField32" src="imagem/icones_ok.gif" alt="Procura LOG (Login e Sennha)" width="20" height="13" border="0">                                    </td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                          <td width="331">&nbsp;</td>
                          <td width="51"><div align="right"></div>
                            <div align="center"></div>
                            <div align="center"><strong> </strong></div>
                            <div align="center"></div>
                          <div align="center"></div></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
    </tr>
  </table>
  <%
'SE É A PRIMEIRA VEZ QUE A PÁGINA É CARREGADA
'--------------------------------------------
If Session("PrimeiraVez_usuarios_log") <> "Nao" then 
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	objrs01.CursorLocation = 3
	call rotinasql
	if objrs01.RecordCount<>"0" then
		session("Pagina_usuarios_log") = 1 
		CALL MostraDados 
		Session("PrimeiraVez_usuarios_log") = "Nao"
	else
		CALL Vazio
	end if

'SE A PÁGINA JÁ FOI CARREGADA
'----------------------------
Else
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	objrs01.CursorLocation = 3
	call rotinasql
	if objrs01.RecordCount<>"0" then
		if Request("Navegacao") = "" then
			Session("Pagina_usuarios_log") = 1 
		end if
		if Request("Navegacao") = "Primeira" then
			Session("Pagina_usuarios_log") = 1 
		end if
		if Request("Navegacao") = "Proxima" then
			Session("Pagina_usuarios_log") = Session("Pagina_usuarios_log") + 1 
		end if
		If Request("Navegacao") = "Anterior" then
			Session("Pagina_usuarios_log") = Session("Pagina_usuarios_log") - 1
		End If
		if Request("Navegacao") = "Ultima" then
			Session("Pagina_usuarios_log") = objrs01.PageCount
		end if
		CALL MostraDados
	else
		CALL Vazio
	end if
End If 
%>
  <%
Sub MostraDados()
	objrs01.AbsolutePage = Session("Pagina_usuarios_log") ' vai para o número da página que está armazenado em session("pagina")
%>
  <%
if (Session("Pagina_usuarios_log") <> 1 )or(Session("Pagina_usuarios_log") <> 1)or(Session("Pagina_usuarios_log") <> objrs01.PageCount )or(Session("Pagina_usuarios_log") <> objrs01.PageCount) then
%>
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
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) cadastrado(s) - Mostrando página " & Session("Pagina_usuarios_log") & " de " & objrs01.PageCount 
	%>
                            </strong></div>
                          <div align="center"><strong></strong></div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <div align="right"> </div></td>
          </tr>
        </table></td>
      <td width="110"><div align="right"> 
          <%
if (Session("Pagina_usuarios_log") <> 1 )or(Session("Pagina_usuarios_log") <> 1)or(Session("Pagina_usuarios_log") <> objrs01.PageCount )or(Session("Pagina_usuarios_log") <> objrs01.PageCount) then

	If Session("Pagina_usuarios_log") <> 1 then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> 1 then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><table width="700" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class=linka>
                <tr> 
                  <td width="50%" bgcolor="#F4F4F4"><div align="center"><strong> 
                      <%
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) cadastrado(s) - Mostrando página " & Session("Pagina_usuarios_log") & " de " & objrs01.PageCount 
	%>
                      </strong></div>
                    <div align="center"><strong></strong></div></td>
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
  <div align="center">
    <div align="center">
      <table width="200" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
        </tr>
      </table>
      <table width="700" border="0" cellspacing="0" cellpadding="0" >
        <tr> 
          <td><table width="700" border="0" cellspacing="0" cellpadding="0"CLASS=LINKA>
              <tr> 
                <td width="100"><div align="center"><strong>LEGENDA :</strong></div></td>
                <td width="18"><table width="100%" height="18" border="1" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td bgcolor="#FFD7D7"></td>
                    </tr>
                  </table></td>
                <td>&nbsp;&nbsp;ACESSO NEGADO OU BLOQUEADO</td>
              </tr>
            </table>
            <div align="right"> </div></td>
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
          <td bgcolor="#999999"><table width="100%" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
              <tr valign="top" bgcolor="#CCCCCC"> 
                <td width="80"><div align="center"><strong>DATA</strong></div></td>
                <td width="80"><div align="center"><strong>IN&Iacute;CIO</strong></div></td>
                <td width="80"><div align="center"><strong>T&Eacute;RMINO</strong></div></td>
                <td width="80"><div align="center"><strong>TEMPO </strong></div></td>
                <td width="70"><div align="center"><strong>LOGIN</strong></div></td>
                <td width="100"><div align="center"><strong>RESULTADO</strong></div></td>
                <td><div align="center"><strong>S.O. - BROWSER</strong></div></td>
                <%if (nivel_aux="2") then%>                <td width="30"><div align="center"><strong>AP</strong></div></td>
<% 
end if
%>				
              </tr>
              <%
For contador = 1 to listar
	call variaveis	

'ESCONDE SE O NÍVEL FOR "HELPDESK"
'---------------------------------
if login <> "d4w" then
%>
 <%if (resultado="ACESSO NEGADO") or (resultado="ACESSO BLOQUEADO")then%>
             <tr valign="middle" bgcolor="#FFC8C8" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFC8C8')"> 
			 <%else
			 %>
                <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
                <%end if%>
				
				
				
                <td width="80"><div align="center"> 
<%

arrData = Split(data,"/")
strData = Right("0" & arrData(0),2) & "/" & Right("0" & arrData(1),2) & "/" & arrData(2)
response.write strData

%>
                  </div></td>
                <td width="80"><div align="center"> 
                    <%response.write inicio%>
                  </div></td>
                <td width="80"><div align="center"> 
                    <%
if cStr(cont) = cstr(Session("log")) then
	response.Write "ON-LINE"
else
	if termino <> "" then
		response.write termino
	else
		response.Write "-"
	end if
end if


%>
                  </div></td>
                <td width="80"><div align="center"> 
<%
'CALCULA O TEMPO TOTAL DO ACESSO
'-------------------------------
if cStr(cont) = cstr(Session("log")) then
	response.Write "ON-LINE"
else
	entrada = inicio
	intervalo = "s"
	Saida = termino
	dia = DateDiff (intervalo,Entrada,Saida)'pega a diferença em segundos
	segundos = (dia) mod  60
	minutos = (( dia - segundos) / 60) mod 60
	horas = int((dia) / 3600)
	arrHora = Split(horas &":"&minutos&":"&segundos,":")
	strHora = Right("0" & arrHora(0),2) & ":" & Right("0" & arrHora(1),2) & ":" & Right("0" & arrHora(2),2)
	response.write strHora	
end if
%>
                  </div></td>
                <td width="70"><div align="center"> 
<%
	response.write login

%>
                  </div></td>

                <td width="100"><div align="center"> 
                    <%
if resultado = "IP INVÁLIDO" then
	response.write "<font color='#FF0000'><strong>"
	response.write "IP<br>INVÁLIDO"&"<br>"
	response.write "</strong></font>"
'	response.write "Usuário ou Senha inválido"
end if
if resultado = "ACESSO NEGADO" then
	response.write "<font color='#FF0000'><strong>"
	response.write "NEGADO"&"<br>"
	response.write "</strong></font>"
'	response.write "Usuário ou Senha inválido"
end if
if resultado = "ACESSO BLOQUEADO" then
	response.write "<font color='#000099'><strong>"
	response.write "BLOQUEADO"&"<br>"
	response.write "</strong></font>"
'	response.write "Excedido Nº de tentativas"
end if
if resultado = "LIBERADO" then
	response.write "<strong>"
	response.write "LIBERADO"&"<br>"
	response.write "</strong>"
end if


%>
                  </div></td>
                <td><div align="center"> 
  <%
if browser = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.1.4322)" then
	response.write "WIN 2000 - MSIE6.0"
else
	if browser = "Mozilla/4.0 (compatible; MSIE 5.01; Windows NT 5.0; .NET CLR 1.1.4322)" then
		response.write "WIN 2000 - MSIE5.01"
	else
		if browser = "Mozilla/4.0 (compatible; MSIE 6.0; Windows 98)" then
			response.write "WIN 98 - MSIE6.0"
		else	 
			if browser = "Mozilla/4.0 (compatible; MSIE 5.01; Windows 98)" then
				response.write "WIN 98 - MSIE5.01"
			else	 
				if browser = "Mozilla/4.0 (compatible; MSIE 6.0; Windows 98; .NET CLR 1.1.4322)" then
					response.write "WIN 98 - MSIE6.0"
				else					 
					if browser = "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98; Win 9x 4.90)" then
						response.write "WIN ME - MSIE5.5"
					else					 
						if browser = "Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" then
							response.write "WIN 98 - MSIE5"
						else					 
							if browser = "Mozilla/4.0 (compatible; MSIE 5.5; Windows 98)" then
								response.write "WIN 98 - MSIE5"
							else					 
								if browser = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" then
									response.write "WINNT5.1 - MSIE6"
								else					 
									if browser = "Mozilla/4.0 (compatible; MSIE 5.5; Windows NT 5.0)" then
										response.write "WINNT5.0 - MSIE5.5"
									else					 
										if browser = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)" then
											response.write "WINNT5.0 - MSIE6.0"
										else					 
											response.write browser
										end if
									end if
								end if
							end if
						end if
					end if
				end if
			end if
		end if
	end if
end if

end if
%>
                </div></td>
                <%if (nivel_aux="2") or nomeresumido="d4w" then%>

				  
				  
 <td width="30"><div align="center"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <form name="form4"  action="usuarios_log.asp"   method="get" >
    					 <input name="auxapaga2" type="hidden" value="sim">
                        <input name="auxcont" type="hidden" value="<%=cont%>">
                        <input name="listar" type="hidden" value="<%=listar%>">
                        <input name="ordem" type="hidden" value="<%=ordem%>">
                        <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                        <tr> 
                          <td><div align="center"> 
                              <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja apagar este registro?'))
	}
	</script>
                              <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga este registro" WIDTH="20" HEIGHT="20">
                            </div></td>
                        </tr>
                      </form>
                    </table>
                    
                  </div></td>
				  
				  
				  
<% 

end if
%>
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
if (Session("Pagina_usuarios_log") <> 1 )or(Session("Pagina_usuarios_log") <> 1)or(Session("Pagina_usuarios_log") <> objrs01.PageCount )or(Session("Pagina_usuarios_log") <> objrs01.PageCount) then

	If Session("Pagina_usuarios_log") <> 1 then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> 1 then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("Pagina_usuarios_log") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""usuarios_log.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
objrs01.close
CALL fecha_conexao01
'Session.Contents.Remove ("Pagina_usuarios_log") 
'Session.Contents.Remove ("PrimeiraVez_usuarios_log") 
	%>

    </div>
  </div>
</div>
</body>
</html>