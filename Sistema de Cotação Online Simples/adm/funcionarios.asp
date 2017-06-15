<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
if nivel_usuarios = "0" then
	session("pg") = request.ServerVariables("SCRIPT_NAME")
    Response.Redirect ("default_acessonegado.asp")
end if
call abre_conexao01
%>
<!--#include file="inc_css.asp"-->

<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0">
<div align="center"> 
<%
'RESGATA AS VARIÁVEIS
'--------------------
auxbloqueado	= request("auxbloqueado")
auxcodfunc      = request("auxcodfunc")
auxusu			= request("auxusu")
auxapaga    	= request("auxapaga")
pesquisa     	= request("pesquisa")
ordem 			= request("ordem")
listar			= request("listar")

call abre_conexao01
call abre_conexao02		

if pesquisa = "" then
	pesquisa = ""
end if
if ordem="" then
   ordem = "nome"
end if
if listar="" then
   listar = "30"
end if

'ALTERA OS REGISTROS
'-------------------
if auxbloqueado="sim" then
	set objrs01 = objConn01.execute("update funcionarios set bloqueado='não' WHERE codfunc='"&auxcodfunc&"' ")

	set objrs01 = objConn01.execute("select login from funcionarios WHERE codfunc='"&auxcodfunc&"' ")
	if not objrs01.eof then
		auxusuario = objrs01("login")
		set objrs01 = objConn01.execute("select login from funcionarios_cont WHERE login='"&auxusuario&"' ")
		if not objrs01.eof then
			set objrs01 = objConn01.execute("delete from funcionarios_cont WHERE login='"&auxusuario&"' ")
		end if
	end if
end if
if auxbloqueado="não" then
	set objrs01 = objConn01.execute("update funcionarios set bloqueado='sim' WHERE codfunc='"&auxcodfunc&"' ")
end if
'ROTINA DE SQL
'-------------
sub rotinasql

	sql = "Select * from funcionarios where nomeresumido <> 'Mossoró' "
	
	if pesquisa <> "" then
	  nometratado =  preparapalavra(pesquisa)
	  sql = sql & "and (nome Like '%"&(nometratado)&"%' or login Like '%"&(nometratado)&"%')"
	end if
	
	sql = sql & " order by " & ordem
	
	objrs01.Open Sql, objConn01
	
end sub
'GRAVA VARIÁVEIS DO BD
'---------------------
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
                  <td bgcolor="#BBECB1"><div align="center"><strong>USU&Aacute;RIOS</strong></div></td>
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
                          <td width="30"><div align="center"><strong><a href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"></a></strong></div></td>
                          <td width="30"><div align="center"><strong>
						  				  <%if (nivel_usuarios) = "2" then%>
						  <a href="funcionarios_inserir.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"><img src="imagem/icones_insere.gif" alt="insere funcionario nesta tabela" width="20" height="20" border="0"></a>
						  <%end if%>
						  </strong></div></td>
                          <td width="120"> <div align="center"> 
                              <table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="funcionarios.asp" method="get" name="form2">
                                  <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <tr> 
                                    <td height="18"> <div align="center"> 
                                        <select name="listar" id="listar" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                                          <%if listar="30" then%>
                                          <option value="funcionarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" selected>LISTAR:30</option>
                                          <option value="funcionarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                          <option value="funcionarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                          <option value="funcionarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                          <%end if%>
                                          <%if listar="100" then%>
                                          <option value="funcionarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                          <option value="funcionarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:100</option>
                                          <option value="funcionarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                          <option value="funcionarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                          <%end if%>
                                          <%if listar="200" then%>
                                          <option value="funcionarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                          <option value="funcionarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                          <option value="funcionarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:200</option>
                                          <option value="funcionarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:500</option>
                                          <%end if%>
                                          <%if listar="500" then%>
                                          <option value="funcionarios.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >LISTAR:30</option>
                                          <option value="funcionarios.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:100</option>
                                          <option value="funcionarios.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">LISTAR:200</option>
                                          <option value="funcionarios.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>LISTAR:500</option>
                                          <%end if%>
                                        </select>
                                      </div></td>
                                  </tr>
                                </form>
                              </table>
                            </div>
                            <div align="right"></div></td>
                          <td width="130"> <div align="center"> 
                              <table width="125" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="funcionarios.asp" method="get" name="form3">
                                  <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <tr> 
                                    <td>&nbsp;</td>
                                  </tr>
                                </form>
                              </table>
                            </div></td>
                          <td width="130"> <div align="right"> 
                              <table width="120" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="funcionarios.asp" method="get" name="form1">
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <tr> 
                                    <td> <input name="pesquisa" type="text" id="pesquisa"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #EAEAEA" value="<%=pesquisa%>" size="15"> 
                                      <input name="imageField3" type="image" id="imageField33" src="imagem/icones_ok.gif" alt="Procura Funcion&aacute;rio (Nome, Login)" width="20" height="13" border="0"> 
                                    </td>
                                  </tr>
                                </form>
                              </table>
                            </div></td>
                          <td width="370">&nbsp;</td>
                          <td width="50"><div align="center"></div>
                          </td>
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
  <div align="center"> 
    <%

	'set objrs01 = objConn01.execute("update funcionarios set senha='1064',resposta='1064' WHERE codfunc='1001'  ")
	'set objrs01 = objConn01.execute("delete from Funcionarios_Cont ")
	
	if auxapaga="sim" then
			set objrs01 = objConn01.execute("delete from funcionarios WHERE codfunc='"&auxcodfunc&"'  ")
	%>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong>O USU&Aacute;RIO DE COD  <%=auxcodfunc%> FOI APAGADO COM SUCESSO!</strong></div></td>
      </tr>
    </table>
    <br>
    <%	
end if
'SE É A PRIMEIRA VEZ QUE A PÁGINA É CARREGADA
'--------------------------------------------
If Session("primeiravez_geral_funcionarios") <> "Nao" then 
	Set objrs01 = Server.CreateObject("ADODB.Recordset")
	objrs01.CacheSize = listar ' tamanho do cache
	objrs01.PageSize = listar ' tamanho da página de registros
	objrs01.CursorLocation = 3
	call rotinasql
	if objrs01.RecordCount<>"0" then
		session("pagina_geral_funcionarios") = 1 
		CALL MostraDados 
		Session("primeiravez_geral_funcionarios") = "Nao"
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
			Session("pagina_geral_funcionarios") = 1 
		end if
		if Request("Navegacao") = "Primeira" then
			Session("pagina_geral_funcionarios") = 1 
		end if
		if Request("Navegacao") = "Proxima" then
			Session("pagina_geral_funcionarios") = Session("pagina_geral_funcionarios") + 1 
		end if
		If Request("Navegacao") = "Anterior" then
			Session("pagina_geral_funcionarios") = Session("pagina_geral_funcionarios") - 1
		End If
		if Request("Navegacao") = "Ultima" then
			Session("pagina_geral_funcionarios") = objrs01.PageCount
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
Sub MostraDados()
	objrs01.AbsolutePage = Session("pagina_geral_funcionarios") ' vai para o número da página que está armazenado em session("pagina")
if (Session("pagina_geral_funcionarios") <> 1 )or(Session("pagina_geral_funcionarios") <> 1)or(Session("pagina_geral_funcionarios") <> objrs01.PageCount )or(Session("pagina_geral_funcionarios") <> objrs01.PageCount) then
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
							  auxusu = objrs01.recordcount
	Response.Write "Existe(m) " & auxusu & " usuários(s) cadastrado(s) - Mostrando página " & Session("pagina_geral_funcionarios") & " de " & objrs01.PageCount 
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
if (Session("pagina_geral_funcionarios") <> 1 )or(Session("pagina_geral_funcionarios") <> 1)or(Session("pagina_geral_funcionarios") <> objrs01.PageCount )or(Session("pagina_geral_funcionarios") <> objrs01.PageCount) then

	If Session("pagina_geral_funcionarios") <> 1 then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> 1 then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
						auxusu = objrs01.recordcount
						Response.Write "Existe(m) " & auxusu & " usuários(s) cadastrado(s) - Mostrando página " & Session("pagina_geral_funcionarios") & " de " & objrs01.PageCount 
						%>
					</strong></div></td>
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
              <td width="36"><div align="center"><strong><strong><a href="funcionarios.asp?ordem=<% if ordem1 = 1 then %>codfunc&ordem1=0<%else%>codfunc desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Cód</a></strong>
                      <%if ordem="codfunc desc" then%>
                      <img src="imagem/cima.gif">
                      <%else if ordem="codfunc" then%>
                      <img src="imagem/baixo.gif">
                      <%end if%>
                      <%end if%>
              </strong></div></td>
              <td width="189"><div align="center"><strong><span class="linkA5"><strong><a href="funcionarios.asp?ordem=<% if ordem1 = 1 then %>nomeresumido&ordem1=0<%else%>nomeresumido desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Nome</a></strong>
                      <%if ordem="nome desc" then%>
                      <img src="imagem/cima.gif">
                      <%else if ordem="nome" then%>
                      <img src="imagem/baixo.gif">
                      <%end if%>
                      <%end if%>
              </span></strong></div>                <div align="center"></div></td>
              <td width="76"><div align="center"><strong></strong><strong><span class="linkA5"><strong><a href="funcionarios.asp?ordem=<% if ordem1 = 1 then %>login&ordem1=0<%else%>login desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">Login</a></strong>
                      <%if ordem="login desc" then%>
                      <img src="imagem/cima.gif">
                      <%else if ordem="login" then%>
                      <img src="imagem/baixo.gif">
                      <%end if%>
                      <%end if%>
              </span></strong></div></td>
              <td width="88"><div align="center" class="linkA5"><strong>Senha</strong></div></td>
              <td width="95"><div align="center"><strong></strong><strong><strong><a href="funcionarios.asp?ordem=<% if ordem1 = 1 then %>nacesso&ordem1=0<%else%>nacesso desc&ordem1=1<%end if%>&listar=<%=listar%>&pesquisa=<%=pesquisa%>&pesquisa2=<%=pesquisa2%>&tipo=<%=tipo%>&filtro=<%=filtro%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro2=<%=filtro2%>" class="linkA5">N° de Acessos</a></strong>
                    <%if ordem="nacesso desc" then%>
                    <img src="imagem/cima.gif">
                    <%else if ordem="nacesso" then%>
                    <img src="imagem/baixo.gif">
                    <%end if%>
                    <%end if%>
              </strong></div></td>
              <td width="117"><div align="center"><strong>Ult.Acesso</strong></div></td>
              <%'if ((nivel="2")) then%><td width="34"><div align="center"><strong>Ed</strong></div></td>
              <td width="40"><div align="center"><strong>Ap</strong></div></td><%'end if%>
            </tr>
            <%
			
	For contador = 1 to listar
	
	codfunc				= objrs01 ("codfunc")
	nome				= objrs01 ("nome")
	nomeresumido2 		= objrs01 ("nomeresumido")
	login 				= objrs01 ("login")
	senha   			= objrs01 ("senha")
	'bloqueado			= objrs01 ("bloqueado")
	nacesso  			= objrs01 ("nacesso")
	'resposta    		= objrs01 ("resposta")
	'dica				= objrs01 ("dica")
	'rg					= objrs01 ("rg")
	'cpf				= objrs01 ("cpf")
	'rastreado   		= objrs01 ("rastreado")
	'rastreado_hora	    = objrs01 ("rastreado_hora")	
	'rastreado_data	    = objrs01 ("rastreado_data")
	ult_acesso  		= objrs01 ("ultimoacesso")
'	cargo  				= objrs01 ("cargo")
	'permitebloqueio	= objrs01 ("permitebloqueio")

if nome <> "Mossoró" then
%>
		 		<%if bloqueado="sim" then%>
			<tr valign="middle" bgcolor="#FFC8C8" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFC8C8')"> 
	<%else%>
            <%if (contador mod 2)<>0 then%>
          <tr valign="middle" bgcolor="#FFFFFF" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#FFFFFF')"> 
              <%else%>
            <tr valign="middle" bgcolor="#E6E6E6" onMouseOver="setPointer(this, '#CCFFCC')" onMouseOut="setPointer(this, '#E6E6E6')"> 
              <%end if%>
	<%end if%>
              <td width="36"><div align="center"> 
                  <%response.write codfunc%>
              </div></td>
              <td>
                <div align="left">
                  <% 
response.write nomeresumido2
%>
                </div></td>
				
              <td width="76"> <div align="left"> 
                  <% 
response.write login
%>
                </div></td>
				
		      <td> 
		        <div align="center">
		          ******
	                </div></td><td width="95"> 
                
                <div align="center">
                  <% 
if login = "d4w" then
	response.write "-"
else	
	response.write nacesso
end if
%>
              </div></td><td width="117"><div align="center"><%=ult_acesso%></div></td>
             <td width="34"><div align="center"> 
                <%if nivel_usuarios = "2" then%>   <a href="funcionarios_editar.asp?codfunc=<%=codfunc%>&listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"><img src="imagem/icones_edita.gif" alt="Edita o Registro deste funcionario" width="20" height="20" border="0"></a> <%end if%>
              </div></td>
              <td width="40"><div align="center"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <form name="form4"  action="funcionarios.asp"   method="get" >
                      <input name="auxapaga" type="hidden" value="sim">
                      <input name="auxcodfunc" type="hidden" value="<%=codfunc%>">
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
                            <%if nivel_usuarios = "2" then%> <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga o registro deste funcionario" WIDTH="20" HEIGHT="20"><%end if%>
                          </div></td>
                      </tr>
                    </form>
                  </table>
                  <%
'else
'	response.write "-"
'end if
%>
                </div></td>
				
            </tr>
<%end if

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
        <td>&nbsp;</td>
        <td width="110"><div align="right"> 
            <%
if (Session("pagina_geral_funcionarios") <> 1 )or(Session("pagina_geral_funcionarios") <> 1)or(Session("pagina_geral_funcionarios") <> objrs01.PageCount )or(Session("pagina_geral_funcionarios") <> objrs01.PageCount) then

	If Session("pagina_geral_funcionarios") <> 1 then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> 1 then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_geral_funcionarios") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""funcionarios.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
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
CALL fecha_conexao01
CALL fecha_conexao02	


		
'Session.Contents.Remove ("pagina_geral_funcionarios") 
'Session.Contents.Remove ("primeiravez_geral_funcionarios") 
		

	%>

  </div>
</div>
</body>
</html>