<!--#include file="inc_login.asp"-->
<!--#include file="inc_css.asp"-->
<%
if nivel_avisos = "0" then
	session("pg") = request.ServerVariables("SCRIPT_NAME")
    Response.Redirect ("default_acessonegado.asp")
end if
%>

<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body leftmargin="0" topmargin="00" marginwidth="0" marginheight="0"  onLoad="form3.pesquisa.focus();">
<div align="center"> 
  <%
 
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
Dim sql, auxcod, auxapaga, Contador, rastreado
Dim cod,categoria,titulo,status2,data,total,contato,email,prazo




'RESGATA AS VARIÁVEIS
'--------------------
auxcod 		 = request("auxcod")
auxapaga     = request("auxapaga")
pesquisa     = request("pesquisa")
auxpaginacao = request("auxpaginacao")
datainicio   = request("datainicio")
datafim  	 = request("datafim")
ordem 	     = request("ordem")
listar       = request("listar")
primeira   	 = request("primeira")
filtro 		 = request("filtro")


if filtro="" then
   filtro = "enviado"
end if
if ordem="" then
   ordem = "data"
end if
if listar="" then
   listar = "30"
end if


if (datainicio="") and (datafim="") then
	datainicio = date
	datainicio 	= fdata(datainicio)

	datafim = date + 10
	datafim		= fdata(datafim)
end if

'ROTINA DE SQL
'-------------
sub rotinasql
	if primeira <> "sim" then
	
		sql = "SELECT 	cod ,data,hora,mensagem from comunicador where cod > '0'"

		nometratado= preparapalavra(pesquisa)
		
		if datainicio <> "" then
			auxdatainicio 	= cdata(datainicio)
		end if
		if datafim <> "" then
			auxdatafim		= cdata(datafim)
		end if
	   
	    
	   if filtro = "recebido" then
	 
	 		sql= "SELECT * FROM comunicador C,comunicador_func F where C.cod=F.cod and (F.codfunc = '"&codfunc&"' or F.codfunc = '0000' )  "
			if pesquisa <> "" then
				sql = sql & " and(mensagem Like '%"&(nometratado)&"%') "
			end if
		
			if (datainicio <> "") and (datafim<>"") then
				 sql = sql & " and data>= #"&auxdatainicio&"# and data<= #"&auxdatafim&"# "
			end if
		else
		   sql= "SELECT * FROM comunicador where codfunc = '"&codfunc&"'"
			if pesquisa <> "" then
				sql = sql & " and(mensagem Like '%"&(nometratado)&"%') "
				'response.Write(sql)
			end if
		
		end if
        
		
		sql= sql & " order by '"&ordem&"'"
'		response.Write(sql)
		set objrs01 = objConn01.execute(sql)
'	 	objrs01.Open Sql, objConn01
	end if

end sub
'AVISO PARA RESULTADO DA PESQUISA EM BRANCO
'------------------------------------------
sub vazio
		if primeira = "sim" then
	else
		response.write "<br><br><br><br><br><br><br>"		
		response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class=linka><tr><td><div align='center'><b><font face='Verdana, Arial, Helvetica, sans-serif'>Não foi encontrado nenhum registro.</font> </b></div></td></tr></table>"
	end if
end sub
%>
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
    </tr>
  </table>
  <table width="700" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="104" valign="top"> <table width="100" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1" CLASS=LINKC>
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" >
                      <tr>
                        <td class=LINKC><div align="center"><strong>AVISOS</strong></div></td>
                      </tr>
                    </table>
                  </div></td>
                </tr>
              </table></td>
          </tr>
      </table></td>
      <td width="596"><table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td bgcolor="#000000"><table width="100%" height="30" border="0" cellpadding="1" cellspacing="1">
                <tr> 
                  <td bgcolor="#BBECB1"><div align="center"> 
                      
					  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="25%"> <div align="center"> 
                              <table width="103%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="fr_mensageiro.asp" method="post" name="form1"> 
                                  <input name="pesquisa2" type="hidden" value="<%=pesquisa%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
                                  <tr> 
								      <td width="6%"> <div align="center"><strong> <a class="linkB2" href="javascript:history.go(-1)"><img src="imagem/icones_volta.gif" alt="Volta &agrave; tela anterior" width="20" height="20" border="0"><br>
								      </a></strong></div></td>
                          <td width="6%"><div align="center"><strong> <a class="linkB2"  href="javascript:location.reload()"> 
                              </a></strong><strong><a class="linkB2" href="JavaScript: location.reload()"><img src="imagem/icones_atualiza.gif" alt="Atualiza esta tela" width="20" height="20" border="0"><br>
                              </a></strong></div></td>
                          <td width="6%">
						  
						  
						  <div align="center">
                              <div align="center">
                                <div align="center"><strong>
								
								<%if nivel_avisos = "2" then%>
								<a class="linkB2" href="fr_mensageiro_inserir.asp?listar=<%=listar%>&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>&filtro=<%=filtro%>"><img src="imagem/icones_insere.gif" alt="Insere um novo registro " width="20" height="20" border="0"><br>
                                </a>
								<%end if%>
								</strong> </div>
                              </div>
                              </div></td>
                                    <td height="12"> <div align="center">Listar:
                                        <select name="select" id="select" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFFFFF" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                                           <%if listar="30" then%>
                                          <option value="fr_mensageiro.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" selected>30</option>
                                          <option value="fr_mensageiro.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">100</option>
                                          <option value="fr_mensageiro.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">200</option>
                                          <option value="fr_mensageiro.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">500</option>
                                          <%end if%>
                                          <%if listar="100" then%>
                                          <option value="fr_mensageiro.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >30</option>
                                          <option value="fr_mensageiro.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>100</option>
                                          <option value="fr_mensageiro.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">200</option>
                                          <option value="fr_mensageiro.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">500</option>
                                          <%end if%>
                                          <%if listar="200" then%>
                                          <option value="fr_mensageiro.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >30</option>
                                          <option value="fr_mensageiro.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">100</option>
                                          <option value="fr_mensageiro.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>200</option>
                                          <option value="fr_mensageiro.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">500</option>
                                          <%end if%>
                                          <%if listar="500" then%>
                                          <option value="fr_mensageiro.asp?listar=30&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>" >30</option>
                                          <option value="fr_mensageiro.asp?listar=100&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">100</option>
                                          <option value="fr_mensageiro.asp?listar=200&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>">200</option>
                                          <option value="fr_mensageiro.asp?listar=500&ordem=<%=ordem%>&pesquisa=<%=pesquisa%>"selected>500</option>
                                          <%end if%>
                                        </select>
                                      </div></td>
                                  </tr>
                                </form>
                              </table>
                            </div>
                          <div align="right"></div></td>
                         
                          <td width="23%"><div align="left"> 
                              <script language="JavaScript">
function ValidateOrder(form)
{
  if (form1.pesquisa.value == "")
  { 
	  alert("Por favor digite a palavra de deseja pesquisar."); form1.pesquisa.focus(); return; 
  }
  form.submit();
}
</script>
                              <table width="100%" border="0" cellspacing="0" cellpadding="0" class=linka>
                                <form action="fr_mensageiro.asp" method="post" name="form3"> 
                                  <input name="listar" type="hidden" value="<%=listar%>">
                                  <input name="ordem" type="hidden" value="<%=ordem%>">
								  <input name="datainicio" type="hidden" value="<%=datainicio%>">
								   <input name="datafim" type="hidden" value="<%=datafim%>">
								   <input name="filtro" type="hidden" value="<%=filtro%>">
                                  <tr> 
                                    <td><p align="right">Pesq: 
                                        <input name="pesquisa" type="text" id="pesquisa"  style=" text-transform: uppercase; FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFFFFF" value="<%=pesquisa%>"  size="10">
                                        <input name="imageField322" type="image" id="imageField32" src="imagem/icones_ok.gif" alt="Procura Usu&aacute;rio (Mensagem)" width="20" height="13" border="0">
                                      </p></td>
                                  </tr>
                                </form>
                              </table>
                          </div></td>
                          <td width="36%"><div align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA">
									 <form action="fr_mensaGEIRO.asp" method="post" name="form4"> 
                                <input name="listar" type="hidden" value="<%=listar%>">
                                <input name="ordem" type="hidden" value="<%=ordem%>">
							    <input name="pesquisa" type="hidden" value="<%=pesquisa%>">
							    <input name="filtro" type="hidden" value="<%=filtro%>">
	 						    <input name="filtro2" type="hidden" value="<%=filtro2%>">
                                <tr> 
                                  <td> De 
                                    <input name="datainicio" type="text" id="datainicio"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFFFFF" value="<%=datainicio%>" size="11" maxlength="10">
                                    At&eacute; 
                                    <input name="datafim" type="text" id="datafim"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFFFFF" value="<%=datafim%>" size="11" maxlength="10"> 
                                    <input name="imageField32" type="image" id="imageField3" src="imagem/icones_ok.gif" alt="Procura Usu&aacute;rio (Data emissão,Data vencimento)" width="20" height="13" border="0"> 
                                  </td>
                                </tr>
                              </form>
													</table>
						  </div></td>
							<td width="16%"><div align="center">
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td> <div align="right">
                                    <select name="filtro" id="filtro" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #FFFFFF" onChange="MM_jumpMenu('parent.frames[\'mainFrame\']',this,0)" >
                                      
                                      <%if filtro="enviado" then%>
                                      <option value="fr_mensageiro.asp?filtro=enviado&listar=<%=listar%>&pesquisa=<%=pesquisa%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>" selected>Enviado</option>
                                      <option value="fr_mensageiro.asp?filtro=recebido&listar=<%=listar%>&pesquisa=<%=pesquisa%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>" >Recebido</option>
                                      <%end if%>
                                      <%if filtro="recebido" then%>
                                      <option value="fr_mensageiro.asp?filtro=enviado&listar=<%=listar%>&pesquisa=<%=pesquisa%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>" >Enviado</option>
                                      <option value="fr_mensageiro.asp?filtro=recebido&listar=<%=listar%>&pesquisa=<%=pesquisa%>&datainicio=<%=datainicio%>&datafim=<%=datafim%>" selected>Recebido</option>
                                      <%end if%>
                                    </select>
                                  </div></td>
                                </tr>
                              </table>
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
  <table width="200" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
    </tr>
  </table>
  <div align="center"> 
<%
if auxapaga="sim" then
call abre_conexao01
	set objrs01 = objConn01.execute("delete from comunicador WHERE cod='"&auxcod&"' ")
	
	set objrs01 = objConn01.execute("delete from comunicador_func WHERE cod='"&auxcod&"' ")
call fecha_conexao01
%>
    <table width="700" border="0" cellspacing="1" cellpadding="1" class=linka>
      <tr> 
        <td bgcolor="#FF9999"><div align="center"> <strong><font color="#FF0000"> 
            </font>A REGISTRO DE C&Oacute;DIGO <%=auxcod%> FOI APAGADO COM SUCESSO!</strong></div></td>
      </tr>
    </table>
    <br>
    <%	
end if
%>
<%if primeira <> "sim" then%>
  <%call abre_conexao01
				   
	'SE É A PRIMEIRA VEZ QUE A PÁGINA É CARREGADA
	'--------------------------------------------
	If Session("pagina_agenda2") <> "Nao" then 
		Set objrs01 = Server.CreateObject("ADODB.Recordset")
		objrs01.CacheSize = listar ' tamanho do cache
		objrs01.PageSize = listar ' tamanho da página de registros
		objrs01.CursorLocation = 3
		call rotinasql
		if objrs01.RecordCount<>"0" then
			session("pagina_agenda") = 1 
			CALL MostraDados 
			Session("pagina_agenda2") = "Nao"
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
				Session("pagina_agenda") = 1 
			end if
			if Request("Navegacao") = "Primeira" then
				Session("pagina_agenda") = 1 
			end if
			if Request("Navegacao") = "Proxima" then
				Session("pagina_agenda") = Session("pagina_agenda") + 1 
			end if
			If Request("Navegacao") = "Anterior" then
				Session("pagina_agenda") = Session("pagina_agenda") - 1
			End If
			if Request("Navegacao") = "Ultima" then
				Session("pagina_agenda") = objrs01.PageCount
			end if
			
			if auxpaginacao <> "" then
				Session("pagina_agenda") = auxpaginacao
			end if
			
			CALL MostraDados
		else
			CALL Vazio
		end if
	End If 
call fecha_conexao01
'MENU DE PAGINAÇÃO
'-----------------
sub menu
End sub
Sub MostraDados()
	
if (Session("pagina_agenda") <> 1 )or(Session("pagina_agenda") <> 1)or(Session("pagina_agenda") <> objrs01.PageCount )or(Session("pagina_agenda") <> objrs01.PageCount) then
%>
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td bgcolor="#999999"><table width="100%" border="0" cellspacing="1" cellpadding="1" class=linka>
                        <tr> 
                          <td width="50%" bgcolor="#FFFFFF"><div align="center"><strong> 
                              <%
	if filtro  = "enviado" then
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) ENVIADO(S) - Mostrando página " & Session("pagina_agenda") & " de " & objrs01.PageCount 
	ELSE
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) RECEBIDOS(S) - Mostrando página " & Session("pagina_agenda") & " de " & objrs01.PageCount 
	END IF
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
if objrs01.PageCount >1 then			
if (Session("pagina_agenda") <> 1 )or(Session("pagina_agenda") <> 1)or(Session("pagina_agenda") <> objrs01.PageCount )or(Session("pagina_agenda") <> objrs01.PageCount) then

	If Session("pagina_agenda") <> 1 then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> 1 then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'>"
	End If
	
End if
end if
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
                    <td width="50%" bgcolor="#FFFFFF"><div align="center"><strong> 
                        <%
	                        
	if filtro  = "enviado" then
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) ENVIADO(S) - Mostrando página " & Session("pagina_agenda") & " de " & objrs01.PageCount 
	ELSE
	Response.Write "Existe(m) " & objrs01.RecordCount & " registro(s) RECEBIDOS(S) - Mostrando página " & Session("pagina_agenda") & " de " & objrs01.PageCount 
	END IF

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
    <table width="200" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="10"></div></td>
      </tr>
    </table>
	
    <div align="center">
      <div align="center"></div>
    </div>
</div>
  <div align="center"> 
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td bgcolor="#999999"><table width="700" border="0" cellpadding="1" cellspacing="1" CLASS=LINKA>
            <tr valign="top" bgcolor="#CCCCCC"> 
              <td width="100"><div align="center">
                  <p><strong>DATA<br>
                    </strong></p>
                </div></td>

              <td width="600"><div align="center"><strong>MENSAGEM</strong></div></td>

              <td width="300"><div align="center"><strong>DESTINAT&Aacute;RIO</strong></div></td>

              <td width="20"><div align="center"><strong>EDT</strong></div></td>
              <td width="20"><div align="center"><strong>APG</strong></div></td>
            </tr>
<%
     'response.Write(listar)
	For contador = 1 to listar
	cod 		= objrs01 ("cod")
	data		= objrs01 ("data")
	hora		= objrs01 ("hora")
	mensagem	= objrs01 ("mensagem")
call abre_conexao02
	set objrs02 = objconn02.execute("select * from  funcionarios where codfunc='"&codfunc&"' ")	
	if not objrs02.eof then
		inic = objrs02 ("nomeresumido")
	end if
call fecha_conexao02
%>


   	<tr valign="middle" bgcolor="#FFFFFF"> 
			  

				
<td width="100"><div align="center">
<%
if data<>"" then
	data 	= fdata(data)
	response.write data
else
	response.write "-"
end if
%>
</div></td>				
				

              <td width="600"> <div align="left">
                  <%
if mensagem<>"" then
	response.write ucase(replace(mensagem,"" & chr(13) & "","<br>"))
else
	response.write "-"
end if
%>
                </div></td>
              <td width="300"> <div align="center"> 
                  <%call abre_conexao02
				    call abre_conexao03
								set objrs02 = objConn02.execute("select * from comunicador_func where cod='"&cod&"' ")
								if not objrs02.eof then
'									if objrs02("codfunc") = "0000" then
'										response.Write "Todos"
'									else	  
'	'									set objrs01 = objConn01.execute("select * from comunicador_func where cod='"&cod&"' ")
										while not objrs02.eof
											codfunc = objrs02 ("codfunc")
											set objrs03 = objConn03.execute("select * from funcionarios where codfunc='"&codfunc&"' ")
											if not objrs03.eof then
												response.write LEFT(objrs03("nomeresumido"),20) & "<BR>"
											end if
										objrs02.movenext
										wend
'									end if
					call fecha_conexao02
				    call fecha_conexao03
								else
									response.write "-"
								end if
							%>
                </div></td>

              <td width="20"><div align="center">
			  <%if nivel_avisos = "2" then%>
			  
			  <a href="fr_mensageiro_editar.asp?cod=<%=cod%>&auxpaginacao=<%=Session("pagina_agenda")%>&listar=<%=listar%>"><img src="IMAGEM/icones_edita.gif" alt="Altera este registro" width="20" height="20" border="0"></a>
			  <%end if%>
			  </div></td>
              <td width="20"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <form name="form5" method="post" action="fr_mensageiro.asp">
                      <input name="auxapaga" type="hidden" value="sim">
					  <input name="auxpaginacao" type="hidden" value="<%=Session("pagina_agenda")%>">
                      <input name="auxcod" type="hidden" value="<%=cod%>">
                      <input name="listar" type="hidden" value="<%=listar%>">
                      <tr> 
                        <td><div align="center"> 
                            <script language="JavaScript">
	function confirmar()
	{
	   return (confirm('Deseja apagar este registro?'))
	}
	</script>
                      <%if nivel_avisos = "2" then%>
					  
					        <input NAME="delete" TYPE="image" OnClick="return confirmar()" SRC="imagem/icones_apaga.gif" alt="Apaga este Registro" WIDTH="20" HEIGHT="20">
							<%end if%>
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
    <table width="700" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td>&nbsp;</td>
        <td width="110"><div align="right"> 
            <%
if objrs01.PageCount >1 then				
if (Session("pagina_agenda") <> 1 )or(Session("pagina_agenda") <> 1)or(Session("pagina_agenda") <> objrs01.PageCount )or(Session("pagina_agenda") <> objrs01.PageCount) then

	If Session("pagina_agenda") <> 1 then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Primeira&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_primeiro.gif' alt='Primeira tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> 1 then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Anterior&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_anterior.gif' alt='Tela anterior' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Proxima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_proxima.gif' alt='Próxima tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_proxima.gif'  alt='Próxima tela' width='20' height='20' border='0'>"
	End If
	If Session("pagina_agenda") <> objrs01.PageCount then
		response.write "&nbsp;"&"<a href=""geral_agenda.asp?Navegacao=Ultima&ordem="&ordem&"&pesquisa="&pesquisa&"&listar="&listar&"""><img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'></a>"
	else
		response.write "&nbsp;"&"<img src='imagem/icones_ultima.gif' alt='Última tela' width='20' height='20' border='0'>"
	End If
	
End if
end if
End Sub
%>
          </div></td>
      </tr>
	  <%end if%>
    </table>


  </div>
</div>
</body>
</html>