<!--#include file="inc_login.asp"-->
<%topo = "sim"%>
<!--#include file="inc_css.asp"-->
<!--#include file="inc_tempo.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%

call abre_conexao01
call abre_conexao02

%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onFocus="javascript:reload();">
<BGSOUND id="sound" src="">
<table width="583" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" bgcolor="#F2F2F2"> <div align="center"><br>
        <table width="95%" border="0" cellspacing="0" cellpadding="0"  class="linkA2">
          <tr> 
            <td> 
			
<table width="100%" border="0" cellspacing="0" cellpadding="0" class="linkA2">
  <tr>
                  <td><div align="center">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="linkA2">
                        <tr>
                          <td width="94%"><div align="center"><strong><u>AVISOS</u></strong></div></td>
                          <td width="6%"><div align="right"></div></td>
                        </tr>
                      </table>
                      <strong> <br>
                      </strong></div></td>
  </tr>
</table>
			
              <%


		set objrs02 = objconn02.execute(" select * from comunicador_func where (codfunc='"&codfunc&"')  and lido <= 5 order by codcomdufunc")
		if not objrs02.eof then
			while not objrs02.eof 
			
				set objrs01 = objconn01.execute("update comunicador_func set lido = lido + 1 where cod='"&objrs02("cod")&"' and (codfunc='"&codfunc&"')")

				set objrs01 = objConn01.execute("select * from comunicador order by data desc, hora desc")
				cod			= objrs01("cod")
				mensagem 	= objrs01("mensagem")
				data	 	= objrs01("data")
				hora		= objrs01("hora")
				codfunc2	= objrs01("codfunc")
				
				set objrs01 = objconn01.execute("select * from  Funcionarios where codfunc='"&codfunc2&"' ")	
				if not objrs01.eof then
					nome = objrs01("nomeresumido")
				end if	
						
				response.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class=linka><tr><td bgcolor='#EAEAEA'>"
				response.write "<b>" & data & " - "  & hora & "h / Remetente : " & nome & "</b><br>"
				response.write "</td></tr></table>"
				response.write "<font class=linka3>" & replace(mensagem,"" & chr(13) & "","<br>") & "</font>"
				response.write "<BR><BR>"
				
			objrs02.movenext
			wend
		else
%>
              <div align="center"></div>
              <table width="100%" border="0" cellspacing="0" cellpadding="0"  class="linkA2">
                <tr> 
                  <td><div align="center"><br>
                      <br>
                      <br>
                      <br>
                      <br>
                      NÃO EXISTE NENHUMA MENSAGEM NA AGENDA PARA HOJE!</div></td>
                </tr>
              </table>
              <%	
	end if
call fecha_conexao01
call fecha_conexao02
%>
            </td>
          </tr>
        </table>
        
      </div></td>
  </tr>
</table>
</body>
</html>
