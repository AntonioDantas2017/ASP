<!--#include file="inc_login.asp"-->
<%topo = "sim"%>
<!--#include file="inc_tempo.asp"-->
<!--#include file="inc_css.asp"-->

<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<div align="center"><br>
  <strong> </strong> 
  <table width="700" border="0" cellspacing="0" cellpadding="0" class=linka>
    <tr> 
      <td bgcolor="#CCCCCC"><div align="center"> <strong>BEM-VINDO</strong></div></td>
    </tr>
  </table>
  <table width="700" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC" CLASS=LINKA>
    <tr> 
      <td><div align="center"><strong> </strong>
          <table width="95%" height="1" border="0" cellpadding="0" cellspacing="0" class=linka>
            <tr> 
              <td width="73%"><div align="center"><strong>
                  <% 
'response.write "<strong><font color='#FF0000'>"

'Set bc = Server.CreateObject("MSWC.BrowserType") 
'browser = bc.browser 
'versao = bc.version	
'if (browser <> "IE") then
'	response.write "<BR>"
'	response.write "<BR>"
'	response.write "O seu browser � o "& browser & "."
'	response.write "<BR>"
'	response.write "O browser recomendado � o Internet Explorer vers�o 6.0"
'	response.write "<BR>"
'	response.write "A utiliza��o deste browser pode apresentar erros durante a utiliza��o do sistema."
'	response.write "<BR>"
'else
'	if (versao < 6 ) then
'		response.write "<BR>"
'		response.write "<BR>"
'		response.write "A vers�o do seu Internet Explorer � a "& versao & "."
'		response.write "<BR>"
'		response.write "A vers�o recomendada � o Internet Explorer 6.0"
'		response.write "<BR>"
'		response.write "A utiliza��o desta vers�o pode apresentar erros durante a utiliza��o do sistema."
'		response.write "<BR>"
'	end if
'end if

'response.write "</font></strong>"

%>
                  <br>
                  <br>
                  <%
response.Write "Ol� "& nomeresumido&", "

if (Cint(hour(time))>=0) then 
	saudacao= "Boa Madrugada."
end if
if (Cint(hour(time))>=6) then 
	saudacao = "Bom Dia."
end if
if (Cint(hour(time))>=12) then 
	saudacao = "Boa Tarde."
end if
if (Cint(hour(time))>=18) then 
	saudacao = "Boa Noite."
end if
response.write saudacao
response.write "<BR><BR>Este � seu "& nacesso & "� acesso ao sistema."
%>
                  <br>
                  <br>
                  </strong> 
           
		          <%

dim aviso, ultimoacesso

call abre_conexao01
sql = "select ultimoacesso from funcionarios where codfunc= '"&codfunc&"'"
set objrs01 = objConn01.execute(sql)


if objrs01("ultimoacesso") <> "" then
	response.write " Seu �ltimo acesso foi dia "
	response.write objrs01 ("ultimoacesso")
%>
                  <br>
                  caso n�o tenha sido voc� que realizou o �ltimo acesso, <br>
                  entre em contato com o suporte t�cnico do sistema." 
                  <%	
else
'	response.write " <strong> Este � seu primeiro acesso ao sistema.</strong>"
end if

'ATUALIZA O ULTIMO ACESSO 
'------------------------
strData 	= fdata(date)
ultimoacesso = strData&" �s "&time
set objrs01 = objConn01.execute("update funcionarios set ultimoacesso='"&ultimoacesso&"' WHERE codfunc='"&codfunc&"'")

call fecha_conexao01
%>
                  <strong> <br>
                  <br>
                  </strong>
                  <table width="600" height="290" border="1" cellpadding="0" cellspacing="0">
                    <tr>
                      <td valign="top"><div align="center">
                          <iframe src="default-centro-agenda.asp" width="600" marginwidth="0" height="290" marginheight="0" scrolling="yes" frameborder="0"></iframe>
                      </div></td>
                    </tr>
                  </table>
                  <strong><br>
                  </strong> <strong> <br>
                  <br>
                  </strong><strong> </strong></div></td>
            </tr>
          </table>
          <strong> </strong><strong> </strong> </div></td>
    </tr>
  </table>
  <strong> </strong> <br>
</div>
</body>
</html>
