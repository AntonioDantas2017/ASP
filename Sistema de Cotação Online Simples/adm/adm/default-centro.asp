
<!--#include file="inc_login.asp"-->
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
             
                  <br>
                  <br>
                  <%
response.Write "Olá "& nomeresumido&", "

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
response.write "<BR><BR>Este é seu "& nacesso & "º acesso ao sistema."
%>
                  <br>
                  <br>
                  </strong> 
           
		          <%

dim aviso, ultimoacesso

call abre_conexao01
data2 = split(date,"/")
data = right("0"&data2(1),2) & "/" & right("0"&data2(0),2) & "/" & data2(2)
'sql = "select data from gas where data= #"&data&"#"
'set objrs01 = objConn01.execute(sql)
'response.Write sql
'if objrs01.eof then
'	set objRs01 = objConn01.execute("select max(codgas) as maxcodgas from gas")
'	if IsNull(objRs01("maxcodgas")) then
'	   codgas = 100000
'	else
'	   codgas = objRs01("maxcodgas") + 1
'	end if
'	
'	set objRs01 = objConn01.execute("select estoque from gas_configuracao")	
'	estoque = objrs01("estoque")
'	
'	sql = "INSERT INTO gas (codgas,data,aberta,estoque) values ("&codgas&",'"&date&"','sim',"&estoque&") "
'	'response.write sql	
'	set objRs01 = objConn01.execute(sql)
'end if



sql = "select ultimoacesso from funcionarios where codfunc= '"&codfunc&"'"
set objrs01 = objConn01.execute(sql)


if objrs01("ultimoacesso") <> "" then
	response.write " Seu último acesso foi dia "
	response.write objrs01 ("ultimoacesso")
%>
                  <br>
                  caso não tenha sido você que realizou o último acesso, <br>
                  entre em contato com o suporte técnico do sistema." 
                  <%	
else
'	response.write " <strong> Este é seu primeiro acesso ao sistema.</strong>"
end if

'ATUALIZA O ULTIMO ACESSO 
'------------------------
arrData = Split(date,"/")
strData = Right("0" & arrData(0),2) & "/" & Right("0" & arrData(1),2) & "/" & arrData(2)
ultimoacesso = strData&" às "&time
set objrs01 = objConn01.execute("update funcionarios set ultimoacesso='"&ultimoacesso&"' WHERE codfunc='"&codfunc&"'")

call fecha_conexao01
%>
                  <strong> <br>
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
