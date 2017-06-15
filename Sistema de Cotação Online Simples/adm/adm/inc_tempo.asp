<%
'GRAVA O HORÁRIO DE SAÍDA DO SISTEMA
'-----------------------------------
call abre_conexao01
set objRs01 = objConn01.execute("update log set termino='"&time&"' WHERE cont='"&Session("log")&"' ")
call fecha_conexao01
%>