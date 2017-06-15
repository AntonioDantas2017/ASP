<!--#include file="inc_conexao.asp"-->
<%
call abre_conexao01
dim objrs01,objconn01

set objrs01 = objConn01.execute("delete from prod_cot_forn")	
set objrs01 = objConn01.execute("delete from produtos_cot")		
set objrs01 = objConn01.execute("delete from temp_cotacao")		
set objrs01 = objConn01.execute("delete from cotacao")
%>
