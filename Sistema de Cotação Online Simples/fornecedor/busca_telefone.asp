<!--#include file="inc_conexao.asp"-->

<%
telefone = request("telefone")

call abre_conexao01
dim objrs01,objconn01

set objrs01 = objconn01.execute("select * from gas_itens where telefone='"&telefone&"'")
if not objrs01.eof then
	rua 	= objrs01("rua")
	numero	= objrs01("numero")
	bairro	= objrs01("bairro")
%>
<script language="javascript" type="text/javascript">
	rua 	= "<%=rua%>";
	numero	= "<%=numero%>";
	bairro	= "<%=bairro%>";
	
	opener.form2.rua.value = rua;
	opener.form2.numero.value = numero;
	opener.form2.bairro.value = bairro;
</script>
<%end if%>
<script language="javascript" type="text/javascript">
	self.window.close();
</script>