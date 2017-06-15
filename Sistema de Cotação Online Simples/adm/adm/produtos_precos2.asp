<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<!--#include file="inc_css.asp"-->
<%
call abre_conexao01
call abre_conexao02
nivel = nivel_cotacao
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
nivel_aux = nivel
'nivel_aux = 2

codpro = request("codpro")

set objrs01 = objconn01.execute("select * from cotacao where aberta='sim'")
if not objrs01.eof then
	codcot = objrs01("codcot")
	set objrs01 = objconn01.execute("select codprc from produtos_cot where codcot="&codcot&" and codpro="&codpro&" ")
	if not objrs01.eof then
		codprc = objrs01("codprc")
		set objrs01 = objconn01.execute("select * from prod_cot_forn where codprc="&codprc&" order by preco")		
		while not objrs01.eof
			set objrs02 = objconn02.execute("select * from fornecedores where codfor="&objrs01("codfor")&"")		
			if not objrs02.eof then
				preco = formatnumber(objrs01("preco"),2)
				empresa = objrs02("empresa")
				for i=len(empresa)-50 to 50 
					empresa = empresa & "&nbsp;"
				next
				response.write "<div class='linkA'>"&empresa&" R$ "&preco&"<BR></div>"
			else
				response.write "<div class='linkA'>Erro! Fornecedor não encontrado</div>"
			end if
			objrs01.movenext
		wend		
	else
		response.write "<div class='linkA'>Erro! Produto não Encontrado</div>"
	end if
else
	response.write "<div class='linkA'>Erro! Não existe cotação aberta</div>"
end if


call fecha_conexao01
call fecha_conexao02
%>

</div></div>
</body>
</html>