<!--#include file="inc_conexao.asp"-->
<!--#include file="inc_css.asp"-->

<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<title>Untitled Document</title>
</head>
<%
if session("codforn") = "" then
%>
<script language="javascript" type="text/javascript">
	alert("Você não está logado favor pedir para algum administrador digitar a senha")
	window.location.href = "sair.asp"
</script>
<%end if%>

<% 
call abre_conexao01
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02

codfor = session("codforn")
codcot = request("codcot")

'criamos o nome do arquivo 
set objrs01 = objConn01.execute("select empresa from fornecedores WHERE codfor="&codfor&"")
if not objrs01.eof then
	arquivo2 = "fornecedores/"&codcot&"/"&codfor&" - "& objrs01("empresa") &".txt" 
	auxforn = codfor&" - "& objrs01("empresa")
	arquivo= request.serverVariables("APPL_PHYSICAL_PATH") & "geral/caixa/"&arquivo2
	
	'conectamos com o FSO 
	set confile = createObject("scripting.filesystemobject") 
	
	'criamos o objeto TextStream 
	set fich = confile.CreateTextFile(arquivo) 
	
	'escrevemos 
	set objrs01 = objConn01.execute("select * from prod_cot_forn2 WHERE codfor="&codfor&" and preco >= 0 and codcot = "&codcot&" ORDER BY cont ")
	if not objrs01.eof then
		fich.write("Cotação n°: "&codcot&" Forn.: "&auxforn) 
		fich.WriteBlankLines(2)
	end if
	while not objrs01.eof 
		cont 		= objrs01("cont")
		codprc   	= objrs01("codprc")
		preco   	= objrs01("preco")
	
		set objrs02= objConn02.execute("select * from produtos_cot WHERE codprc="&codprc&"")
		if not objrs02.eof then
			codpro  = objrs02("codpro")
		else
			produto = ""
		end if
	
		set objrs02= objConn02.execute("select * from produtos WHERE codpro="&codpro&"")
		if not objrs02.eof then
			produto  = objrs02("produto")
		else
			produto = ""
		end if
	
		if not isnumeric(preco) then
			preco = 0
		end if
		
		if (preco <> 0) and (produto <> "") then
			produto = ucase(right(produto,35))
			if len(produto) < 35 then
				for i=0 to 35 - len(produto)
					produto = produto & " "
				next
			end if
			preco = formatnumber(preco,2)
			descricao = produto & preco
			fich.write(descricao) 
			fich.WriteBlankLines(1)
		end if
		objrs01.movenext
	wend
	'fechamos o arquivo 
	fich.close() 
	
	'voltamos a abrir o arquivo para leitura 
	'set fich = confile.OpenTextFile(arquivo) 
	
	'lemos o conteúdo do arquivo 
	'texto_arquivo = fich.readAll() 
	
	'imprimimos na página o conteúdo do arquivo 
	'response.write(texto_arquivo) 
	
	'fechamos o arquivo 
	'fich.close() 


	
end if


%> 

<!--<body onLoad="javascript:MM_openBrWindow('<%=arquivo2%>','','status=yes,scrollbars=yes,width=100,height=100');window.location.href = 'sair.asp';">-->
<body onLoad="javascript:window.location.href = 'sair.asp';">
</body>
</html>
