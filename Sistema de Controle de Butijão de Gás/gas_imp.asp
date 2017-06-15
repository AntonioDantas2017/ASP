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
call abre_conexao01
dim objrs01,objconn01

call abre_conexao02
dim objrs02,objconn02

coditem = request("coditem")

'criamos o nome do arquivo 
set objrs01 = objConn01.execute("select * from gas_itens WHERE coditem="&coditem&"")
if not objrs01.eof then
	arquivo2 = "fornecedores/gas/gas"&coditem&".txt"
	arquivo= request.serverVariables("APPL_PHYSICAL_PATH") & "geral/caixa/"&arquivo2
	
	'conectamos com o FSO 
	set confile = createObject("scripting.filesystemobject") 
	
	'criamos o objeto TextStream 
	set fich = confile.CreateTextFile(arquivo) 
	
	'escrevemos 
	fich.write("Venda n°: "&right(coditem,5))
	fich.WriteBlankLines(2)

	hora	 		= objrs01("hora")
	forma	 		= objrs01("forma")
	valor	 		= objrs01("valor")
	op		 		= objrs01("op")

	data2		= split(date,"/")
	data		= right("0"&data2(0),2)&"/"&right("0"&data2(1),2)&"/"&data2(2)

	fich.write("Data: "&data&" Hora: "&hora&" Operador: "&op) 
	fich.WriteBlankLines(1)
	fich.write("Valor: "&formatnumber(valor,2)&" Forma de Pagamento: "&forma) 
	fich.WriteBlankLines(1)	

	aux_dados		= ""
	rua		 		= objrs01("rua")
	bairro	 		= objrs01("bairro")
	numero	 		= objrs01("numero")
	nome	 		= objrs01("nome")
	telefone 		= objrs01("telefone")
	
	if (rua<>"") or (numero<>"") then
		aux_dados = "1"
		fich.write("Rua: "&rua&" n°: "&numero) 
		fich.WriteBlankLines(1)	
	end if
	if (bairro<>"") or (telefone<>"") then
		aux_dados = "1"	
		fich.write("Bairro: "&bairro&" Telefone: "&telefone) 
		fich.WriteBlankLines(1)			
	end if
	if (nome<>"") then
		aux_dados = "1"	
		fich.write("Nome Entregador: "&nome) 
		fich.WriteBlankLines(1)			
	end if
	
	if aux_dados = "" then
		fich.write("LEVOU") 
		fich.WriteBlankLines(1)				
	end if

	'fechamos o arquivo 
	fich.close() 
end if
%> 

<script>
//	MM_openBrWindow('gas.asp','','status=yes,scrollbars=yes,width=100,height=100');
	self.window.close();
</script>
<body>
</body>
</html>
