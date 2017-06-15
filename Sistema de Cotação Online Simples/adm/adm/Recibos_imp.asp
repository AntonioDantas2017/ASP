<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_contratos
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if

nivel_aux = nivel
'nivel_aux = 2
%>

<%
function unidades(num)
	dim vet_unidades, str, aux
	vet_unidades = Array("um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove")
	aux = Right(num,2)
	if int(aux) < 20 then
	if int(aux) > 0 then
	str = vet_unidades(int(aux)-1)
	else
	str = ""
	end if
	else
	str = vet_unidades(int(right(num,1)) - 1)
	end if
	unidades = str
	end function
	function dezenas(num)
	dim vet_dezenas, aux, str
	vet_dezenas = Array(VBNullString, "vinte", "trinta", "quarenta", "cinqüenta", "sessenta", "setenta", "oitenta", "noventa")
	aux = Right(num,2)
	aux = Left(aux,1)
	if len(num) >= 2 then
	if int(aux) >= 2 then
	str = vet_dezenas(int(aux)-1)
	if right(num,1) > 0 then
	str = str &" e "&unidades(num)
	end if
	else
	str = unidades(num)
	end if
	else
	str = unidades(num)
	end if
	dezenas = str
	end function
	function centenas(num, numero)
	dim vet_centenas, aux, str
	vet_centenas = Array("cem", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")
	if int(num) > 99 then
	aux = Right(num, 3)
	aux = Left(aux, 1)
	if int(right(num,2)) > 0 then
	vet_centenas(0) = "cento"
	end if
	str = vet_centenas(aux-1)
	if int(right(num, 2)) > 0 then
	str = str & " e "
	end if
	else
	str = ""
	end if
	centenas = str & dezenas(num)
	end function
	function milhares(num, numero)
	dim str, aux, auxstr
	aux = right(numero,3)
	aux = left(aux,1)
	if int(aux) > 0 then
	auxstr = ", "
	else
	auxstr = " e "
	end if
	if int(num) <> 0 then
	str = centenas(num, numero)&" mil"&auxstr
	else
	str = ""
	end if
	milhares = str
	end function
	function milhoes(num, numero)
	dim str, aux, strmilhao
	aux = int(num)
	if aux > 0 then
	if aux = 1 then
	strmilhao = "milhão,"
	else
	strmilhao = "milhões,"
	end if
	str = centenas(num, numero)&" "&strmilhao&" "
	else
	str = ""
	end if
	milhoes = str
	end function
	function bilhoes(num, numero)
	dim str, aux, strbilhao
	aux = int(num)
	if aux > 0 then
	if aux = 1 then
	strbilhao = "bilhão,"
	else
	strbilhao = "bilhões,"
	end if
	str = centenas(num, numero)&" "&strbilhao&" "
	else
	str = ""
	end if
	bilhoes = str
	end function
	function centavos(num)
	dim str, aux, strcent
	num = "0"&num
	if int(num) > 0 then
	if int(num) = 1 then
	strcent = " centavo"
	else
	strcent = " centavos"
	end if
	str = centenas(num, num)&strcent
	else
	str = ""
	end if
	centavos = str
	end function
	function extenso(num)
	dim sizenum, strsinal, inteiro, cents, aux_array, aux, bilhar, milhao, milhar, centena, dezena
	dim strreal, strvirgula
	dim strextenso
	
	num = Replace(num, ".", "")
	aux_array = split(num, ",")
	if num = "" then
	extenso = ""
	exit function
	end if
	if UBound(aux_array) > 0 then
	inteiro = aux_array(0)
	cents = left(aux_array(1),2)
	else
	inteiro = aux_array(0)
	cents = "00"
	end if
	if InStr(1, inteiro, "-") > 0 then
	inteiro = right(inteiro, len(inteiro) - 1)
	strsinal = "menos "
	else
	strsinal = ""
	end if
	sizenum = len(inteiro)
	aux = string(12 - sizenum, "0")
	inteiro = aux & inteiro
	bilhar = mid(inteiro, 1, 3)
	milhao = mid(inteiro, 4, 3)
	milhar = mid(inteiro, 7, 3)
	centena = mid(inteiro, 10, 3)
	if int(inteiro) = 1 then
	strreal = " real "
	else
	if int(inteiro) = 0 then
	strreal = ""
	else
	strreal = " reais "
	end if
	end if
	if int(cents) > 0 and int(inteiro) > 0 then
	strvirgula = "e "
	else
	strvirgula = ""
	end if
	strextenso = bilhoes(bilhar, inteiro)
	strextenso = strextenso & milhoes(milhao, inteiro)
	strextenso = strextenso & milhares(milhar, inteiro)
	strextenso = strextenso & centenas(centena, inteiro)
	strextenso = strsinal & strextenso & strreal & strvirgula & centavos(cents)
	extenso = strextenso
end function
%>


<%
codrec = request("codrec")
set objrs01 = objconn01.execute("select * from recibos where codrec="&codrec&"")
if not objrs01.eof then
	valor		=	objrs01("valor")
	data		=	objrs01("data")
	nome_rec	=	objrs01("nome_rec")
	cpf_rec		=	objrs01("cpf_rec")
	rg_rec		=	objrs01("rg_rec")	
	nome_ent	=	objrs01("nome_ent")
	cpf_ent		=	objrs01("cpf_ent")
	rg_ent		=	objrs01("rg_ent")			
	referente	=	objrs01("referente")			
end if
%><%
call abre_conexao02

'nivel = nivel_bd
'if ((nivel="1")or(nivel="2")or(nivel="3")) then

'else
 '   Response.Redirect ("default_acessonegado.asp")
'end if
nivel_aux = nivel
%>
<!--#include file="inc_css.asp"-->


<html><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-family: "Times New Roman", Times, serif;
	font-size: 36px;
	font-weight: bold;
}
.style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 16px;
}
-->
</style><head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


</head>


<body onLoad="javascript:print();">

<div align="center"> 
<%
'DECLARA TODAS AS VARIÁVEIS
'--------------------------
fornecedor = request("fornecedor")
codcot = request("codcot")

call abre_conexao01
call abre_conexao02
call abre_conexao03


%>
  <table width="600" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
    <tr> 
      <td><div align="center" class="style1">RECIBO N&deg; <%=codrec%> <br>
        <br>
      </div></td>
    </tr>
  </table>
  <br>
  <div align="center"></div>
  <div align="center">
<table width="600" border="0" cellspacing="0" cellpadding="0" CLASS="linkA">
      <form action="COTACAO_imp.asp" method="post" name="form1">
        <input name="tipo2" type="hidden" value="<%=tipo2%>">
		
		
		<tr> 
          <td valign="top"><div align="left" class="style2">Recebi do Sr. <%=ucase(nome_ent)%> CPF <%=cpf_ent%> RG <%=rg_ent%> a quantia de R$ 
		  <%
		  if isnumeric(valor) then
		  	response.Write formatnumber(valor,2)
		  end if
		  %>
		  (<%=Extenso(valor)%>)
		  referente a <%=referente%>.<br>
		  Sem mais nada a declarar.<br>
		  <br>
		  <%
		  if isdate(data) then
		  	data2 = split(data,"/")
		  	dia		= right("0"&data2(0),2)
		  	mes		= right("0"&data2(1),2)			
		  	ano		= data2(2)			
		  end if
		  %>
		  S&atilde;o Jos&eacute; dos Campos, <%=dia%> de <%=MonthName(mes, False)%> de <%=ano%><br>
		  <br>
		  <br>
		  <br>
		  __________________________________________________<br>
		  <%=ucase(nome_rec)%> CPF <%=cpf_rec%> RG <%=rg_rec%>
		  <br>
            <br>
          </div></td>
        </tr>
      </form>
    </table>
<%call fecha_conexao01%>    
  </div>
  <div align="center"></div>
</div>
</body>
</html>