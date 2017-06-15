<!--#include file="inc_login.asp"-->
<!--#include file="inc_tempo.asp"-->
<%
call abre_conexao01
nivel = nivel_contratos
if ((nivel="1")or(nivel="2")) then

else
    Response.Redirect ("default_acessonegado.asp")
end if
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
	vet_dezenas = Array(VBNullString, "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa")
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
nivel_aux = nivel
'nivel_aux = 2
codcon = request("codcon")
set objrs01 = objconn01.execute("select * from contratos where codcon="&codcon&"")
if not objrs01.eof then
	auxcodloc 	= objrs01 ("codloc")
	auxcodimo 	= objrs01 ("codimo")
	periodo 	= objrs01 ("periodo")
	dataini 	= objrs01 ("dataini")
	datafim 	= objrs01 ("datafim")
	aluguel 	= objrs01 ("aluguel")
	vencimento 	= objrs01 ("vencimento")
	multa 		= objrs01 ("multa")
	destinado 	= objrs01 ("destinado")
	reajuste 	= objrs01 ("reajuste")
	impressao 	= objrs01 ("impressao")
	testemunha1 = objrs01 ("testemunha1")
	testemunha2 = objrs01 ("testemunha2")
	set objrs01 = objconn01.execute("select * from locatarios where codloc="&auxcodloc&"")
	if not objrs01.eof then
		nome 			= objrs01 ("nome")
		cpf				= objrs01 ("cpf")
		rg				= objrs01 ("rg")
		qualificacao	= objrs01 ("qualificacao")
	end if
	set objrs01 = objconn01.execute("select * from imoveis where codimo="&auxcodimo&"")
	if not objrs01.eof then
		rua 		= objrs01 ("rua")
		numero		= objrs01 ("numero")
		bairro		= objrs01 ("bairro")
		cidade		= objrs01 ("cidade")
	end if

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
<!--#include file="inc_css2.asp"-->


<html><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-family: verdana;
	font-weight: bold;
	font-size: 36px;
}
.style2 {color: #FF0000}
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


if fornecedor = "" then
%>
  <table width="600" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
    <tr> 
      <td><div align="center" class="style1">CONTRATO<br>
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
          <td valign="top"><div align="justify"><strong>
            
              </strong>Os signat&aacute;rios deste  instrumentos, de um lado <span class="style2">ANTONIO JOSIVALDO DANTAS</span> qualifica&ccedil;&atilde;o <span class="style2">comerciante</span> CPF <span class="style2">480939124-87</span> e RG <span class="style2">959462</span>, de outro lado,<span class="style2"> <%=UCASE(nome)%></span> qualifica&ccedil;&atilde;o<span class="style2"> <%=qualificacao%></span> CPF <span class="style2"><%=cpf%> </span>e RG<span class="style2"> <%=rg%></span> t&ecirc;m justo e contratado o seguinte, que mutuamente  aceitam e outorgam, a saber:<br>
              O primeiro nomeado, aqui  chamado &ldquo;locador&rdquo;, sendo propriet&aacute;rio do im&oacute;vel sito  nesta<span class="style2"> 
              <%
			  response.Write "Rua: "&rua& " n°: "&numero&" Cidade: "&cidade%>
              </span>              loca-o ao segundo, aqui designado  &ldquo;locat&aacute;rio&rdquo;, mediante as cl&aacute;usulas e condi&ccedil;&otilde;es adiante estipuladas, ou seja:<br>
              <br>
              1&deg;)O prazo de loca&ccedil;&atilde;o &eacute; de <span class="style2"><%=periodo%> </span>meses a partir de <span class="style2"><%=dataini%>  </span>e a terminar em <span class="style2"><%=datafim%> </span>data em que o locat&aacute;rio se  obriga a restituir o im&oacute;vel completamente desocupado, no estado em que o  recebeu, independentemente de Notifica&ccedil;&atilde;o ou Interpela&ccedil;&atilde;o Judicial, ressalvada  a hip&oacute;tese de prorroga&ccedil;&atilde;o da loca&ccedil;&atilde;o, o que somente se far&aacute; por escrito.<br>
&nbsp; &sect;&nbsp;  &uacute;nico: Caso o locat&aacute;rio n&atilde;o restitua o im&oacute;vel no fim do prazo  contratual, pagar&aacute; enquanto estiver na posse do mesmo o aluguel mensal  reajustado nos termos da Cl&aacute;usulas D&eacute;cima Oitava, at&eacute; a efetiva desocupa&ccedil;&atilde;o do  im&oacute;vel objeto deste instrumento;<br>
<br>
2&deg;)O aluguel mensal &eacute; de R$<span class="style2"> <%=formatnumber(aluguel)%></span> (<span class="style2"><%=Extenso(aluguel)%></span>) que o locat&aacute;rio se compromete a pagar pontualmente,  at&eacute; o dia<span class="style2"> <%=vencimento%></span>, na resist&ecirc;ncia do locador ou de  seu representante;<br>
<br>
3&deg;)O locat&aacute;rio, salvo as  obras que importem na seguran&ccedil;a do im&oacute;vel, obriga-se por todas as outras,  devendo trazer o im&oacute;vel locado em boas condi&ccedil;&otilde;es de higiene e limpeza, com os  aparelhos sanit&aacute;rios e de ilumina&ccedil;&atilde;o, fog&atilde;o, pap&eacute;is, pinturas, telhados,  vidra&ccedil;as, m&aacute;rmores, fechos, torneiras, pias, banheiros, ralos e demais  acess&oacute;rios em perfeito estado de conserva&ccedil;&atilde;o e funcionamento, para assim,  restitu&iacute;-los quando findo ou rescindido este contrato sem direito a reten&ccedil;&atilde;o ou  indeniza&ccedil;&atilde;o por quaisquer benfeitoria, ainda que necess&aacute;rias, as quais ficar&atilde;o  desde logo incorporados ao im&oacute;vel;<br>
<br>
4&deg;)Obriga-se mais o  locat&aacute;rio a satisfazer a todas as exig&ecirc;ncias dos Poderes P&uacute;blicos, a que der  causa, e a n&atilde;o transferir este contrato, nem fazer modifica&ccedil;&atilde;o ou  transforma&ccedil;&otilde;es no im&oacute;vel sem autoriza&ccedil;&atilde;o escrita do locador;<br>
<br>
5&deg;)O locat&aacute;rio desde j&aacute;  faculta ao locador examinar ou vistoriar o im&oacute;vel locado quando entender  conveniente;<br>
<br>
6&deg;)O locat&aacute;rio tamb&eacute;m n&atilde;o  poder&aacute; sublocar nem emprestar o im&oacute;vel no todo ou em parte, sem preceder  consentimento por escrito do locador; devendo no caso desde ser dado, agir  oportunamente junto aos ocupantes, a fim de que o im&oacute;vel esteja desimpedido no  termino do presente contrato;<br>
<br>
7&deg;)Nenhuma intima&ccedil;&atilde;o do  Servi&ccedil;o Sanit&aacute;rio ser&aacute; motivo para o locat&aacute;rio abandonar o im&oacute;vel ou pedir a  rescis&atilde;o deste contrato, salvo procedendo vistoria judicial, que apure estar a  constru&ccedil;&atilde;o amea&ccedil;ando ru&iacute;na;<br>
<br>
8&deg;)Para todas as quest&otilde;es  resultantes deste contrato, ser&aacute; competente o foro da situa&ccedil;&atilde;o do im&oacute;vel, seja  qual for o domicilio dos contratantes;<br>
<br>
9&deg;)Tudo quando for devido em  raz&atilde;o deste contrato e que n&atilde;o comporte o processo executivo, ser&aacute; cobrado em  a&ccedil;&atilde;o competente, ficando a cargo do devedor, em qualquer caso, os honor&aacute;rios do  advogado que o credo constituir para ressalva dos seus direitos;<br>
<br>
10&deg;)Fica estipulada multa de um aluguel na qual incorrer&aacute; a parte  que infligir qualquer clausula deste contrato; com a faculdade, para a  parte inocente, de poder considerar simultaneamente rescindida a loca&ccedil;&atilde;o,  independentemente de qualquer formalidade;<br>
<br>
11&deg;)Assina  tamb&eacute;m o  presente, solidariamente com o locat&aacute;rio por todas as obriga&ccedil;&otilde;es acima  exaradas, a Sr.(a)  CPF X RG X c&ocirc;njuge X cuja responsabilidade,  entretanto, perdurar&aacute; at&eacute; a entrega real e efetiva real e efetiva das chaves do  im&oacute;vel locado;<br>
<br>
12&deg;)Quaisquer estragos  ocasionados ao im&oacute;vel e suas instala&ccedil;&otilde;es, bem como despesas a que o  propriet&aacute;rio for obrigado por eventuais modifica&ccedil;&otilde;es feitas no im&oacute;vel, pelo  locat&aacute;rio, n&atilde;o ficam compreendidas no multa da clausula 12&deg;, mas ser&atilde;o pagas &aacute;  parte;<br>
<br>
13&deg;)Em caso de falecimento  de qualquer parte contratante, os herdeiro&nbsp;  da parte falecida ser&atilde;o obrigados ao cumprimento integral deste  contrato, at&eacute; a sua termina&ccedil;&atilde;o;<br>
<br>
14&deg;)Estabelecem as partes  contratantes que, para reformar ou renova&ccedil;&atilde;o deste contrato, as parte  interessadas se notificar&atilde;o mutuamente, com anteced&ecirc;ncia nunca inferior a cento  e vinte dias, findo este prazo, considera-se como desinteressante para o  locat&aacute;rio, a sua continua&ccedil;&atilde;o no im&oacute;vel ora locado, devendo o mesmo entregar as  suas chaves ao locador; impreterivelmente no dia do vencimento deste&nbsp; contrato;<br>
<br>
15&deg;)O im&oacute;vel, objeto de loca&ccedil;&atilde;o,  destina-se exclusivamente a<span class="style2"> <%=destinado%></span> n&atilde;o  podendo ser mudada a sua destina&ccedil;&atilde;o sem consentimento expresso do locador;<br>
<br>
16&deg;)O valor locat&iacute;cio  expresso na clausula 2&deg;, ser&aacute; reajustado em<span class="style2"> <%=reajuste%></span> de  conformidade com a varia&ccedil;&atilde;o da infra&ccedil;&atilde;o, ou  outro &iacute;ndice que viver oficialmente &aacute; substitu&iacute;-lo, inclusive na hip&oacute;tese de  ocorr&ecirc;ncia de prorroga&ccedil;&atilde;o do presente contrato;<br>
<br>
17&deg;)Todos os impostos e taxas  que atualmente recaem sobre o im&oacute;vel locado, bem como qualquer aumento dos  mesmo, ou novos que venham a ser criados pelo Poder Publico, ser&atilde;o da inteira  responsabilidade do locat&aacute;rio as contas de &aacute;gua, luz, for&ccedil;a e g&aacute;s, assim como  as despesas de condom&iacute;nio, se houver.<br>
&sect; &uacute;nico:O locat&aacute;rio ser&aacute;  respons&aacute;vel pelas despesas e multas decorrentes de eventuais reten&ccedil;&otilde;es dos avisos  de impostos, taxas e outros que j&aacute; incidem ou venham a incidir sobre o im&oacute;vel objeto da presente loca&ccedil;&atilde;o;<br>
<br>
18&deg;)A falta de pagamento, nas  &eacute;pocas supra determinada, dos alugueis e encargos, por si constituir&aacute; o  locat&aacute;rio em mora, independentemente de qualquer Notifica&ccedil;&atilde;o, Interpela&ccedil;&atilde;o ou  aviso extra-judicial;<br>
<br>
19&deg;)Se o locador admitir, em  beneficio do locat&aacute;rio, qualquer&nbsp; atraso  no pagamento do aluguel e demais despesas que lhe incumba, ou no cumprimento de  qualquer outra obriga&ccedil;&atilde;o contratual, essa toler&acirc;ncia n&atilde;o poder&aacute; ser considerada  como altera&ccedil;&atilde;o das condi&ccedil;&otilde;es deste contrato, nem dar&aacute; ensejo &agrave; invoca&ccedil;&atilde;o do  Artigo 1,503-Inciso I do c&oacute;digo Civil  Brasileiro, por parte do fiador, pois se constituir&aacute; em alto de mera  liberalidade do locador;<br>
E por assim terem contratado,  assinaram o presente, em 2 vias, em presen&ccedil;a das testemunhas abaixo:<br>
<br>
<%
data2 = split(impressao,"/")
dia = right("0"&data2(0),2)
mes = right("0"&data2(1),2)
ano = data2(2)
%>
S&atilde;o Jos&eacute; dos Campos,<span class="style2"> <%=dia%> de <%=MonthName(mes, False)%> de <%=ano%> &nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Contrato n&deg;<span class="style2"> <%=codcon%></span></div>
            <div>
              <div align="justify"><br>
                <br>
                <br>
                <strong>Locador: <br>
                </strong><br>
                _____________________________________________________<br>
                ANTONIO JOSIVALDO DANTAS<br>
                <br>
              <strong>Locat&aacute;rio:</strong><br>
              <br>
              _____________________________________________________<br>
              <%=UCASE(nome)%><br>
<%IF testemunha1 <> "" or testemunha2 <> "" then%>
              <br>
              <br>
<strong>Testemunhas:<br></strong>
              <br>
_____________________________________________________<br>
<%response.Write ucase(testemunha1)%>
<br>
<br>


<%IF testemunha2 <> "" then%>
              <br>
             _____________________________________________________<br>
<%response.Write ucase(testemunha2)%>
<br>
<br>

<%end if%>
<%end if%>
<%end if%>
</div>
            </div>
            <p align="justify">&nbsp;</p>
            <div align="justify"><br>
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