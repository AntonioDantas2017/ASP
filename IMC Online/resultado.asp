<%@ LANGUAGE="VBScript"%>
<!--#include file="inc_css.asp"-->
<%
peso = replace(request("peso"),".",",")
altura = replace(request("altura"),".",",")
'imc = Cdbl(Cdbl(peso) / (sqr(Cdbl(altura))))
imc = Cdbl(Cdbl(peso) / (Cdbl(altura)*Cdbl(altura)))
nome = request("nome")
sexo = request("sexo")

dim situacao(6)
situacao(0) = "Você está com o peso abaixo do normal. Consulte um especialista para saber se há algum problema com sua saúde que esteja causando este peso abaixo do normal, ou se o peso abaixo do normal pode estar de alguma forma ameaçando sua saúde."
situacao(1) = "Seu peso está dentro da faixa considerada normal pela Organização Mundial de Saúde. Algumas pessoas, no entanto, já podem ter um maior risco de problemas metabólicas mesmo dentro desta faixa, principalmente se acumularem gordura na região interna do abdome. Uma maneira simples de avaliar isto é medir a circunferência da cintura. Mais de 80 cm de cintura em mulheres e 94 cm em homens podem indicar um possível excesso de gordura no interior do abdome. Riscos à parte, muita gente que se encontra nesta faixa de peso considerada normal tenta emagrecer, principalmente por razões estéticas. Para estas pessoas recomenda-se apenas um planejamento alimentar saudável e a prática regular de atividades físicas. Os remédios para emagrecer em princípio não estão indicados."
situacao(2) = "Você está dentro da faixa chamada de pré-obesidade pela Organização Mundial de Saúde. Sabe-se que este peso já pode representar um risco considerável para a saúde. Se você também está com a pressão alta, diabetes ou aumento de colesterol, os remédios para emagrecer serão indicados se você não conseguir emagrecer sem eles. Nunca utilize nenhum deles sem um acompanhamento médico cuidadoso. Mesmo utilizando medicamentos, lembre-se de que é muito importante um planejamento alimentar e a prática de atividades físicas. A eficácia dos medicamentos aumenta muito quando estes fundamentos não são esquecidos."
situacao(3) = "Você está na faixa de peso denominada 'obesidade classe I' pela Organização Mundial de Saúde. Seu peso já está causando um risco aumentado para várias doenças, incluindo o diabetes, a hipertensão arterial, o infarto do miocárdio e diversos tipos de câncer. Sua obesidade, por si só, já é considerada uma doença e necessita ser tratada com remédios. Procure seu médico para uma orientação adequada. Nunca utilize medicamentos sem acompanhamento médico, e lembre-se de que é muito importante um planejamento alimentar e a prática de atividades físicas. A eficácia dos medicamentos aumenta muito quando estes fundamentos não são esquecidos."
situacao(4) = "Você está na faixa de peso denominada 'obesidade classe II' pela Organização Mundial de Saúde. Seu peso já está causando um risco muito aumentado para várias doenças, incluindo o diabetes, a hipertensão arterial, o infarto do miocárdio e diversos tipos de câncer. Sua obesidade, por si só, já é considerada uma doença e necessita ser tratada com remédios. Se você além disso tem também diabetes, hipertensão arterial ou outra complicação importante da obesidade, a cirurgia para emagrecer pode ser a melhor solução para o seu caso. Procure seu médico para uma orientação adequada."
situacao(5) = "Você está na faixa de peso denominada 'obesidade classe III' pela Organização Mundial de Saúde. Esta categoria engloba todos as pessoas com mais de 40 kg/m2 de IMC. Muitas vezes é chamada também de obesidade mórbida. Seu peso já está causando um risco altíssimo para várias doenças, incluindo o diabetes, a hipertensão arterial, o infarto do miocárdio e diversos tipos de câncer. A obesidade neste grau é considerada uma doença grave e necessita ser tratada com todos os recursos disponíveis, incluindo os remédios e a cirurgia para emagrecer. Procure seu médico para uma orientação adequada."

if imc < 18.5 then
	sit = situacao(0)
	foto = "anorexia.jpg"
else
	if imc >= 18.5 and imc < 24.9 then
		sit = situacao(1)
		foto = "normal.jpg"		
	else
		if imc >=25 and imc < 29.9 then
			sit = situacao(2)
			foto = "pre.bmp"
		else
			if imc >= 30 and imc < 34.9 then
				sit = situacao(3)
				foto = "obesidade.jpg"
			else
				if imc >= 35 and imc < 39.9 then
					sit = situacao(4)
					foto = "obesos.jpg"		
				else
					sit = situacao(5)
					foto = "morbida.jpg"
				end if
			end if
		end if
	end if
end if
%>

<html>
<head>
<TITLE>Índice de Massa Corpórea On-Line</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body background="img/bgfull.jpg" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="21%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="98%" height="617" align="center" valign="top" background="img/bgsquare.gif"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td align="center" valign="top"><!--#include file="toplogo.asp"-->
          </td>
        </tr>
        <tr> 
          <td align="center" valign="top"><img src="img/1x6v.gif" width="1" height="6"></td>
        </tr>
        <tr> 
          <td align="left" valign="top">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="1%" rowspan="9" align="right" valign="top"><img src="img/6x1h.gif" width="7" height="1"></td>
                <td align="center" valign="top" bgcolor="#FFFFFF"><img src="img/tpcadastro.jpg" width="752" height="200"></td>
                <td width="1%" rowspan="9" align="left" valign="top"><img src="img/6x1h.gif" width="6" height="1"></td>
              </tr>
              <tr> 
                <td align="left" valign="top" bgcolor="#FFFFFF">
<div align="left"><font color="#727292"><br>
                    </font></div></td>
              </tr>

              <tr> 
                <td align="left" valign="top" bgcolor="#FFFFFF"><font color="#FF0000"><img src="img/blank25x8h.gif" width="25" height="8"></font></td>
              </tr>
              <tr>
                <td align="left" valign="top" bgcolor="#FFFFFF" class="linkB"><div align="center"><br>
                    <br>
                    O seu resultado final &eacute;:<br>
                    <br>
                    <table width="75%" border="0" class="linkA">
                      <tr> 
                        <td width="21%"><div align="right">Nome:</div></td>
                        <td width="2%">&nbsp;</td>
                        <td width="77%"><%=nome%></td>
                      </tr>
                      <tr> 
                        <td><div align="right">Sexo:</div></td>
                        <td>&nbsp;</td>
                        <td><%=sexo%></td>
                      </tr>
                      <tr> 
                        <td><div align="right">IMC:</div></td>
                        <td>&nbsp;</td>
                        <td><% response.Write mid(imc,1,4)%></td>
                      </tr>
                      <tr> 
                        <td><div align="right">Situa&ccedil;&atilde;o:</div></td>
                        <td>&nbsp;</td>
                        <td><%=sit%></td>
                      </tr>
                      <tr>
                        <td colspan="3"><div align="center"><img src="img/<%=foto%>"></div></td>
                      </tr>
                    </table>
                    <br>
                    <br>
                    C&aacute;lculo do IMC:<br>
                    <br>
                    IMC = PESO/(ALTURA)&sup2;<br>
                    <br>
                    Tabelo do IMC:<br>
                    <table width="75%" border="1" class="linkA">
                      <tr> 
                        <td>Categoria</td>
                        <td>IMC</td>
                      </tr>
                      <tr> 
                        <td bgcolor="#FFFF00"><font color="#000000">Abaixo do 
                          peso</font></td>
                        <td bgcolor="#FFFF00"><font color="#000000">Abaixo de 
                          18,5</font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#00FF00">Peso Normal</td>
                        <td bgcolor="#00FF00">18,5 - 24,9</td>
                      </tr>
                      <tr> 
                        <td bgcolor="#FFFF00">Sobrepeso</td>
                        <td bgcolor="#FFFF00">25,0 - 29,9</td>
                      </tr>
                      <tr bgcolor="#FFB08A"> 
                        <td>Obesidade Grau I</td>
                        <td>30,0 - 34,9</td>
                      </tr>
                      <tr> 
                        <td bgcolor="#FF8000">Obesidade Grau II</td>
                        <td bgcolor="#FF8000">35,0 - 39,9</td>
                      </tr>
                      <tr bgcolor="#FF0000"> 
                        <td><font color="#000000">Obesidade Grau III</font></td>
                        <td><font color="#000000">40,0 e acima</font></td>
                      </tr>
                    </table>
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>
                    <br>
                  </div></td>
              </tr>

           </table></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
