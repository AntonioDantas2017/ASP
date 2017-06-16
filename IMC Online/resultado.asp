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
situacao(0) = "Voc� est� com o peso abaixo do normal. Consulte um especialista para saber se h� algum problema com sua sa�de que esteja causando este peso abaixo do normal, ou se o peso abaixo do normal pode estar de alguma forma amea�ando sua sa�de."
situacao(1) = "Seu peso est� dentro da faixa considerada normal pela Organiza��o Mundial de Sa�de. Algumas pessoas, no entanto, j� podem ter um maior risco de problemas metab�licas mesmo dentro desta faixa, principalmente se acumularem gordura na regi�o interna do abdome. Uma maneira simples de avaliar isto � medir a circunfer�ncia da cintura. Mais de 80 cm de cintura em mulheres e 94 cm em homens podem indicar um poss�vel excesso de gordura no interior do abdome. Riscos � parte, muita gente que se encontra nesta faixa de peso considerada normal tenta emagrecer, principalmente por raz�es est�ticas. Para estas pessoas recomenda-se apenas um planejamento alimentar saud�vel e a pr�tica regular de atividades f�sicas. Os rem�dios para emagrecer em princ�pio n�o est�o indicados."
situacao(2) = "Voc� est� dentro da faixa chamada de pr�-obesidade pela Organiza��o Mundial de Sa�de. Sabe-se que este peso j� pode representar um risco consider�vel para a sa�de. Se voc� tamb�m est� com a press�o alta, diabetes ou aumento de colesterol, os rem�dios para emagrecer ser�o indicados se voc� n�o conseguir emagrecer sem eles. Nunca utilize nenhum deles sem um acompanhamento m�dico cuidadoso. Mesmo utilizando medicamentos, lembre-se de que � muito importante um planejamento alimentar e a pr�tica de atividades f�sicas. A efic�cia dos medicamentos aumenta muito quando estes fundamentos n�o s�o esquecidos."
situacao(3) = "Voc� est� na faixa de peso denominada 'obesidade classe I' pela Organiza��o Mundial de Sa�de. Seu peso j� est� causando um risco aumentado para v�rias doen�as, incluindo o diabetes, a hipertens�o arterial, o infarto do mioc�rdio e diversos tipos de c�ncer. Sua obesidade, por si s�, j� � considerada uma doen�a e necessita ser tratada com rem�dios. Procure seu m�dico para uma orienta��o adequada. Nunca utilize medicamentos sem acompanhamento m�dico, e lembre-se de que � muito importante um planejamento alimentar e a pr�tica de atividades f�sicas. A efic�cia dos medicamentos aumenta muito quando estes fundamentos n�o s�o esquecidos."
situacao(4) = "Voc� est� na faixa de peso denominada 'obesidade classe II' pela Organiza��o Mundial de Sa�de. Seu peso j� est� causando um risco muito aumentado para v�rias doen�as, incluindo o diabetes, a hipertens�o arterial, o infarto do mioc�rdio e diversos tipos de c�ncer. Sua obesidade, por si s�, j� � considerada uma doen�a e necessita ser tratada com rem�dios. Se voc� al�m disso tem tamb�m diabetes, hipertens�o arterial ou outra complica��o importante da obesidade, a cirurgia para emagrecer pode ser a melhor solu��o para o seu caso. Procure seu m�dico para uma orienta��o adequada."
situacao(5) = "Voc� est� na faixa de peso denominada 'obesidade classe III' pela Organiza��o Mundial de Sa�de. Esta categoria engloba todos as pessoas com mais de 40 kg/m2 de IMC. Muitas vezes � chamada tamb�m de obesidade m�rbida. Seu peso j� est� causando um risco alt�ssimo para v�rias doen�as, incluindo o diabetes, a hipertens�o arterial, o infarto do mioc�rdio e diversos tipos de c�ncer. A obesidade neste grau � considerada uma doen�a grave e necessita ser tratada com todos os recursos dispon�veis, incluindo os rem�dios e a cirurgia para emagrecer. Procure seu m�dico para uma orienta��o adequada."

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
<TITLE>�ndice de Massa Corp�rea On-Line</TITLE>
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
