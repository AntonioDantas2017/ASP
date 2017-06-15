
<%
Response.clear 'Limpa tudo o que estava armazenado no buffer 
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl = "no-cache"
Response.CacheControl = "Private"
Response.ExpiresAbsolute = #January 1, 1990 00:00:01#
Response.Expires= 0
Response.Buffer = False

%>

<!--#include file="adm/inc_css.asp"-->



<html>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<head>
<title><!--#include file="adm/inc_title.asp"--></title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script language=javascript>
function proibido_voltar()
{ 
	history.go(-1) == false;
}

if (screen.width == "720")
{
	location.href = 'default3.asp'
} 
if (screen.width == "640")
{
	location.href = 'default3.asp'
} 

if (screen.width == "800")
{
//	location.href = 'default3.asp'
} 
if (screen.width == "1024")
{
	if (top.frames.length!=0)
		top.location=self.document.location;
	self.moveTo(0,0);
	self.resizeTo(screen.availWidth,screen.availHeight);
} 
</script>

<script language="JavaScript" type="text/JavaScript">
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<div align="left">
  <script language="JavaScript">
function keypaddown (keyvalue)
{
	if (keyvalue == "") {
		document.FLogin.senha.value = ""
		return;
	}
	var pass = document.FLogin.senha.value;
	if (pass.length >= 6)
	{
		return;
	}
	document.FLogin.senha.value = document.FLogin.senha.value + keyvalue;
}

function teclaclick(tecla)
{
	return false
}

function teclaup(tecla)
{
	keypaddown(tecla);
}
</script>
  <table width="870" height="100%" border="0" cellpadding="0" cellspacing="0">
    <form name="FLogin" action="default2.asp" method="post" >
      <tr> 
        <td valign="top" background="IMAGEM/home01.jpg"><div align="center"> 
            <div align="center"> 
              <div align="center">
                <table width="100%" height="190" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td valign="bottom"><div align="center">
                        <div align="center"></div>
              </div></td>
                  </tr>
                </table>
	 <%
acesso = request("acesso")
cookie_login = request.cookies("piemme")("login")

if acesso <> "" then
 %>
                <br>
                <div align="center"> <b><font color="#FF3333" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
                  </font></b> 
                  <table width="90%" border="0" cellspacing="0" cellpadding="0" class=linka>
                    <tr> 
                      <td><div align="center">
                          <table width="500" border="0" cellspacing="0" cellpadding="0" class=linka>
                            <tr valign="top"> 
                              <td width="150"><br> </td>
                              <td> <div align="center">
                                  <%
if acesso="vazio" then
%>
                                  <br>
                                  <b><font color="#CC3333">FAVOR PREENCHER TODOS 
                                  OS CAMPOS.</font><font color="#888888"><br>
                                  <br>
                                  </font><br>
                                  Lembrando que 3 tentativas<br>
                                  inv&aacute;lidas bloqueiam seu acesso.</b><font color="#666666"><br>
                                  </font><br>
                                  <br>
                                  <br>
                                  <strong><a href="DEFAULT1.asp" class=linkC>Tentar 
                                  Novamente</a></strong> 
                                  <%
end if
%>
                                  <%
if acesso="negado" then
%>
                                  <b> <br>
                                  <font color="#CC3333">ACESSO NEGADO</font></b> 
                                  <font color="#CC3333"><b><br>
                                  </b></font><font color="#888888"><b><br>
                                  </b></font><b>Lembrando que 3 tentativas<br>
                                  inv&aacute;lidas bloqueiam seu acesso.<br>
                                  </b><br>
                                  <br>
                                  <br>
                                  <strong><a href="default.asp" class=linkC>Tentar 
                                  Novamente</a></strong> 
                                  <%
end if
%>
                                  <%
if acesso="bloqueado" then
%>
                                  <b> <br>
                                  <font color="#CC3333">ACESSO BLOQUEADO<br>
                                  </font><font color="#888888"><br>
                                  <br>
                                  </font>Excedido N&ordm; de Tentativas</b> <br>
                                  <strong><br>
                                  Entre em contato com o administrador <br>
                                  do program para desbloquear seu acesso.</strong> 
                                  <br>
                                  <br>
                                  <br>
                                  <strong><a href="default.asp" class=linkC>Voltar</a></strong> 
                                  <%
end if
%>
                                  <%
if acesso="5" then
%>
                                  <b> <br>
                                  <font color="#CC3333">ACESSO BLOQUEADO</font><font color="#FFFFFF"><br>
                                  </font><br>
                                  Este computador n&atilde;o est&aacute; <br>
                                  cadastrado para acessar o program.</b><br>
                                  <strong><br>
                                  <br>
                                  Entre em contato com o <br>
                                  administrador do program para <br>
                                  cadastrar este computador.</strong><font color="#888888"> 
                                  </font> 
                                  <%
end if
%>
                                </div></td>
                            </tr>
                          </table>
                          
                        </div></td>
                    </tr>
                  </table>
                  <%
else
%>
                </div>
              <table width="100%" height="190" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td><div align="center"> 
                        <div align="center"> <br>
                          <table width="600" border="0" cellspacing="0" cellpadding="0" class=linka>
                            <tr valign="top"> 
                              <td width="279"><br>
                              </td>
                              <td width="71"> <div align="left"> <img src="imagem/TECLADO2.jpg" width="65" height="90" border="0" name="Teclado" usemap="#tecladoMapMap"> 
                                  <map name="tecladoMapMap">
                                    <area shape="rect" coords="0,1,21,22" href="#" onClick="return teclaclick('7')" onMouseUp="teclaup('7')">
                                    <area shape="rect" coords="21,1,46,22" href="#" onClick="return teclaclick('8')" onMouseUp="teclaup('8')">
                                    <area shape="rect" coords="45,2,67,22" href="#" onClick="return teclaclick('9')" onMouseUp="teclaup('9')">
                                    <area shape="rect" coords="0,23,21,45" href="#" onClick="return teclaclick('4')" onMouseUp="teclaup('4')">
                                    <area shape="rect" coords="22,24,43,44" href="#" onClick="return teclaclick('5')" onMouseUp="teclaup('5')">
                                    <area shape="rect" coords="43,23,70,45" href="#" onClick="return teclaclick('6')" onMouseUp="teclaup('6')">
                                    <area shape="rect" coords="1,46,21,67" href="#" onClick="return teclaclick('1')" onMouseUp="teclaup('1')">
                                    <area shape="rect" coords="22,46,43,66" href="#" onClick="return teclaclick('2')" onMouseUp="teclaup('2')">
                                    <area shape="rect" coords="44,46,70,66" href="#" onClick="return teclaclick('3')" onMouseUp="teclaup('3')">
                                    <area shape="rect" coords="0,67,21,96" href="#" onClick="return teclaclick('0')" onMouseUp="teclaup('0')">
                                    <area shape="rect" coords="21,69,75,96" href="#" onClick="return teclaclick('')" onMouseUp="teclaup('')">
                                  </map>
                              </div></td>
                              <td width="250"><strong><b> <font color="#666666">LOGIN 
                                :</font></b><br>
                                <font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"> 
                                <input name="login" type="text" id="login"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #F2F2F2" tabindex="1" value="<%=cookie_login%>" size="10" maxlength="10" >
                                </font> </strong> 
                                <table width="30%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td height="7"><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
                                  </tr>
                                </table>
                                <strong> <font color="#666666">SENHA</font></strong><font color="#666666">: 
                                </font><br> 
<input  name="senha" type="PASSWORD" id="senha" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #F2F2F2"  tabindex="2" onBlur="if(!bchange){window.document.FLogin.IDPRESP.focus()}" size="8" maxlength="6" onKeyPress="return false;">
                                <table width="30%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
                                  </tr>
                                </table>
                                <table width="250" border="0" cellspacing="0" cellpadding="0" CLASS=LINKA>
                                  <tr> 
                                    <td><strong><font color="#666666"><!--RESPOSTA 
                                      SECRETA : --></font></strong></td>
                                  </tr>
                                  <tr> 
                                    <td><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"> 
                                      <input name="resposta" type="HIDDEN"  value="1" id="resposta"  style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #F2F2F2" tabindex="3" size="40" maxlength="30" >
                                      <script Language="JavaScript">
	//	window.document.write('<input type="text" name="IDPCONT" maxlength="2" size="2" value="0" style="FONT-FAMILY: Verdana; FONT-SIZE: 10px; BACKGROUND-COLOR: #F2F2F2" tabindex="3" readonly>');
</script>
                                      <script language="JavaScript">
function ContaFrase()
{
	window.document.FLogin.IDPCONT.value = window.document.FLogin.resposta.value.length;
	return true;
}
window.document.FLogin.resposta.onkeyup = ContaFrase;
</script>
                                      <br>
                                      </font></font></font></td>
                                  </tr>
                                </table>
                                <table width="30%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td><div align="center"><img src="imagem/img_transp.gif" width="1" height="5"></div></td>
                                  </tr>
                                </table>
                                <font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"> 
                                </font></font></font> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td width="80%"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"> 
                                      <input name="imageField" tabindex="4" type="image" src="IMAGEM/BOTAO.gif" width="71" height="22" border="0"   onClick="ValidateOrder(this.form)"  >
                                      </font></font></font></td>
                                  </tr>
                              </table></td>
                            </tr>
                          </table>
                          
                        </div>
              </div></td>
                  </tr>
                </table>
                <%
end if
%>
                <table width="100%" height="100" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td valign="top"><div align="center"> <br>
                        <br>
                      </div></td>
                  </tr>
                </table>
                
              </div>
            </div>
          </div></td>
      </tr>
    </form>
  </table>
</div>
</body>
</html>
