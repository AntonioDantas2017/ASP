<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">


<html>
<head>
<title><!--#include file="inc_title.asp"--></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!--<meta http-equiv="Refresh" content="900">-->
</head>

<script language="JavaScript" type="text/JavaScript">
<!--
	



	//rotina para fechar a tela que abre esta
	var wName = window.name;
	if (wName != "teste")
	{
		window.name = "teste";
	}
	
    try
    {
	    if (opener) 
		{
			opener.opener = opener; 
			opener.close(); 
		}
	}
	catch(err)
	{
		//Do nothing
	}
	
	
	//rotina que pega a resolução do usuário e redimenciona a tela
	
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
		if (top.frames.length!=0)
			top.location=self.document.location;
		self.moveTo(0,0);
		self.resizeTo(screen.availWidth,screen.availHeight);
	//	window.resizeTo(800,600);
	//	self.moveTo(0,0);
	} 
	if (screen.width == "1024")
	{
		if (top.frames.length!=0)
			top.location=self.document.location;
		self.moveTo(0,0);
		self.resizeTo(screen.availWidth,screen.availHeight);
	} 
	else
	{
	//	window.resizeTo(1024,740);
	//	self.moveTo(0,0);
	}
	
//-->
</script>

<frameset rows="90,*" frameborder="NO" border="0" framespacing="0">
  <frame src="default-topo.asp?primeira=sim" name="topFrame" frameborder="no" scrolling="NO" noresize >
  <frame src="default-centro.asp" name="mainFrame" frameborder="no" scrolling="yes">
</frameset>

<noframes>

<body>

</body>
</noframes>

</html>
