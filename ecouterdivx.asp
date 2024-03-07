<HTML>
<HEAD>
<TITLE>Base de données MP3</TITLE>
<script type="text/javascript">
<!--
	function redimensionne()
	{
		window.resizeTo(window.screen.availWidth/1.5, window.screen.availHeight/1.5);
	}
	function fermer()
	{
		window.close(); 
	}
//-->
</script>
</HEAD>
<BODY onLoad="redimensionne()" bgcolor="#FFFFFF"> 
<%
URL = Request.QueryString("url")
%> 
<table width="85%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#D3D1FA"> 
  <tr> 
    <td bgcolor="#F8EEE4"><h3 align="center"><%=URL%></h3></td> 
  </tr> 
  <tr> 
    <td align="center"><EMBED width="380" height="400" src="F:\MP3 Techno\<%=URL%>" autostart=true loop=infinite volume=100% show=true></td> 
  </tr> 
  <tr> 
    <td bgcolor="#F8EEE4" align="center"><input name="Fermer" type="button" onClick="fermer()" value="Fermer"></td> 
  </tr> 
</table> 
</BODY>
</HTML>
