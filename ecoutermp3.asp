<%
	'On récupère la variable url de l'url
	URL = Request.QueryString("url")
%>
<HTML>
<HEAD>
<TITLE>MP3&nbsp;&ETH;ance&nbsp;&dagger;echno&nbsp;H&loz;use...&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;d(&not;_&not;)b&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TITLE>
</HEAD>
<BODY bgcolor="#FFCCFF"> 
<table style="border:#9094c9 1pt solid;" width="85%" border="0" align="center" cellpadding="0" cellspacing="0"> 
  <tr> 
    <td align="center" height="30" bgcolor="#B7C9F7"><%=URL%></td> 
  </tr> 
  <tr> 
    <td align="center" bgcolor="#B7C9F7"> 

<object classid="clsid:6BF52A52-394A-11D3-B153-00C04F79FAA6" id="WindowsMediaPlayer1" height="42" type="audio/mpeg" height="45" width="380">
	<param name="URL" ref value="D:\\mp3\<%=URL%>">
      <param name="autostart" value="true"> 
      <param name="loop" value="false"> 
</object>

    </td> 
  </tr> 
  <tr> 
    <td bgcolor="#B7C9F7" align="center"><input style="background-color:#FECFE4; height:18; width:50; border-width:0;" name="Fermer" type="button" onClick="window.close()" value="Fermer"></td> 
  </tr> 
</table> 
</BODY>
</HTML>
