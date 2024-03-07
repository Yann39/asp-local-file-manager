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
<object classid="clsid:6BF52A52-394A-11D3-B153-00C04F79FAA6" id="WindowsMediaPlayer1">
	<param name="URL" ref value="D:\\MP3 Tek\<%=URL%>">
	<param name="rate" value="1">
	<param name="balance" value="0">
	<param name="currentPosition" value="0">
	<param name="defaultFrame" value>
	<param name="playCount" value="1">
	<param name="autoStart" value="-1">
	<param name="currentMarker" value="0">
	<param name="invokeURLs" value="-1">
	<param name="baseURL" value>
	<param name="volume" value="50">
	<param name="mute" value="0">
	<param name="uiMode" value="full">
	<param name="stretchToFit" value="0">
	<param name="windowlessVideo" value="0">
	<param name="enabled" value="-1">
	<param name="enableContextMenu" value="-1">
	<param name="fullScreen" value="0">
	<param name="SAMIStyle" value>
	<param name="SAMILang" value>
	<param name="SAMIFilename" value>
	<param name="captioningID" value>
	<param name="enableErrorDialogs" value="0">
</object> </td> 
  </tr> 
  <tr> 
    <td bgcolor="#B7C9F7" align="center"><input style="background-color:#FECFE4; height:18; width:50; border-width:0;" name="Fermer" type="button" onClick="window.close()" value="Fermer"></td> 
  </tr> 
</table> 
</BODY>
</HTML>
