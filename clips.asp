<%
	Set Conn = Server.CreateObject ("ADODB.Connection") 'Création de la connection
	Conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";" 'Ouverture de la connection
	lettre = Request.QueryString("lettre") 'On récupère la variable lettre de l'url
	if Request.QueryString("plus") = "true" Then 'Quand on clique sur modifier, la variable de l'url passe à true, on affiche le formulaire
	imgpm = "moins" 'Pour afficher le petit moins (l'image)
	display = "" 'On montre le div contenant le formulaire
	else 'Si la variable et pas à true on affiche pas le formulaire
	imgpm = "plus" 'Pour afficher le petit plus (l'image)
	display="none"'On cache le div contenant le formulaire
	end if 'et voilà :)
%>
<HTML>
<HEAD>
<TITLE>Base de données Clips</TITLE>
<!-- styles -->
<LINK media=screen href="scripts/css_styles.css" type=text/css rel=stylesheet>
<!-- scripts -->
<script language="Javascript" src="scripts/js_scripts.js" type="text/javascript"></SCRIPT>
<script language="Javascript" src="scripts/js_sorttable.js" type="text/javascript" ></script>
</HEAD>
<BODY id="id1" background="images/fond.gif"> 
<!-- div indispensable pour pouvoir afficher les infobulles --> 
<div style="visibility:hidden" id="curseur" class="infobulle"></div> 
<table style="border:#9094c9 0pt solid;" width="880" align="center" cellpadding="0" cellspacing="0"> 
  <tr> 
    <td width="880" align="center"><img src="images/ban1.gif" width="880" height="130"></td> 
  </tr> 
  <tr> 
    <td> <table height="35" style="border-left:#000000 1pt solid; border-right:#000000 1pt solid;" width="880" align="center" cellpadding="0" cellspacing="0"> 
        <tr height="40"> 
          <td align="center" class="Style3" width="125"><SCRIPT language="javascript">
var months=new Array(13);
months[1]="Janvier";
months[2]="Fevrier";
months[3]="Mars";
months[4]="Avril";
months[5]="Mai";
months[6]="Juin";
months[7]="Juillet";
months[8]="Aout";
months[9]="Septembre";
months[10]="Octobre";
months[11]="Novembre";
months[12]="Decembre";

var time=new Date();
var lmonth=months[time.getMonth() + 1];
var date=time.getDate();
var year=time.getYear();

document.write(date +" ");
document.write(lmonth + ", " + year);
</SCRIPT> 
      <!-- balise vide mais indispensable pour afficher l'heure --> 
      <a id="clock"></a> </td> 
          <td width="600"><span class="Style3"> 
            <div align="center"><a href="mp3.asp?lettre=ALL"><img hspace=7 align=absMiddle border="0" src="images/icone-CD-7.gif" width="32" height="32">Mp3</a> | <a href="clips.asp?lettre=ALL"><img border="0" hspace="7" align=absMiddle src="images/icone-CD-1.gif" width="32" height="32">Clips</a> | <a href="divx.asp?lettre=ALL"><img hspace="7" align=absMiddle border="0" src="images/movie.png" width="32" height="32">DivX</a> | <a href="jeux.asp?lettre=ALL"><img hspace=7 align=absMiddle border="0" src="images/icone-CD-2.gif" width="32" height="32">Jeux</a> </div> 
            </span></td> 
          <td width="125" align="center" class="Style3"> <SCRIPT LANGUAGE="JavaScript1.2">
		document.write("Votre r&eacute;solution : <b>"+screen.width+"x"+screen.height+"</b>.")
		</SCRIPT></td> 
        </tr>
		<tr bgcolor="#FFEFFF"><td style="border-top:#000000 1pt solid; border-bottom:#000000 1pt solid;">&nbsp;</td>
		<td style="border-top:#000000 1pt solid; border-bottom:#000000 1pt solid;"><span class="Style3"> 
            <div align="center"> <A HREF="clips.asp?lettre=0-9">0-9</A> <A HREF="clips.asp?lettre=A">A</A> <A HREF="clips.asp?lettre=B">B</A> <A HREF="clips.asp?lettre=C">C</A> <A HREF="clips.asp?lettre=D">D</A> <A HREF="clips.asp?lettre=E">E</A> <A HREF="clips.asp?lettre=F">F</A> <A HREF="clips.asp?lettre=G">G</A> <A HREF="clips.asp?lettre=H">H</A> <A HREF="clips.asp?lettre=I">I</A> <A HREF="clips.asp?lettre=J">J</A> <A HREF="clips.asp?lettre=K">K</A> <A HREF="clips.asp?lettre=L">L</A> <A HREF="clips.asp?lettre=M">M</A> <A HREF="clips.asp?lettre=N">N</A> <A HREF="clips.asp?lettre=O">O</A> <A HREF="clips.asp?lettre=P">P</A> <A HREF="clips.asp?lettre=Q">Q</A> <A HREF="clips.asp?lettre=R">R</A> <A HREF="clips.asp?lettre=S">S</A> <A HREF="clips.asp?lettre=T">T</A> <A HREF="clips.asp?lettre=U">U</A> <A HREF="clips.asp?lettre=V">V</A> <A HREF="clips.asp?lettre=W">W</A> <A HREF="clips.asp?lettre=X">X</A> <A HREF="clips.asp?lettre=Y">Y</A> <A HREF="clips.asp?lettre=Z">Z</A> <A HREF="clips.asp?lettre=ALL">ALL</A></div> 
            </span></td>
		<td style="border-top:#000000 1pt solid; border-bottom:#000000 1pt solid;">&nbsp;</td>
		</tr> 
        <tr> 
          <td style="border-bottom:#000000 1pt solid;" height="50" align="center"> 
            <span class="Style3"><a class=current id="tlbk_3">Rechercher un mot :</a></span><br>
            <input style="background-color:#F7CBF7; border-color:#9094C9; border-style:solid; border-width:1;" class="Style1" name="string" type="text" size=15 onChange="n = 0;"> 
            <input name="search" onClick="return findInPage(string.value);" class="Style3" type="submit" value="Go">          </td> 
          <td style="border-bottom:#000000 1pt solid;" align="center"><a onclick="visibilite('tlbk_form');" href="javascript:changerimg('img4');"><img id="img4" border="0" src="images/<% = imgpm %>.jpg" width="9" height="9"></a> Ajouter/Modifier un clip</td> 
          <td style="border-bottom:#000000 1pt solid;" align="center">
            <span class="Style3"><a class=current id="tlbk_1">Selection par style :</a></span>
            <select style="background-color:#F7CBF7; border-color:#9094C9; border-style:solid; border-width:1;" class="Style5" onChange="ChangeUrl(this)"> 
              <option selected>Choix...</option> 
              <option value="clips.asp?lettre=RapRNB">Rap/R&amp;B</option> 
              <option value="clips.asp?lettre=Makina">Makina</option>  
              <option value="clips.asp?lettre=Hardcore">Hardcore</option> 
              <option value="clips.asp?lettre=Pop">Pop</option> 
              <option value="clips.asp?lettre=Autre">Autre</option> 
              <option value="clips.asp?lettre=Dance">Dance</option> 
            </select>            </td> 
        </tr> 
      </table></td> 
  </tr> 
</table>
<table width="880" border="0" align="center" cellspacing="0" style="border-left:#000000 1pt solid; border-right:#000000 1pt solid;">
  <tr>
  <td style="border-left:#000000 1pt solid;"  width="250"></td> 
    <td width="380"><div align="center" id="tlbk_form" style="display:<% = display %>;"> 
        <%
dim aute, titre, dure, styl, tail, exte, lien
aute = Request.QueryString ("aut")
titr = Request.QueryString ("tit")
dure = Request.QueryString ("dur")
styl = Request.QueryString ("sty")
tail = Request.QueryString ("tai")
exte = Request.QueryString ("ext")
lien = Request.QueryString ("lie")
%> <br>
        <FORM ACTION="sent_clip.asp" METHOD="post"> 
          <table width="300" border="0" align="center" cellspacing="0" style="border:#9094c9 1pt solid;"> 
            <tr> 
              <td align="center" colspan="2" style="border-bottom:#9094c9 1pt solid;">Ajouter/Modifier un clip :</td> 
            </tr> 
            <tr> 
              <td>&nbsp;Auteur :</td> 
              <td><input value="<% = aute %>" class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="auteur_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Titre : </td> 
              <td><input value="<% = titr %>" class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="titre_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Durée : </td> 
              <td><input value="<% = dure %>" class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="duree_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Style : </td> 
              <td><input value="<% = styl %>" class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="style_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Taille : </td> 
              <td><input value="<% = tail %>" class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="taille_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Extension : </td> 
              <td><input value="<% = exte %>" class="Style1"style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="extension_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td>&nbsp;Lien : </td> 
              <td><input value="<% = lien %>" class="Style1"style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="lien_clip" type="text"></td> 
            </tr> 
            <tr> 
              <td align="center" colspan="2"><input style="background-color:#FECFE4; height:18; width:50; border-width:0;" name="submit" type="submit" value="valider"></td> 
            </tr> 
          </table> 
        </FORM> 
        <%
If Request.QueryString ("sent") = "erreur" Then
 Response.Write "<span style=color:#FF0000;>Erreur, vérifiez vos valeurs et remplissez tous les champs <br>&nbsp;</span>"
elseif Request.QueryString ("sent") = "ok" Then
 Response.Write "<span style=color:#FF0000;>Les données ont bien été enregistrées dans la base <br>&nbsp;</span>"
elseif Request.QueryString ("sent") = "ok_suppr" Then
 Response.Write "<span style=color:#FF0000;>Les données ont bien été supprimées de la base <br>&nbsp;</span>"
End If  
%> 
      </div></td> 
    <td style="border-right:#000000 1pt solid;" width="250"></td>
  </tr>
</table> 
<table width="880" align="center" style="border-left:#000000 1pt solid; border-right:#000000 1pt solid; border-bottom:#000000 1pt solid;" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td> <TABLE width="100%" align="center" cellpadding="0" cellspacing="0"> 
              <TR height="18" bgcolor="#F7CBF7"> 
                <TD style="border-top:#000000 1pt solid;" class="Style3"> <b>&nbsp;<a href="clips.asp?lettre=SortAuteur">Auteur</a></b> </TD> 
                <TD bgcolor="#F7CBF7" class="Style3" style="border-top:#000000 1pt solid;"> <b><a href="clips.asp?lettre=SortTitre">Titre</a></b> </TD> 
                <TD style="border-top:#000000 1pt solid;" class="Style3"> <b><a href="clips.asp?lettre=SortDuree">Durée</a></b> </TD> 
                <TD style="border-top:#000000 1pt solid;" class="Style3"> <b><a href="clips.asp?lettre=SortStyle">Style</a></b> </TD>  
                <TD style="border-top:#000000 1pt solid;" class="Style3"> <b><a href="#">Options</a></b> </TD> 
            </TR> 
              <% 
'---------------------------------D - Récupération des valeurs dans la base-------------------------------->
'variable pour compter le nombre de lignes dans chaque répertoires
i=0
'on selectionne les Clips à l'aide des requêtes
if (lettre="0-9") then
 SQL = "SELECT * FROM [Clips] WHERE Auteur LIKE '1%' OR Auteur LIKE '2%' OR Auteur LIKE '3%' OR Auteur LIKE '4%' OR Auteur LIKE '5%' OR Auteur LIKE '6%' OR Auteur LIKE '7%' OR Auteur LIKE '8%' OR Auteur LIKE '9%' ORDER by Auteur ASC"
elseif (lettre="ALL") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Auteur ASC"
elseif (lettre="RapRNB") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Rap/R&B%' ORDER by Auteur ASC"
elseif (lettre="Makina") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Makina%' ORDER by Auteur ASC"   
elseif (lettre="Classique") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Classique%' ORDER by Auteur ASC"  
elseif (lettre="Test") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Test%' ORDER by Auteur ASC"
 elseif (lettre="Hardcore") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Hardcore%' ORDER by Auteur ASC" 
elseif (lettre="Pop") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Pop%' ORDER by Auteur ASC"
elseif (lettre="Autre") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Autre%' ORDER by Auteur ASC"
elseif (lettre="Dance") then
 SQL = "SELECT ALL * FROM [Clips] WHERE Style LIKE '%Dance%' ORDER by Auteur ASC"  
elseif (lettre="SortAuteur") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Auteur ASC"
elseif (lettre="SortTitre") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Titre ASC"
elseif (lettre="SortDuree") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Durée ASC" 
elseif (lettre="SortStyle") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Style ASC" 
elseif (lettre="SortTaille") then
 SQL = "SELECT ALL * FROM [Clips] ORDER by Taille ASC"           
else
 SQL = "SELECT * FROM [Clips] WHERE Auteur LIKE '"&lettre&"%"&"' ORDER by Auteur ASC"
end if
'on crée un recordset dans RS
 Set RS = Server.CreateObject ("ADODB.RecordSet")
'on ouvre le recordset
 RS.Open SQL, Conn
'si on est pas à la fin des enregistrements
 If Not RS.EOF Then
'tant qu'on est pas à la fin des enregistrements
 While Not RS.EOF
 'on incrémente i pour compter le nb d'enregistrements
 i=i+1
 'on récupère les valeurs des champs dans la base
 dim AuteurClip, TitreClip, DureeClip, StyleClip, TailleClip, ExtensionClip, LienClip

 AuteurClip = RS ("Auteur")
 TitreClip = RS ("Titre")
 DureeClip = RS ("Durée")
 StyleClip = RS ("Style")
 TailleClip = RS ("Taille")
 ExtensionClip = RS ("Extension")
 LienClip = RS ("Lien")
 LienClip = server.urlencode(LienClip)
 '---------------------------------F - Récupération des valeurs dans la base-------------------------------->
%>
              <TR onMouseOver="changeCouleur(this);" onMouseOut="remetCouleur(this);"> 
                <TD class="Style3">&nbsp;<img onMouseOver="montre('<table class=Style3 bgcolor=#EBEBEB><tr><td><u>Taille :</u></td><td align=center><%=TailleClip/1000%> Mo</td></tr><tr><td><u>Extension :</u></td><td align=center><%=ExtensionClip%></td></tr></table>')" onMouseOut="cache()" src="images/clip.gif" width="17" height="9"> 
                <% = AuteurClip %> </TD> 
                <TD class="Style3"> <% = TitreClip %> </TD> 
                <TD class="Style3"> <% = DureeClip/100 %> 
                  min</TD> 
                <TD class="Style3"> <% = StyleClip %> </TD> 
                <TD class="Style3"><a target="_blank" style="cursor:pointer; vertical-align:middle;" onclick="window.open('regarderclip.asp?url=<% = LienClip %>', 'Sample', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=yes,width=400,height=334')"><img src="images/play.gif" alt="écouter" width="15" height="15" border="0"></a>&nbsp;<a href="sent_modif_clip.asp?auteur=<% = server.URLEncode(AuteurClip) %>&titre=<% = server.URLEncode(TitreClip) %>"><img src="images/modifier.gif" alt="modifier" width="15" height="14" border="0"></a>&nbsp;<a href="sent_supp_clip.asp?auteur=<% = server.URLEncode(AuteurClip) %>&titre=<% = server.URLEncode(TitreClip) %>"><img src="images/supprimer.gif" alt="supprimer" width="15" height="14" border="0"></a></TD> 
              </TR> 
              <%
 '------------------------------------D - Affichages & fermeture de la base--------------------------------->
'on passe à l'enregistrement suivant
 RS.MoveNext
 Wend
	 Response.Write "<span class="&"Style2"&"> &nbsp;<img style="&"vertical-align:middle"&" src="&"images/icone_dossier.gif"&">"
	 Response.Write "&nbsp;Il y a actuellement "
 	 Response.Write "<b>"&(i)&"</b>"
 	 Response.Write " clips dans ce répertoire. "
	 Response.Write " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Options : "
	 Response.Write "<a onclick=visibilite('tlbk_5'); href=javascript:changerimg('img3');>"	 
	 Response.Write "<img id=img3 border=0 src=images/moins.jpg width=9 height=9></a> "
	 Response.Write "<span id=tlbk_5><a href="&"javascript:history.back()"&"><img align=middle border="&"0"&" src="&"images/avant.gif"&"></a> "
	 Response.Write "<a href="&"javascript:window.print()"&"><img align=middle border="&"0"&" src="&"images/print.gif"&"></a> "
	 Response.Write "<a href="&"javascript:favoris()"&"><img align=middle border="&"0"&" src="&"images/favoris.gif"&"></a> "
	 Response.Write "<a href="&"#"&" onClick="&"HomePage(this);"&"><img align=middle border="&"0"&" src="&"images/demarage.gif"&"></a> "
	 Response.Write "<a href="&"javascript:history.go(0)"&"><img align=middle border="&"0"&" src="&"images/actualise.gif"&"></a> "
	 Response.Write "<a href="&"javascript:history.forward()"&"><img align=middle border="&"0"&" src="&"images/apres.gif"&"></a></span>"
	 Response.Write "</span>"

 End If

'on ferme l'enregistrement et la connexion
 RS.Close : Set RS = Nothing
 Conn.Close : Set Conn = Nothing
  '----------------------------------F - Affichages & fermeture de la base----------------------------------> 
%> 
            </TABLE></td> 
        </tr> 
</table> 
      <table align="center"> 
        <tr> 
          <td><span class="Style1">&copy; Copyright <a href="mailto:admin@example.com"><img src="images/icone-Mail-6.gif" width="14" height="10" border="0"></a> Yann39 2005-2006 </span></td>
        </tr> 
      </table> 
</BODY>
</HTML>
