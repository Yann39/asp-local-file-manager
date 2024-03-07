<%
	Set Conn = Server.CreateObject ("ADODB.Connection")'Création de la connection
	Conn.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"'Ouverture de la connection
	lettre = Request.QueryString("lettre")'On récupère la variable lettre de l'url
%>
<HTML>
<HEAD>
<TITLE>Base de données Jeux</TITLE>
<!--Styles-->
<LINK media=screen href="scripts/css_styles.css" type=text/css rel=stylesheet>
<!--Scripts-->
<script language="Javascript" src="scripts/js_scripts.js" type="text/javascript"></SCRIPT>
<script language="Javascript" src="scripts/js_sorttable.js" type="text/javascript" ></script>
</HEAD>
<BODY id="id1" background="images/fond.gif"> 
<table style="border:#9094c9 1pt solid;" width="85%" align="center" cellpadding="0" cellspacing="0"> 
  <tr> 
    <td colspan="3" align="center"><span class="Style4">Base de données Jeux</span></td> 
  </tr> 
  <tr> 
    <td colspan="3"> <table bgcolor="#EBEBEB" height="35" style="border:#9094c9 1pt solid; border-left:none; border-right:none" width="100%" align="center" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td width="15" rowspan="2" align="center"><a onclick="visibilite('tlbk_3'); visibilite('tlbk_4');" href="javascript:changerimg('img1');"><img id="img1" border="0" src="images/moins.jpg" width="9" height="9"></a></td> 
          <td width="150" align="center" class="Style3"><A class=current id="tlbk_3">Couleur arri&egrave;re plan :</a></td> 
          <td><span class="Style3"> 
            <div align="center"><a href="mp3.asp?lettre=ALL">Mp3</a> | <a href="clips.asp?lettre=ALL">Clips</a> | <a href="divx.asp?lettre=ALL">DivX</a> | <a href="jeux.asp?lettre=ALL">Jeux</a> </div> 
            </span></td> 
          <td width="150" align="center" class="Style3"><A class=current id="tlbk_1">Selection par style :</a></td> 
          <td width="15" rowspan="2" align="center" class="Style3"><a onclick="visibilite('tlbk_1'); visibilite('tlbk_2');" href="javascript:changerimg('img2');"><img id="img2" border="0" src="images/moins.jpg" width="9" height="9"></a></td> 
        </tr> 
        <tr> 
          <td align="center"><A class=current id="tlbk_4"> 
            <select onChange="ChangeArrierePlan(id1, this)" class="Style5" style="border-color:#000000; border-style:solid; border-width:1;"> 
              <option value="images/fond.gif" selected>Choix...</option> 
              <option value="images/fondred.gif">rouge</option> 
              <option value="images/fondblue.gif">bleu</option> 
              <option value="images/fondyellow.gif">jaune</option> 
              <option value="images/fondgreen.gif">vert</option> 
              <option value="">aucun</option> 
            </select> 
            </a></td> 
          <td> <span class="Style3"> 
            <div align="center"> <A HREF="jeux.asp?lettre=0-9">0-9</A> <A HREF="jeux.asp?lettre=A">A</A> <A HREF="jeux.asp?lettre=B">B</A> <A HREF="jeux.asp?lettre=C">C</A> <A HREF="jeux.asp?lettre=D">D</A> <A HREF="jeux.asp?lettre=E">E</A> <A HREF="jeux.asp?lettre=F">F</A> <A HREF="jeux.asp?lettre=G">G</A> <A HREF="jeux.asp?lettre=H">H</A> <A HREF="jeux.asp?lettre=I">I</A> <A HREF="jeux.asp?lettre=J">J</A> <A HREF="jeux.asp?lettre=K">K</A> <A HREF="jeux.asp?lettre=L">L</A> <A HREF="jeux.asp?lettre=M">M</A> <A HREF="jeux.asp?lettre=N">N</A> <A HREF="jeux.asp?lettre=O">O</A> <A HREF="jeux.asp?lettre=P">P</A> <A HREF="jeux.asp?lettre=Q">Q</A> <A HREF="jeux.asp?lettre=R">R</A> <A HREF="jeux.asp?lettre=S">S</A> <A HREF="jeux.asp?lettre=T">T</A> <A HREF="jeux.asp?lettre=U">U</A> <A HREF="jeux.asp?lettre=V">V</A> <A HREF="jeux.asp?lettre=W">W</A> <A HREF="jeux.asp?lettre=X">X</A> <A HREF="jeux.asp?lettre=Y">Y</A> <A HREF="jeux.asp?lettre=Z">Z</A> <A HREF="jeux.asp?lettre=ALL">ALL</A></div> 
            </span> </td> 
          <td align="center"><A class=current id="tlbk_2"> 
            <select class="Style5" onChange="ChangeUrl(this)"> 
              <option selected>Choix...</option> 
              <option value="jeux.asp?lettre=Action">Action</option> 
              <option value="jeux.asp?lettre=Strategie">Stratégie</option> 
              <option value="jeux.asp?lettre=Course">Course</option> 
              <option value="jeux.asp?lettre=Avion">Avion</option> 
              <option value="jeux.asp?lettre=Helicopter">Hélicopter</option> 
              <option value="jeux.asp?lettre=Billard">Billard</option> 
            </select> 
            </a></td> 
        </tr> 
      </table> 
      <table width="100%" align="center"> 
        <tr> 
          <td align="center"><a onclick="visibilite('tlbk_form');" href="javascript:changerimg('img4');"><img id="img4" border="0" src="images/plus.jpg" width="9" height="9"></a> Ajouter un jeu          <a onclick="visibilite('tlbk_form2');" href="javascript:changerimg('img5');"><img id="img5" border="0" src="images/plus.jpg" width="9" height="9"></a> Effacer un jeux
            <div id="tlbk_form2" style="display:none;"> 
              <FORM ACTION="sent_supp_jeux.asp" METHOD="post"> 
                <table width="300" border="0" align="center" cellspacing="0" style="border:#9094c9 1pt solid;"> 
                  <tr> 
                    <td align="center" colspan="2" style="border-bottom:#9094c9 1pt solid;">Effacer un jeu :</td>
                  </tr> 
                  <tr> 
                    <td>&nbsp;Titre : </td> 
                    <td><input class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="titre_jeux_supp" type="text"></td> 
                  </tr> 
                  <tr> 
                    <td align="center" colspan="2"><input style="background-color:#FECFE4; height:18; width:50; border-width:0;" name="submit" type="submit" value="valider"></td> 
                  </tr> 
                </table> 
                <br> 
              </FORM> 
              <%
If Request.QueryString ("sent") = "erreur_suppr_jeux" Then
 Response.Write "Erreur, il n'y a pas d'enregistrement portant ce nom"
elseif Request.QueryString ("sent") = "ok_suppr_jeux" Then
 Response.Write "L'enregistrement a bien été supprimé de la base"
End If  
%> 
            </div> 
            <div id="tlbk_form" style="display:none;"> 
              <FORM ACTION="sent_jeux.asp" METHOD="post"> 
                <table width="300" border="0" align="center" cellspacing="0" style="border:#9094c9 1pt solid;"> 
                  <tr> 
                    <td align="center" colspan="2" style="border-bottom:#9094c9 1pt solid;">Ajouter un jeu :</td>
                  </tr> 
                  <tr> 
                    <td>&nbsp;Titre : </td> 
                    <td><input class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="titre_jeux" type="text"></td> 
                  </tr> 
                  <tr> 
                    <td>&nbsp;Genre : </td> 
                    <td><input class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="genre_jeux" type="text"></td> 
                  </tr> 
                  <tr> 
                    <td>&nbsp;NbCD : </td> 
                    <td><input class="Style1" style="width:200; background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" name="nbcd_jeux" type="text"></td> 
                  </tr> 
                  <tr> 
                    <td align="center" colspan="2"><input style="background-color:#FECFE4; height:18; width:50; border-width:0;" name="submit" type="submit" value="valider"></td> 
                  </tr> 
                </table> 
                <br> 
              </FORM> 
              <%
If Request.QueryString ("sent") = "erreur_jeux" Then
 Response.Write "Erreur, remplissez tous les champs"
elseif Request.QueryString ("sent") = "ok_jeux" Then
 Response.Write "Les données ont bien été enregistrées dans la base"
End If  
%> 
            </div> 
            <br></td> 
        </tr> 
      </table></td> 
  </tr> 
  <tr> 
    <td class="Style3" width="250" align="right"> <SCRIPT language="javascript">
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
      <a id="clock"></a> </td> 
    <td height="35" align="center"> <p><span class="Style3">Rechercher : </span> 
        <input style="background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:1;" class="Style1" name="string" type="text" size=15 onChange="n = 0;"> 
        <input style="background-color:#EBEBEB; border-color:#000000; border-style:solid; border-width:0; " name="search" onClick="return findInPage(string.value);" class="Style3" type="submit" value="Go"> 
      </p></td> 
    <td class="Style3" width="250" align="center"> <SCRIPT LANGUAGE="JavaScript1.2">
		document.write("r&eacute;solution : <b>"+screen.width+"x"+screen.height+"</b>.")
		</SCRIPT> </td> 
  </tr> 
  <tr> 
    <td colspan="3"> <table width="96%" align="center" style="border:#9094c9 1pt solid;" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td> <TABLE width="100%" align="center" cellpadding="0" cellspacing="0" class="sortable" id="youhou" style="border:#9094c9 1pt solid; border-left:none; border-right:none; border-bottom:none"> 
              <TR height="18" bgcolor="#E8E8F9"> 
                <TD width="500" class="Style3"> <b>Titre</b> </TD> 
                <TD width="400" class="Style3"> <b>Genre</b> </TD> 
                <TD width="100" class="Style3"> <b>Nb CD</b> </TD> 
              </TR> 
              <% 
'---------------------------------D - Récupération des valeurs dans la base-------------------------------->
'variable pour compter le nombre de lignes dans chaque répertoires
i=0
'on selectionne les Jeux à l'aide des requêtes
if (lettre="0-9") then
 SQL = "SELECT * FROM [Jeux] WHERE Titre LIKE '1%' OR Titre LIKE '2%' OR Titre LIKE '3%' OR Titre LIKE '4%' OR Titre LIKE '5%' OR Titre LIKE '6%' OR Titre LIKE '7%' OR Titre LIKE '8%' OR Titre LIKE '9%' ORDER by Titre ASC"
elseif (lettre="ALL") then
 SQL = "SELECT ALL * FROM [Jeux] ORDER by Titre ASC"
elseif (lettre="Action") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Action%' ORDER by Titre ASC"
 elseif (lettre="Strategie") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Stratégie%' ORDER by Titre ASC"
 elseif (lettre="Course") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Course%' ORDER by Titre ASC"
 elseif (lettre="Avion") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Avion%' ORDER by Titre ASC"
 elseif (lettre="Helicopter") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Hélicopter%' ORDER by Titre ASC"
 elseif (lettre="Billard") then
 SQL = "SELECT ALL * FROM [Jeux] WHERE Genre LIKE '%Billard%' ORDER by Titre ASC"
else
 SQL = "SELECT * FROM [Jeux] WHERE Titre LIKE '"&lettre&"%"&"' ORDER by Titre ASC"
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
 dim TitreJeux, GenreJeux, NbCDJeux
 TitreJeux = RS ("Titre")
 GenreJeux = RS ("Genre")
 NbCDJeux = RS ("NbCD")
 '---------------------------------F - Récupération des valeurs dans la base-------------------------------->
%> 
              <TR height="15" bgcolor="#EBEBEB" onMouseOver="changeCouleur(this);" onMouseOut="remetCouleur(this);"> 
                <TD><span class="Style3">&nbsp;<img src="images/imgcd.JPG" width="12" height="14"> 
                  <% = TitreJeux %> 
                  </span></TD> 
                <TD><span class="Style3"> 
                  <% = Genrejeux %> 
                  </span></TD> 
                <TD><span class="Style3"> 
                  <% = NbCDJeux %> 
                  cd(s)</span></TD> 
              </TR> 
              <%
 '------------------------------------D - Affichages & fermeture de la base--------------------------------->
'on passe à l'enregistrement suivant
 RS.MoveNext
 Wend

	 Response.Write "<span class="&"Style2"&"> &nbsp;<img src="&"images/icone_dossier.gif"&">"
	 Response.Write "&nbsp;Il y a actuellement "
 	 Response.Write "<b>"&(i)&"</b>"
 	 Response.Write " jeux dans ce répertoire. "
	 Response.Write " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Options : "
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
      </table></td> 
  </tr> 
</table> 
</BODY>
</HTML>
