<%
'D�claration Variables
 dim tit_jeux_supp, SQLjs, Connjs, RS

'Cr�ation de la connection
 Set Connjs = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Connjs.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'R�cup�ration des donn�es du formulaire
 tit_jeux_supp = Request.Form ("titre_jeux_supp")

'V�rification du contenu des champs
 If (tit_jeux_supp = "" ) Then

'Si un des champs est vide on affiche une erreur
	Response.Redirect "jeux.asp?sent=erreur_suppr_jeux"

'Si les champs sont remplis on continu
 Else	

' Requ�te SQL pour s�l�ctionner les jeux dont le titre correspond a la valeur du formulaire
 SQLjs = "SELECT * FROM [Jeux] where (Titre = '"&tit_jeux_supp&"')"

'Cr�ation de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLjs, Connjs
 RS.Open SQLjs, Connjs, 1, 2, 1

'On efface l'enregistrement
 RS.delete

'On affiche un message comme quoi ca s'est bien pass�
 Response.Redirect "jeux.asp?sent=ok_suppr_jeux"

'Fermeture et Destruction RS & Connjs
 RS.Close : Set RS = Nothing
 Connjs.Close : Set Connjs = Nothing
 
 End If
%> 