<%
'Déclaration Variables
 dim tit_jeux_supp, SQLjs, Connjs, RS

'Création de la connection
 Set Connjs = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Connjs.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'Récupération des données du formulaire
 tit_jeux_supp = Request.Form ("titre_jeux_supp")

'Vérification du contenu des champs
 If (tit_jeux_supp = "" ) Then

'Si un des champs est vide on affiche une erreur
	Response.Redirect "jeux.asp?sent=erreur_suppr_jeux"

'Si les champs sont remplis on continu
 Else	

' Requête SQL pour séléctionner les jeux dont le titre correspond a la valeur du formulaire
 SQLjs = "SELECT * FROM [Jeux] where (Titre = '"&tit_jeux_supp&"')"

'Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLjs, Connjs
 RS.Open SQLjs, Connjs, 1, 2, 1

'On efface l'enregistrement
 RS.delete

'On affiche un message comme quoi ca s'est bien passé
 Response.Redirect "jeux.asp?sent=ok_suppr_jeux"

'Fermeture et Destruction RS & Connjs
 RS.Close : Set RS = Nothing
 Connjs.Close : Set Connjs = Nothing
 
 End If
%> 