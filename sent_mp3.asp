<%
'Déclaration des Variables
 dim aut_mp3, tit_mp3, dur_mp3, sty_mp3, tai_mp3, dat_mp3, lie_mp3, SQLe, Conne, RS

'Création de la connection
 Set Conne = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Conne.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'Récupération des données du formulaire
 aut_mp3 = Request.Form ("auteur_mp3")
 tit_mp3 = Request.Form ("titre_mp3")
 dur_mp3 = Request.Form ("duree_mp3")
 sty_mp3 = Request.Form ("style_mp3")
 tai_mp3 = Request.Form ("taille_mp3")
 dat_mp3 = Request.Form ("date_mp3")
 lie_mp3 = Request.Form ("lien_mp3")

'Vérification du contenu des champs
 If (aut_mp3 = "" OR tit_mp3 = "" OR dur_mp3 = "" OR sty_mp3 = "" OR tai_mp3 = "" OR dat_mp3 = "" OR lie_mp3 = "") Then

'Si un des champs est vide, on affiche une erreur
 Response.Redirect "mp3.asp?sent=erreur&plus=true"

'Si les champs sont remplis on créé le nouvel enregistrement  
 Else	

'Requête SQL pour séléctionner toute la table Musique
 SQLe = "SELECT * FROM [Musique] ORDER by Auteur ASC"

' Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

' Ouverture RS, SQLe, Conne
 RS.Open SQLe, Conne, 3, 3

' Ajout des informations dans la base et mise à jour
 RS.AddNew
 RS.Fields ("Auteur").Value = aut_mp3
 RS.Fields ("Titre").Value = tit_mp3
 RS.Fields ("Durée").Value = dur_mp3
 RS.Fields ("Style").Value = sty_mp3
 RS.Fields ("Taille").Value = tai_mp3
 RS.Fields ("Date").Value = dat_mp3
 RS.Fields ("Lien").Value = lie_mp3
 RS.Update

'On affiche un message comme quoi ca s'est bien passé
 Response.Redirect "mp3.asp?sent=ok&plus=true"

' Fermeture et Destruction RS & Conne
 RS.Close : Set RS = Nothing
 Conne.Close : Set Conne = Nothing
 
 End If
%> 