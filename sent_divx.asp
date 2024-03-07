<%
'Déclaration des Variables
 dim tit_divx, dur_divx, gen_divx, tai_divx, dat_divx, ext_divx, lie_divx, SQLd, Connd, RS

'Création de la connection
 Set Connd = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Connd.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'Récupération des données du formulaire
 tit_divx = Request.Form ("titre_divx")
 dur_divx = Request.Form ("duree_divx")
 gen_divx = Request.Form ("genre_divx")
 tai_divx = Request.Form ("taille_divx")
 dat_divx = Request.Form ("date_divx")
 ext_divx = Request.Form ("extension_divx")
 lie_divx = Request.Form ("lien_divx")

'Vérification du contenu des champs
 If (tit_divx = "" OR dur_divx = "" OR gen_divx = "" OR tai_divx = "" OR dat_divx = "" OR lie_divx = "" OR ext_divx = "") Then

'Si un des champs est vide, on affiche une erreur
 Response.Redirect "divx.asp?sent=erreur_divx"

'Si les champs sont remplis on créé le nouvel enregistrement  
 Else	

'Requête SQL pour séléctionner toute la table Divx
 SQLd = "SELECT * FROM [Divx] ORDER by Titre ASC"

' Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

' Ouverture RS, SQLd, Connd
 RS.Open SQLd, Connd, 3, 3

' Ajout des informations dans la base et mise à jour
 RS.AddNew
 RS.Fields ("Titre").Value = tit_divx
 RS.Fields ("Durée").Value = dur_divx
 RS.Fields ("Genre").Value = gen_divx
 RS.Fields ("Taille").Value = tai_divx
 RS.Fields ("Date").Value = dat_divx
 RS.Fields ("Extension").Value = ext_divx
 RS.Fields ("Lien").Value = lie_divx
 RS.Update

'On affiche un message comme quoi ca s'est bien passé
 Response.Redirect "divx.asp?sent=ok_divx&plus=true"

' Fermeture et Destruction RS & Connd
 RS.Close : Set RS = Nothing
 Connd.Close : Set Connd = Nothing
 
 End If
%> 