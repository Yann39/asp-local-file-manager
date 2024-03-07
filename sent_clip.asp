<%
'Déclaration des Variables
 dim aut_clip, tit_clip, dur_clip, sty_clip, tai_clip, ext_clip, lie_clip, SQLc, Connc, RS

'Création de la connection
 Set Connc = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Connc.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'Récupération des données du formulaire
 aut_clip = Request.Form ("auteur_clip")
 tit_clip = Request.Form ("titre_clip")
 dur_clip = Request.Form ("duree_clip")
 sty_clip = Request.Form ("style_clip")
 tai_clip = Request.Form ("taille_clip")
 ext_clip = Request.Form ("extension_clip")
 lie_clip = Request.Form ("lien_clip")

'Vérification du contenu des champs
 If (aut_clip = "" OR tit_clip = "" OR dur_clip = "" OR sty_clip = "" OR tai_clip = "" OR ext_clip = "" OR lie_clip = "") Then

'Si un des champs est vide, on affiche une erreur
 Response.Redirect "clips.asp?sent=erreur_clip"

'Si les champs sont remplis on créé le nouvel enregistrement  
 Else	

'Requête SQL pour séléctionner toute la table Clips
 SQLc = "SELECT * FROM [Clips] ORDER by Auteur ASC"

' Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

' Ouverture RS, SQLc, Connc
 RS.Open SQLc, Connc, 3, 3

' Ajout des informations dans la base et mise à jour
 RS.AddNew
 RS.Fields ("Auteur").Value = aut_clip
 RS.Fields ("Titre").Value = tit_clip
 RS.Fields ("Durée").Value = dur_clip
 RS.Fields ("Style").Value = sty_clip
 RS.Fields ("Taille").Value = tai_clip
 RS.Fields ("Extension").Value = ext_clip
 RS.Fields ("Lien").Value = lie_clip
 RS.Update

'On affiche un message comme quoi ca s'est bien passé
 Response.Redirect "clips.asp?sent=ok_clip&plus=true"

' Fermeture et Destruction RS & Connc
 RS.Close : Set RS = Nothing
 Connc.Close : Set Connc = Nothing
 
 End If
%> 