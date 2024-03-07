<%
'D�claration des Variables
 dim tit_jeux, gen_jeux, nbc_jeux, SQLj, Connj, RS

'Cr�ation de la connection
 Set Connj = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Connj.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'R�cup�ration des donn�es du formulaire
 tit_jeux = Request.Form ("titre_jeux")
 gen_jeux = Request.Form ("genre_jeux")
 nbc_jeux = Request.Form ("nbcd_jeux")

'V�rification du contenu des champs
 If (tit_jeux = "" OR gen_jeux = "" OR nbc_jeux = "") Then

'Si un des champs est vide, on affiche une erreur
 Response.Redirect "jeux.asp?sent=erreur_jeux"

'Si les champs sont remplis on cr�� le nouvel enregistrement  
 Else	

'Requ�te SQL pour s�l�ctionner toute la table Jeux
 SQLj = "SELECT * FROM [Jeux] ORDER by Titre ASC"

' Cr�ation de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

' Ouverture RS, SQLj, Connj
 RS.Open SQLj, Connj, 3, 3

' Ajout des informations dans la base et mise � jour
 RS.AddNew
 RS.Fields ("Titre").Value = tit_jeux
 RS.Fields ("Genre").Value = gen_jeux
 RS.Fields ("NbCD").Value = nbc_jeux
 RS.Update

'On affiche un message comme quoi ca s'est bien pass�
 Response.Redirect "jeux.asp?sent=ok_jeux"

' Fermeture et Destruction RS & Connj
 RS.Close : Set RS = Nothing
 Connj.Close : Set Connj = Nothing
 
 End If
%> 