<%
'Déclaration Variables
 dim aut_mp3_modif, tit_mp3_modif, SQLs, Conns, RS
 aut_mp3_modif = Request.QueryString("auteur")
 tit_mp3_modif = Request.QueryString("titre")

'Création de la connection
 Set Conns = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Conns.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

' Requête SQL pour séléctionner les musiques dont l'auteur ou le titre correspond à ceux demandés
 SQLs = "SELECT * FROM [Musique] where (Auteur = '"&aut_mp3_modif&"') and (Titre = '"&tit_mp3_modif&"')"

'Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLs, Conns
 RS.Open SQLs, Conns, 1, 2, 1

' déclaration des variables
 dim aut_mp3_mod, tit_mp3_mod, dur_mp3_mod, sty_mp3_mod, tai_mp3_mod, dat_mp3_mod, lie_mp3_mod
 
'récupération des valeurs dans la base
 aut_mp3_mod = RS.Fields ("Auteur").Value
 tit_mp3_mod = RS.Fields ("Titre").Value 
 dur_mp3_mod = RS.Fields ("Durée").Value 
 sty_mp3_mod = RS.Fields ("Style").Value 
 tai_mp3_mod = RS.Fields ("Taille").Value 
 dat_mp3_mod = RS.Fields ("Date").Value
 lie_mp3_mod = RS.Fields ("Lien").Value 

'on efface l'ancien enregistrement
 'RS.delete

'On redirectionne vers l'adresse avec les parametres de la musique afin de remplir automatiquement le formulaire
 Response.Redirect "mp3.asp?aut="&server.URLEncode(aut_mp3_mod)&"&tit="&server.URLEncode(tit_mp3_mod)&"&dur="&server.URLEncode(dur_mp3_mod)&"&sty="&server.URLEncode(sty_mp3_mod)&"&tai="&server.URLEncode(tai_mp3_mod)&"&dat="&server.URLEncode(dat_mp3_mod)&"&lie="&server.URLEncode(lie_mp3_mod)&"&plus=true"
 
'Fermeture et Destruction RS & Conns
 RS.Close : Set RS = Nothing
 Conns.Close : Set Conns = Nothing
%> 