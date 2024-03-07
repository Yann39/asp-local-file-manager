<%
'Déclaration Variables
 dim aut_clip_modif, tit_clip_modif, SQLs, Conns, RS
 aut_clip_modif = Request.QueryString("auteur")
 tit_clip_modif = Request.QueryString("titre")

'Création de la connection
 Set Conns = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Conns.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

' Requête SQL pour séléctionner les musiques dont l'auteur ou le titre correspond au valeurs du formulaire
 SQLs = "SELECT * FROM [Clips] where (Auteur = '"&aut_clip_modif&"') and (Titre = '"&tit_clip_modif&"')"

'Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLs, Conns
 RS.Open SQLs, Conns, 1, 2, 1

' déclaration des variables
 dim aut_clip_mod, tit_clip_mod, dur_clip_mod, sty_clip_mod, tai_clip_mod, ext_clip_mod, lie_clip_mod
 
'récupération des valeurs dans la base
 aut_clip_mod = RS.Fields ("Auteur").Value
 tit_clip_mod = RS.Fields ("Titre").Value 
 dur_clip_mod = RS.Fields ("Durée").Value 
 sty_clip_mod = RS.Fields ("Style").Value 
 tai_clip_mod = RS.Fields ("Taille").Value 
 ext_clip_mod = RS.Fields ("Extension").Value
 lie_clip_mod = RS.Fields ("Lien").Value 

'on efface l'ancien enregistrement
 'RS.delete

'On redirectionne vers l'adresse avec les parametres de la musique afin de remplir automatiquement le formulaire
 Response.Redirect "clips.asp?aut="&server.URLEncode(aut_clip_mod)&"&tit="&server.URLEncode(tit_clip_mod)&"&dur="&server.URLEncode(dur_clip_mod)&"&sty="&server.URLEncode(sty_clip_mod)&"&tai="&server.URLEncode(tai_clip_mod)&"&ext="&server.URLEncode(ext_clip_mod)&"&lie="&server.URLEncode(lie_clip_mod)&"&plus=true"
 
'Fermeture et Destruction RS & Conns
 RS.Close : Set RS = Nothing
 Conns.Close : Set Conns = Nothing
%> 