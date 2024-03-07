<%
'Déclaration Variables
 dim tit_divx_modif, SQLs, Conns, RS
 tit_divx_modif = Request.QueryString("titre")

'Création de la connection
 Set Conns = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Conns.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

' Requête SQL pour séléctionner les divx ayant l'auteur demandé
 SQLs = "SELECT * FROM [Divx] where (Titre = '"&tit_divx_modif&"')"

'Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLs, Conns
 RS.Open SQLs, Conns, 1, 2, 1

' déclaration des variables
 dim tit_divx_mod, dur_divx_mod, gen_divx_mod, tai_divx_mod, ext_divx_mod, dat_divx_mod, lie_divx_mod
 
'récupération des valeurs dans la base
 tit_divx_mod = RS.Fields ("Titre").Value 
 dur_divx_mod = RS.Fields ("Durée").Value 
 gen_divx_mod = RS.Fields ("Genre").Value 
 tai_divx_mod = RS.Fields ("Taille").Value
 ext_divx_mod = RS.Fields ("Extension").Value 
 dat_divx_mod = RS.Fields ("Date").Value
 lie_divx_mod = RS.Fields ("Lien").Value 

'on efface l'ancien enregistrement
 'RS.delete

'On redirectionne vers l'adresse avec les parametres de la musique afin de remplir automatiquement le formulaire
 Response.Redirect "divx.asp?tit="&server.URLEncode(tit_divx_mod)&"&dur="&server.URLEncode(dur_divx_mod)&"&gen="&server.URLEncode(gen_divx_mod)&"&tai="&server.URLEncode(tai_divx_mod)&"&dat="&server.URLEncode(dat_divx_mod)&"&ext="&server.URLEncode(ext_divx_mod)&"&lie="&server.URLEncode(lie_divx_mod)&"&plus=true"
 
'Fermeture et Destruction RS & Conns
 RS.Close : Set RS = Nothing
 Conns.Close : Set Conns = Nothing
%> 