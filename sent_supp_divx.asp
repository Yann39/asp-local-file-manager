<%
'Déclaration Variables
 dim tit_divx_supp, SQLs, Conns, RS
 tit_divx_supp = Request.QueryString("titre")

'Création de la connection
 Set Conns = Server.CreateObject ("ADODB.Connection")

'Ouverture de la connection
 Conns.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " & "DBQ=" & Server.MapPath("bd1.mdb") & ";"

'Requête SQL pour séléctionner les clips dont l'auteur ou le titre correspond au valeurs du formulaire
 SQLs = "SELECT * FROM [Divx] where (Titre = '"&tit_divx_supp&"')"

'Création de l'objet Serveur
 Set RS = Server.CreateObject ("ADODB.RecordSet")

'Ouverture RS, SQLs, Conns
 RS.Open SQLs, Conns, 1, 2, 1

'On efface l'enregistrement
 RS.delete

'On affiche un message comme quoi ca s'est bien passé
 Response.Redirect "divx.asp?sent=ok_suppr&plus=true"

'Fermeture et Destruction RS & Conns
 RS.Close : Set RS = Nothing
 Conns.Close : Set Conns = Nothing
%> 