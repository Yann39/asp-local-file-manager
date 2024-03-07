//==============================================================
// Fonction pour cacher ou rendre visible des objets ou du texte
//==============================================================
function visibilite(thingId)
{
	var targetElement;
	targetElement = document.getElementById(thingId) ;
	if (targetElement.style.display == "none")
	{
		targetElement.style.display = "" ;
	} 
	else 
	{
		targetElement.style.display = "none" ;
	}
}
	
//=================================================================
// Fonction pour rendre visible tout objets ou texte contenant l'id
//=================================================================
function openall(thingId)
{
	targetElement = document.getElementById(thingId) ;
	targetElement.style.display = "" ;
}
	
//=========================================================
// Fonction pour cacher tout objets ou texte contenant l'id
//=========================================================
function closeall(thingId)
{
	targetElement = document.getElementById(thingId) ;
	targetElement.style.display = "none" ;
}

//==================================================
// Fonction pour mettre le site en page de d�marrage
//==================================================	
function HomePage(obj)
{
	obj.style.behavior='url(#default#homepage)';
	obj.setHomePage('http://localhost/X/mp3.asp');
}
	
//==============================================
// Fonction pour mettre le site dans les favoris
//==============================================	
function favoris()
{
	browserName = navigator.appName;
	browserVer = parseInt(navigator.appVersion);
	if (browserName == "Microsoft Internet Explorer" & browserVer >= 4) 
	{
		window.external.AddFavorite('http://localhost/X/mp3.asp', 'Webby.free.fr');
	}
} 

//=======================================================
// Fonction pour recherche des mots dans la page courante
//=======================================================		
var n = 0;
function findInPage(str)
{
	var txt, i, found;
	if (str == "") return false;
	if (document.layers) {
		if (!window.find(str))
		    while (window.find(str, false, true))
		        n++;
		else n++;
		if (n == 0)
			alert("Not found.");
	}
	if (document.all) {
		txt = window.document.body.createTextRange();
		for (i = 0; i <= n & (found = txt.findText(str)) != false; i++) {
			txt.moveStart("character", 1);
			txt.moveEnd("textedit");
		}
		if (found) {
			txt.moveStart("character", -1);
			txt.findText(str);
			txt.select();
			txt.scrollIntoView();
			n++;
		}
		else {
			if (n > 0) {
				n = 0;
				findInPage(str);
			}
			else
				alert("Not found.");
		}
	}
	return false;
}

// -----------------------------------------------------------------
// pour changer les petites images de dossiers + en - et inversement
// -----------------------------------------------------------------
function changerimg(thingId)
{
	var target;
	target = document.getElementById(thingId);
	if (target.src == "http://localhost/X/images/moins.jpg") {
	    target.src = "http://localhost/X/images/plus.jpg";}
	else {
	    target.src = "http://localhost/X/images/moins.jpg";
	}
}
  
//===============================================================================
// Fonctions pour changer les couleur des lignes du tableau quand on passe dessus
//===============================================================================	
function changeCouleur(ligne)
{
	ligne.bgColor = '#CED3FF';
}

function remetCouleur(ligne)
{
	ligne.bgColor = '';
}  
  
//===================================================
// Fonction pour afficher l'heure (horloge dynamique)
//===================================================
function HorlogeDynamique() {
	var DateActuel = new Date();
	var heure = DateActuel.getHours();
	var minutes = DateActuel.getMinutes();
	var secondes = DateActuel.getSeconds();

	if (heure == 0) {
		heure = "0" + heure;
	}
	if (minutes <= 9) {
		minutes = "0" + minutes;
	}
	if (secondes <= 9) {
		secondes = "0" + secondes;
	}
	
	Horloge = "<b>"+ heure + ":" + minutes + ":" + secondes + "</b>";

	if (document.getElementById) {
		document.getElementById("clock").innerHTML = Horloge;
	}

	if (document.layers) { 
		document.clock.document.write("<br>&nbsp;&nbsp;"+Horloge); 
		document.clock.document.close(); 
	}

	if ((document.all)&&(!document.getElementById)) { 
		document.all["clock"].innerHTML = Horloge;
	}

	setTimeout("HorlogeDynamique()", 1000)
}
	
window.onload = HorlogeDynamique;

//==================================================
// Fonction pour les liens dans la liste d�roulante
//==================================================
function ChangeUrl(option) 
{ 
	location.href = option.value; 
} 

//=======================================================
// Fonction pour les changer la couleur de l'arri�re plan
//=======================================================
function ChangeArrierePlan(id, option)
{
	id.background = option.value;
}

//======================================
// Fonctions pour afficher les infobulles
//======================================
function GetId(id)
{
	return document.getElementById(id);
}

var i=false; // La variable i nous dit si la bulle est visible ou non

function move(e) {
	if(i) {  // Si la bulle est visible, on calcul en temps reel sa position ideale
		if (navigator.appName!="Microsoft Internet Explorer") { // Si on est pas sous IE
			GetId("curseur").style.left=e.pageX + 5+"px";
			GetId("curseur").style.top=e.pageY + 10+"px";
		}
		else {
			GetId("curseur").style.left=window.event.x + 5+"px";
			GetId("curseur").style.top=window.event.y + 10 + document.body.scrollTop+"px"; // Sous IE, voici un petit hack pour que lors du scroll la position reste bonne !
		}
	}
}

function montre(text) {
	if (i==false) {
		GetId("curseur").style.visibility="visible"; // Si il est cacher (la verif n'est qu'une securit�) on le rend visible.
		GetId("curseur").innerHTML = text; // Cette fonction est a am�liorer, il parait qu'elle n'est pas valide (mais elle marche)
		i=true;
	}
}

function cache() {
	if(i==true) {
		GetId("curseur").style.visibility="hidden"; // Si la bulle était visible on la cache
		i=false;
	}
}
document.onmousemove=move; // des que la souris bouge, on appelle la fonction move pour mettre a jour la position de la bulle.