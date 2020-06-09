/* these are also used on V5 and V7, modify on V5 and copy to V7 */

var reAlphaNumeric = new RegExp(/^[0-9A-z]+$/);
var reAlpha = new RegExp( /^[A-z]+$/ );
var reNumeric = new RegExp( /^[0-9]+$/ ); // this only allows pure integers, ie NOT negative numbers or floatint values
var rePassword = new RegExp(/^[0-9A-z\!\@\$\%\^\*\(\)\_\+\-\{\}\[\]\;\<\>\,\.\:]+$/);
var reEmail = new RegExp( /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,6})+$/ ); // basic email edit 

function isWhitespace(charToCheck) {
  var whitespaceChars = " \t\n\r\f";
  return (whitespaceChars.indexOf(charToCheck) !== -1);
}
function ltrim(str) {
  for (var k = 0; k < str.length && isWhitespace(str.charAt(k)) ; k++);
  return str.substring(k, str.length);
}
function rtrim(str) {
  for (var j = str.length - 1; j >= 0 && isWhitespace(str.charAt(j)) ; j--);
  return str.substring(0, j + 1);
}
function trim(str) {
  return ltrim(rtrim(str));
}
function left(str, n) {
  if (n <= 0)
    return "";
  else if (n > String(str).length)
    return str;
  else
    return String(str).substring(0, n);
}
function right(str, n) {
  if (n <= 0)
    return "";
  else if (n > String(str).length)
    return str;
  else {
    var iLen = String(str).length;
    return String(str).substring(iLen, iLen - n);
  }
}

/* end of V5/V7 functions */

function isNumber( n ) { // returns true if numeric
	return !isNaN( parseFloat( n ) ) && isFinite( n );
}
function fMax( lang, theElement, maxLength ) {
	if ( theElement.value.length > maxLength ) {
		var message = "You cannot enter more than " + maxLength + " characters."
		if ( lang === "FR" ) message = "Vous ne pouvez pas entrer plus de " + maxLength + " caractères.";
		if ( lang === "ES" ) message = "Usted no puede entrar más de " + maxLength + " caracteres.";
		alert( message );
		theElement.value = ( theElement.value.substr( 0, maxLength ) );
		return false;
	}
	return true;
}
function fMin( lang, theElement, minLength ) {
	if ( theElement.value.length < minLength ) {
		var message = "You must enter at least " + minLength + " characters."
		if ( lang === "FR" ) message = "Vous devez entrer au moinse " + minLength + " caractères.";
		if ( lang === "ES" ) message = "Debe introducir al menos " + minLength + " caracteres.";
		alert( message );
		theElement.focus;
		return false;
	}
	return true;
}
function fMinO( lang, theElement, minLength ) {// if you enter anything then it must be at least minLength characters

	if ( theElement.value.length < minLength && theElement.value.length > 0 ) {
		var message = "You must enter at least " + minLength + " characters."
		if ( lang === "FR" ) message = "Vous devez entrer au moinse " + minLength + " caractères.";
		if ( lang === "ES" ) message = "Debe introducir al menos " + minLength + " caracteres.";
		alert( message );
		theElement.focus;
		return false;
	}
	return true;
}
function fCap( theElement ) {
	theElement.value = theElement.value.toUpperCase()
	return true;
}
function jconfirm( url, msg ) {
	if ( confirm( msg ) ) {
		location = url
	}
}
function bconfirm( msg ) {
	var ok = confirm( msg );
	if ( ok )
		return true;
	else
		return false;
}
function toggle( theDiv ) {
	var divStyle = document.getElementById( theDiv ).style;
	if ( divStyle.display !== 'block' ) {
		divStyle.display = 'block';
	}
	else {
		divStyle.display = 'none';
	}
}
function divOn( theDiv ) {
	var theElement = document.getElementById( theDiv );
	if ( theElement !== null ) theElement.style.display = "block";
}
function divOff( theDiv ) {
	var theElement = document.getElementById( theDiv );
	if ( theElement !== null ) theElement.style.display = "none";
}
function disable( theElement ) {
	document.getElementById( theElement ).disabled = true;
}
function enable( theElement ) {
	document.getElementById( theElement ).disabled = false;
}
function jPrint() {
	// hide any buttons
	var theButtons = document.getElementsByTagName( "input" );
	for ( var i = 0; i < theButtons.length; i++ ) {
		if ( theButtons[i].type === "button" || theButtons[i].type === "submit" ) {
			theButtons[i].style.visibility = 'hidden';
		}
	}
	// enable any disabled fields so they will print better
	document.getElementsByTagName( "body" )[0].innerHTML = document.getElementsByTagName( "body" )[0].innerHTML.replace( /disabled/g, "pooh-disabled" );
	// convert any textareas to a div so the entire text will be displayed (IE converts all tags to upper case, FF is lower)
	document.getElementsByTagName( "body" )[0].innerHTML = document.getElementsByTagName( "body" )[0].innerHTML.replace( /textarea/g, "pooh:div" );
	document.getElementsByTagName( "body" )[0].innerHTML = document.getElementsByTagName( "body" )[0].innerHTML.replace( /TEXTAREA/g, "pooh:div" );
	// print the page
	parent.main.print();
	// return the div to the original textarea
	document.getElementsByTagName( "body" )[0].innerHTML = document.getElementsByTagName( "body" )[0].innerHTML.replace( /pooh:div/g, "textarea" );
	// re-disable any disabled fields so they will print better
	document.getElementsByTagName( "body" )[0].innerHTML = document.getElementsByTagName( "body" )[0].innerHTML.replace( /pooh-disabled/g, "disabled" );
	// return any buttons
	theButtons = document.getElementsByTagName( "input" );
	for ( i = 0; i < theButtons.length; i++ ) {
		if ( theButtons[i].type === "button" || theButtons[i].type === "submit" ) {
			theButtons[i].style.visibility = 'visible';
		}
	}
}
function hideElement( theElementId ) {
	var theElement = document.getElementById( theElementId );
	if ( theElement !== null ) theElement.style.visibility = 'hidden';
}
function showElement( theElementId ) {
	var theElement = document.getElementById( theElementId );
	if ( theElement !== null ) theElement.style.visibility = 'visible';
}
function openDivs( divId ) {
	var divs = document.getElementsByTagName( "div" );
	var j = divs.length;
	for ( i = 0; i < j; i++ ) {
		if ( divs[i].id.substring( 0, divId.length ) === divId ) {
			document.getElementById( divs[i].id ).style.display = "block";
		}
	}
}
function hideDivs( divId ) {
	var divs = document.getElementsByTagName( "div" );
	var j = divs.length;
	for ( i = 0; i < j; i++ ) {
		if ( divs[i].id.substring( 0, divId.length ) === divId ) {
			document.getElementById( divs[i].id ).style.display = "none";
		}
	}
}
function emptyField( theElement ) {
	document.getElementById( theElement ).value = "";
}
function fillField( theElement, theValue ) {
	document.getElementById( theElement ).value = theValue;
}
function refillField( theElement, theValue ) {
	if ( document.getElementById( theElement ).value === "" ) {
		fillField( theElement, theValue )
	}
}
function WebService(vUrl, vMsg) {
	var agt = navigator.userAgent.toLowerCase();
	var ie = ( agt.indexOf( "msie" ) !== -1 );
	if ( ie )
		oXmlHttp = new ActiveXObject( "Microsoft.XMLHTTP" );
	else
		oXmlHttp = new XMLHttpRequest();
	try {
		oXmlHttp.open( "POST", vUrl, false );
		oXmlHttp.setRequestHeader( "Content-Type", "application/x-www-form-urlencoded" );
		oXmlHttp.send( vMsg );
		return oXmlHttp.responseText;
	}
	catch ( err ) {
		alert( err );
		return "error using web service";
	}
}
function jsonWebService( vUrl, vMsg ) {
	var agt = navigator.userAgent.toLowerCase();
	var ie = ( agt.indexOf( "msie" ) !== -1 );
	if ( ie )
		oXmlHttp = new ActiveXObject( "Microsoft.XMLHTTP" );
	else
		oXmlHttp = new XMLHttpRequest();
	try {
		oXmlHttp.open( "POST", vUrl, false );
		oXmlHttp.setRequestHeader( "Content-Type", "application/json; charset=utf-8" );
		oXmlHttp.send( vMsg );
		return oXmlHttp.responseText;
	}
	catch ( err ) {
		alert( err );
		return "error using json web service";
	}
}
function renderInfo( theElementNo ) {//   this is used to render an info item, it displays the appropriate div near the info (exclamation) mark 

	// this defines the DIV that contains the message
	document.getElementById( 'div_' + theElementNo ).style.padding = "5px";
	document.getElementById( 'div_' + theElementNo ).style.color = "#008000";
	document.getElementById( 'div_' + theElementNo ).style.borderStyle = "solid";
	document.getElementById( 'div_' + theElementNo ).style.borderWidth = "1px";
	document.getElementById( 'div_' + theElementNo ).style.borderColor = "#008000";
	document.getElementById( 'div_' + theElementNo ).style.position = "absolute";
	document.getElementById( 'div_' + theElementNo ).style.width = "150px";
	document.getElementById( 'div_' + theElementNo ).style.backgroundColor = "#D7FFD7";
	document.getElementById( 'div_' + theElementNo ).style.textAlign = "left";


	// display the div so we can get the size
	toggle( 'div_' + theElementNo );

	// determine if we need to position the DIV to the left or right of the info mark (ie the mouse)
	var xPlus = 10;
	if ( ( document.body.scrollWidth - window.event.clientX ) < 200 ) {
		xPlus = -150 - 10;
	}

	// determine if we need to position the DIV to the left or right of the info mark (ie the mouse)
	var yPlus = 10;
	//  if ((document.body.scrollHeight - window.event.clientY) < document.getElementById('div_' + theElementNo).clientHeight) { 
	//    yPlus = - document.getElementById('div_' + theElementNo).clientHeight - 10;
	//  }

	//  alert('Screen Height: ' + document.body.scrollHeight + '\n Mouse Height: ' + window.event.clientY);

	// rendered the DIV
	document.getElementById( 'div_' + theElementNo ).style.left = window.event.clientX + xPlus;
	document.getElementById( 'div_' + theElementNo ).style.top = window.event.clientY + yPlus;

}
function idOk( vId ) {//  this is used on landing pages or anywhere that we want to ensure that the ID/Password is valid
	var charsOk = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_@.";
	var allValid = true;
	for ( i = 0; i < vId.length; i++ ) {
		ch = vId.charAt( i );
		for ( j = 0; j < charsOk.length; j++ ) {
			if ( ch === charsOk.charAt( j ) ) break;
		}
		if ( j === charsOk.length ) {
			allValid = false;
			break;
		}
	}
	return allValid;
}
function getParameter(name) {//  this returns a querystring value and strips off any "+" representing spaces  
	var pair = location.search.substring( 1 ).split( "&" );
	for ( var i = 0; i < pair.length; i++ ) {
		var a = pair[i].split( "=" );
		i, n = "", v = "";
		if ( a.length > 0 ) {
			n = a[0];
			if ( n === name ) {
				if ( a.length > 1 ) {
					v = unescape( a[1] );
					for ( i = 0; i < v.length; i++ ) {
						v = v.replace( '+', ' ' );
					}
					return v;
				}
			}
		}
	}
}
function jYN( i, lang ) {// turns y/n into language values
	var yn;
	switch ( lang ) {
		case "ES": yn = ( i === "y" ) ? "si" : "¡no"; break;
		case "FR": yn = ( i === "y" ) ? "oui" : "non"; break;
		default: yn = ( i === "y" ) ? "Yes" : "No"; break;
	}
	return yn
}
function isDate( dateStr, lang ) {// month must be either empty or formatted as per v5 fFormatDate routines and be between 2000 and current year plus 1

	var today = new Date();
	var year = today.getFullYear();
	dateStr = dateStr.replace( " 0", " " );
	dateStr = dateStr.replace( ", ", " " );
	var datePts = dateStr.split( " " );
	switch ( lang ) {
		case "FR":
			mmm = "janv.   févr.   mars    avril   mai     juin    juillet août    sept.   oct.    nov.    déc.";
			mm = mmm.indexOf( datePts[1] ) / 8 + 1;
			dd = isNumber( datePts[0] ) ? parseInt( datePts[0] ) : 0;
			yy = isNumber( datePts[2] ) ? parseInt( datePts[2] ) : 0;
			break;
		case "ES":
			mmm = "ene.    feb.    mar.    abr.    may.    jun.    jul.    ago.    sept.   oct.    nov.    dic.";
			mm = mmm.indexOf( datePts[1] ) / 8 + 1;
			dd = isNumber( datePts[0] ) ? parseInt( datePts[0] ) : 0;
			yy = isNumber( datePts[2] ) ? parseInt( datePts[2] ) : 0;
			break;
		default:
			mmm = "Jan     Feb     Mar     Apr     May     Jun     Jul     Aug     Sep     Oct     Nov     Dec";
			mm = mmm.indexOf( datePts[0] ) / 8 + 1;
			dd = isNumber( datePts[1] ) ? parseInt( datePts[1] ) : 0;
			yy = isNumber( datePts[2] ) ? parseInt( datePts[2] ) : 0;
			break;
	}
	if ( yy < 2000 || yy > year + 1 ) return false;
	if ( mm < 1 || mm > 12 ) return false;
	if ( dd < 1 || dd > 31 ) return false;
	if ( ( mm === 4 || mm === 6 || mm === 9 || mm === 11 ) && dd === 31 ) return false;
	if ( mm === 2 ) {
		var isleap = ( year % 4 === 0 && ( year % 100 !== 0 || year % 400 === 0 ) );
		if ( dd > 29 || ( day === 29 && !isleap ) ) {
			return false;
		}
	}
	return true;
}
function jSubmitPlus( formId, hideId, showId ) {// this will submit the form but first hide the submit button and show the Progress Bar
	$( "#" + hideId ).hide();
	$( "#" + showId ).show();
	$( "#" + formId ).submit();
}

if (typeof $ !== "undefined") {
  $(document).ready(function () { // this is used to highlight a button (if jQuery is loaded)
    var elements = ".button, .button040, .button070, .button085, .button100, .button150, .button200, .shite";
    $(elements).bind('mouseenter', function () {
      this.style.backgroundColor = "black"
      this.style.border = "1px solid #000000";
      this.style.color = "white"
    });
    $(elements).bind('mouseleave', function () {
      this.style.backgroundColor = "white"
      this.style.border = "1px solid navy";
      this.style.color = "navy";
    });
  });
}
