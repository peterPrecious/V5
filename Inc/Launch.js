
	// create a popup alert
	function popUpAlert () {
		if (parent.lang == "FR") {
			alert("Nous avons tenté de lancer votre cours dans une nouvelle fenêtre mais un pop-up blocker a empêché de s'ouvrir. S'il vous plaît désactiver pop-up blockers pour ce site.");
		} else if (parent.lang == "ES") {
			alert("Hemos tratado de poner en marcha su curso en una nueva ventana, pero un bloqueador de pop-up que impide la apertura. Por favor, desactive bloqueadores de ventanas emergentes para este sitio.");
		} else {
			alert("We attempted to launch your course in a new window but a pop-up blocker prevented it from opening. Please disable pop-up blockers for this site.");
		}	
	}

	function launchCert(vParms) {	  // this is called by any program to launch a certificate, with the parms encoded
    var certWindow = window.open('','Certificate','width='+screen.width+',height='+screen.height+',top=0,left=0,resizable=yes');
    var vHTML = ""
    vHTML += '			  <form method="POST" action="/certservice/default.aspx" id="formPost">'
    vHTML += '			    <input type="hidden" name="format" id="format" value="PDF">'
    vHTML += '			    <input type="hidden" name="vParms" id="vParms" value="' + vParms + '">'
    vHTML += '			  </form>'
    certWindow.document.write(vHTML);
    certWindow.document.getElementById("formPost").submit();
  }
  
	function fullScreen(vUrl) {
    if (parent.popupBlockerOn) {
      popUpAlert()
    } else {
      //  if we are just passing a ProgId|ModId (P1234EN|19876EN|Y|Y|Y) then launching a full screen popup for this content, 
      //  else this will be a complete URL
      if (vUrl.substring(0,1) == "P" && vUrl.length >= 14 && vUrl.length <= 21) {
        vUrl = "../LaunchObjects.asp?vModId=" + vUrl + "&vNext=CloseWindow.asp"
      }  
      var modwindow = window.open(vUrl,'FullSceen','width='+screen.width+',height='+screen.height+',top=0,left=0,resizable=yes');
	    try {top.addWindowToArray(modwindow);} catch (err) {}; 
	    modwindow.focus();
	    parent.bodyFocus = false;
	    parent.vModWindow = modwindow;
	    parent.vModWindowOpen = true;
	  }
  }

  function fmodulewindow(vModId) {
    if (parent.popupBlockerOn) {
      popUpAlert(); 
    } else {   
      var vmodule = "../LaunchObjects.asp?vModId=" + vModId
      var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=no,resizable=yes');
      try {top.addWindowToArray(modwindow);} catch (err) {}; 
      modwindow.focus();
	    parent.bodyFocus = false;
      parent.vModWindow = modwindow;
      parent.vModWindowOpen = true;
    }
  }

  function amodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=yes,resizable=yes');
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function assessmentwindow(vId) {
    var vmodule = "../Assessments/" + vId
    var modwindow = window.open(vmodule,'Assessment','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=no,resizable=yes');
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function fmodulestaging(vModId) {
    var vmodule = "../LaunchStaging.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function finfoarch(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=668,height=590,left=10,top=10,status=no,scrollbars=yes,resizable=yes').focus()
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function ccimodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=780,height=546,left=10,top=10,status=yes,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function ldmodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'ldModule','toolbar=no,width=800,height=600,left=10,top=10,status=yes,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }
 
  function lsmodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'lsModule','toolbar=no,width=700,height=480,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function zmodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function zmodulestaging(vModId) {
    var vmodule = "../LaunchStaging.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function cciciscomodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=780,height=540,left=25,top=25,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function zmodulebookmark(vModId,vpageno)  {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId + "&vPageNo=" + vpageno
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function testwindow(vModId) {
    var modwindow = window.open(vModId,'Module','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function examwindow(vModId) {
    var modwindow = window.open(vModId,'Assessment','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function SiteWindow(vurl) {
    var modwindow = window.open(vurl,'Site','toolbar=no,width=800,height=600,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function SiteWindow2(vurl) {
    var modwindow = window.open(vurl,'Site','toolbar=no,width=750,height=475,left=25,top=25,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function FlashIntro(vurl) {
    var modwindow = window.open(vurl,'Site','toolbar=no,width=750,height=275,left=25,top=100,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function Privacy(vurl) {
    var modwindow = window.open(vurl,'Privacy','toolbar=no,width=500,height=475,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function PublicSite(vurl) {
    var modwindow = window.open(vurl,'Public','toolbar=no,width=775,height=600,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function ebizwindow() {
    var modwindow = window.open('//vubiz.com/ebiz','Ebiz','toolbar=no,width=825,height=485,left=25,top=25,status=no,scrollbars=yes,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function ObjectWindow(vurl,vwindow) { 
    var modwindow = window.open(vurl,vwindow,'toolbar=no,width=500,height=300,left=25,top=25,status=no,scrollbars=yes,resizable=yes')   
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function programwindow(vprogram) {
    var vmodule = "ContentEdit.asp?vProgram=" + vprogram
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=400,height=300,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
  }

  function articles(vartsno) {
    var varticle
    varticle = "ArtsView.asp?vArts_No=" + vartsno
    var modwindow = window.open(varticle,'article','toolbar=no,width=700,height=600,left=100,top=100,status=no,scrollbars=yes,resizable=yes')
  }

  function CCOHSmodulewindow(vModId)  {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=790,height=575,left=10,top=10,status=no,scrollbars=no,resizable=no')  
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function AlignMediaNugget(vModId)   {
    var modwindow = window.open('../LaunchObjects.asp?vModId='+vModId,'eLearning','toolbar=no,width=640,height=515,left=10,top=10,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }
  
  function AlignMediaExercise(url, vModId)  {
    var url     = "../xModules/AlignMedia/exercises.htm?exercise=" + url + "&vModId=" + vModId
  	exercise = window.open(url,'eLearning','toolbar=no,width=657,height=515,left=10,top=10,status=no,scrollbars=no,resizable=no')
    top.addWindowToArray(exercise) 
    exercise.focus()
    parent.vModWindow = exercise 
    parent.vModWindowOpen = true;
  } 
  
  function xmodulewindow(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=790,height=515,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }
  
  function Pentad(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=800,height=625,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }  
 
  function vuwindow(vUrl,vWidth,vHeight,vLeft,vTop,vStatus,vScrollbars,vResizable) {
    var vStatus = "toolbar=no,width="+vWidth+",height="+vHeight+",left="+vLeft+",top="+vTop+",status="+vStatus+",scrollbars="+vScrollbars+",resizable="+vResizable
    var modwindow = window.open(vUrl,'Vubiz',vStatus)
		if (!modwindow) {
			popUpAlert(); 
		} else {
	    top.addWindowToArray(modwindow); 
	    modwindow.focus();
	    parent.vModWindow = modwindow;
	    parent.vModWindowOpen = true;
	  }
  }

  function W575x775(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=775,height=575,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W740x540(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule, 'Module', 'toolbar=no,width=740,height=540,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow);
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W545x745(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=745,height=545,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W545x905(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=805,height=545,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W1050x625(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=1050,height=625,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W1000x600(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule, 'Module', 'toolbar=no,width=1000,height=600,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow);
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function W805x640(vModId) {
    if (parent.popupBlockerOn) {
      popUpAlert();
    } else {
      var vmodule = "../LaunchObjects.asp?vModId=" + vModId
      var modwindow = window.open(vmodule, 'Module', 'toolbar=no,width=805,height=640,left=10,top=10,status=no,scrollbars=no,resizable=yes');
      try { top.addWindowToArray(modwindow); } catch (err) { };
      modwindow.focus();
      parent.bodyFocus = false;
      parent.vModWindow = modwindow;
      parent.vModWindowOpen = true;
    }
  }

  function Redwood(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=760,height=550,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  function RedwoodAssess(vModId) {
    var vmodule = "../LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=810,height=550,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    top.addWindowToArray(modwindow); 
    modwindow.focus();
    parent.vModWindow = modwindow;
    parent.vModWindowOpen = true;
  }

  //...this prints a platform certificate from My World Status Link, the Assessment Report and ExamGrade
  //...vLang no longer used
  function jCertificate(vLang, vCertId, vCertTitle, vCertDate, vCertMark, vCertType, vCertName) {
    var certForm = ''
    certForm += '<form method="POST" action="Certificate.asp" id="fCertificate">';
    certForm += '  <input type="hidden" name="vHidden"    value="y">';
    certForm += '  <input type="hidden" name="vModId"     value="' + vCertId    + '">';
    certForm += '  <input type="hidden" name="vMark"      value="' + vCertMark  + '">';
    certForm += '  <input type="hidden" name="vCertTitle" value="' + vCertTitle + '">';
    certForm += '  <input type="hidden" name="vCertType"  value="' + vCertType  + '">';
    certForm += '  <input type="hidden" name="vCertDate"  value="' + vCertDate  + '">';
    certForm += '  <input type="hidden" name="vCertName"  value="' + vCertName  + '">';
    certForm += '</form>';
    var modwindow = window.open('','Certificate','toolbar=no,width=650,height=425,left=100,top=100,status=no,scrollbars=no,resizable=no');
    modwindow.document.write (certForm)
    modwindow.fCertificate.submit()
  }

  //...this prints a VuAssess certificate from My World Status Link, the Assessment Report and ExamGrade
  function jVuCertificate(vCertId, vCertTitle, vCertDate, vCertMark, vCertFirstName, vCertLastName, vCertLogoUrl)  {
//  var vCertUrl = "//vubiz.com/v5/Assessments/components/V5CertMaker.html?vMods_Id=" + vCertId + "&vMods_Title=" + vCertTitle + "&vLastScore=" + vCertDate + "&vScore=" + vCertMark + "&vFirstName=" + vCertFirstName + "&vLastName=" + vCertLastName + "&vLogoUrl=//vubiz.com/v5/images/logos/" + vCertLogoUrl
    var vCertUrl = "/v5/Assessments/components/V5CertMaker.html?vMods_Id=" + vCertId + "&vMods_Title=" + vCertTitle + "&vLastScore=" + vCertDate + "&vScore=" + vCertMark + "&vFirstName=" + vCertFirstName + "&vLastName=" + vCertLastName + "&vLogoUrl=//vubiz.com/v5/images/logos/" + vCertLogoUrl
    var modwindow = window.open(vCertUrl,'Certificate','toolbar=no,width=750,height=450,left=10,top=10,status=no,scrollbars=no,resizable=yes')
//  top.addWindowToArray(modwindow); 
//  modwindow.focus();
//  parent.vModWindow = modwindow;
//  parent.vModWindowOpen = true;
  }

  //...this prints a Custom certificate from My World Status Link, the Assessment Report and ExamGrade
  function jCustomCertificate(vCertFolder, vCertId, vCertTitle, vCertDate, vCertMark, vCertFirstName, vCertLastName, vCertLogoUrl)  {
//  var vCertUrl = "//vubiz.com/v5/Assessments/CustomCerts/" + vCertFolder + "/Default.htm?vMods_Id=" + vCertId + "&vMods_Title=" + vCertTitle + "&vLastScore=" + vCertDate + "&vScore=" + vCertMark + "&vFirstName=" + vCertFirstName + "&vLastName=" + vCertLastName + "&vLogoUrl=//vubiz.com/v5/images/logos/" + vCertLogoUrl
    var vCertUrl = "/v5/Assessments/CustomCerts/" + vCertFolder + "/Default.htm?vMods_Id=" + vCertId + "&vMods_Title=" + vCertTitle + "&vLastScore=" + vCertDate + "&vScore=" + vCertMark + "&vFirstName=" + vCertFirstName + "&vLastName=" + vCertLastName + "&vLogoUrl=//vubiz.com/v5/images/logos/" + vCertLogoUrl
    var modwindow = window.open(vCertUrl,'Certificate','toolbar=no,width=750,height=450,left=10,top=10,status=no,scrollbars=no,resizable=yes')
//  top.addWindowToArray(modwindow); 
//  modwindow.focus();
//  parent.vModWindow = modwindow;
//  parent.vModWindowOpen = true;
  }

  //...this prints a Custom certificate from My World Status Link, the Assessment Report and ExamGrade
  function Certificate(vUrl) {
    var modwindow = window.open(vUrl,'Certificate','toolbar=no,width=750,height=450,left=10,top=10,status=no,scrollbars=yes,resizable=yes')
//  top.addWindowToArray(modwindow); 
//  modwindow.focus();
//  parent.vModWindow = modwindow;
//  parent.vModWindowOpen = true;
  }

  //...this prints a Custom certificate from My World Status Link, the Assessment Report and ExamGrade
  function jCert(vFolder, vId, vTitle, vDate, vMark, vFirstName, vLastName, vLogo) {
    var vUrl = "/V5/Assessments/CustomCerts/" + vFolder + "/Default.htm?vMods_Id=" + vId + "&vMods_Title=" + escape(vTitle) + "&vLastScore=" + escape(vDate) + "&vScore=" + vMark + "&vFirstName=" + escape(vFirstName) + "&vLastName=" + escape(vLastName) + "&logo=" + vLogo;
    var certwindow = window.open(vUrl,'Certificate','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=yes,resizable=yes');
  }
