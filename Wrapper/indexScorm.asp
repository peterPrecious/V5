  <script>
    var objAPI;
    var SCORM_FALSE = "false";
    var blnFoundAPI = false;
    var blnCalledFinish = false;
    var blnCalledComplete = false;
    var blnStandAlone = false;
    var findAPITries = 0;
    var vDeployment = 'scorm';
    var vlesson_location, vstudent_name;
    var vFirstName, vLastName, vScore, vPass;
  

    // Parses the querystring
    function getParameter(name) {
      var pair=location.search.substring(1).split("&");    
      for (var i = 0; i < pair.length; i++) {
        var a = pair[i].split("=");
        var n="",v="";
        if (a.length > 0)
          n = unescape(a[0]);
        if (a.length > 1)
          v = unescape(a[1]);
        if (n.toLowerCase() == name.toLowerCase()) return v;
      }
      return null;       
    } 

    //Some LMSs give the name in all caps, make title case
    function TitleCase(str){ 
      var first; 
      var the_rest; 
      var stringArray; 
      var newString;       
      stringArray = str.split(" "); 
      for (var i=0; i < stringArray.length; i++) { 
        first = stringArray[i].charAt(0); 
        the_rest = stringArray[i].substring(1); 
        if (i == 0) { 
          newString = ""; 
        } 
        newString = newString + first.toUpperCase() + the_rest.toLowerCase() + " "; 
      } 
      return newString; 
    } 
    
    function SCORM_Start(){    
    	var strStandAlone;
    	var strResult;    	
    	strStandAlone = getParameter("StandAlone");
    	if (ConvertStringToBoolean(strStandAlone)){
    		blnStandAlone = true;
      }	
    	
    	if (! blnStandAlone){
    		objAPI = getAPI();
    		if ((objAPI == null) || (typeof(objAPI) == "undefined")){
          vDeployment = '3rdParty';
    			alert("SCORM API could not be found. Results may not record properly.");
    			return false;
    		}
    		else{
    			blnFoundAPI = true;
    		}
    	
    		strResult = objAPI.LMSInitialize("");
          startTimer();
          vlesson_location = objAPI.LMSGetValue("cmi.core.lesson_location");
          if(objAPI.LMSGetLastError() != '0'){
            alert('Could not retrieve bookmark from LMS');
          }             
          vstudent_name    = objAPI.LMSGetValue("cmi.core.student_name");
          vstudent_name    = TitleCase(vstudent_name);
          if(objAPI.LMSGetLastError() != '0'){
            alert('Could not retrieve name from LMS');
          }             
    
    		//MOE
    		/*
    		if(objAPI.LMSGetValue("cmi.core.lesson_status") == "not attempted"){  
          objAPI.LMSSetValue("cmi.core.lesson_status", "incomplete");
    		}
    		else if(objAPI.LMSGetValue("cmi.core.lesson_status") == "completed"){  
          alert('This module has already been completed, no more results will stored.');
    			SCORM_Complete();
    		}
    		*/
    		//End MOE
    
        var vCommaLoc, vLength;
        
        if(vstudent_name != ''){
          vLength        = vstudent_name.length;   
          vCommaLoc      = vstudent_name.indexOf(',');
          vFirstName     = vstudent_name.substring(vCommaLoc+2, vLength);
          vLastName      = vstudent_name.substring(0, vCommaLoc);
        }
    
    		if (strResult == SCORM_FALSE){
    			alert("Content failed to establish communication with the LMS. Results may not record properly.");
    			return false;
    		}
    	}
    	
    	return true;
    }
    	
    	
    function SCORM_Finish(){
      if ((! blnStandAlone) && (blnFoundAPI) && (!blnCalledFinish)){
        blnCalledFinish = true;
        var myTime = computeTime(); 
        objAPI.LMSSetValue("cmi.core.session_time",myTime ); 
        if(objAPI.LMSGetLastError() != '0'){
          alert('Could not set session time to LMS');
        }             
        objAPI.LMSCommit("");
        if(objAPI.LMSGetLastError() != '0'){
          alert('Could not commit to LMS');
        }             
        strResult = objAPI.LMSFinish("");
        if (strResult == SCORM_FALSE){
          alert("Content failed to communicate your results to the LMS.");
          return false;
        }
      }
      return true;
    }
    
    
    function SCORM_Complete(){
      if ((!blnStandAlone) && (blnFoundAPI) && (!blnCalledFinish) && (!blnCalledComplete)){
        objAPI.LMSSetValue("cmi.core.lesson_status", "completed");
        objAPI.LMSSetValue("cmi.core.exit",""); 
        if(objAPI.LMSGetLastError() != '0'){
          alert('Could not set session as complete.');
        }             
        else{
          blnCalledComplete = true;
        }  
        objAPI.LMSCommit("");
        if(objAPI.LMSGetLastError() != '0'){
          alert('Could not commit to LMS');
        }             
      }
    
      SCORM_Finish()
      //return true;
    }  
    
    function ConvertStringToBoolean(str){    	        
    	var intTemp;   	        
    	if (EqualsIgnoreCase(str, "true") || EqualsIgnoreCase(str, "t")){
    		return true;   
    	}
    	else{
    		intTemp = parseInt(str);
    		if (intTemp == 1){
    		  return true;
    		}
    		else{
    		  return false;
    		}
    	}
    }    
    
    
    function EqualsIgnoreCase(str1, str2){
    	str1 = new String(str1);
    	str2 = new String(str2);    	
    	return (str1.toLowerCase() == str2.toLowerCase());
    }
    
    
    function findAPI(win){
      while ((win.API == null) && (win.parent != null)){    
        findAPITries++;
        // Note: 7 is an arbitrary number, but should be more than sufficient
        if (findAPITries > 7) {
          return null;
        }     
        win = win.parent;
      }
      return win.API;
    }
    
    function getAPI() {
      var theAPI = findAPI(window);
      if ((theAPI == null) && (window.opener != null) && (typeof(window.opener) != "undefined")) {
        theAPI = findAPI(window.opener);
      }
      return theAPI;
    }
    
    function SCORM_setBookmark(pg) {
      objAPI.LMSSetValue("cmi.core.lesson_location", pg)
    }
    
    function SCORM_setScore(scr) {
    
    alert("setting score... " + scr);
    
    	if(objAPI.LMSGetValue("cmi.core.lesson_status") == "completed"){
    		alert('The results of this assessment attempt have already been recorded.');   
    	}
    	else {   
    	  var scrArray   = scr.split(",");
    	  vScore         = parseInt(scrArray[0]);



    	  vPass          = parseInt(scrArray[1]);



    	  vActualAnswers = scrArray[2];
    	  aActualAnswers = vActualAnswers.split("|");
      	//  alert(vPass);
      	//  alert(vScore);
      	//  alert(vActualAnswers);
        var answerComponents;
        for (key in aActualAnswers) {
          // seperate question identification number and users's selected answer
          answerComponents = aActualAnswers[key].split(':');          
          // if not answered (ie "nil") then set to -1  
          if (isNaN(answerComponents[1])) {
//          alert(answerComponents[1]);
            answerComponents[1] = -1;
          }         
          
          // store interaction element id, type and student_response
          objAPI.LMSSetValue("cmi.interactions." + key + ".id", answerComponents[0]);
          objAPI.LMSSetValue("cmi.interactions." + key + ".type", "choice");
          objAPI.LMSSetValue("cmi.interactions." + key + ".student_response", answerComponents[1]);
				}
    
    	  //MOE
    	  if(vScore >= vPass) {
    	    objAPI.LMSSetValue("cmi.core.score.raw",vScore);
    	    // first call overwrites the second, so only one is necessary
    	    //objAPI.LMSSetValue("cmi.core.lesson_status", "passed");
    	    objAPI.LMSSetValue("cmi.core.lesson_status", "completed");
    	  }
    	  else {
    	    objAPI.LMSSetValue("cmi.core.score.raw", vScore); 
    	    // first call overwrites the second, so only one is necessary
    	    //objAPI.LMSSetValue("cmi.core.lesson_status", "failed");
    	    //objAPI.LMSSetValue("cmi.core.lesson_status", "completed");
    	  }
    	}
    }
    
    function SCORM_initTest() {
    	//MOE
    	if(objAPI.LMSGetValue("cmi.core.lesson_status") == "completed"){
    		alert('The results of this assessment attempt have already been recorded.');   
    	}
    	// only need to set lesson_status to incomplete if for some reason it is set at not attempted
    	else if(objAPI.LMSGetValue("cmi.core.lesson_status") == "not attempted"){
    		objAPI.LMSSetValue("cmi.core.lesson_status", "incomplete");
    	}
    }
    
    function startTimer() {
      startDate = new Date().getTime();
    }
    
    function computeTime() {
      var formattedTime = "00:00:00.0";
      if ( startDate != 0 ){
        var currentDate = new Date().getTime();
        var elapsedSeconds = ( (currentDate - startDate) / 1000 );
        formattedTime = convertTotalSeconds( elapsedSeconds );
      }
      return formattedTime;		
    } 
    
    function convertTotalSeconds(ts) {	
      var Sec = (ts % 60);
      ts -= Sec;
      var tmp = (ts % 3600);  //# of seconds in the total # of minutes
      ts -= tmp;              //# of seconds in the total # of hours
      if ( (ts % 3600) != 0 ) var Hour = "00" ;
      else var Hour = ""+ (ts / 3600);
      if ( (tmp % 60) != 0 ) var Min = "00";
      else var Min = ""+(tmp / 60);
      Sec=""+Sec;
      Sec=Sec.substring(0,Sec.indexOf("."));
      if (Hour.length < 2)Hour = "0"+Hour;
      if (Min.length < 2)Min = "0"+Min;
      if (Sec.length <2)Sec = "0"+Sec;
      var rtnVal = Hour+":"+Min+":"+Sec;
      return rtnVal;
    }
    
    //Run this here instead of onload to get the params to pass to the flash player
    SCORM_Start();

    var vUrl = "/V5/fModules/" + getParameter("vModId") + "/Default.asp" + location.href.substring(location.href.indexOf("?"));
//  alert("indexScorm vUrl: " + vUrl);

    document.write   ('<html>');
    document.write   ('  <head>');
    document.write   ('    <title>' + getParameter("vTitle") + '</title>');
    document.write   ('  </head>');
    document.write   ('  <frameset rows="100%" onbeforeunload="SCORM_Finish()" onunload="SCORM_Finish()">');
    document.write   ('    <frame name="content" src="' + vUrl  + '" frameborder="no" scrolling="yes" noresize marginwidth="0" marginheight="0">');
    document.write   ('  </frameset>');
    document.write   ('</html>');

  </script>