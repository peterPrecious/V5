
  //set up handle for API 
  var adl_API = null;
  
  // getAPI, which calls findAPI as needed
  function getAPI() {
    var myAPI = null;
    var tries = 0, triesMax = 10;

    while (tries < triesMax && myAPI == null) {
      window.status = 'Looking for API object ' + tries + '/' + triesMax;
  
      myAPI = findAPI(window);
  	
      if (myAPI == null && typeof(window.parent) != 'undefined') myAPI = findAPI(window.parent)
      if (myAPI == null && typeof(window.top)    != 'undefined') myAPI = findAPI(window.top);
      if (myAPI == null && typeof(window.opener) != 'undefined') if (window.opener != null && !window.opener.closed) myAPI = findAPI(window.opener);
      tries++;
    }

    if (myAPI == null)
      window.status = 'LMS not found';
    else {
      adl_API = myAPI;
      window.status = 'LMS Initialized';
      adl_API.LMSInitialize("");
      adlOnload();
    }
  }
  
  function findAPI(win) {
    // look in this window
    if (typeof(win) != 'undefined' ? typeof(win.API) != 'undefined' : false)
      if (win.API != null )  return win.API;
  
    // look in this window's frameset kin (except opener)
    if (win.frames.length > 0) {
      for (var i = 0 ; i < win.frames.length ; i++) {
        if (typeof(win.frames[i]) != 'undefined' ? typeof(win.frames[i].API) != 'undefined' : false)
          if (win.frames[i].API != null)  return win.frames[i].API;
      }
    }
    return null;
  }
  
  function adlOnload() {
    if(adl_API != null) {
      var cur_status = adl_API.LMSGetValue("cmi.core.lesson_status");
      if(cur_status == "not attempted" || cur_status == "" || cur_status == null) {
        adl_API.LMSSetValue("cmi.core.lesson_status", "incomplete");
        adl_API.LMSCommit("");
      }
      else {
        bm = adl_API.LMSGetValue("cmi.core.lesson_location");
        if(bm != 0)
          vBookmarkPage = bm;
        else
          vBookmarkPage = 0;
      }
    }
  }
  
  function set_bm(bm) {
    if (vDebug_All) alert("...setting bookmark");
    if(adl_API != null) {
      adl_API.LMSSetValue("cmi.core.lesson_location", bm);
      adl_API.LMSCommit("");
    }
  }
  
  
  // this should be set by the content (and thus not used in this wrapper)
  function set_score(vScore) {
    if (vDebug_All) alert("...setting score");
    if(adl_API != null) {
      adl_API.LMSSetValue("cmi.core.score.raw", vScore);
      adl_API.LMSCommit("");
    }
  }
  
  function set_complete() {
    if (vDebug_All) alert("...setting status");
    if(adl_API != null) {
      adl_API.LMSSetValue("cmi.core.lesson_status", "completed");
      adl_API.LMSCommit("");
    }	
  }
  
  function adlOnunload() {
    if(adl_API != null) {
      // build time in propper format...HHHH:MM:SS
      var vTime = new Date()
      var vTimeString
      vTime.setHours(0)
      vTime.setMinutes(vTotalModTime)
      vTime.setSeconds(0)
      vTimeString = ("0000" + vTime.getHours()).slice(-4) + ":" + ("00" + vTime.getMinutes()).slice(-2) + ":00"
      adl_API.LMSSetValue("cmi.core.session_time", vTimeString)
      adl_API.LMSCommit("")
      adl_API.LMSFinish("")
    }	
  }
  
