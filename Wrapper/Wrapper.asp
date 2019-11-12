<html>

<head>
  <title>Wrapper</title>
  <script>

    //  use to debug key or all events (best to set both to true for full testing)

    var vDebug_All = true; 
    var vDebug_Key = true;

    var vDebug_Key = false;
    var vDebug_All = false;


    // use this to ensure we only commit/send results to LMS once
    var vCommitted = false

    if (vDebug_All) alert ("starting");


 
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
  
     function findGetParamValue(vData,vField) {
      // Find the value within a getParam result
      var vToFind = vField.toLowerCase();
      vData = vData.toLowerCase();
      var vStartPos = vData.indexOf(vToFind);
    
      if (vStartPos > -1) {
        var vEndPos = vData.indexOf(String.fromCharCode(13),vStartPos)
        if (vEndPos == -1) {
          return vData.substr(vStartPos+vToFind.length+1)
        }  
        else {
          return vData.substr(vStartPos+vToFind.length+1,vEndPos-(vStartPos+vToFind.length+1))
        }  
      }
      return null;
    }

    function Left(str, n){
    	if (n <= 0)
    	  return "";
    	else if (n > String(str).length)
    	  return str;
    	else
    	  return String(str).substring(0, n);
    }

    function Right(str, n){
      if (n <= 0)
        return "";
      else if (n > String(str).length)
        return str;
      else {
        var iLen = String(str).length;
        return String(str).substring(iLen, iLen - n);
      }
    }

  </script>
  <script src="ScormAPI.js"></script>
  <script src="ScormWrapper.js"></script>
  <script>

    if ((opener != null) && (!opener.closed)) {
      if (opener.vChildMod != null) opener.vChildMod=window
    }  

    var i = 0

    var vScorm = false;
    var vTotalModTime = 0;
    var vLastPageNo = 0;
    var vModAutoClose;
  
  <%
    If Len(Request.QueryString("vPageNo")) > 0 Then 
'     Response.Write "vBookmarkPage=" & (Request.QueryString("vPageNo")-1) & vbCr
    End If
  
    '...Set whether to ping the server to keep the Session Alive
    If Session("ModAutoClose") Then
      Response.Write "vModAutoClose=true" & vbCr
    End If    
  %>

    // this is triggered when either the content or assessment player closes
    // it simply calls a black page before exiting
    function contentClosed() {
      location.href = "/V5/Wrapper/Black.asp";
    }  
  
    function bookmarkSCORM(vPage) {
      if (vScorm) {
        set_bm(vPage)
      }
    }

    function moduleCompleteSCORM() {
      if (vScorm) {
        if (typeof(vPage) != 'undefined')
          set_bm(vPage)
      }
    }
    
    function closeSCORM() {
      if (vScorm) {
        if (typeof(vPage) != 'undefined') {
          set_bm(vPage)
        }  
      }
    }
    
    function updateModTime() {
      // Increment the time within this Module
      vTotalModTime = vTotalModTime + 1
      // Ping the server every 15 min. to keep the session alive
      if ((vTotalModTime < 121) && ((vTotalModTime % 15) == 0)) {
        frames["ping"].location.href = vPingPage + '?vModId=' + vParamModId + '&vProgId=' + vParamProgId + '&vTimeSpent=' + vTotalModTime
      }
      var timerID = setTimeout("updateModTime()",60000)
      timerRunning = true
    }
  
    function jumpToBookmark() {
      if (vScorm) {
        bookmarkSCORM(i+1)
        alert('When you return to this module, it will begin at this page.  Note: this bookmark will remain active for 90 days.')
      }
    }

  
    function SignOffSession() {
      // use the hidden frame to communicate to the server rather than using popups
      
      if (vDebug_Key) alert("...content has been closed...");      
      if (vScorm) closeSCORM(); //...this sets the bookmark
  
      //Create a dynamic form in the wrapperSubmit frame...populated with all data to send to server
      var submitForm = "";
      submitForm    += '\n<form name="submitForm" id="submitForm" action="/V5/CloseModule.asp" method="post">';
      submitForm    += '\n  <input type="hidden" name="vProgId"     value="' + vParamProgId + '">';
      submitForm    += '\n  <input type="hidden" name="vModId"      value="' + vParamModId + '">';
      submitForm    += '\n  <input type="hidden" name="vTimeSpent"  value="' + vTotalModTime + '">';
      submitForm    += '\n  <input type="hidden" name="vLastPageNo" value="' + vLastPageNo + '">';
      submitForm    += '\n</form>';

      if (vDebug_Key) alert(submitForm);

      if (vDebug_Key) alert("wrapper status: " + wrapperSubmit.document.getElementById('wrapperDIV'));
      if (wrapperSubmit.document.getElementById('wrapperDIV')) {
        wrapperSubmit.document.getElementById('wrapperDIV').innerHTML = submitForm;
      }
  
      //Submit form + data to server
      wrapperSubmit.document.getElementById('submitForm').submit();
    }
  

    function InitializePage(vPage) {
      var url = "";
      i=vPage
      url = urls[vPage];
      main.location.href = vModId + '/' + url;
      setPageInfo();
      synchCombo();
    }
  

    function synchCombo() {
      var comboObject = header.document.forms[0].vSection
      // Uncomment the following 2 lines to kill Combo Synching
      //comboObject.selectedIndex=0
      //return
      if (comboObject!=null) {
        for (var comboCount=0; comboCount < comboObject.options.length; comboCount++) {
          if (i>=comboObject.options[comboCount].value-1) {
            comboObject.selectedIndex = comboCount;
          }
        }
      }
    }
    
    
    // _____________________________________________________________

    // Main Logic starts here  
    // grab the Module ID and load it into the frame.
    // _____________________________________________________________


    var vModId = getParameter("vModId");
    if (vModId == null) {
      alert('This Module has been called incorrectly.\n\nModule will be terminated.')
      var vHackWindow = window.open('','','left=100,right=100,width=500,height=150,menubar=no,resizable=no,scrollbars=no')
      vHackWindow.document.write('<p align="center"><b><font face="Arial" size="3">This Module as been called incorrectly.<br><br>Module has terminated.</font></b></p>')
      opener = vHackWindow
      window.close()
    }
    
    // pull the Type of Module (f,u,x,z) so we load the appropriate frameset(s)
    var vModType = getParameter('vModType').toLowerCase();
    
    // pull the Mod url which determines where to load the X Module into the frames.
    var vModUrl = getParameter('vModUrl');
    
    // pull the ProgID
    var vProgId = getParameter('vProgId');
    
    // pull the Title which will be the caption of the Browser
    var vTitle = getParameter('vTitle');
    vTitle = vTitle.replace(/\+/gi,' '); // Replace + with a space
    
    // pull the entire URL parameter list (if any)
    var vParams = location.href.substring(location.href.indexOf("?"));
    if (vDebug_Key) alert("Parameters received: " + vParams);
    
    // content type can be "f" modules (coded in the module table as "fs", "x" modules (local 3rd party) or "u" modules (remote 3rd party)
    var vUrl;
    if (vModType == 'f') {
      vUrl = "/V5/Wrapper/indexScorm.asp" + vParams
    } 

    else if (vModType == 'u') {
      vUrl = vModUrl
    } 
    else {
      vUrl = "/V5/" + vModUrl + vParams
    }
    if (vDebug_Key) alert("Content located at: " + vUrl);

    var       vOutput = '<frameset onUnload="SignOffSession()" framespacing="0" border="0" rows="0,0,100%,0" frameborder="0">\n';
    vOutput = vOutput + '  <frame marginheight="0" scrolling="no"   noresize marginwidth="0" name="wrapperSubmit" target="main" src="WrapperSubmit.htm">\n';
    vOutput = vOutput + '  <frame marginheight="0" scrolling="no"   noresize marginwidth="0" name="header"        target="main" src="ScormSubmit.htm">\n';
    vOutput = vOutput + '  <frame marginheight="0" scrolling="auto" noresize marginwidth="0" name="main"          target="main" src="' + vUrl  + '">\n';
    vOutput = vOutput + '  <frame marginheight="0" scrolling="no"   noresize marginwidth="0" name="ping"          target="ping" src="/V5/Refresh.asp">\n';
    vOutput = vOutput + '</frameset>\n';
  
    document.writeln(vOutput);
  
    // determine if called from a SCORM compliant LMS
    if (typeof(findAPI)!='undefined' && findAPI(parent) != null) {
      vScorm = true
      getAPI()

<%    '...initialize api (note: lesson_status is one of: Passed, Completed, Failed, Incomplete, Browsed, and Not Attempted)
      Dim lesson_status, lesson_location, student_name

      lesson_status   = Request.QueryString("vLessonStatus")
      lesson_location = Request.QueryString("vPageNo")
      student_name    = Request.QueryString("vFirstName") & " " & Request.QueryString("vLastName")
    
      If Instr("passed completed failed incomplete browsed not attempted", lesson_status) = 0 Then lesson_status = "incomplete"
%>
      if (vDebug_Key) alert("Wrapper is initializing api data from Vubiz LMS for content to access...");
  
      adl_API.LMSSetValue("cmi.core.lesson_status",  "<%=lesson_status%>")
      adl_API.LMSSetValue("cmi.core.lesson_location","<%=lesson_location%>")
      adl_API.LMSSetValue("cmi.core.student_name",   "<%=student_name%>")
    }
  
    var vParamModId  = getParameter('vModId');
    var vParamProgId = getParameter('vProgId');
  
    var timerID = setTimeout("updateModTime()",60000)
    timerRunning = true
  
  </script>
</head>

<body></body>

</html>
