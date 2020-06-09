<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% 
	If Len(Session("shit")) = 0 Then Session("shit") = 0
	Session("shit") = Session("shit") + 1
%>

<html>
<head>
<script>

  function fullScreen(vUrl) {
    //  if we are just passing a ProgId|ModId (P1234EN|9876EN|Y|Y|Y) then launching a full screen popup for this content, 
    //  else this will be a complete URL
    if (vUrl.substring(0,1) == "P" && vUrl.length >= 14 && vUrl.length <= 20) {
      vUrl = "/V5/LaunchObjects.asp?vModId=" + vUrl + "&vNext=CloseWindow.asp"
    }  
    var modwindow = window.open(vUrl,'FullSceen','width='+screen.width+',height='+screen.height+',top=0,left=0,resizable=yes');
		if (!modwindow) {
			popUpAlert(); 
		} else {
	    try {top.addWindowToArray(modwindow)} catch (err) {}; 
	    parent.vModWindow = modwindow
	    parent.vModWindowOpen = true
	  }
  }

  function fmodulewindow(vModId) {
    var vmodule = "/V5/LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left=10,top=10,status=no,scrollbars=no,resizable=yes');
		if (!modwindow) {
			popUpAlert(); 
		} else {
	    try {top.addWindowToArray(modwindow)} catch (err) {}; 
	    modwindow.focus()
	    parent.vModWindow = modwindow
	    parent.vModWindowOpen = true
    }
  }

    // This refreshes the status of the caller when a module is launched
    var vModWindow, vModWindowOpen = false;
    function fDoRefresh() {
      // check to see if module ever opened within session
      if (vModWindow != null) {
        // if module has been opened and now close, refresh the details frame (note IE will not go to the anchor)  
        if (vModWindow.closed && vModWindowOpen) {
          vModWindowOpen = false;
//        top.frames['main'].location.reload(true);
//        window.location.reload(true); '...this does NOT work in half the browsers, need document
          document.location.reload(true);
        }
      } 
    }

    // This keeps track of any windows opened in the Functions.js script file so they can be closed if the platform is shut down
    var openWindows = new Array()      
    var curWindow = 0

    function addWindowToArray(handle) { 
      openWindows[curWindow++] = handle
    } 

    function closeAllWindows() {      
      for(i=0; i<openWindows.length; i++)
        if (!openWindows[i].closed) {
          openWindows[i].close()      
        }
    } 

</script>
</head>

<body onfocus="fDoRefresh()" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

<p><a href="javascript:fmodulewindow('P1234EN|0876EN|N|Y|N')">fmod - popup</a></p>
<p><a href="javascript:fullScreen('P1234EN|0876EN|N|Y|N')">fmod - full screen popup</a></p>
<p><a href="javascript:fullScreen('P2486EN|4598EN|N|Y|N')">xmod - full screen popup</a></p>
<p><a href="//localhost/V5/LaunchObjects.asp?vModId=P1689EN|2274EN|Y|Y|Y&vNext=MyWorld.asp">xmod - imbedded</a></p>
<%=Session("shit")%>

</body>

</html>
