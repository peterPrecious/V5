<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
	Dim vTabs, vTitle, vLingo

	vPage = Request("vPage")

	If fNoValue(vPage) Then
		If svSecure Then 
			'...determine starting page
			sGetCust svCustId     
			If vCust_Tab1 Then 
'       vPage = svCustCluster & ".asp"
				vPage = "Info.asp"
			ElseIf vCust_Tab2 Then
				'...If there's an intro page, then launch that first
				If Len(vCust_MyWorldLaunch) > 4 Then
					vPage = "/V5/Repository/" & svHostDb & "/" & svCustAcctId & "/Tools/" & vCust_MyWorldLaunch
				Else
					vPage = "MyWorld.asp"
				End If 
			ElseIf vCust_Tab3 Then
				vPage = "RTE_MyContent.asp"
			ElseIf vCust_Tab4 Then
				If svMembLevel = 2 Then       
					vPage = "RC_Home.asp"
				Else
					vPage = "Menu.asp"
				End If
			Else
'       vPage = svCustCluster & ".asp"
				vPage = "Info.asp"
			End If         
		Else  
			Response.Redirect "../Public/Default.asp"      
		End If      
	End If
	
	If svSecure Then   
		vTabs = "TabsLive.asp?vTab=" & Request("vTab") & "&vMode=" & Request("vMode") 
		vTitle = Session("CustTitle")    
	Else  
		'...ensure this is not a bookmark
		If Len(Session("HostDb")) = 0 Then 
			vLingo = "EN"
			If Len(Request.ServerVariables("Path_Info")) > 5 Then 
				vLingo = Mid(Request.ServerVariables("Path_Info"), 5,2)
			End If
			Response.Redirect "../Default.asp?vLang=" & vLingo
		End If
		vTabs = "TabsPublic.asp?vTab=" & Request("vTab")
		vTitle = "<!--{{-->Welcome<!--}}-->"
	End If      

%>

<html>

<head>
	<title>:: <%=vTitle%></title>
	<meta charset="UTF-8">
<!--	<script src="/V5/Inc/jQuery.js"></script>-->
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>

	<link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
	<script src="/V5/Inc/Functions.js"></script>
	<% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
	<script>
			<!--  This calls a web service to grab the location of the MultiUserManuals -->
			<!-- get the MultiUserManual repository path and put into a session variable - note nothing is returned -->
			<% If Len(svMultiUserManual) = 0 Then %>
			var vWs = WebService("/V5/Repository/Documents/MultiUserManual/MultiUserManual_ws.asp", "")
			<% End If %>
			// This provides scripts with the language (accessed as parent.lang)
			var lang = "<%=svLang%>";

			// This determines if there's a popup blocker on (accessed in launch.js as parent.popupBlockerOn)
			try {
				var dummyWindow = window.open('/v5/inc/popupBlocker.htm','','width=10,height=10,left=1,top=1,scrollbars=no,location=no,menubar=no,toolbar=no,statusbar=no');
				if (dummyWindow) {
					popupBlockerOn = false;
					dummyWindow.close();
				} 
				else {
					popupBlockerOn = true;
//			  alert("popups are blocked");
				}
			} 
			catch(e) {}	

			// This parameter determines if we need to reload the content page to render new log status
			// code is in MyContent.asp
			var bodyFocus = true;      // start with focus ON
	
			// This keeps track of any windows opened in the Functions.js script file so they can be closed if the platform is shut down
			var vModWindow, vModWindowOpen = false;
			var openWindows = new Array()      
			var curWindow = 0
			function addWindowToArray(handle) { 
				openWindows[curWindow++] = handle
			} 

			function closeAllWindows() {      
				for (i=0; i<openWindows.length; i++) {
					if (!openWindows[i].closed) {
						openWindows[i].close()      
					}
				}
			}     
	</script>
</head>

<frameset id="v5-frameset" onunload="closeAllWindows()" border="0" frameborder="0" framespacing="0" rows="76,*">
	<frame marginheight="0" marginwidth="0" id="tabs" name="tabs" src="<%=vTabs%>" target="main" scrolling="no" noresize>
	<frame marginheight="0" marginwidth="0" id="frameMain" name="main" src="<%=vPage%>" target="_self" scrolling="auto" noresize>
</frameset>

</html>
