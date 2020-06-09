<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vUrl, vRole

  Select Case svMembLevel
    Case 5    : vRole = 1
    Case 4,3  : vRole = 2
    Case Else : vRole = 3
  End Select

' vUrl = "//74.213.175.110/lz/VubizForum/" _

  vUrl = "//aspobjectlive.vubiz.com/lz/VubizForum/" _
       & "?vRole=" & vRole _
       & "&vFirstName=" & svMembFirstName _
       & "&vLastName=" & svMembLastName _
       & "&vUserID=" & svMembNo _ 
       & "&vClient=" & svCustId 
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body style="overflow: hidden;" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <!--
    Possible vRole values:
    1  
    2
    3
    
    1 – System Administrator
    2 – Administrator
    3 – User (any user with role level less then “Administrator”)
    
    Example: //74.213.175.110/lz/VubizForum/index.html?vUserID=100&vRole=1&vFirstName=Peter&vLastName=Bulloch
  -->

  <div id="divDiscuss" style="position: absolute; left:48px; top:39px; width:100%; height:100%;">
    <iframe id="iDiscuss" name="iDiscuss" style="width:100%;height:100%;"
		border="0" frameborder="0" src="<%=vUrl%>" style="background-color: #FFFFFF" scrolling="no" marginwidth="1" marginheight="1"></iframe></div>

</body>

</html>



