<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>:: Sign Off</title>
</head>

<body>

  <% 
	  	Server.Execute vShellHi	
	    Dim vLang, vLogo, vCust, vId
	    vLang = Request("vLang")
	    vLogo = Request("vLogo")
	    vCust = Request("vCust")
	    vId		= Request("vId")
  %>

  <div style="text-align: center;">
    <% If Not fNoValue(vLogo) Then %>
    <%   If Len(vCust) > 0 Then %>
    <p><a href="/ChAccess/SignIn/Default.asp?vCust=<%=vCust%>&vLang=<%=vLang%>&vId=<%=vId%>"><img border="0" src="../Images/Logos/<%=vLogo%>"></a></p>
    <p><!--[[-->Thank you, your session has been terminated<!--]]-->.<br><!--[[-->Click on the logo to return<!--]]-->.</p>
    <%   Else %>
    <p><img border="0" src="../Images/Logos/<%=vLogo%>"></p>
    <p><!--[[-->Thank you, your session has been terminated<!--]]-->.<br></p>
    <%   End If %>
    <% Else %>
    <p><!--[[-->Thank you, your session has been terminated<!--]]-->.</p>
    <% End If %>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
