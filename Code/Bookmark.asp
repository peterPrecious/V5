<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

<% Server.Execute vShellHi %>

<p align="left"><font face="Arial Black" size="2" color="#3977b6">::&nbsp; Bookmark this service</font></p>
<p align="left">If you are currently on your own computer you can bookmark this service for speedy access.&nbsp; However, please note that your Customer Id and Password will be stored in your Favorites List which may be a security breach.&nbsp; <font color="#FF0000"><br><br>Obviously, do not use this feature if you are on a public computer.</font></p>

<script language="JavaScript">
  if ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4)) 
  {
    var url="/V5/default.asp?vCust=<%=svCustId%>&vId=<%=svMembId%>";
    var title="Vubiz Access for <%=svMembFirstName & " " & svMembLastName%>";
    
    document.write('To bookmark, click here...<br> <A HREF="javascript:window.external.AddFavorite(url,title);" ');
    document.write('onMouseOver=" window.status=');
    document.write("'Add Vubiz to your favorites!'; return true ");
    document.write('"onMouseOut=" window.status=');
    document.write("' '; return true ");
    document.write('">Vubiz Access for <%=svMembFirstName & " " & svMembLastName%></a>');
   }
</script>
<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body></html>


