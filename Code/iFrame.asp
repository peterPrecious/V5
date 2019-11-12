<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  '...this will launch a Gold app in a V5 shell
  Dim vSrc 
  vSrc = Request("vGold") & "?" & Request.QueryString() 
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <p>
  <iframe 
    name="iActivity_Adv" 
    src="<%=vSrc%>" 
    width="99%" 
    height="750" 
    marginwidth="1" 
    marginheight="1" 
    border="0" 
    frameborder="0" 
    align="center"
  >
  </iframe>
  </p>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


