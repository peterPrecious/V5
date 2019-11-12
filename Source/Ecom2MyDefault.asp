<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  '...set to null so programs can differentiate between My and More Content
  Session("Ecom_Media") = ""
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>


  </head>
  <frameset cols="50%,*" framespacing="0" border="0" frameborder="0">
    <frame name="Contents" src="Ecom2MyPrograms.asp" target="Details" scrolling="auto">
    <frame name="Details"  src="Ecom2MyModules.asp"  target="_self"   scrolling="auto">
  </frameset>
</html>