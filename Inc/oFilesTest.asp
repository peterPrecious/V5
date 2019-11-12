<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/CustomCertRoutines.asp"-->


<%
  '...used to test file drop downs
  Dim vCustomCert
  vCustomCert = Request("vCustomCert")
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <title>Custom Certificate Drop Down</title>
</head>

<body>

  <form method="POST" action="oFilesTest.asp">
    <select size="1" name="vCustomCert">
    <%=fCustomCertOptions(vCustomCert)%>
    </select> 
    <input type="submit" value="Submit" name="B1">
  </form>

</body>

</html>