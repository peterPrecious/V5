<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <title>New Page 1</title>
</head>

<body>
<% svLang = "EN" %>
<br>EN: <%=fFormatDate(now)%>
<% svLang = "FR" %>
<br>FR: <%=fFormatDate(now)%>
<% svLang = "ES" %>
<br>ES: <%=fFormatDate(now)%>


</body>

</html>
