<%
  '...clear password sessions
  Session("MyWorld_Password") = ""
  Session("MyWorld_PasswordAttemps") = ""
%>
<html><body>
<p>Cleared</p>
<p>Session("MyWorld_Password") = <%=Session("MyWorld_Password")%></p>
<p>Session("MyWorld_PasswordAttemps") = <%=Session("MyWorld_PasswordAttemps")%></p>
</body></html>

