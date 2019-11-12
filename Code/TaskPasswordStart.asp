<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Password_Routines.asp"-->

<%
  '...store the decoded password
  Session("MyWorld_Password") = "123"
' Session("MyWorld_PasswordAttemps") = ""

  If fPasswordOK Then
    response.write "<P>OK"
  Else
    Response.Redirect "TaskPassword.asp?vUrl=TaskPasswordStart.asp"
  End If
%>

