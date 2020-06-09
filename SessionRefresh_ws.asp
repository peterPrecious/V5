<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  '...Refresh the server and set the time this session started
  If Session("Secure") Then
    Response.Write "ok"    
    Session("SessionStarted") = Now()
  Else  
    Response.Write "err"    
  End If
%>