<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  '...return the number of minutes left in this session 
  Dim vMins
  If Session("Secure") Then
    vMins = Session.Timeout - DateDiff("n", Session("SessionStarted"), Now())   
    If vMins < 0 Then vMins = 0
    Response.Write vMins
  Else  
    Response.Write -1    
  End If
%>