<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Sessions.asp"-->

<%
  '...used in assessment player

  Dim vSessions
  vSessions = Request.Form("vSessions")

  '...do we need to recreate the session? 
  If Len(Session("CustId")) = 8 Then
    Response.Write "SessionOk"
  '...did we get the session string?
  ElseIf Len(vSessions) = 0 Then
    Response.Write "SessionDead"   
  '...recreate the sessions
  Else
    sReCreateSessions vSessions
    Response.Write "SessionRecreated"
  End If
%>
