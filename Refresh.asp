<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->

<%
  Dim vModId, vProgId, vTimeSpent
  '...allow bypass security, if timed out, then just ignore code below
  If Session("Secure") Then
    '...get prog | module info and log timespent in the module - only if this is part of a program, not a single module
    vProgId    = Request.QueryString("vProgId")
    vModId     = Request.QueryString("vModId")
    vTimeSpent = Request.QueryString("vTimeSpent")  
    If Len(vProgId) = 7 Then fLogTimespent vProgId, vModId, vTimeSpent    
    If Request("sessionActive") = "Y" Then Response.Write "sessionActive=Y"    
    Session("SessionStarted") = Now()
  Else  
    If Request("sessionActive") = "Y" Then Response.Write "sessionActive=N"            
  End If
%>