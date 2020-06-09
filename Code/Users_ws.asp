<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  '...determine if the user ID and Password are valid when signing in
  If Request.Form("vFunction") = "ResendEmail" Then
    Set oDb = Server.CreateObject("ADODB.Connection")
    Set oCmd = Server.CreateObject("ADODB.Command")
'   oDb.ConnectionString = "Provider=SQLOLEDB.1;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=vuGold;Data Source=" & svSQL
    oDb.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=" & svHostDbId & ";Initial Catalog=vuGold;Data Source=" & svSQL
    oDb.Open
    Set oCmd.ActiveConnection = oDb
    oCmd.CommandType = adCmdStoredProc
    With oCmd
      .CommandText = "spResetLearnerProgramAssignment"
      .Parameters.Append .CreateParameter("@MemberID", adInteger,  adParamInput, , Request("vMemberId"))
    End With
    oCmd.Execute()
    Set oCmd= Nothing
    sCloseDb
    Response.Write "ok"
  Else
    Response.Write "error"
  End If  
%>


