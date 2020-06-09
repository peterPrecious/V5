<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vModId, vMark, vProgId

  vProgId   = Request("vProgId")
  vModId    = Request("vModId")
  vMark     = Request("vMark") 

  '...If mark is not passed in then grade test
  If fNoValue(vMark) Then vMark = GradeTest (vModId)

  '...log test results?
  If Len(vProgId) = 7 Then
    sGetProg vProgId
    If vProg_LogTestResults = "Y" Then
      vLogs_Item = vModId & "_" & Right("000" & vMark * 100, 3)
      sLogTestResults    
    End If
  End If

  '...need 80% for diploma (note, only passes get this far)
  If vMark >= .8 Then

    Session("CertType")     = "Test"
    Session("CertId")       = vModId 
    Session("CertMark")     = vMark
    Session("CertTitle")    = fModsTitle(vModId)

    '...transfer to certificates
    Response.Redirect "CertificateOk.asp?vModId=" & vModId
    
  End If
%>

<html>
  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title><!--[[-->Description<!--]]--></title>
  </head>

  <body>

  <% Server.Execute vShellHi %>
  <table border="0" cellspacing="0" width="100%">
    <tr>
      <td width="100%" align="center">
      <h6 align="center">
      <% If Not fNoValue(svMembFirstName) Then %> <%=svMembFirstName%>, <% End If %>
      <!--[[-->You correctly answered<!--]]--> <%=FormatPercent(vMark,1)%>
      <!--[[-->of the questions.<!--]]--></h6>
      <h6 align="center">
      <!--[[-->You require 80% for a Certificate of Completion<!--]]-->.</h6>
      <br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp;</td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>

