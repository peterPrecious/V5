<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Exam_Routines.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Db_Keys.asp"-->

<%
  Dim vModId, vMark, oRsCheck, vAttempt, vJustViewingCert, vNoDisplay, vExpires, vCustCert

  vModId = Request.QueryString("vModId")
  vMark  = Request.QueryString("vTotalGrade")
  vMark  = Round(vMark * 100) / 100
  vExpires = DateAdd("yyyy", -1, Now) '...used to isolate only current exam info

  '...if custom cert (from prog/cust table) do not display anything
  vNoDisplay = False
  vCustCert  = False
  sGetProg (Session("CertProg"))
  If vProg_CustomCert Then vNoDisplay = True '...legacy
  If Len(vProg_AssessmentCert) > 0 Then vCustCert = True
  sGetCust (svCustId)
  If Len(vCust_AssessmentCert) > 0 Then vNoDisplay = True : vCustCert = True

  '...catch if trying to force a Certificate without passing
  If Not Session(vModId & "ExamPassed") Then
    '...if so, redirect to Cheat page
'   Response.Redirect "ExamCheat.asp"
  End If

  Session("CertType")     = "Exam"
  Session("CertId")       = vModId 
  Session("CertMark")     = vMark

  If Len(Session("CertTitle")) = 0 Then
    Session("CertTitle")    = GetExamTitle(vModId)
  End If

  Session("CertSample")   = ""   '...ensure certificate is NOT a sample

  '...insert Id & course/page number into Audit - ONE TIME ONLY IF PASSED
  ' NOTE: make this an optional exercise ie vTestAudit=y
  '...allow multiple certificates !!!!
  
  sOpenDb
  vLogs_Item = vModId & "_" & Session(vModId & "Attempt") & "_" & Right("000" & Int(vMark * 100) ,3)
  vSql = "SELECT * FROM Logs WHERE Logs_AcctId='" & svCustAcctId  & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item = '" & vLogs_Item & "'"
  Set oRsCheck = oDb.Execute(vSql)

  If oRsCheck.EOF Then
    vSql = "INSERT INTO Logs"
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & svMembNo & ")"
    oDb.Execute(vSql)
    Session("CertDate") = ""
    vJustViewingCert = False
  Else
    '...Grab the orginal Certificate Date
    'Session("CertDate") = FormatDateTime (oRsCheck("Logs_Posted"), vbShortDate)
    vJustViewingCert = True
  End If

  oRsCheck.Close
  Set oRsCheck = Nothing

  If vJustViewingCert Then
    '...Need to get the HIGHEST scored exam in case of multiple passed exams
    vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctId='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & Left(vLogs_Item,6) & "%' ORDER BY RIGHT(Logs_Item, 3) DESC"

    Set oRsCheck = oDb.Execute(vSql)
    vMark = Right(oRsCheck("Logs_Item"),3) / 100
    Session("CertMark") = vMark
    Session("CertDate") = FormatDateTime (oRsCheck("Logs_Posted"), vbShortDate)
    
    vAttempt = Mid(oRsCheck("Logs_Item"),8,1)
    Session("AttemptNo") = vAttempt

    oRsCheck.Close
    Set oRsCheck = Nothing
  Else

    vAttempt = Session(vModId & "Attempt")
    Session("AttemptNo") = vAttempt
  End If

  sCloseDB

  '...Unlock Processes - ie 03050866 (process number string must be between 001-999)
  If Len(Session("ExamUnlock")) > 0 Then
    i = Session("ExamUnlock")
    For j = 1 To Len(i) Step 3
      k = Cint(Mid(i, j, 3))
      sUnlock k, svMembNo
    Next
    Session("ExamUnlock") = ""
  End If  

  If (Request.QueryString("vTotalGrade")*100) < CInt(Request.QueryString("vPassGrade")) Then
    Session(vModId & "ExamPassed") = False
  End If
  
%>
<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Exam Complete</title>
  <script src="/V5/Inc/Functions.js"></script>

  <!-- put this script in any page that must refresh the contents panel -->
  <script for="window" event="onload">
    //08/17/05 j.b. - Code added for exam string addition to scrollbar option. for custom ccohs certificate.
    <%
      If vCustCert Then
    %>
  	  var vCertWindow = window.open('Certificate.asp?','Certificate','toolbar=no,width=800,height=500,left=100,top=100,status=no,scrollbars=yes,resizable=yes')
    <%
      ElseIf Session(vModId & "Scrollbar") = "yes" Then
    %>
  	  var vCertWindow = window.open('Certificate.asp?','Certificate','toolbar=no,width=650,height=425,left=100,top=100,status=no,scrollbars=yes,resizable=no')
    <% 
      Else
    %>
  	  var vCertWindow = window.open('Certificate.asp','Certificate','toolbar=no,width=650,height=425,left=100,top=100,status=no,scrollbars=no,resizable=no')
    <%
      End If
    %>
    if (this != parent) {
      parent.frames["contents"].location.href = parent.frames["contents"].location.href;
      <% If vNoDisplay Then %>
        vCertWindow.opener = window.opener
        window.close()
      <% End If %>
    }
  </script>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080" onload="showtime()">

  <% Server.Execute vShellHi %>

  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center">
        <h1>
        <!--[[-->Congratulations!<!--]]--></h1>
        <h3>&nbsp;<%=vModId & " - " & GetExamTitle(vModId)%></h3>
        <h2>
        <% If vJustViewingCert Then %>
          <!--[[-->Your highest achieved Certificate is now displayed in a separate window.<!--]]--><br>
        <% Else %>
          <!--[[-->Your Certificate is now displayed in a separate window.<!--]]--><br>
        <% End If %> 

        <%
           sGetQueryString
           If vPlatform = "win" Then 
        %>
          <!--[[-->This Certificate may now be printed by pressing &lt;Ctrl&gt;+p simultaneously.<!--]]--><br>
        <%   Else %>
          <!--[[-->This Certificate may now be printed by pressing &lt;Command&gt;+p simultaneously.<!--]]--><br>&nbsp; </h2>
        <% End If %>
  
        <h2>
        
        <a <%=fstatx%> href="javascript:jCertificate('<%=svLang%>','','','','','Exam', '')">
        <!--[[-->Click here if your certificate did not appear in a separate window.<!--]]--></a>
        <br>
  
        <a <%=fstatx%> href="#" onclick="fullScreen('<%=fCertificateUrl("", "", vMark, "", vModId, Session("CertTitle"), "", "", "", vProg_Id, "", "", "")%>')">
        <!--[[-->Click here if your certificate did not appear in a separate window.<!--]]--></a> [New Certificate]
  
        &nbsp;&nbsp;&nbsp; 
        <% If vMark <> 1 Then %> 
          <br><br><br>
          <!--[[-->Click below to review the Questions that were answered incorrectly.<!--]]--> </h2>
          <p align="center">&nbsp;
          <a href="ExamReportReview.asp?vModID=<%=vModId & "&vTotalGrade=" & vMark & "&vAttempt=" & vAttempt%>" target="_self"><img border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif"></a>
          </p>
        <% End If %> 
        <br><br>&nbsp; 
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
