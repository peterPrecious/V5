<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>

</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" width="100%" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9">
    <tr>
      <td valign="top" colspan="5" align="center">
      <form method="POST" action="AssessmentRenewalReport.asp">
        <h1>Assessment Renewal Report</h1>
        <h2>This lists learners who may require an annual assessment renewal.<br>Assessment dates in <font color="#008000">green</font> were taken at 10 months ago.</h2>
      </form>
      </td>
    </tr>
    <tr>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Learner <span style="font-weight: 400">(with email link)</span></th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Assessment <br>Date</th>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Assessment Title</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Score<br><span style="font-weight: 400">(Certificate)</span>&nbsp; </th>
    </tr>
    <% 
      Dim vBold, vGrade, vTestExam, vTitle, vOk, vCertUrl, vCertType
    	'...exams ids that are renewable are in the RENW file

      vSql = "SELECT Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Email, "
      vSql = vSql & " Left(Logs.Logs_Item, 6) AS Logs_Module, Right(Logs.Logs_Item,3) AS Logs_Grade, Logs.Logs_Posted "
      vSql = vSql & " FROM Logs WITH (nolock) "
      vSql = vSql & " INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "
      vSql = vSql & " INNER JOIN Renw ON LEFT(Logs.Logs_Item, 6) = Renw.Renw_Id "      
      vSql = vSql & " WHERE Logs_AcctId= '" & svCustAcctId & "' AND Logs_Type = 'T' AND LEN(Logs_Item) > 10 AND CAST(RIGHT(Logs_Item, 3) AS FLOAT) >= 80 AND Memb_Level < 4 "
      vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Logs.Logs_Posted  "

'     sDebug
      sOpenDB
      Set oRs = oDB.Execute(vSql)

      '...read until either eof or end of group
      Do While Not oRs.Eof  

        vLogs_Module                = oRs("Logs_Module")
        vLogs_Grade                 = oRs("Logs_Grade")
        vLogs_Posted                = oRs("Logs_Posted")
        vMemb_Id                    = oRs("Memb_Id")
        vMemb_FirstName             = oRs("Memb_FirstName")
        vMemb_LastName              = oRs("Memb_LastName")
        vMemb_Email                 = oRs("Memb_Email")
 
        vTitle = fExamTitle(vLogs_Module)  '...get title
        Session("CertProg") = fProgCert (vLogs_Module)  '...Flag ProgCerts
        vCertUrl = "javascript:jCertificate('" & svLang & "','" & vLogs_Module & "','" & fjUnquote(vTitle) & "','" & vLogs_Posted & "','" & vLogs_Grade/100 & "','Exam', '" & vMemb_FirstName & " " & vMemb_LastName & "')"

        '...display expiry date in green if expiry is 11 months or more
        i = fIf(DateDiff("d", vLogs_Posted, Now) > 305, "<font color='#008000'>" & fFormatDate (vLogs_Posted) & "</font>", fFormatDate (vLogs_Posted))
    %>
    <tr>
      <td valign="top">
        <% If Len(Trim(vMemb_Email)) > 5 Then %>
        <a href="mailto:<%=vMemb_Email%>"><%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%></a>&nbsp; 
        <% Else %>
        <%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%>
        <% End If %>
      </td>
      <td valign="top" align="center"><%=i%>&nbsp; </td>
      <td valign="top"><%=fLeft(vTitle, 42)%>&nbsp; </td>
      <td valign="top" align="center"><a <%=fstatx%> href="<%=vCertUrl%>"><%=vLogs_Grade%></a></td>
    </tr>
    <%
       oRs.MoveNext
      Loop 
      sCloseDb      
    %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

</html>

