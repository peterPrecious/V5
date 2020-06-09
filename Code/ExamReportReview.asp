<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Exam_Routines.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Exam Report</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center"><center>
    <table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1" bordercolor="#FFFFFF">
      <tr>
        <td width="100%" align="center">
        <%
          Dim vModID, vStr, aQue, vQue, aAns, vAns, vCheck, vChecked, vFormOK, vCountZero, vExpires
          Const cAlpha = "abcdefg" 
          vExpires = DateAdd("yyyy", -1, Now) '...used to isolate only current exam info
          vModID = Request.QueryString("vModID")
        %>
        <h1>
        <!--webbot bot='PurpleText' PREVIEW='Exam Review'--><%=fPhra(000131)%>
        </h1>
        <h2>
        <!--webbot bot='PurpleText' PREVIEW='View Incorrect Exam Responses'--><%=fPhra(000022)%> - <%=vModID%>
        </h2>
        <h2>
        <!--webbot bot='PurpleText' PREVIEW='Final grade achieved for this exam in Attempt number'--><%=fPhra(000309)%>&nbsp;<%=Request.QueryString("vAttempt")%> : <%=Round(Request.QueryString("vTotalGrade")*100) %>%.<br><br>
        <!--webbot bot='PurpleText' PREVIEW='Below is a list of all questions that were answered <b>incorrectly</b>.'--><%=fPhra(000077)%> <br><br></h2>
        <table border="0" width="100%" cellspacing="1" cellpadding="0">
        <%
          Dim oRSExamRes, oRSExamDef, aExamRes, aExamDef, vSqlRes, vSqlDef, vCount
          sOpenDB
          sOpenDBBase
          vCount = 0
          '...get all responses for this Exam
'         vSqlRes = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='H' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_" & Request.QueryString("vAttempt") & "_%" & "'"
          vSqlRes = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='H' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_" & Request.QueryString("vAttempt") & "_%" & "'"
          Set oRSExamRes = oDB.Execute(vSqlRes)
          If oRSExamRes.Eof Then
        %>
          <tr>
            <td bgcolor="#DDEEF9" align="center" height="30"><p class="c6">
            <!--webbot bot='PurpleText' PREVIEW='There is no History information for this examination.'--><%=fPhra(000004)%> </p></td>
          </tr>
        </table>
        <%
          Else
            While Not oRSExamRes.Eof
            aExamRes = Split(oRSExamRes("Logs_Item"),"@@")
            For i = 1 To UBound(aExamRes)
        %> </td>
      </tr>
      <tr>
        <td bgcolor="#DDEEF9" width="30" height="30">&nbsp;<%=i%>.</td>
        <td bgcolor="#DDEEF9" height="30"><%=aExamRes(i)%></td>
      </tr>
      <%
          Next
          oRSExamRes.MoveNext
        Wend
      %>
    </table>
    </center>
    <h2>
    <%
      '...display all banks that have 0 time (timed out, closed browser, manipulated, etc.)
      '...these questions are pulled from TstQ, so they might be outdated...all we can do !!!
      '...get all banks for this Exam with time of 0
'     vSqlRes = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='E' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_" & Request.QueryString("vAttempt") & "_%" & "' AND (RIGHT(Logs_Item, 2) = '_0')"
      vSqlRes = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='E' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_" & Request.QueryString("vAttempt") & "_%" & "' AND (RIGHT(Logs_Item, 2) = '_0')"

      Set oRSExamRes = Nothing
      Set oRSExamRes = oDB.Execute(vSqlRes)

      If Not oRSExamRes.Eof Then
    %> 
    <br><br>
    <!--webbot bot='PurpleText' PREVIEW='Below is a list of all questions that were unanswered because of a timeout or browser manipulation.'--><%=fPhra(000078)%> <br><br></h2>
    <center><p></p><center>
    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#FFFFFF">
      <%     
        vCountZero = 0
        While Not oRSExamRes.Eof
          aExamRes = Split(oRSExamRes("Logs_Item"),"_")
          For i = 3 To (UBound(aExamRes)-1) Step 2
            vCountZero = vCountZero + 1
      %>
      <tr>
        <td bgcolor="#DDEEF9" valign="top" width="30">&nbsp;<%=vCountZero%>.</td>
        <td bgcolor="#DDEEF9" valign="top"><%=GetQuestion(vModID, aExamRes(i))%></td>
      </tr>
      <%
          Next
          oRSExamRes.MoveNext
        Wend
      %>
    </table>
    <%
        End If
  
      End If
  
      sCloseDB
      sCloseDBBase
  
      %> 
      </td>
    </tr>
    </table>
    </center></center></div>
  <p align="center"><a href="javascript:window.close()"><img border="0" src="../Images/Buttons/Close_<%=svLang%>.gif"></a> </p><!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>




