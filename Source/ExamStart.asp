<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Exam Start</title>
</head>

<body>

  <!----------------   DO NOT REFORMAT <li> get's screwed up </li>  -------------------------->
  
  <% 
    Server.Execute vShellHi 

  
    Dim vModId, vMinQue, vBankTLimit, vPassGrade, vMaxAttempts, vAllowPassRetry
  
    vModId       = Request.QueryString("vModId")
    vMinQue      = Request.QueryString("vMinQue")
    vBankTLimit  = Request.QueryString("vBankTLimit")
    vPassGrade   = Request.QueryString("vPassGrade")
    vMaxAttempts = Request.QueryString("vMaxAttempts")
    vAllowPassRetry = False
  
    '...for custom certs, some use scrollbar
    If Request.Querystring("vScrollbar") = "yes" Then
      Session(vModId & "Scrollbar")      = "yes"
    End If
  
    '...new feature to allow
    Session("CertTitle") = Request.QueryString("vCertTitle")
  
    '...get prog id in case there are cust certs
    If Len(Request.QueryString("vProgId")) = 7 Then
     Session("CertProg") = Request.QueryString("vProgId")
    End If
  
  
    '...Check this paramenter to allow a user to retry even if they've already Passed
    If Len(Request.QueryString("vAllowPassRetry")) > 0 Then
      If LCase(Request.QueryString("vAllowPassRetry")) = "y" Then vAllowPassRetry = True
    End If
  
    '... Custom Coded Added by Mike (killed by Peter - no longer used)
  ' If UCase(vModId) = "1073EN" AND UCase(svCustId) = "NSRC2321" then
  '   Dim vEmlBody, oEmlEmail
  '   vEmlBody = svMembFirstName & " " & svMembLastName & " (" & svMembEmail & ") Began the Basic Broker Practice Exam at " & Now & vbCrLf
  '   Set oEmlEmail = Server.CreateObject("SMTPsvg.Mailer")
  '   oEmlEmail.FromName       = "VUBIZ"
  '   oEmlEmail.FromAddress    = "info@vubiz.com"
  '   oEmlEmail.RemoteHost     = svMailServer
  '   oEmlEmail.Recipient      = "jmoore@vubiz.com" 
  '   oEmlEmail.Recipient      = "examnotice@ibao.com" 
  '   oEmlEmail.ReturnReceipt  = false
  '   oEmlEmail.ConfirmRead    = false
  '   oEmlEmail.Subject        = "VUBIZ Exam Started"
  '   oEmlEmail.ClearBodyText
  '   oEmlEmail.BodyText       = vEmlBody
  '   oEmlEmail.SendMail
  '   oEmlEmail.ClearRecipients
  '   oEmlEmail.ClearBodyText    
  '   Set oEmlEmail = Nothing  
  ' End If
  
    If Request.QueryString("vStart") Then Session(vModId & "TestStarted") = False
  
    '...determine if already passed within last 12 months
    Dim aMark, oRsCheck, vExpires
    vExpires = DateAdd("yyyy", -1, Now) '...used to isolate only current exam info
  
    sOpenDb
    vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctId='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModId & "_%'"
  ' sDebug

    Set oRsCheck = oDb.Execute(vSql)
    If Not oRsCheck.Eof And Not vAllowPassRetry Then
      While Not oRsCheck.Eof
        aMark = Split(oRsCheck("Logs_Item"), "_")
        If Cint(aMark(2)) >= Cint(vPassGrade) Then
          Session(vModId & "ExamPassed") = True
          Session(vModId & "Attempt") = Right(Left(oRsCheck("Logs_Item"),8),1)
          oRsCheck.Close
          Set oRsCheck = Nothing
          sCloseDb
          Session("ExamUnlock") = Request.QueryString("vExamUnlock")
          Response.Redirect "ExamGrade.asp?" & Request.QueryString & "&vTotalGrade=" & (aMark(2)/100)
        End If
      oRsCheck.MoveNext
      Wend
    End If
  
    oRsCheck.Close
    Set oRsCheck = Nothing
    sCloseDb
  %>

  <table border="1" width="100%" cellpadding="5" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td>
      <h1><!--[[-->Examination Instructions<!--]]--></h1>
      <h2><!--[[-->Welcome,&nbsp; you are about to begin a strictly controlled examination.&nbsp; Please read carefully before you begin:<!--]]--></h2>

      <ul class="c2">
        <% 
          p1 = vMinQue 
          p2 = FormatPercent(vPassGrade/100, 0)
          p3 = vMaxAttempts - 1
          p4 = vBankTLimit
        %>  
        <li><!--[[-->this exam consists of ^1 questions;<!--]]--></li>
        <li><!--[[-->you require ^2 to pass and be granted an online certificate;<!--]]--></li>
        <li><!--[[-->if you do not get ^2 on the first pass, you can try again ^3 times;<!--]]--></li> 

        <% If vAllowPassRetry Then %>
        <li><!--[[-->if you pass, you will have a chance to better your grade;<!--]]--></li> 
        <% End If %>

        <li><!--[[-->if you still do not pass, you cannot take this exam again;<!--]]--></li> 

        <% If vBankTLimit = 0 Then %>
        <li><!--[[-->the exam presents a bank of 5 questions with no time limit;<!--]]--></li> 
        <% Else %>
        <li><!--[[-->the exam presents a bank of 5 questions at a time which must be completed within ^4 minutes;<!--]]--></li>
        <% End If %>

        <li><!--[[-->once a bank is presented you must answer all 5 questions;<!--]]--></li> 

        <% If vBankTLimit > 0 Then %>
        <li><!--[[-->if you run out of time you will be notified accordingly, will score 0 out of 5 for the bank and will need to click <b>Next</b> to continue;<!--]]--></li>
        <% End If %>

        <li><!--[[-->between banks, you may exit the exam to review content or simply sign off;<!--]]--></li>
        <li><!--[[-->when you return to the exam you will be positioned at the next bank;<!--]]--></li>
        <li><!--[[-->upon successful completion of this exam you will be able to print out a copy of your certificate on your local printer - keep this certificate for your reference;<!--]]--></li>
        <li><!--[[-->all activities are logged - please do not try to &quot;trick&quot; the system;<!--]]--></li>
        <li><!--[[-->DO NOT REFRESH ANY PAGE NOR USE YOUR BACK ARROW AT ANY TIME!<!--]]--></li>
      </ul>


      <h2><!--[[-->If you are ready, click <b>Next</b> below and GOOD LUCK!<!--]]--></h2>

      <p align="center" class="c2"> <a href="Exam.asp?<%=Request.QueryString%>" target="_self"><img border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" alt="<%=Server.HtmlEncode("<!--{{-->CLICK WHEN READY TO BEGIN YOUR EXAM - NOT BEFORE!<!--}}-->")%>"></a>
    
      
      <%
        Dim oRsAttempts, aExamAttempt, aExamDef, vSqlAttempt
        sOpenDb

        '...If the user has already passed and does NOT want another attempt, allow to access Certificate directly
        vSqlAttempt = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctId='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModId & "_%' AND (RIGHT(Logs_Item, 3) >= " & Request.QueryString("vPassGrade") & ")"

        Set oRsAttempts = oDb.Execute(vSqlAttempt)
        If vAllowPassRetry And Not oRsAttempts.Eof Then
      %>
      
      </p><p align="center" class="c2">

      <!--[[-->If you do NOT want another attempt and just want to access your Certificate, click below.<!--]]--></p><p align="center" class="c2">
      <a href="ExamStart.asp?<%=Replace(Request.QueryString,"vAllowPassRetry=Y","vAllowPassRetry=N")%>" target="_self"><img border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif"></a></p>

      <%
        End If
        Set oRsAttempts = Nothing

        '...get all previous Attempts from this Exam where grade is NOT perfect
        vSqlAttempt = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctId='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModId & "_%'  AND (RIGHT(Logs_Item, 3) < 100)"
        Set oRsAttempts = oDb.Execute(vSqlAttempt)
        If Not oRsAttempts.Eof Then
      %> 

      <p align="center" class="c2">
      <!--[[-->Click below to review the Questions that were answered incorrectly in previous attempt(s).<!--]]--> </p><p align="center" class="c2">

      <%
        End If

        While Not oRsAttempts.Eof
          aExamAttempt = Split(oRsAttempts("Logs_Item"),"_")
          p1 = aExamAttempt(1)
      %>

      <!--[[-->View Attempt ^1 questions...<!--]]--> <a href="ExamReportReview.asp?<%=Request.QueryString & "&vTotalGrade=" & aExamAttempt(2)/100 & "&vAttempt=" & aExamAttempt(1)%>" target="_self"><img border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif"></a><br>
      
      <%
          oRsAttempts.MoveNext
        Wend
        sCloseDb
      %>

      </p>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
