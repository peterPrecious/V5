<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Exam_Routines.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Exam</title>
</head>

  <%
    Dim vNoTimeLimit, vAllowPassRetry

    vAllowPassRetry = False
    '...Check this paramenter to allow a user to retry even if they've already Passed
    If Len(Request.QueryString("vAllowPassRetry")) > 0 Then
      If LCase(Request.QueryString("vAllowPassRetry")) = "y" Then vAllowPassRetry = True
    End If
  
    '...localize variables
    sGetQueryString
  
    vModID = ""
    
    If Request.QueryString("vModID").Count = 1 then 
      vModID = Request.QueryString("vModID")
    ElseIf Request.Form("vModID").Count = 1 then 
      vModID = Left(Request.Form("vModID"),6)  
    ElseIf Len(Session("ModID")) > 0 Then
      vModID = Session("ModID")
    End If
  
    Session("ModID") = vModID
  
 
    '...check if coming from Tree Menu again
    If (Len(Request.QueryString("vMinQue")) > 0) And (Session(vModID & "TestStarted")) Then
      '...if so, redirect to Cheat page
      Session(vModID & "TestStarted") = False
      Response.Redirect "ExamCheat.asp"
    End If
  %>
  <script>

    var ns,ie
    var vBrowserVer = <%=vBVer%>
    var browser = navigator.appName.indexOf('Netscape') != -1 ? (ns=1) : (ie=1)
  
    var timerID = null
    var timerRunning = false
  
    var minutes = 
    <%
      vNoTimeLimit = False
      If Len(Session(vModID & "BankTLimit")) = 0 Then
        If Cint(Request.QueryString("vBankTLimit")) = 0 Then vNoTimeLimit = True
        Response.Write Request.QueryString("vBankTLimit")
      Else
        If Cint(Session(vModID & "BankTLimit")) = 0 Then vNoTimeLimit = True
        Response.Write Session(vModID & "BankTLimit")
      End If
    %>
  
    var seconds = 01
    var timeValue
    var message = "<%=fPhraH(000018)%>"
  
    function showtime(){
       
      if (!((seconds == 00) && (minutes == 00))) {
        if (seconds == 00) {
          seconds = 59
          minutes = minutes - 1
        }
        else
          seconds = seconds - 1
      }
      
      if ((seconds == 00) && (minutes == 00))
        timeValue = "Time Expired"
      else {
        timeValue = ((minutes < 10) ? "0" : "") + minutes
        timeValue += ((seconds < 10) ? ":0" : ":") + seconds
      }
  
      if (ns && vBrowserVer!=6) {
        with(window.document.ExamCountdown) {
          document.open();
          document.write('<font face="Verdana" size="3"><b>' + message + ': ' + timeValue + '</b></font>\n');
          document.close();
        }
      } else
        document.getElementById("ExamCountdown").innerHTML = message + ': ' + timeValue
  
      timerID = setTimeout("showtime()",1000)
      timerRunning = true
    }
   </script>
  <% If vNoTimeLimit Then %>

  <body link="#000080" vlink="#000080" alink="#000080" bgcolor="#FFFFFF" text="#000080">

  <% Else %>

  <body link="#000080" vlink="#000080" alink="#000080" bgcolor="#FFFFFF" text="#000080" onload="showtime()">

  <% End If %> 
    
  <% Server.Execute vShellHi %> 
    
  <%
    Dim vModID, aStr, aQue, vQue, aAns, vAns, vCheck, vChecked, vFormOK, vTotal, vRandom, vTotalTime, vMess, vExpires
    Const cAlpha = "abcdefg"
    vExpires = DateAdd("yyyy", -1, Now) '...used to isolate only current exam info
 
  
    '...check to see if test already in progress
    If Not Session(vModID & "TestStarted") Then
      '...store Min/Max number of questions and time limit (in min.) for each bank
      If Len(Request.QueryString("vMinQue")) > 0 Then
        Session(vModID & "MinQue")      = Request.QueryString("vMinQue")
        Session(vModID & "MaxQue")      = GetMaxQue(vModID)
        Session(vModID & "BankTLimit")  = Request.QueryString("vBankTLimit")
        Session(vModID & "MaxAttempts") = Request.QueryString("vMaxAttempts")
        Session(vModID & "PassGrade")   = Request.QueryString("vPassGrade")
        Session("ExamUnlock")           = Request.QueryString("vExamUnlock")
      End If
  
      '...starting test from beginning
      '...check if user is restarting a test in progress (LOGS table)
      Dim vBank, aResults, vStartInfo, vAttempt
      If TestInProgress(vModID, vBank, aResults, vAttempt) Then
        vAttempt = Cint(vAttempt) '...Convert to workable number
        '...check if test has already been completed by Min 90% or failure
        Session(vModID & "Bank") = Ubound(aResults) + 1
        Session(vModID & "BankCheat") = Ubound(aResults) + 1
        vTotal = GetTotalResults (vModID, vTotalTime, Session(vModID & "Bank"))
  
        '...Must check if all banks completed AND NO Final Grade recorded.  If not, we must store it.
        If Cint(vBank*5) = Cint(Request.QueryString("vMinQue")) Then
          Dim oRsCheckFinalLog
          sOpenDb
'         vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_%' AND (RIGHT(LEFT(Logs_Item, 8), 1) = '" & vAttempt & "')"
          vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_%' AND (RIGHT(LEFT(Logs_Item, 8), 1) = '" & vAttempt & "')"

          Set oRsCheckFinalLog = oDb.Execute(vSql)
          If oRsCheckFinalLog.Eof Then
            '...Must insert missing Log entry for this completed exam
            Dim vLogs_Item
            vLogs_Item = vModId & "_" & vAttempt & "_" & Right("000" & Int(vTotal * 100) ,3)
            vSql = "INSERT INTO Logs"
            vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
            vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & svMembNo & ")"
            oDb.Execute(vSql)
          End If
          oRsCheckFinalLog.Close
          Set oRsCheckFinalLog = Nothing
          sCloseDb
        End If
  
        If (Cint(Session(vModID & "Bank")) * 5) < Cint(Session(vModID & "MinQue")) Then
          '...haven't completed Min # of Questions yet
          vStartInfo = fPhraH(000284) & "&nbsp;" & vBank + 1 & "&nbsp;" & fPhraH(000285) & ".<br><br>"

        ElseIf (Cint(Session(vModID & "Bank")) * 5) >= Cint(Session(vModID & "MinQue")) AND vTotal < (Cint(Session(vModID & "PassGrade"))/100) And (vAttempt < Cint(Session(vModID & "MaxAttempts"))) Then
          '...completed Min # of questions with a failing grade...another attempt is available
          vBank = 0
          vAttempt = vAttempt + 1
          vStartInfo = fPhraH(000286) & "&nbsp;" & Session(vModID & "PassGrade") & "%.<br>" & fPhraH(000287) & "&nbsp;" & vAttempt & "&nbsp;" & fPhraH(000288) & "<br><br>"

        ElseIf (Cint(Session(vModID & "Bank")) * 5) >= Cint(Session(vModID & "MinQue")) AND vTotal < (Cint(Session(vModID & "PassGrade"))/100) And (vAttempt >= Cint(Session(vModID & "MaxAttempts"))) Then
          '...completed Min # of questions with a failing grade...no more attempts are available
          '...must first check if exam passed in previous attempt.  If so, continue to Certificate Page
          Dim oRsCheckPass, vTotalGradePass
          sOpenDb
'         vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_%' AND (RIGHT(Logs_Item, 3) >= " & Request.QueryString("vPassGrade") & ")"
          vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctID='" & svCustAcctId & "' AND Logs_Type='T' AND Logs_Posted > '"  & vExpires & "' AND Logs_MembNo=" & svMembNo & " AND Logs_Item LIKE '" & vModID & "_%' AND (RIGHT(Logs_Item, 3) >= " & Request.QueryString("vPassGrade") & ")"

          Set oRsCheckPass = oDb.Execute(vSql)
          If Not oRsCheckPass.Eof Then
            Session(vModId & "Attempt") = Mid(oRsCheckPass("Logs_Item"),8,1)
            vTotalGradePass = Right(oRsCheckPass("Logs_Item"),3)/100
            oRsCheckPass.Close
            Set oRsCheckPass = Nothing
            sCloseDb
            '...failed current exam, but has a previous passing grade
            Session(vModID & "ExamPassed") = True
            Response.Redirect "ExamGrade.asp?vModID=" & vModID & "&vTotalGrade=" & vTotalGradePass
          End If
          oRsCheckPass.Close
          Set oRsCheckPass = Nothing
          sCloseDb
  
          vMess = fPhraH(000039)
          Response.Redirect "ExamComplete.asp?vMess=" & Server.URLencode(vMess)
        ElseIf (Cint(Session(vModID & "Bank")) * 5) >= Cint(Session(vModID & "MinQue")) AND vTotal >= (Cint(Session(vModID & "PassGrade"))/100) And (vAttempt < Cint(Session(vModID & "MaxAttempts")) And (Not vAllowPassRetry)) Then
          '...completed Min # of with a passing grade
          Session(vModID & "ExamPassed") = True
          Response.Redirect "ExamGrade.asp?vModID=" & vModID & "&vTotalGrade=" & vTotal
        ElseIf (Cint(Session(vModID & "Bank")) * 5) >= Cint(Session(vModID & "MinQue")) AND vTotal >= (Cint(Session(vModID & "PassGrade"))/100) And (vAttempt < Cint(Session(vModID & "MaxAttempts")) And (vAllowPassRetry)) Then
          '...completed Min # of with a passing grade BUT ALLOW another attempt
          vBank = 0
          vAttempt = vAttempt + 1
          vStartInfo = fPhraH(000289) & "&nbsp;" & vAttempt & "&nbsp;" & fPhraH(000288) & "<br><br>"
        ElseIf (Cint(Session(vModID & "Bank")) * 5) >= Cint(Session(vModID & "MinQue")) Then
          Session(vModID & "ExamPassed") = True
          Response.Redirect "ExamGrade.asp?vModID=" & vModID & "&vTotalGrade=" & vTotal
        End If
  
        '...otherwise continue
        Session(vModID & "Attempt") = vAttempt
        Session(vModID & "TestStarted") = True
        Session(vModID & "Bank") = vBank + 1
        Session(vModID & "BankCheat") = vBank + 1
        Session(vModID & "TestResults") = aResults
      Else
        '...if not, start at first bank
        Session(vModID & "Attempt") = vAttempt + 1
        Session(vModID & "TestStarted") = True
        Session(vModID & "Bank") = 1
        Session(vModID & "BankCheat") = 1
      End If
    Else
      '...increment Bank
      Session(vModID & "Bank") = Session(vModID & "Bank") + 1
    End If
  
    '...check if coming from improper page/instance of page
    If Len(Request.QueryString("FromDisplay")) > 0 Then
      If Cint(Request.QueryString("FromDisplay")) + 1 <> Cint(Session(vModID & "Bank")) Then
        '...if so, redirect to Cheat page
        Session(vModID & "TestStarted") = False
        Response.Redirect "ExamCheat.asp"
      End If
    End If
  
    '...check if trying to start this bank again
  '  If Session("BankStarted" & Session("ModID") & Session(vModID & "Bank")) Then
  '    '...if so, redirect to Cheat page
  '    Session(vModID & "TestStarted") = False
  '    Response.Redirect "ExamCheat.asp"
  '  Else
  '    Session("BankStarted" & Session("ModID") & Session(vModID & "Bank")) = True
  '  End If
  
    '...need to define if next bank is random or not
    If VarType(vTotal) = vbEmpty Then
      If Request.QueryString("vRandom") Then
        vRandom = True
      Else
        vRandom = False
      End If
    End If
  
    '...Check if 1st Bank should be fixed
    If (Request.QueryString("vBankFixed") = "Y") And (Session(vModID & "Bank") = 1) Then
      vRandom = False
    Else
      vRandom = True
    End If
  
    aQue = GetStrBank (vModID, Session(vModID & "Bank"), vRandom) : vQue = Ubound(aQue)
    Session(vModID & "CurrentBank") = aQue
  
    '...default the results (LOGS) to all zeros
    If GradeInitBank (vModID, Session(vModID & "Bank"), aQue, Session(vModID & "Attempt")) Then
      '...set start time ONLY if no entry in Logs table...in case user hits "REFRESH"
      '...store the Start time
      Session(vModID & "TestStart") = Time
    End If
  
    '...increase timeout from 20 to 35
    Session.Timeout = 35
  
  %>
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
      <tr>
        <td>
        <h1>
        <!--webbot bot='PurpleText' PREVIEW='Examination'--><%=fPhra(000132)%></h1>
        <h3><%=vModID & " - " & fExamTitle(vModID)%></h3>
        <h2><%=vStartInfo%><!--webbot bot='PurpleText' PREVIEW='This exam is intended to help you gauge your understanding of the material. When finished all five questions, click the &quot;next&quot; button to continue. If you run out of time you will be notified accordingly, will score 0 out of 5 for the bank, and will need to click &quot;next&quot; to continue.'--><%=fPhra(000013)%>&nbsp;
        <!--webbot bot='PurpleText' PREVIEW='If no timer is present, there is no time limit.'--><%=fPhra(000147)%><br></h2>
        </td>
      </tr>
    </table>
    <form method="POST" action="ExamSummary.asp?vBank=<%=Session(vModID & "Bank")%>" id="form1" name="form1" target="_self">
      <input type="hidden" name="vModID" value="<%=vModID%>">

        <div align="center"><center><font face="Verdana" size="3"><b>
      <% If (Session("Browser") = "msie") Or (Session("Browser") = "mozilla") Then %>
          <div id="ExamCountdown">
      <% Else 
           If Len(vStartInfo) > 0 Then %>
        <div id="ExamCountdown" style="color: #000000; position: absolute; left: 200; top: 180">
      <%   Else %>
        <div id="ExamCountdown" style="color: #000000; position: absolute; left: 200; top: 160">
      <%   End If 
         End If %> 

           </div></b></font><br>
            <table border="1" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="0" cellspacing="0">
           <%  
            vFormOK = False
            For i = 0 To vQue  
              aAns = split(aQue(i),"||"): vAns = Ubound(aAns) 
              If Len(aAns(0)) > 0 Then
                vFormOK = true
           %>
              <tr>
                <td width="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">
                <h1>&nbsp;<%=((Session(vModID & "Bank")-1) * 5) + i + 1%>.</h1>
                </td>
                <td bgcolor="#DDEEF9" bordercolor="#FFFFFF">
                <h1><%=aAns(1)%> 
           <%
                If svMembLevel = 5 Then
                  Response.Write " <font face='Verdana' size='2' color='Red'>" & aAns(2) & "</font>"
                End If
           %>
             &nbsp; </h1>
                </td>
              </tr>
           <%
              For j = 3 To vAns
                If Len(aAns(j)) > 0 then  
           %>
              <tr>
                <td valign="top" width="30" align="right"><input name="Q<%=right("00" & i+1,2)%>" type="radio" value="<%=j-2%>" <%=vchecked%>></td>
                <td valign="top"><p class="c2"><%=aAns(j)%> </p></td>
              </tr>
          <%  
                End if
              Next
            End If
          Next
          %>
            </table>
            <% If Not vFormOK Then%> 
            <h6 align="center">
            <!--webbot bot='PurpleText' PREVIEW='Error: there are no questions on file for this module! Please inform your
            '--><%=fPhra(001193)%>&nbsp;<%=svCustFacilitator%>. </h6>
            <% End If %> 
            <p align="center">&nbsp;<input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I3" type="image"></p></div>
        </div>
        </center></div>
      </div>
      </center></div>
    </form>
    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</body>

</html>

