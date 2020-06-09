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

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Exam Summary</title>
</head>

<body>

<% 
  
  Server.Execute vShellHi 
  
  vModID = Request.Form("vModID") 

  '...check to see if test is in progress
  If Not Session(vModID & "TestStarted") Then
    Response.Redirect "Exam.asp"
    '...check if user is on incorrect bank of questions
  ElseIf CInt(Request.QueryString("vBank")) <> CInt(Session(vModID & "BankCheat")) Then
    '...if so, redirect to Cheat page
    Session(vModID & "TestStarted") = False
    Response.Redirect "ExamCheat.asp"
  End If

  Dim vModID, vMark, vTotal, vTotalTime, vMess

  vMark = GradeTestBank (vModID, Session(vModID & "Bank"), Session(vModID & "TestStart"), Session(vModID & "BankTLimit"), Session(vModID & "CurrentBank"), Session(vModID & "Attempt"))
  '...took too long if -999
  If vMark = -999 Then
    vMess = "<!--{{-->(The time allowed for this bank was exceeded)<!--}}-->"
    vMark = 0
  End If
  vTotal = GetTotalResults (vModID, vTotalTime, Session(vModID & "Bank"))

%>
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%" align="center" class="c2"><p class="c2"><br><br>
      
      <%
        Dim vLogs_Item
  
        If (CInt(Session(vModID & "Bank")) * 5) < CInt(Session(vModID & "MinQue")) Then
          '...haven't completed Min # of Questions yet
          Session(vModID & "BankCheat") = Session(vModID & "BankCheat") + 1
      %>
      <!--[[-->Thank you. Please click <b>Next</b> to continue with Bank<!--]]-->&nbsp<%=Session(vModID & "Bank")+1%>...<br><br>
      <!--[[-->or, you can take a break and leave the site. <br>When you return you will be positioned at the next bank.<!--]]--></p></td>
    </tr>

    <tr>
      <td width="100%" align="center" class="c2"><p>&nbsp; </p>
      <form method="POST" action="Exam.asp?FromDisplay=<%=Session(vModID & "Bank")%>" id="form1" name="form1" target="_self">
        <p>&nbsp;<input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I3" type="image"></p>
      </form>
      </td>
    </tr>

    <center>
    <tr>
      <td width="100%" align="center" class="c2">&nbsp;<br>

      <% ElseIf (CInt(Session(vModID & "Bank")) * 5) >= CInt(Session(vModID & "MinQue")) AND vTotal < (CInt(Session(vModID & "PassGrade"))/100) And (CInt(Session(vModID & "Attempt")) < CInt(Session(vModID & "MaxAttempts"))) Then 
    		'...completed Min # of questions with a failing grade...another attempt is available
    
    		'...insert ID & course/page number into Audit
    		'   make this an optional exercise (ie vTestAudit=y)
    		vLogs_Item = Session("ModID") & "_" & Session(vModID & "Attempt") & "_" & Right("000" & Int(vTotal * 100) ,3)
    		vSql = "INSERT INTO Logs"
    		vSql = vSql & "(Logs_AcctID, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    		vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & svMembNo & ")"
    		sOpenDB
    		oDB.Execute(vSql)
    		sCloseDB
    
    		Session(vModID & "TestStarted") = False
      %>
      <!--[[-->You have completed the exam with a score below<!--]]-->&nbsp;<%=Session(vModID & "PassGrade")%>%.
      <br><!--[[-->You are able to take this Examination again.<!--]]--><br>
      <p align="center"><a href="javascript:window.close()"><img border="0" src="../Images/Buttons/Close_<%=svLang%>.gif"></a></p>
      
      <% ElseIf (CInt(Session(vModID & "Bank")) * 5) >= CInt(Session(vModID & "MinQue")) AND vTotal < (CInt(Session(vModID & "PassGrade"))/100) And (CInt(Session(vModID & "Attempt")) >= CInt(Session(vModID & "MaxAttempts"))) Then 
  
    		'...completed Min # of questions with a failing grade...no more attempts are available
    		'...insert ID & course/page number into Audit
    		'   make this an optional exercise (ie vTestAudit=y)
    		vLogs_Item = Session("ModID") & "_" & Session(vModID & "Attempt") & "_" & Right("000" & Int(vTotal * 100) ,3)
    		vSql = "INSERT INTO Logs"
    		vSql = vSql & "(Logs_AcctID, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    		vSql = vSql & "('" & svCustAcctId & "','T', '" & vLogs_Item & "', " & svMembNo & ")"
    		sOpenDB
    		oDB.Execute(vSql)
    		sCloseDB
    
    		Session(vModID & "TestStarted") = False
      %>
      <!--[[-->You have completed the exam with a score below<!--]]-->&nbsp;<%=Session(vModID & "PassGrade")%>%.
      <br><!--[[-->You DO NOT have any more attempts at this Examination left.<!--]]--><br>
      <p align="center"><a href="javascript:window.close()"><img border="0" src="../Images/Buttons/Close_<%=svLang%>.gif"></a></p>
      
      <% ElseIf (CInt(Session(vModID & "Bank")) * 5) >= CInt(Session(vModID & "MinQue")) AND vTotal >= (CInt(Session(vModID & "PassGrade"))/100) And (CInt(Session(vModID & "Attempt")) < CInt(Session(vModID & "MaxAttempts"))) Then 
    		'...completed Min # of with a passing grade
    		Session(vModID & "ExamPassed") = True
    		Session(vModID & "TestStarted") = False
    		Response.Redirect "ExamGrade.asp?vModID=" & vModID & "&vTotalGrade=" & vTotal
      %>
      <!--[[-->You have completed the exam successfully with a passing grade.<!--]]--><br>
      <% ElseIf (CInt(Session(vModID & "Bank")) * 5) >= CInt(Session(vModID & "MinQue")) Then
    		Session(vModID & "ExamPassed") = True
    		Session(vModID & "TestStarted") = False
    		Response.Redirect "ExamGrade.asp?vModID=" & vModID & "&vTotalGrade=" & vTotal
      %>
      <!--[[-->You have completed the exam successfully with a passing grade; you may now claim a Certificate.<!--]]--><br>
      <% End If %> 
      
      </td>
    </tr>
  </table>
  </center>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
