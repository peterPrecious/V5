<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Exam.asp"-->

<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->


<%
  Dim vModID, vMess, vBank

  vBank = Request.QueryString("vBank")

  '...validate Question and Answers
  If Not ValidateQuestionsBank(vMess, vBank) Then
    Response.Redirect "ExamEdit.asp?vBank=" & vBank & "&vMess=" & vMess
  End If
  '...save test file
  SaveQuestionsBank vBank

  '...if requesting next bank, display it...otherwise, confirm all saved questions
  If Len(Request.Form("I0.x")) > 0 Then
    Session("EditBank") = vBank + 1
    Response.Redirect "ExamEdit.asp?vBank=" & Session("EditBank")
  Else
    Session("EditTest") = False
    Response.Redirect "ExamView.asp?vModID="& vModID
  End If
%>