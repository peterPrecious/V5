<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->

<%
  Dim vModID
  '...save test file  
  sSaveQuestions
  Response.Redirect "TestView.asp?vModID="& vModID
%>