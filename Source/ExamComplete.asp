<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->
<!--#include virtual = "V5/Inc/Exam_Routines.asp"-->

<%
  '...FOR DEBUGGING ONLY !!!
  If Request.QueryString("DeleteScores") Then DeleteScores Session("ModID")
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title>Exam Complete</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>

  <% Server.Execute vShellHi %> 
  
  <h2 align="center"><%=Request.QueryString("vMess")%></h2>

  <% If svMembLevel = 5 Then %>
    
    <h2 align="center">
    Click below to delete Score results for ALL attempts *ADMINISTRATORS/TESTING ONLY*... 
    </h2>
    <form method="POST" action="ExamComplete.asp?DeleteScores=True" id="form1" name="form1" target="_self">
      <p align="center"><input border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif" name="I3" type="image"></p>
    </form>
  <% End If %> 
  
  <p align="center"><a href="javascript:window.close()"><img border="0" src="../Images/Buttons/Close_<%=svLang%>.gif"></a> </p>
  
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

