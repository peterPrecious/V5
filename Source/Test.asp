<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->

<%
  '...this feature will creat on "onunload" auto close feature for QModId access
  If Session("ModAutoClose") Then
    Response.Write "<script FOR='window' EVENT='onunload'>" & vbCrLf
    Response.Write "window.open('ModSignoff.asp','','width=10,height=10,left=0,top=0')" & vbCrLf
    Response.Write "</script>" & vbCrLf
  End If     

  '...increase timeout from 20 to 30
  Session.Timeout = 30

  Dim vModId, vStr, aQue, vQue, aAns, vAns, vCheck, vChecked, vFormOK
  Const cAlpha = "abcdefg" 

  vModId = Request("vModId")
  If Len(vModId) > 6 Then vModId = Left(vModId, 6)  

  Session("ModId") = vModId
  Session("ExamUnlock") = Request.QueryString("vExamUnlock")

  vStr = GetStr (vModId)
  aQue = split(vStr,"~~"): vQue = ubound(aQue)
%>

<html>
  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
    <title><!--[[-->Self Assessment<!--]]--></title>
  </head>

  <body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="TestGrade.asp" id="form1" name="form1" target="_self">
    <input type="hidden" name="vModId" value="<%=vModId%>">
    <input type="hidden" name="vProgId" value="<%=Request("vProgId")%>">

    <table border="1" width="100%" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td colspan="2">
          <h1 align="center"><%=vModId & " - " & fModsTitle(vModId)%></h1>
          <p>
          <!--[[-->Please endeavour to answer all questions by clicking the radio button beside your answer. When finished, click &quot;next&quot; below for grading.<!--]]--> <font color="#FF0000">
          <!--[[-->Important: You must answer all questions and click &quot;next&quot; within 20 minutes to submit your responses. Otherwise, your session will expire and you will have to sign in again and retake the self-assessment.<!--]]--></font>
          <!--[[-->Good luck!<!--]]--><br>&nbsp;</p>
        </td>
      </tr>
      <%  
        vFormOK = False
        For i = 0 To vQue - 1  
          aAns = split(aQue(i),"||"): vAns = Ubound(aAns) 
          If Len(aAns(0)) > 0 Then
            vFormOK = true
      %>
      <tr>
        <th valign="top" width="30" align="left">&nbsp;<%=i+1%>.</th>
        <th bgcolor="#DDEEF9" align="left"><%=aAns(0)%><br>&nbsp;</th>
      </tr>
      <%
            For j = 2 To vAns
              If Len(aAns(j)) > 0 then  
      %>
      <tr>
        <td valign="top" width="30" align="right"><input name="Q<%=right("00" & i+1,2)%>" type="radio" value="<%=j-1%>" <%=vchecked%>></td>
        <td valign="top"><%=aAns(j)%></td>
      </tr>
      <%  
              End if
            Next
      %>
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <%  
          End If
        Next
      %>
      <tr>
        <td colspan="2" align="center">        
          <% If Not vFormOK Then%>
            <h5 align="center">
            <!--[[-->Error: there are no questions on file for this module!<!--]]--></h5>
          <% End If %>
          
          <br>&nbsp;<input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I3" type="image">
          
          <% If (svLang = "EN" Or svLang = "FR") And (Request("vClose") = "Y") Then '...if using a jWindow then mention closing... %> 
          <h1 align="center"><!--[[-->Click on the X to close this window.<!--]]--></h1>
          <% End If %>

        </td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>

