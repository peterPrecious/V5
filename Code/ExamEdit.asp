<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Exam.asp"-->
<!-- include virtual = "V5/Inc/Db_ExamReport.asp"-->

<%
  Dim vModID, vCertOK, vCertEdit
  Dim vStr, aQue, vQue, aAns, vAns, vChecked
  Const cAlpha = "abcdefg"

  '...check if adding a new Exam
  If Request.Form("vAdd.x") > 0 Then
    Response.Redirect "ExamMenu.asp?vMess=" & AddExam()
  End If

  If Not Session("EditTest") Then
    Session("EditTest") = True
    Session("EditModID") = Request.Form("vModID")
    If Len(Request.QueryString("vBank")) <> 0 Then 
      Session("EditBank") = Request.QueryString("vBank")
    Else
      Session("EditBank") = 1
    End If
  Else
    If Len(Request.QueryString("vBank")) <> 0 Then Session("EditBank") = Request.QueryString("vBank")
  End If

  vModID = Session("EditModID")
  If vModID = "Select" Then
    vModID = Ucase(Request.Form("vAddModID"))
  End If
  
  '...must be 6 chars with trailing en/fr/es/pt/pl

  If Len(vModID) <> 6 Or Instr("EN FR ES", Right(vModID, 2)) = 0 Or Not IsNumeric(Left(vModID, 4)) Then 
    Response.Redirect "ExamMenu.asp"
  End If
  
  If Len(Request.QueryString("vMess")) > 0 Then
    aQue = Session("Que") : vQue = Ubound(aQue)
  Else
    '...Get the test, or initialize a new test if not there
    aQue = GetStrBankEdit (vModID, Session("EditBank")) : vQue = Ubound(aQue)
  End If
%>


<html>
  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  </head>

  <body>

  <% Server.Execute vShellHi %>
  

  <form method="POST" action="ExamSave.asp?vBank=<%=Session("EditBank")%>" target="_self">
    <input type="hidden" name="vModID" value="<%=vModID%>">
    <table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <td valign="top" colspan="2"><h1>Edit Exam: <%=vModId%></h1><p class="c2">NOTE:
        </p>
        <ul class=c2>
          <li>Maximum 20 banks of questions</li>
          <li>Minimum 5 questions per bank</li>
          <li>Maximum 5 answers per question</li>
          <li>Leave trailing questions completely empty if not required.</li>
          <li>Click radio button left of correct answer.</li>
          <li>Click save at bottom and view revised test.</li>
        </ul>
          <% If Len(Request.QueryString("vMess")) > 0 Then %> <%=Request.QueryString("vMess")%><br><br><% End If %> </td>
      </tr>
        <% If Session("EditBank") = 1 Then %>
        <tr>
          <td width="100%" colspan="2" valign="middle">Exam Title <input type="text" size="60" name="vTitle" value="<%=GetExamTitle(vModID)%>"> <br><br></td>
      </tr>
        <% 
          End If
    
          For i = 0 To vQue
            aAns = Split(aQue(i),"||"): vAns = Ubound(aAns) 
	      %>
      <tr>
          <td valign="top" colspan="2">&nbsp;</td>
      </tr>
      <tr>
          <td valign="top" colspan="2" bgcolor="#DDEEF9">&nbsp;</td>
      </tr>
      <tr>
        <th valign="top" width="30"><%=aAns(0)%>.</th>
        <td valign="top"><input type="text" size="73" name="Q<%=right("000" & aAns(0),3)%>" value="<%=aAns(1)%>"></td>
      </tr>
        <%
          For j = 3 To vAns - 1  
            vChecked = ""       
            If IsNumeric(aAns(2)) Then
              If cInt(j-2) = cInt(aAns(2)) Then vChecked = " Checked" 
            End If
        %>
        <tr>
          <td valign="top" align="right" width="30"><input name="A<%=right("000" & aAns(0),3)%>" type="radio" value="<%=j-2%>" <%=vchecked%>></td>
          <td valign="top"><input type="text" size="68" name="Q<%=right("000" & aAns(0),3) & ucase(mid(cAlpha,j-2,1))%>" value="<%=aAns(j)%>"></td>
      </tr>
        <%  
            Next
          Next
        %>
      </table>
    <center><br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I0" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Save_<%=svLang%>.gif" name="I1" type="image"><p><a href="ExamList.asp">Exam List</a><br>&nbsp;</p></center>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>


