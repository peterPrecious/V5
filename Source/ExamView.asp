<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Exam.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->

<%
  Dim vModId, vDelModId, vStr, aQue, vQue, aAns, vAns, vCheck, vChecked, vFormOK
  Const cAlpha = "abcdefg" 

  '...delete exam
  vDelModId = Ucase(Request("vDelModId"))
  If Len(vDelModId) = 6 Then

    '...calling 
    If sDeleteExam(vDelModId) Then
      Response.Redirect "ExamList.asp?vMess=Exam " & vDelModId & " has been successfully deleted."
    Else
      Response.Redirect "ExamList.asp?vMess=Errors occured while trying to delete Exam " & vDelModId & "."
    End If
  End If

  vModId = Request.QueryString("vModId")
  aQue = GetStrBankAll(vModId) : vQue = ubound(aQue)
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

  <table border="1" width="100%" cellspacing="0" cellpadding="0" bordercolor="#DDEEF9" style="border-collapse: collapse">

    <tr>
      <td colspan="2">
        <h1>View Exam: <%=vModId%></h1>
        <p>This lists all the exam banks and signifies which question is the designated answer.&nbsp; If you wish to edit any bank, click on the &quot;edit&quot; button beneath the bank in question.<h6>If you wish to delete all elements of this exam, click on the &quot;delete&quot; button at the bottom of the page.&nbsp; NOTE: deleting an exam is an irreversible action!<br>&nbsp;</h6>
      </td>
    </tr>

    <%  
      vFormOK = False
      For i = 0 To vQue  
        aAns = split(aQue(i),"||"): vAns = Ubound(aAns) 
        If Len(aAns(0)) > 0 Then
          vFormOK = true
    %>
    <tr>
      <td bgcolor="#DDEEF9" height="20" align="left" valign="top"><h1><%=i+1%>.&nbsp;&nbsp; </h1> </td>
      <td bgcolor="#DDEEF9" height="20" align="left" valign="top"><h1><%=aAns(0)%>&nbsp;</h1></td>
    </tr>
    <%
          For j = 2 To vAns
            vCheck = 1
            On Error Resume Next
            vCheck =cInt(aAns(1))
            If Cint(j-1) = vCheck Then vChecked = " Checked" Else vChecked = ""
            If Len(aAns(j)) > 0 then  
    %>
    <tr>
      <td align="center" class="c2">
        <input name="Q<%=right("000" & i+1,3)%>" type="radio" value="<%=j+1%>" <%=vchecked%>>
      </td>
      <td class="c2"><%=aAns(j)%>&nbsp; </td>
    </tr>
    <%  
            End If
          Next
        End If

        If ((i+1)/5) = ((i+1)\5) Then 
    %>
    <tr>
      <td bgcolor="#FFFFFF" align="center" colspan="2">
        <form method="POST" name="<%=i%>" action="ExamEdit.asp?vBank=<%=(i+1)/5%>">
          <input type="hidden" name="vModId" value="<%=vModId%>"><br>
          <input border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif" name="I1" type="image"><br>&nbsp;
        </form>
      </td>
    </tr>
    <%  
        End If

      Next
    %>


    <% If Not vFormOK Then %> 
    <tr>
      <td colspan="2">
        <h5 align="center"><br><br>There are no questions on file for this exam!</h5>
      </td>
    </tr>
	  <% End If %> 
 


    <tr>
      <td align="center" colspan="2">
        <br>
        <a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>
        <% If vFormOK Then %> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:jconfirm('ExamView.asp?vDelModId=<%=vModId%>','<%=Server.HtmlEncode("Ok to delete this exam?")%>')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a>
        <% End If %>
        <br><br><a href="ExamList.asp">Exam List</a><br>&nbsp;
      </td>
    </tr>


  </table>

  



  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>
