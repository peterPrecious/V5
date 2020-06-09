<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->

<%
  Dim vModId, vCertOK, vCertEdit
  Dim vStr, aQue, vQue, aAns, vAns, vChecked, vDelModId
  Const cAlpha = "abcdefg"

  '...delete test
  vDelModId = Ucase(Request("vDelModId"))
  If Len(vDelModId) = 6 Then
    sDeleteTest vDelModId
    Response.Redirect "TestMenu.asp"
  End If

  '...must be 6 chars with trailing en/fr/es/pt/pl
  vModId = Ucase(Request("vModId"))
  If Len(vModId) <> 6 Or Instr("EN FR ES PT PL", Right(vModId, 2)) = 0 Or Not IsNumeric(Left(vModId, 4)) Then 
    Response.Redirect "TestMenu.asp"
  End If

  '...Get the test, or initialize a new test if not there
  vStr = GetStr (vModId)
  aQue = Split(vStr,"~~"): vQue = Ubound(aQue)  
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
  
  <form method="POST" action="TestSave.asp" target="_self">
    <input type="hidden" name="vModId" value="<%=vModId%>">
    <table border="1" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <td valign="top" colspan="2">
        <h1>Edit Self Assessment - <%=vModId%></h1>
        <p>NOTE:</p>
        <ul>
          <li>You can enter up to 20 Questions, each with up to 6 Answers.</li>
          <li>Leave trailing questions completely empty if not required.</li>
          <li>Click radio button left of correct answer.</li>
          <li>Click save at bottom and view revised test.</li>
        </ul>
        <h6>If you wish to delete this test, click on the &quot;delete&quot; button at the bottom of the page.&nbsp; NOTE: deleting a test is an irreversible action!</h6>
        </td>
      </tr>
      <%  
        For i = 0 To vQue - 1 
          aAns = Split(aQue(i),"||"): vAns = Ubound(aAns) 
      %>
      <tr>
        <td bgcolor="#DDEEF9" valign="top" width="30"><br><%=i+1%>.</td>
        <td bgcolor="#DDEEF9" valign="top"><br><input type="text" size="73" name="Q<%=right("00" & i+1,2)%>" value="<%=aAns(0)%>"></td>
      </tr>
      <%
          For j = 2 To vAns - 1  
            vChecked = ""       
            If IsNumeric(aAns(1)) Then
              If cInt(j-1) = cInt(aAns(1)) Then vChecked = " Checked" 
            End If
      %>
      <tr>
        <td valign="top" align="right" width="30"><input name="A<%=right("00" & i+1,2)%>" type="radio" value="<%=j-1%>" <%=vchecked%>></td>
        <td valign="top"><input type="text" size="68" name="Q<%=right("00" & i+1,2) & ucase(mId(cAlpha,j-1,1))%>" value="<%=aAns(j)%>"></td>
      </tr>
      <%  
          Next
        Next
      %>
    </table>
    <center><br><input border="0" src="../Images/Buttons/Save_<%=svLang%>.gif" name="I1" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
   
    <a href="javascript:jconfirm('TestEdit.asp?vDelModId=<%=vModId%>','Ok to delete this test?')">
    
    <img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a><p><a href="TestMenu.asp">Self Assessment Menu</a></p></center>


  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>
