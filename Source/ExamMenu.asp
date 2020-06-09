<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"--><%
  Dim vModID
  If Request.Form("vModID").Count > 0 Then
    vModID = Request.Form("vModID")
    If Request.Form("vEdit").Count > 0 Then Response.Redirect "ExamEdit.asp?vModID=" & vModID
  End If
  Session("EditTest") = False
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
  <form method="POST" action="ExamEdit.asp" target="_self">
    <table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3">
      <tr>
        <td colspan="3">
        <h1>Exam Maintenance</h1>
        <p>Either enter a new exam you wish to add (ie 1234EN) then click &quot;add&quot;, OR select an existing exam you wish to edit then click &quot;edit&quot;.<br>&nbsp;&nbsp;&nbsp; <% 
          If Len(Request.QueryString("vMess")) > 0 Then
            Response.Write "<br><b>" & Request.QueryString("vMess") & "</b><br><br>"
          End If
        %> </p></td>
      </tr>
      <tr>
        <th align="right" nowrap>Name :</th>
        <td valign="middle"><input type="text" name="vAddModID" size="13"></td>
        <td align="right"><select size="1" name="vModID">
        <option selected value="Select">Select Module</option>
        <%=fExamOptionsAll%></select><input border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif" name="I3" type="image"></td>
      </tr>
      <tr>
        <th align="right" nowrap>No. Questions :</th>
        <td><input type="text" name="vNumQ" size="5"> (mandatory)</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <th align="right" nowrap>Exam Title :</th>
        <td><input type="text" name="vTitle" size="25">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/add_<%=svLang%>.gif" name="vAdd" type="image"></td>
        <td>&nbsp;</td>
      </tr>
    </table>
  </form>
  <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br><br><a href="ExamList.asp">Exam List</a><br>&nbsp;</p>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
