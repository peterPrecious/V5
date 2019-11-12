<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Skil.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %> <%
    Dim vFunction
    vFunction = ""
  
    '...update tables
    If Request("vFunction") = "add" Then
      sExtractSkil
      sInsertSkil
    ElseIf Request("vFunction") = "edit" Then
      sExtractSkil
      sUpdateSkil
    ElseIf Len(Request("vDelSkilId")) > 0 Then 
      vSkil_Id = Request("vDelSkilId")
      sDeleteSkil
    End If  
      
    If Len(Request("vForm")) = 0 Or Request("vFunction") = "del" Then
  %>
  <form method="POST" action="SkilEdit.asp" target="_self">
    <table border="0" width="100%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td>
        <h1 align="center">The Skills Table</h1>
        <h2>Click &quot;add&quot; to enter a new unique Skills Id OR click on an existing Skills Id you wish to edit or delete.&nbsp; If you try to re-enter an existing Skill, the system will ignore the action.&nbsp; Note, once entered, it cannot be modified - only deleted and re-entered.&nbsp; </h2></td>
        <th align="right" nowrap valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp; Skills Id :&nbsp; <input type="text" name="vSkil_Id" size="20" maxlength="32"><input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="I1" type="image"> </th>
      </tr>
    </table>
    <input type="Hidden" name="vForm" value="Y">
  </form>
  <!---Edit List-->
  <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" height="26">
    <tr>
      <th align="left" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Skills Id</th>
      <th align="left" height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Description</th>
    </tr>
    <%
        '...read Skil
        sGetSkil_Rs
        Do While Not oRs4.Eof
         sReadSkil
      %>
    <tr>
      <td><a href="SkilEdit.asp?vEditSkilId=<%=vSkil_Id%>&vForm=n"><%=vSkil_Id%></a>&nbsp; </td>
      <td><%=vSkil_Desc%>&nbsp; </td>
    </tr>
    <%  
          oRs4.MoveNext
        Loop
        Set oRs4 = Nothing
        sCloseDb4   
      %>
    <tr>
      <td colspan="2" align="center"><br><br><a href="JobsEdit.asp">Jobs Table</a>&nbsp; |&nbsp; <a href="CritEdit.asp">Criteria Table</a><br><br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp;</td>
    </tr>
  </table>
  <%
    Else
      If Len(Request.Form("vAddSkilId")) = 0 And Len(Request.Form("vForm")) > 0 Then 
        vSkil_Id = fNoQuote(Request.Form("vSkil_Id"))
        vFunction = "add"
      ElseIf Len(Request.QueryString("vEditSkilId")) > 0 Then 
        vSkil_Id = Request.QueryString("vEditSkilId")
        vFunction = "edit"
      Else
         Response.Redirect "SkilEdit.asp"          
      End If
  
      '...get the values (even if trying to add)
      If vSkil_Id <> "" Then sGetSkil 
  %>
  <form method="POST" action="SkilEdit.asp" target="_self">
    <input type="Hidden" name="vFunction" value="<%=vFunction%>">
    <input type="Hidden" name="vSkil_Id" value="<%=vSkil_Id%>">
    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2">
        <h1 align="center">Edit the Skills Table</h1>
        <h2 align="left">The Skills table simply allows you to describe the Skill that you have added.&nbsp; You can also edit or delete an existing Skill.&nbsp; Click update when finished.<br>&nbsp;</h2></td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" bordercolor="#DDEEF9">Skills Id :</th>
        <td width="70%" valign="top" bordercolor="#DDEEF9">
        <h1><%=vSkil_Id%></h1>
        </td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top">Description :</th>
        <td width="70%"><textarea rows="4" name="vSkil_Desc" cols="46"><%=vSkil_Desc%></textarea></td>
      </tr>
      <tr>
        <td align="center" width="100%" valign="top" colspan="2" bordercolor="#DDEEF9"><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br><br>Ensure you do <b>NOT </b>delete any skills that may be in use, ie assigned to certain jobs...<a href="javascript:jconfirm('SkilEdit.asp?vDelSkilId=<%=vSkil_Id%>&vFunction=del', '<!--[[-->Ok to delete?<!--]]-->')"><img border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif"></a> <p align="center"><a href="SkilEdit.asp">Skills List</a>&nbsp; |&nbsp; <a href="JobsEdit.asp">Jobs Table</a>&nbsp; |&nbsp; <a href="CritEdit.asp">Criteria Table</a><br>&nbsp;</p></td>
      </tr>
    </table>
  </form>
  <%
    End If  
  %>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
