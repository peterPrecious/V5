<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Request("vDelete").Count = 1 Then
    vMemb_No = Request("vDelete")
    sDeleteMemb 

  ElseIf Request("vInsert").Count = 1 Then
    vMemb_AcctId = svCustAcctId
    vMemb_Level  = 5
    vMemb_Id = Request("vInsert")
    sAddMemb svCustAcctId

  ElseIf Request("vDeleteALL").Count = 1 Then
    vMemb_Id = Request("vDeleteALL")
    sDeleteMembAllById vMemb_Id

  End If

%>
<html>

<head>
  <meta charset="UTF-8">
  <title>ManageAdmins</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Manage <%=svCustId %> Administrators</h1>
  <h2 class="c2">If the Administrator Id (ie Level 5, non-internal) exists in this account it will be displayed in bold allowing you to <b>Delete</b> it from this Account. Otherwise you can <b>Insert</b> it into this Account. </h2>
  <h3 class="c3"><span class="red">Clicking <b>Delete ALL</b> erases this Id from ALL Accounts therefore it will no longer appear as an option for future Inserts.</span></h3>

  <form method="POST" action="ManageAdmins.asp">
    <table class="table">
      <tr>
        <td class="rowshade">Administrator ID</td>
        <td class="rowshade">Action</td>
        <td class="rowshade">Occurs In Accounts</td>
      </tr>
      <%
          vSql = "SELECT DISTINCT Memb_Id FROM Memb WHERE Memb_Level = 5 AND Memb_Internal = 0 AND Memb_Active = 1 ORDER BY Memb_Id"
          sOpenDb
          Set oRs = oDb.Execute(vSql)
          Do While Not oRs.Eof 
            vSql  = "SELECT Memb_No FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & oRs("Memb_Id") & "'"
            Set oRs2 = oDb.Execute(vSql)
      %>
      <tr>
        <td><%=fIf(oRs2.Eof,"","<b>")%><%=oRs("Memb_Id")%><%=fIf(oRs2.Eof,"","</b>")%></td>
        <td style="white-space:nowrap">
          <% If oRs2.Eof Then %>
          <input type="button" value="Insert" name="bInsert" class="button085" onclick="location.href = 'ManageAdmins.asp?vInsert=<%=Server.UrlEncode(oRs("Memb_Id"))%>'">
          <% Else  %>
          <input type="button" value="Delete" name="bDelete" class="button085" onclick="jconfirm('ManageAdmins.asp?vDelete=<%=oRs2("Memb_No")%>', 'Delete ID in this Account?')">
          <% End If %>
          <input type="button" value="Delete ALL" name="bDeleteAll" class="button085" onclick="jconfirm('ManageAdmins.asp?vDeleteALL=<%=Server.UrlEncode(oRs("Memb_Id"))%>', 'Delete this ID in ALL Accounts?')">
        </td>
        <td><%=fMembIdAll(oRs("Memb_Id"))%></td>
      </tr>
      <%
            oRs.MoveNext
          Loop
          Set oRs = Nothing
          Set oRs2 = Nothing
          sCloseDb
      %>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
