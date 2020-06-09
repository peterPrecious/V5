<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  If Request("vInactivate").Count = 1 Then
    vMemb_No = Request("vInactivate")
    sInactivateMemb 

  ElseIf Request("vInsert").Count = 1 Then
    vMemb_AcctId = svCustAcctId
    vMemb_Level  = 4
    vMemb_Auth  = 1
    vMemb_Id = Request("vInsert")
    sAddMemb svCustAcctId

  ElseIf Request("vInactivateALL").Count = 1 Then
    vMemb_Id = Request("vInactivateALL")
    sInactivateMembAllById vMemb_Id

  End If

%>
<html>

<head>
  <meta charset="UTF-8">
  <title>ManageAuthors</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Manage <%=svCustId %> Authors</h1>
  <h2 class="c2">If the Authoring Id exists in this account it will be displayed in bold allowing you to <b>Inactive</b> it in this Account. Otherwise you can <b>Insert</b> it into this Account. </h2>
  <h3 class="c3"><span class="red">Clicking <b>Inactive ALL</b> inactivates this Id from ALL Accounts therefore it will no longer appear as an option for future Inserts.</span></h3>

  <form method="POST" action="ManageAuthors.asp">
    <table class="table">
      <tr>
        <td class="rowshade">Author ID</td>
        <td class="rowshade">Action</td>
        <td class="rowshade">Occurs In Accounts</td>
      </tr>
      <%
          vSql = "SELECT DISTINCT Memb_Id FROM Memb WHERE Memb_Level = 4 AND Memb_Auth = 1 AND Memb_Active = 1 AND Memb_Id NOT LIKE '%_SALES' ORDER BY Memb_Id"
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
          <input type="button" value="Insert" name="bInsert" class="button085" onclick="location.href = 'ManageAuthors.asp?vInsert=<%=Server.UrlEncode(oRs("Memb_Id"))%>'">
          <% Else  %>
          <input type="button" value="Inactivate" name="bInactivate" class="button085" onclick="jconfirm('ManageAuthors.asp?vInactivate=<%=oRs2("Memb_No")%>', 'Inactivate ID in this Account?')">
          <% End If %>
          <input type="button" value="Inactivate ALL" name="bInactivateAll" class="button085" onclick="jconfirm('ManageAuthors.asp?vInactivateALL=<%=Server.UrlEncode(oRs("Memb_Id"))%>', 'Inactivate this ID in ALL Accounts?')">
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


