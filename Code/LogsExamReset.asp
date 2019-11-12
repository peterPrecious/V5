<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vMsg, vReturn, vMembNo, vFirstName, vLastName, vModId, vNext

  vNext = fDefault(Request("vNext"), "LogReport5.asp")

  '...values must come from the Assessment report via URL which will not interfere with the form values 
  If Request.Form.Count > 0 Then
    vMembNo		 = Request("vMembNo")
    vFirstName = Request("vFirstName")
    vLastName  = Request("vLastName")
    vModId     = Request("vModId")
    sReset
    Response.Redirect "Error.asp?vClose=N&vErr=" & Server.HtmlEncode(vMsg) & "&vReturn=" & Server.UrlEncode(vReturn)
  End If


  Sub sReset   
    sOpenDb3
    vSql =        " SELECT * FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
    vSql = vSql & " WHERE Logs_AcctId = '" & svCustAcctId & "'"
    vSql = vSql & " AND Logs_MembNo = " & vMembNo 
    vSql = vSql & " AND Memb_FirstName = '" & fUnQuote(vFirstName) & "'"
    vSql = vSql & " AND Memb_LastName = '" & fUnQuote(vLastName) & "'"
    vSql = vSql & " AND Left(Logs_Item, 6) = '" & vModId & "'"
    vSql = vSql & " AND (Logs_Type = 'E' OR Logs_Type = 'H' OR Logs_Type = 'T' OR Logs_Type = 'L')"
'   sDebug
    Set oRs3 = oDb3.Execute (vSql)
    If oRs3.Eof Then 
      vReturn = ""
      vMsg = "There were either no exam entries to reset/deleted or the exam entries could NOT be reset/deleted successfully.  Please submit details to support@vubiz.com"
    Else
      vSql =        " DELETE Logs "
      vSql = vSql & " WHERE Logs_AcctId = '" & svCustAcctId & "'"
      vSql = vSql & " AND Logs_MembNo = " & vMembNo 
      vSql = vSql & " AND Left(Logs_Item, 6) = '" & vModId & "'"
      vSql = vSql & " AND (Logs_Type = 'E' OR Logs_Type = 'H' OR Logs_Type = 'T' OR Logs_Type = 'L')"
  '   sDebug
      oDb3.Execute (vSql)
      vReturn = vNext
      vMsg = "The exam entries were reset/deleted successfully."
    End If

    Set oRs3 = Nothing
    sCloseDb3
  End Sub
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Edit Log Entries</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <form method="POST" action="LogsExamReset.asp">
      <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
        <tr>
          <td align="center" valign="top"><h1 align="center">Reset an Assessment</h1><p class="c2">This allows you to <font color="#FF0000">DELETE/RESET</font> all log entries for this assessment for this learner.&nbsp; <br><span class="c6">Note: this action cannot be reversed!</span></p>
          <table border="0" id="table2" cellspacing="0" cellpadding="2">
            <tr>
              <th nowrap align="right" valign="top">Vubiz Learner No :</th>
              <td><%=Request("vMembNo")%></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">First Name :</th>
              <td><%=Request("vFirstName")%></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Last Name :</th>
              <td><%=Request("vLastName")%></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Assessment Id :</th>
              <td><%=Request("vModId")%></td>
            </tr>
          </table>
          <h2>&nbsp;</h2><h2>If you do NOT wish to DELETE/RESET the logs for this assessment click <b>Return</b>.</h2><p>
            <input onclick="location.href='<%=vNext%>'" type="button" value="Return" name="bReturn" id="bReturn"class="button"></p><p>&nbsp;</p><h2>If you DO wish to DELETE/RESET the logs for this assessment click <b>Reset</b>.<br><br>
            <input type="submit" value="Reset" name="bContinue" class="button"><br>&nbsp;</h2></td>
        </tr>
      </table>
      <input type="hidden" name="vMembNo"    value="<%=Request("vMembNo")%>">
      <input type="hidden" name="vFirstName" value="<%=Request("vFirstName")%>">
      <input type="hidden" name="vLastName"  value="<%=Request("vLastName")%>">
      <input type="hidden" name="vModId"     value="<%=Request("vModId")%>">
      <input type="hidden" name="vNext"      value="<%=vNext%>">

    </form>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

