<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Sub sRefreshSearchResults()
    '...this routine simulates a submit on the Search Page (simple form) to refresh Search Results
    Response.Write "<form method='POST' action='LogsEdit.asp' name='fSearch'>"
    Response.Write "  <input type='hidden' name='vLogs_MembNo' value='" & Session("vLogs_MembNo") & "'>"
    Response.Write "  <input type='hidden' name='vLogs_Type' value='" & Session("vLogs_Type") & "'>"
    Response.Write "  <input type='hidden' name='vStrDate' value='" & Session("vLogsStartDate") & "'>"
    Response.Write "  <input type='hidden' name='vEndDate' value='" & Session("vLogsEndDate") & "'>"
    Response.Write "  <input type='hidden' name='bSearch.x' value='1'>"
    Response.Write "</form>"
    Response.Write "<script>fSearch.submit()</script>"
  End Sub
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Edit Log Entries</title>
  <script src="/V5/Inc/Functions.js"></script>
  <script language="vbscript">
    Function ValidateDate(vDate)
      ValidateDate = IsDate(vDate)
    End Function
  </script>

  <script>

    var vAction = ''
    
    // Validation for Search part of form
    function fSearch_Validator(theForm) {
      // Validate the Member No
      // Check if empty
      if (theForm.vLogs_MembNo.value == "") {
        alert("Please enter a value for the \"Member No\" field.");
        theForm.vLogs_MembNo.focus();
        return (false);
      }
      // Check if blank
      if (theForm.vLogs_MembNo.value.length < 1) {
        alert("Please enter at least 1 characters in the \"Member No\" field.");
        theForm.vLogs_MembNo.focus();
        return (false);
      }
      // Check if proper length
      if (theForm.vLogs_MembNo.value.length > 8) {
        alert("Please enter at most 8 characters in the \"Member No\" field.");
        theForm.vLogs_MembNo.focus();
        return (false);
      }
      // Check for only numbers
      var checkOK = "0123456789-.,";
      var checkStr = theForm.vLogs_MembNo.value;
      var allValid = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0;  i < checkStr.length;  i++) {
        ch = checkStr.charAt(i);
        for (j = 0;  j < checkOK.length;  j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        if (ch == ".") {
          allNum += ".";
          decPoints++;
        }
        else if (ch != ",")
          allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"Member No\" field.");
        theForm.vLogs_MembNo.focus();
        return (false);
      }
      if (decPoints > 1) {
        alert("Please enter a valid number in the \"vLogs_MembNo\" field.");
        theForm.vLogs_MembNo.focus();
        return (false);
      }
    
      // Validate the Start/End Dates
      if (theForm.vStrDate.value != '')
        if (!ValidateDate(theForm.vStrDate.value)) {
          alert('Please enter a valid Start Date.')
          return (false);
        }
      if (theForm.vEndDate.value != '')
        if (!ValidateDate(theForm.vEndDate.value)) {
          alert('Please enter a valid End Date.')
          return (false);
        }
    
      // If we get here, all data correct...continue with Results
      return (true);
    }
    
    // Validation for Update/Delete/Add part of form
    function fUpdate_Validator(theForm) {
      // Vaidate the Logs Posted Date
      if (theForm.vLogsPosted.value.length == 0) {
        alert('Please enter a value for Date.')
        return (false);
        }
      if (theForm.vLogsPosted.value != '')
        if (!ValidateDate(theForm.vLogsPosted.value)) {
          alert('Please enter a valid Date.')
          return (false);
        }
      // Vaidate Logs Item...cannot be empty
      if (theForm.vLogsItem.value.length == 0) {
        alert('Please enter a value for Item.')
        return (false);
      }
    
      // If Delete, confirm
      if (vAction=='d')
        if (!confirm('Are you sure you want to Delete this entry?')) return (false)
    
      // If we get here, all data correct...continue with Results
      return (true);
    }
    
  </script>
</head>

<body>

  <% Server.Execute vShellHi %> 

  <% If Request.Form = Empty Then %>
  <table width="100%" border="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td><h1 align="center">Edit Log Entries</h1><h2>This allows you to add, edit or delete a log entry <b>within this account</b>.&nbsp; You must narrow down the selection to the Learner No, Transaction Type and you can optionally select a time period.&nbsp; Once entered, click &quot;go&quot; and any logs on file will be presented for your review.</h2>
      <h5 align="center">Be very careful - these actions cannot be reversed !</h5>
      <p>&nbsp;</p></td>
    </tr>
    <tr>
      <td align="center"><p>&nbsp;</p>
      <div align="center">
        <form method="POST" action="LogsEdit.asp" onsubmit="return fSearch_Validator(this)" name="fSearch">
          <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
            <tr>
              <th nowrap colspan="2" bgcolor="#DDEEF9" height="30">Select log entries for review...</th>
            </tr>
            <tr>
              <th nowrap align="right" valign="top" width="35%">Account Id :</th>
              <td valign="top" width="65%"><p class="c1"><%=svCustAcctId%> </p></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top" width="35%">Learner No :</th>
              <td valign="top" width="65%"><input type="text" name="vLogs_MembNo" size="20" maxlength="8"> Mandatory.&nbsp; Note: learner no can be derived from the learner report.</td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top" width="35%">Type:</th>
              <td valign="top" width="65%">
              <select size="1" name="vLogs_Type">
              <option value="S">S: Completion</option>
              <option value="B">B: Bookmark</option>
              <option value="P">P: TimeSpent</option>
              <option value="E">E: Exam Attempts</option>
              <option value="H">H: Exam History</option>
              <option value="U">U: Survey Results</option>
              <option value="T">T: Test/Exam Grades</option>
              </select></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top" width="35%">Start Date:</th>
              <td valign="top" width="65%"><input type="text" name="vStrDate" size="20"> ie. Jan 1, 2004.&nbsp; Leave empty to select first date on file.</td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top" width="35%">End Date:</th>
              <td valign="top" width="65%"><input type="text" name="vEndDate" size="20"> ie. Dec 31, 2004. Leave empty to select last date on file.</td>
            </tr>
            <tr>
              <td align="center" colspan="2"><br><input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" name="bSearch" type="image"><br>&nbsp;</td>
            </tr>
          </table>
        </form>
      </div>
      </td>
    </tr>
  </table>

  <% 
    ElseIf Request.Form("bSearch.x") > 0 Then
      Session("vLogs_MembNo")   = Request("vLogs_MembNo")
      Session("vLogs_Type")     = Request("vLogs_Type")
      Session("vLogsStartDate") = Request("vStrDate")
      Session("vLogsEndDate")   = Request("vEndDate")

      '...query the Logs table with entered items
      sOpenDb
      vSql = "SELECT * FROM Logs WITH (nolock) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_MembNo = " & Session("vLogs_MembNo") & " AND Logs_Type = '" & Session("vLogs_Type") & "'"
      If Len(Session("vLogsStartDate")) > 0 Then vSql = vSql & " AND Logs_Posted >= '" & Session("vLogsStartDate") & "'"
      If Len(Session("vLogsEndDate")) > 0 Then vSql = vSql & " AND Logs_Posted <= '" & Session("vLogsEndDate") & "'"
      vSql = vSql & " ORDER BY Logs_Posted "
'     sDebug
      Set oRs = oDb.Execute (vSql)
  %>

  <table border="1" style="border-collapse: collapse" id="table2" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
    <tr>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Acct Id </th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Learner No</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Type</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Date</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Item *</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Action</th>
    </tr>
      <!-- Edit an entry portion -->
      <tr>
        <td align="left" colspan="6"><p align="center"><br>To edit an existing entry, edit the date and/or item , then click &quot;add&quot;, or simply &quot;delete&quot;.&nbsp; <br>Remember that deleting a log item is an irreversible action.<br>&nbsp;</td>
    </tr>
      <%
        '...if no entries found, display a message in red
        If oRs.EOF Then
      %>
      <tr>
        <td align="left" colspan="6">
        <h6 align="center">No Items found!</h6>
        </td>
    </tr>
      <%
        End If
       
        '...loop through all records and display Edit/Delete items for each
        While Not oRs.EOF
      %>
      <form method="POST" action="LogsEdit.asp" onsubmit="return fUpdate_Validator(this)" name="fResults">
        <tr>
          <th nowrap align="left"><p class="c1"><%=svCustAcctId%></p></th>
          <th nowrap align="left"><p class="c1"><%=Session("vLogs_MembNo")%></p></th>
          <th nowrap align="left"><p class="c1"><%=Session("vLogs_Type")%></p></th>
          <th nowrap align="left"><input type="text" name="vLogsPosted" size="20" maxlength="22" value="<%=oRs("Logs_Posted")%>"></th>
          <td align="left"><input type="text" name="vLogsItem" size="26" maxlength="8000" value="<%=oRs("Logs_Item")%>"></td><input type="hidden" name="vLogsNo" value="<%=oRs("Logs_No")%>">
          <td align="left"><input onclick="vAction='u'" border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="bUpdate" type="image"> <input onclick="vAction='d'" border="0" src="../Images/Buttons/Delete_<%=svLang%>.gif" name="bDelete" type="image"> </td>
        </tr>
    </form>
      <%
          oRs.MoveNext
        Wend
        sCloseDb
      %>
      <!-- Add a New entry portion -->
      <tr>
        <td align="left" colspan="6"><p align="center"><br>To add a new log entry, enter date and the item below, then click &quot;add&quot; <br>&nbsp;</td>
    </tr>
    <form method="POST" action="LogsEdit.asp" onsubmit="return fUpdate_Validator(this)" name="fAdd">
      <tr>
        <th nowrap align="left"><p class="c1"><%=svCustAcctId%></p></th>
        <th nowrap align="left"><p class="c1"><%=Session("vLogs_MembNo")%></p></th>
        <th nowrap align="left"><p class="c1"><%=Session("vLogs_Type")%></p></th>
        <th nowrap align="left"><input type="text" name="vLogsPosted" size="20" maxlength="22"></th>
        <td align="left"><input type="text" name="vLogsItem" size="26" maxlength="8000"></td>
        <td align="left"><input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="bAdd" type="image"> </td>
      </tr>
    </form>
    <tr>
      <th nowrap colspan="6">&nbsp;<p><a href="LogsEdit.asp">Restart</a></p><p>&nbsp;</p></th>
    </tr>
  </table><p>&nbsp;</p>
  <table border="1" style="border-collapse: collapse" id="table3" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
    <tr>
      <th colspan="4" nowrap height="30"><h1>Item Formats</h1></th>
    </tr>
    <tr>
      <th nowrap align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Type of Item</th>
      <th nowrap align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Format</th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">#Chars</th>
      <th nowrap align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Example</th>
    </tr>
    <tr>
      <td valign="top">B - Bookmarks</td>
      <td valign="top">ModId_PageNo</td>
      <td valign="top" align="center">10</td>
      <td valign="top">5500EN_009</td>
    </tr>
    <tr>
      <td valign="top">P - Time Spent</td>
      <td valign="top">ProgId|ModId_Minutes</td>
      <td valign="top" align="center">21</td>
      <td valign="top">P3426EN|5390EN_000033</td>
    </tr>
    <tr>
      <td valign="top">S - Completion</td>
      <td valign="top">ProgId or ModId</td>
      <td valign="top" align="center">6 or 7</td>
      <td valign="top">P3426EN or 5390EN</td>
    </tr>
    <tr>
      <td valign="top">H - Exam History</td>
      <td valign="top">ExamId_AttemptNo_Ques...</td>
      <td valign="top" align="center">n/a</td>
      <td valign="top">1001EN_1_@@Viral marketing works because <br>&amp;#8230;. @@Costs involved in setting up a <br>website include technological...</td>
    </tr>
    <tr>
      <td valign="top">E - Exam Attempts</td>
      <td valign="top">ExamId_AttemptNo_Answers...</td>
      <td valign="top" align="center">n/a</td>
      <td valign="top">1001EN_1_1_7_0_6_0_56_0_99_0_31_0_0</td>
    </tr>
    <tr>
      <td valign="top">T - Test/Exam Results </td>
      <td valign="top">TestId_Mark or <br>ExamId_Attempt_Mark</td>
      <td valign="top" align="center">8 or 10</td>
      <td valign="top">0002EN_100<br>&nbsp; or<br>1001EN_1_088</td>
    </tr>
    <tr>
      <td valign="top">A - Assessment Results by Question</td>
      <td valign="top">ModuleId_Results</td>
      <td valign="top" align="center">n/a</td>
      <td valign="top">5390EN_4|Bite me</td>
    </tr>
    <tr>
      <td valign="top">U - Survey Results</td>
      <td valign="top">ProgramId|ModuleId_Results</td>
      <td valign="top" align="center">n/a</td>
      <td valign="top">P3426EN|5390EN_1|4|Bite me</td>
    </tr>
  </table>
  <%
    '...Update the requested item
    ElseIf Request.Form("bUpdate.x") > 0 Then
      '...update item
      sOpenDb
      vSql = "UPDATE Logs SET Logs_Posted = '" & Request.Form("vLogsPosted") & "', Logs_Item = '" & Request.Form("vLogsItem") & "' WHERE Logs_No = " & Request.Form("vLogsNo")

      oDb.Execute (vSql)
      sCloseDb
      '...refresh Results
      sRefreshSearchResults()
    '...Delete the requested item
    ElseIf Request.Form("bDelete.x") > 0 Then
      '...delete item
      sOpenDb
      vSql = "DELETE Logs WHERE Logs_No = " & Request.Form("vLogsNo")

      oDb.Execute (vSql)
      sCloseDb
      '...refresh Results
      sRefreshSearchResults()
    '...Add the new item
    ElseIf Request.Form("bAdd.x") > 0 Then
      '...add item
      sOpenDb
      vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES ("
      vSql = vSql & "'" & svCustAcctId & "', '" & Session("vLogs_Type") & "', " & Session("vLogs_MembNo") & ", '" & Request.Form("vLogsPosted") & "', '" & Request.Form("vLogsItem") & "')"

      oDb.Execute (vSql)
      sCloseDb
      '...refresh Results
      sRefreshSearchResults()
    End If 
  %> 
  
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


