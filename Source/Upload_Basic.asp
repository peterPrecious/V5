<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<%
  Dim vFname, vLname, vEmail, vPassw, vProgs, vMemos, vCrit
  Dim vFieldOpt1, vFieldOpt2, vFieldOpt3, vFieldOpt4, vFieldOpt5
  Dim vRecOpt1, vRecOpt2, vRecOpt3, vRecOpt4, vRecOpt5, vRecOpt6

  '...Grab all info if coming back her from an error
  If Request.QueryString("vFname") = "True" Then vFname = "checked"
  If Request.QueryString("vLname") = "True" Then vLname = "checked"
  If Request.QueryString("vEmail") = "True" Then vEmail = "checked"
  If Request.QueryString("vPassw") = "True" Then vPassw = "checked"
  If Request.QueryString("vProgs") = "True" Then vProgs = "checked"
  If Request.QueryString("vMemos") = "True" Then vMemos = "checked"

  If Len(Request.QueryString("vCrit")) > 0 Then vCrit = Request.QueryString("vCrit")

  Select Case Request.QueryString("vFieldSep")
    Case "comma"
      vFieldOpt1 = "checked"
    Case "semi"
      vFieldOpt2 = "checked"
    Case "pipe"
      vFieldOpt3 = "checked"
    Case "tilde"
      vFieldOpt4 = "checked"
    Case "tab"
      vFieldOpt5 = "checked"
    Case Else
      vFieldOpt5 = "checked"
  End Select

  Select Case Request.QueryString("vRecordSep")
    Case "comma"
      vRecOpt1 = "checked"
    Case "semi"
      vRecOpt2 = "checked"
    Case "pipe"
      vRecOpt3 = "checked"
    Case "tilde"
      vRecOpt4 = "checked"
    Case "tab"
      vRecOpt5 = "checked"
    Case "enter"
      vRecOpt6 = "checked"
    Case Else
      vRecOpt6 = "checked"
  End Select

%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Bulk Upload Users</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <form method="POST" action="Upload_Basic_Verify.asp">
      <tr>
        <td align="center" valign="top" colspan="2"><h1>Upload Learner Profiles (Basic)</h1>
        <h2 align="left">This allows you to upload up to 500 new learners profiles.&nbsp; You cannot upload/overwrite existing learner profiles.&nbsp; Note all learners must have a unique Learner ID.&nbsp; If you are inputting names from an Excel spreadsheet, highlight the values (without any headers) then pasted them into the Bulk Input text box using the default Field and Record Delimiters.&nbsp; Once you have entered all the fields carefully then click <b>Continue</b>. You will be returned to this page if you neglect or enter invalid or duplicate fields.&nbsp; </h2>
        <p align="left">&nbsp;</p>
          <% If Len(Request.QueryString("vErrMess")) > 0 Then %>
          <table border="0" style="border-collapse: collapse" id="table1" cellspacing="1">
            <tr>
              <td><h5><br>Error...</h5><h6><%=Request.QueryString("vErrMess")%></h6></td>
            </tr>
          </table>
          <% End If %> 
        </td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Fields :</th>
        <td valign="top" width="80%">Select the fields that you wish to import.<br><br>&nbsp;<img border="0" src="../images/Icons/CheckMark.gif"> Learner Id (this is mandatory and must be unique)<br>
          <input type="checkbox" name="cFname" value="True" <%=vfname%>>First name<br>
          <input type="checkbox" name="cLname" value="True" <%=vlname%>>Last name<br>
          <input type="checkbox" name="cEmail" value="True" <%=vemail%>>Email Address<br>
          <input type="checkbox" name="cPassw" value="True" <%=vpassw%>>Password (must be upper case)<br>
          <input type="checkbox" name="cProgs" value="True" <%=vprogs%>>Programs (if more than one, separate with spaces, ie P1234EN P1235EN)<br>
          <input type="checkbox" name="cMemos" value="True" <%=vmemos%>>Memo (if more than one, separate with pipes, ie Paris|France)
        </td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Group :</th>
        <td width="80%">If groups are used, select the group that applies to all uploaded learners,<br>or you can assign a group individually after input via the Learner Profile.<br><select size="5" name="dCrit" multiple><%=fCriteriaList(svCustAcctId, vCrit)%></select> </td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Field Delimiter :</th>
        <td valign="top" width="80%">Select the delimiter that will separate each field. Do not confuse tabs with spaces.<br>If you are cutting/pasting from a spreadsheet use the default delimiters.<br>
          <input type="radio" name="oField" value="tab"   <%=vfieldopt5%>>Tab (default)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ie john.jenkins&nbsp;&nbsp;&nbsp; john&nbsp;&nbsp;&nbsp; jenkins&nbsp;&nbsp;&nbsp; jjenkins@email.com<br>
          <input type="radio" name="oField" value="comma" <%=vfieldopt1%>>Comma/CSV&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ie john.jenkins,john,jenkins,jjenkins@email.com</td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Record Delimiter :</th>
        <td valign="top" width="80%">&nbsp;Ensure your selection is different from the field delimiter:<br>
        <input type="radio" name="oRecord" value="enter" <%=vrecopt6%>>Enter Key / CR (default)<br>
        <input type="radio" name="oRecord" value="comma" <%=vrecopt1%>>Comma</td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Bulk Input :</th>
        <td valign="top" width="80%"><h2><textarea rows="19" name="tList" cols="77"><%=Session("BulkImportList")%></textarea><br>Maximum 500 names.&nbsp; Ensure there are no trailing spaces or lines after the last name.</h2></td>
      </tr>
      <tr>
        <td colspan="2" valign="top" align="center">&nbsp;<p><input type="submit" value="Continue" name="bContinue" class="button"></p><h2><a href="Users.asp">Learner List</a></h2><p>&nbsp;</p></td>
      </tr>
    </form>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>