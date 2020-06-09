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
  <title>UsersBulkInput</title>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>

  <script>
    function validate(theForm) {
      var id = theForm.tList.value.toUpperCase();
      if (id.match(rePassword)==null) {
        alert("You are using certain characters that are not allowed.\n\nEnsure all fields, particularly the Learner Ids (and Passwords) only use\nA-Z 0-9 and !@$%^*()_+-{}[];<>,.");
        theForm.tList.focus();
        return (false);
      }  
    }  
  </script>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>Upload Learner Profiles (Basic)</h1>
  <div class="c2">Enter all the fields carefully then click <b>Continue</b>.&ensp;You will be returned to this page if you neglect or enter invalid or duplicate fields.&nbsp; Do not upload more than 500 names at a time!&nbsp; Note all learners must be new and have a unique Learner ID.&nbsp; If you are inputting names from an Excel spreadsheet, highlight the values (without any headers) then pasted them into the Bulk Input text box using the default Field and Record Delimiters.</div>
  <h6>Note: this service should only be used for sites containing less than 500 learners.&ensp;For larger uploads Advanced Upload.</h6>
  <% If Len(Request.QueryString("vErrMess")) > 0 Then %>
  <h6><br />Error...<br /><%=Request.QueryString("vErrMess")%></h6>
  <% End If %>

<!--  <form method="POST" action="UsersBulkInputVerify.asp" onsubmit="return validate(this)">-->
  <form method="POST" action="UsersBulkInputVerify.asp">
    <table class="table">

      <tr>
        <th>Fields :</th>
        <td>Select the fields that you wish to import (note the first is mandatory).<br><br>&nbsp;<img border="0" src="../images/Icons/CheckMark.gif">
          Learner Id (called Password in Self Service Accounts)<br>
          <blockquote style="color: red;">
            Must be unique using only English alpha, numeric and !@$%^*()_+-{}[];<>,.: characters.&nbsp;Value is NOT case sensitive.
          </blockquote>
          <input type="checkbox" name="cFname" value="True" <%=vfname%>>First name<br>
          <input type="checkbox" name="cLname" value="True" <%=vlname%>>Last name<br>
          <input type="checkbox" name="cEmail" value="True" <%=vemail%>>Email Address (must be unique or left empty)<br>
          <input type="checkbox" name="cPassw" value="True" <%=vpassw%>>Password (available in certain Custom Accounts with same constraints as Learner Id)<br>
          <input type="checkbox" name="cProgs" value="True" <%=vprogs%>>Programs (if more than one, separate with spaces, ie P1234EN P1235EN)<br>
          <input type="checkbox" name="cMemos" value="True" <%=vmemos%>>Memo (separate items with pipes, ie Ontario|Canada)
        </td>
      </tr>

      <tr>
        <th>Group :</th>
        <td>If groups are used, select the group that applies to all uploaded learners,<br>or you can assign a group individually after input via the Learner Profile.<br>Otherwise leave as &quot;All&quot;<br>
          <select size="5" name="dCrit" multiple><%=fCriteriaList(svCustAcctId, vCrit)%></select>
        </td>
      </tr>

      <tr>
        <th>Field Delimiter :</th>
        <td>Select the delimiter that will separate each field. Do not confuse tabs with spaces.<br>If you are cutting/pasting from a spreadsheet use the default delimiters.<br>
          <input type="radio" name="oField" value="tab" <%=vfieldopt5%>>Tab (default)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ie john.jenkins&nbsp;&nbsp;&nbsp; john&nbsp;&nbsp;&nbsp; jenkins&nbsp;&nbsp;&nbsp; jjenkins@email.com<br>
          <input type="radio" name="oField" value="comma" <%=vfieldopt1%>>Comma/CSV&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ie john.jenkins,john,jenkins,jjenkins@email.com</td>
      </tr>
      <tr>
        <th>Record Delimiter :</th>
        <td>&nbsp;Ensure your selection is different from the field delimiter:<br>
          <input type="radio" name="oRecord" value="enter" <%=vrecopt6%>>Enter Key / CR (default)<br>
          <input type="radio" name="oRecord" value="comma" <%=vrecopt1%>>Comma</td>
      </tr>
      <tr>
        <th>Bulk Input :</th>
        <td>
          <textarea rows="19" id="tList" name="tList" cols="77"><%=Session("BulkImportList")%></textarea><br>Maximum 500 names.&nbsp; Ensure there are no trailing spaces or lines after the last name.
        </td>
      </tr>
    </table>

    <div style="text-align: center;">
      <input type="submit" value="Continue" name="bContinue" class="button"><br /><br />
      <a href="Users.asp">Learner List</a>
    </div>

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
