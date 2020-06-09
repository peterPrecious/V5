<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vMessage, vModsId, vModsTitle, vId, vScore, vScoreDate
  vMessage = ""

  If Request("vForm") = "post" Then

    '...check user
    vId    = Ucase(Request("vId"))
    sGetMembById svCustAcctId, vId
    If vMemb_Eof Then
      vMessage = vMessage & "<br>That Learner ID/Password is not on file."
    End If
        
    '...check exam
    vModsId  = Ucase(Request("vModsId"))
    vModsTitle = fModsTitle(vModsId) 
    If Len(vModsTitle) = 0 Then
      vMessage = vMessage & "<br>That assessment is not on file (ie it is not a VuAssess launch module defined in the module table)."
    End If
    
    '...check score
    vScore = Request("vScore")
    If Not IsNumeric(vScore) Then
      vMessage = vMessage & "<br>The score must be a number between 0 and 100."
    ElseIf vScore < 0 or vScore > 100 Then
      vMessage = vMessage & "<br>The score must be a number between 0 and 100."
    End If

    '...check scoredate
    vScoreDate = Request("vScoreDate")
    If fFormatSqlDate(Request("vScoreDate")) = " " Then
      vMessage = vMessage & "<br>The date must be in a valid Enlish date format."
    End If
    
    '...if all ok then update file (note enter bank 0 to show this was an offline entry)
    If vMessage = "" Then
      '...if score is zero then delete any posting for this exam 
      If vScore = 0 Then
        sDeleteLogs vMemb_No, vModsId
        vMessage = "<br>The score for assessment: " & vModsId & " (" & vModsTitle & ")<br>was deleted for user: " & vId & " (" & vMemb_FirstName & " " & vMemb_LastName & ")."
      Else
        sDeleteLogs vMemb_No, vModsId     

        vLogs_Item = vModsId & "_" & Right(000 & vScore, 3)
        vSql = " INSERT INTO Logs " _
             & " (Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo, Logs_Posted) " _
             & " VALUES " _
             & " ('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & vMemb_No & ", '" & vScoreDate & "')"
'       sDebug
        sOpenDb
        oDb.Execute(vSql)
        sCloseDb

        vMessage = "<br>Score of: " & vScore & "% on assessment: " & vModsId & " (" & vModsTitle & ")<br>was posted successfully for user: " & vId & " (" & vMemb_FirstName & " " & vMemb_LastName & ")"
      End If
    End If
   
  End If
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
  <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vId.value == "")
  {
    alert("Please enter a value for the \"Learner Id/Password\" field.");
    theForm.vId.focus();
    return (false);
  }

  if (theForm.vModsId.value == "")
  {
    alert("Please enter a value for the \"Assessment Id\" field.");
    theForm.vModsId.focus();
    return (false);
  }

  if (theForm.vModsId.value.length < 6)
  {
    alert("Please enter at least 6 characters in the \"Assessment Id\" field.");
    theForm.vModsId.focus();
    return (false);
  }

  if (theForm.vModsId.value.length > 6)
  {
    alert("Please enter at most 6 characters in the \"Assessment Id\" field.");
    theForm.vModsId.focus();
    return (false);
  }

  if (theForm.vScoreDate.value == "")
  {
    alert("Please enter a value for the \"Score Date\" field.");
    theForm.vScoreDate.focus();
    return (false);
  }

  if (theForm.vScoreDate.value.length < 11)
  {
    alert("Please enter at least 11 characters in the \"Score Date\" field.");
    theForm.vScoreDate.focus();
    return (false);
  }

  if (theForm.vScoreDate.value.length > 12)
  {
    alert("Please enter at most 12 characters in the \"Score Date\" field.");
    theForm.vScoreDate.focus();
    return (false);
  }

  if (theForm.vScore.value == "")
  {
    alert("Please enter a value for the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  if (theForm.vScore.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  if (theForm.vScore.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vScore.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= 0 && prsVal <= 100))
  {
    alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"100\" in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form1" method="POST" action="PostVuAssessScores.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
      <tr>
        <td width="100%" valign="top" colspan="2" align="center">
        <h1>Post VuAssess Scores</h1>
        <h2 align="left">This allows you to post an assessment score for a learner regardless of when the assessment was taken.&nbsp; If you enter a score for an assessment that has already had a score posted, the latest score will overwrite the original.&nbsp; If you wish to delete a score for an assessment, enter a score of zero and it will be removed.&nbsp; Note: even if there was no previous score to be removed, the message will say that the score was removed.&nbsp; Once a new score is recorded you can check this on the Assessment Report where you can also display the certificate.</h2>
        <% If vMessage <> "" Then %><font color="#FF0000"><%=vMessage%></font><br>&nbsp;<% End If %> 
        </td>
      </tr>
      <tr>
        <th align="right" width="40%" valign="top" nowrap><%=fIf(svCustPwd, "Learner Id", "Learner Password")%> :</th>
        <td width="60%" valign="top"><!--webbot bot="Validation" s-display-name="Learner Id/Password" b-value-required="TRUE" --><input type="text" size="30" name="vId" value="<%=vId%>"></td>
      </tr>
      <tr>
        <th align="right" width="40%" valign="top" nowrap>Assessment Id :</th>
        <td width="60%" valign="top"><!--webbot bot="Validation" s-display-name="Assessment Id" b-value-required="TRUE" i-minimum-length="6" i-maximum-length="6" --><input type="text" size="16" name="vModsId" value="<%=vModsId%>" maxlength="6"> Launch Module Id, ie 1234EN</td>
      </tr>
      <tr>
        <th align="right" width="40%" valign="top" nowrap>Date :</th>
        <td width="60%" valign="top"><!--webbot bot="Validation" s-display-name="Score Date" b-value-required="TRUE" i-minimum-length="11" i-maximum-length="12" --><input type="text" size="16" name="vScoreDate" maxlength="12" value="<%=vScoreDate%>"> In English format, ie: <%=fFormatSqlDate(Now)%></td>
      </tr>
      <tr>
        <th align="right" width="40%" valign="top" nowrap>Score :</th>
        <td width="60%" valign="top"><!--webbot bot="Validation" s-display-name="Score" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="3" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="100" --><input type="text" name="vScore" size="4" maxlength="3" value="<%=vScore%>">%&nbsp; (0-100) <br>Note: entering a value of 0 will remove any existing score on file for this assessment!</td>
      </tr>
      <tr>
        <td align="center" width="100%" valign="top" colspan="2">&nbsp;<p><input type="submit" value="Update" name="bUpdate" class="button"></p><h2><a href="LogReport5.asp">Assessment Report</a></h2></td>
      </tr>
      <input type="hidden" name="vForm" value="post">
    </form>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
