<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vNoIds.value == "")
  {
    alert("Please enter a value for the \"Number of Learner Passwords\" field.");
    theForm.vNoIds.focus();
    return (false);
  }

  if (theForm.vNoIds.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Number of Learner Passwords\" field.");
    theForm.vNoIds.focus();
    return (false);
  }

  if (theForm.vNoIds.value.length > 4)
  {
    alert("Please enter at most 4 characters in the \"Number of Learner Passwords\" field.");
    theForm.vNoIds.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vNoIds.value;
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
    alert("Please enter only digit characters in the \"Number of Learner Passwords\" field.");
    theForm.vNoIds.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= 1 && prsVal <= 1000))
  {
    alert("Please enter a value greater than or equal to \"1\" and less than or equal to \"1000\" in the \"Number of Learner Passwords\" field.");
    theForm.vNoIds.focus();
    return (false);
  }

  if (theForm.vDuration.value.length > 4)
  {
    alert("Please enter at most 4 characters in the \"No of Days\" field.");
    theForm.vDuration.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vDuration.value;
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
    alert("Please enter only digit characters in the \"No of Days\" field.");
    theForm.vDuration.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="IssueIdsOk.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
      <tr>
        <td colspan="2"><h1 align="center">Generate Multiple Learner Passwords</h1><h2>This will generate a Learner Passwords and allow you to the First Name, Last Name and Email address for each learner record created - in the next screen.&nbsp; Enter when this record will expire.&nbsp; It can be, for example, 90 days from today or 90 days from date learner first enters the site.&nbsp; The latter option is used when generating Access Ids for third party ecommerce providers.</h2><h6 align="center">Note: Please conserve resources - do not issue unnecessary passwords.&nbsp; </h6></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Generate : </th>
        <td valign="top"><!--webbot bot="Validation" s-display-name="Number of Learner Passwords" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="4" s-validation-constraint="Greater than or equal to" s-validation-value="1" s-validation-constraint="Less than or equal to" s-validation-value="1000" --><input type="text" name="vNoIds" size="4" value="1" maxlength="4" class="c2"> Password(s)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (Maximum 1000) </td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>All of which can access the follow program(s) : </th>
        <td valign="top"><input type="text" name="vPrograms" size="56" class="c2"><br>Leave empty or enter in CAPS separated by spaces, ie<br>P1002EN P1204EN</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>Access will expire in : </th>
        <td valign="top"><!--webbot bot="Validation" s-display-name="No of Days" s-data-type="Integer" s-number-separators="x" i-maximum-length="4" --><input type="text" name="vDuration" size="4" maxlength="4" value="0" class="c2"> days.&nbsp; Leave empty or enter no of days.</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap>From : </th>
        <td valign="top">
          <input type="radio" value="Ignore" name="vDateFrom" checked>Not Required (leave expiry field empty above)<br>
          <input type="radio" value="Today"  name="vDateFrom">Today (ensure above field contains the number of days), or<br>
          <input type="radio" value="Access" name="vDateFrom">Date Learner first enters system </td>
      </tr>
      <tr>
        <td align="center" colspan="2"><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Next_<%=svLang%>.gif" name="I1" type="image"><br>&nbsp;</td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
