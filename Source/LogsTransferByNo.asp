<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vMsg, vNo_From, vNo_To, vFirstName_From, vFirstName_To, vLastName_From, vLastName_To, vCnt

  sGetCust(svCustId)
  
  If Request.Form.Count > 0 Then
    vNo_From					= Request("vNo_From")
    vNo_To            = Request("vNo_To")
    vFirstName_From   = Request("vFirstName_From")
    vFirstName_To     = Request("vFirstName_To")
    vLastName_From    = Request("vLastName_From")
    vLastName_To      = Request("vLastName_To")
    sTransfer
    Response.Redirect "Error.asp?vClose=Y&vErr=" & Server.HtmlEncode(vMsg) 
  End If

  Sub sTransfer   
    sOpenDb

    vSql =        " SELECT * FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_AcctId = Logs.Logs_AcctId"
    vSql = vSql & " WHERE Logs_AcctId = '" & svCustAcctId & "'"
    vSql = vSql & " AND Logs_MembNo = " & vNo_From 
    vSql = vSql & " AND Memb_FirstName = '" & vFirstName_From & "'"
    vSql = vSql & " AND Memb_LastName = '" & vLastName_From & "'"
'   sDebug
    Set oRs = oDb.Execute (vSql)
    If oRs.Eof Then 
      vMsg = "There are either no Log items for that learner or the From Name is incorrect.  Log entries could NOT transferred."
      Exit Sub
    End If
    
    vSql =        " SELECT * FROM Memb WITH (nolock) "
    vSql = vSql & " WHERE Memb_AcctId = '" & svCustAcctId & "'"
    vSql = vSql & " AND Memb_No = '" & vNO_To & "'"
    vSql = vSql & " AND Memb_FirstName = '" & vFirstName_To & "'"
    vSql = vSql & " AND Memb_LastName = '" & vLastName_To & "'"
'   sDebug
    Set oRs = oDb.Execute (vSql)
    If oRs.Eof Then 
      vMsg = "The From learner is not on file.  Log entries could NOT transferred."
      Exit Sub
    End If

    '...determine how many log entries will be transferred
    vSql =        " SELECT COUNT(*) AS Memb_Count "
    vSql = vSql & " FROM Logs WITH (nolock) "
    vSql = vSql & " WHERE Logs_MembNo = " & vNo_From
    Set oRs = oDb.Execute (vSql)
    vCnt = oRs("Memb_Count")
    If vCnt = 0 Then
      vMsg = "There are no log entries to transfer."
      Set oRs = Nothing
      sCloseDb
      Exit Sub
    End If


    vSql =        " UPDATE Logs Set Logs_Membno = " & vNo_To
    vSql = vSql & " WHERE Logs_MembNo = " & vNo_From
'   sDebug
    oDb.Execute (vSql)
    vMsg = vCnt & "The log entries were transferred successfully."
    Set oRs = Nothing
    sCloseDb
    
  End Sub


%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vNo_From.value == "")
  {
    alert("Please enter a value for the \"From Learner No\" field.");
    theForm.vNo_From.focus();
    return (false);
  }

  if (theForm.vNo_From.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"From Learner No\" field.");
    theForm.vNo_From.focus();
    return (false);
  }

  if (theForm.vNo_From.value.length > 8)
  {
    alert("Please enter at most 8 characters in the \"From Learner No\" field.");
    theForm.vNo_From.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vNo_From.value;
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
    alert("Please enter only digit characters in the \"From Learner No\" field.");
    theForm.vNo_From.focus();
    return (false);
  }

  if (theForm.vNo_To.value == "")
  {
    alert("Please enter a value for the \"To Learner No\" field.");
    theForm.vNo_To.focus();
    return (false);
  }

  if (theForm.vNo_To.value.length < 4)
  {
    alert("Please enter at least 4 characters in the \"To Learner No\" field.");
    theForm.vNo_To.focus();
    return (false);
  }

  if (theForm.vNo_To.value.length > 8)
  {
    alert("Please enter at most 8 characters in the \"To Learner No\" field.");
    theForm.vNo_To.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vNo_To.value;
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
    alert("Please enter only digit characters in the \"To Learner No\" field.");
    theForm.vNo_To.focus();
    return (false);
  }

  if (theForm.vFirstName_From.value == "")
  {
    alert("Please enter a value for the \"From First Name\" field.");
    theForm.vFirstName_From.focus();
    return (false);
  }

  if (theForm.vFirstName_From.value.length > 32)
  {
    alert("Please enter at most 32 characters in the \"From First Name\" field.");
    theForm.vFirstName_From.focus();
    return (false);
  }

  if (theForm.vFirstName_To.value == "")
  {
    alert("Please enter a value for the \"To First Name\" field.");
    theForm.vFirstName_To.focus();
    return (false);
  }

  if (theForm.vFirstName_To.value.length > 32)
  {
    alert("Please enter at most 32 characters in the \"To First Name\" field.");
    theForm.vFirstName_To.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="LogsTransferByNo.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
      <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
        <tr>
          <td align="center" valign="top">
            <h1 align="center">Transfer Log Entries by Learner NO</h1>
            <p align="left">This allows you to transfer log records from one Learner to another within this Account.&nbsp; This issue can occur when a Learner registers more than once with different Learner IDs.&nbsp; Enter all 6 fields then click <b>Continue</b>.&nbsp; <span class="c6">&nbsp;Note: these actions cannot be reversed!</span></p>

            <% If Len(vMsg) > 0 Then %>            
            <p class="c5"><%=vMsg%></p>
            <% End If %>
            
            <table border="0" id="table2" cellspacing="0" cellpadding="2">
              <tr>
                <td>&nbsp;</td>
                <th align="left">From</th>
                <th align="left">To</th>
              </tr>
              <tr>
                <th nowrap align="right" valign="top">Vubiz Learner No :</th>
                <td><!--webbot bot="Validation" s-display-name="From Learner No" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="8" --><input type="text" name="vNo_From" size="20" maxlength="8"></td>
                <td><!--webbot bot="Validation" s-display-name="To Learner No" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="8" --><input type="text" name="vNo_To" size="20" maxlength="8"></td>
              </tr>
              <tr>
                <th nowrap align="right" valign="top">First Name :</th>
                <td><!--webbot bot="Validation" s-display-name="From First Name" b-value-required="TRUE" i-maximum-length="32" --><input type="text" name="vFirstName_From" size="20" maxlength="32"></td>
                <td><!--webbot bot="Validation" s-display-name="To First Name" b-value-required="TRUE" i-maximum-length="32" --><input type="text" name="vFirstName_To" size="20" maxlength="32"></td>
              </tr>
              <tr>
                <th nowrap align="right" valign="top">Last Name :</th>
                <td><input type="text" name="vLastName_From" size="20" maxlength="32"></td>
                <td><input type="text" name="vLastName_To" size="20" maxlength="32"></td>
              </tr>
            </table>

            <br><br> 
            <input type="submit" value="Continue" name="bContinue" class="button">
            <br><br><h2 align="center"><%=vCust_Id & "  (" & vCust_Title & ")"%></h2>
          </td>

        </tr>
      </table>
    </form>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>