﻿<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vMsg, vFrom_Id, vTo_Id, vFrom_No, vTo_No, vFrom_Cnt, vTo_Cnt, vFinal_Cnt, vAction

  sGetCust(svCustId)

  vAction  					= fDefault(Request("vAction"), "a")
  
  If Request.Form.Count > 0 Then
    vFrom_Id					= Request("vFrom_Id")
    vTo_Id            = Request("vTo_Id")
    sOpenDb

    vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & vFrom_Id & "' "
'   sDebug
    Set oRs = oDb.Execute (vSql)
    If oRs.Eof Then 
      vMsg = "The 'From' Id is not on file.  Processing was interrupted.  Log entries could NOT be transferred."
    Else
      vFrom_No = oRs("Memb_No")

      vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & vTo_Id & "' "
'     sDebug
      Set oRs = oDb.Execute (vSql)
      If oRs.Eof Then 
        vMsg = "The 'To' Id is not on file.  Processing was interrupted.  Log entries could NOT be transferred."
      Else
        vTo_No = oRs("Memb_No")
    
        '...determine how many log entries will be transferred
        vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vFrom_No
        Set oRs = oDb.Execute (vSql)
        vFrom_Cnt = oRs("Count")
        vMsg = vMsg & "The 'From' learner contains " & vFrom_Cnt & " log entries to transfer."
      
      
        '...determine how many log entries will be transferred
        vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vTo_No
        Set oRs = oDb.Execute (vSql)
        vTo_Cnt = oRs("Count")
        vMsg = vMsg & "<br>The 'To' learner contains " & vTo_Cnt & " log entries before the transfer."
      
      
        '...do the transfer
        If vFrom_Cnt > 0 Then
          vSql = "UPDATE Logs Set Logs_Membno = " & vTo_No & " WHERE Logs_MembNo = " & vFrom_No
'         sDebug
          oDb.Execute (vSql)
          vMsg = vMsg & "<br>" & vFrom_Cnt & " log entries were transferred successfully."

          '...determine how many log entries will be transferred
          vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vTo_No
          Set oRs = oDb.Execute (vSql)
          vFinal_Cnt = oRs("Count")
          vMsg = vMsg & "<br>The 'To' learner now contains " & vFinal_Cnt & " log entries."

        Else
          vMsg = vMsg & "<br>No log entries were transferred."
        End If    

        If vAction = "i" Then
          vSql =  "UPDATE Memb Set Memb_Active = '0' WHERE Memb_No = " & vFrom_No
'         sDebug
          oDb.Execute (vSql)
          vMsg = vMsg & "<br>The 'From' learner has been inactivated."
        ElseIf vAction = "d" Then
          vSql = "DELETE Memb WHERE Memb_No = " & vFrom_No
'         sDebug
          oDb.Execute (vSql)
          vMsg = vMsg & "<br>The 'From' learner has been deleted."
        End If

      End If
    
    End If
 
    Set oRs = Nothing
    sCloseDb


  End If

%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %> 
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vFrom_Id.value == "")
  {
    alert("Please enter a value for the \"From Learner ID\" field.");
    theForm.vFrom_Id.focus();
    return (false);
  }

  if (theForm.vTo_Id.value == "")
  {
    alert("Please enter a value for the \"To Learner ID\" field.");
    theForm.vTo_Id.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="LogsTransferById.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
      <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
        <tr>
          <td align="center" valign="top">
            <h1>Transfer Log Entries by Learner ID</h1>
            <p class="c2" align="left">This allows you to transfer log records from one Learner to another within this Account.&nbsp; This issue can occur when a Learner registers more than once with different Learner IDs.&nbsp; Enter the From and To Learner IDs and what you would like to do with the original &quot;From&quot; Learners record. <span class="c6">&nbsp;Note: these actions cannot be reversed!</span></p>
  
            <% If Len(vMsg) > 0 Then %>            
            <p class="c5"><%=vMsg%></p>
            <% End If %>
  
            <table border="0" id="table2" cellspacing="0" cellpadding="5">
              <tr>
                <th nowrap align="right" valign="top">Transfer logs From Learner Id :</th>
                <td valign="top">&nbsp;<!--webbot bot="Validation" s-display-name="From Learner ID" b-value-required="TRUE" --><input type="text" name="vFrom_Id" size="48" value="<%=vFrom_Id%>"></td>
              </tr>
              <tr>
                <th nowrap align="right" valign="top">To the logs of Learner Id :</th>
                <td valign="top">&nbsp;<!--webbot bot="Validation" s-display-name="To Learner ID" b-value-required="TRUE" --><input type="text" name="vTo_Id" size="48" value="<%=vTo_Id%>"></td>
              </tr>
              <tr>
                <th nowrap align="right" valign="top">Then do the&nbsp;following to the &quot;From&quot; learner record :<br>
                <span style="font-weight: 400">(even in there are no log items to transfer) </span>&nbsp; </th>
                <td valign="top">
                <input type="radio" value="a" name="vAction" <%=fCheck("a", vAction)%>>Leave the Learner Active in the system.<br>
                <input type="radio" value="i" name="vAction" <%=fCheck("i", vAction)%>>Make the Learner Inactive.<br>
                <input type="radio" value="d" name="vAction" <%=fCheck("d", vAction)%>>Delete the Learner from the Learner table (most common).</td>
              </tr>
            </table>
  
            <br><br> 
            <input type="submit" value="Continue" name="bContinue" class="button">
            <br><br><h2 align="center"><%=vCust_Id & "  (" & vCust_Title & ")"%></h2>

          </td>
        </tr>
      </table>
    </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

