<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vFrom_Id, vTo_Id, vFrom_No, vTo_No, vFrom_CustId, vTo_CustId
  Dim vMsg, vFrom_Cnt, vTo_Cnt, vFinal_Cnt, vAction

  vFrom_CustId			= fDefault(Request("vFrom_CustId"), svCustId)
  vTo_CustId        = fDefault(Request("vTo_CustId"), svCustId)
  vFrom_Id					= Request("vFrom_Id")
  vTo_Id            = Request("vTo_Id")

  If Request.Form.Count > 0 Then

    sOpenDb

    vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & fCustAcctId (vFrom_CustId) & "' AND Memb_Id = '" & vFrom_Id & "' "
'   sDebug
    Set oRs = oDb.Execute (vSql)
    If oRs.Eof Then 
      vMsg = "The 'FROM' Learner Id is not on file.  Processing was interrupted.  Log entries could NOT be transferred."
    Else
      vFrom_No = oRs("Memb_No")

      vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & fCustAcctId (vTo_CustId) & "' AND Memb_Id = '" & vTo_Id & "' "
'     sDebug
      Set oRs = oDb.Execute (vSql)
      If oRs.Eof Then 
        vMsg = "The 'TO' Learner Id is not on file.  Processing was interrupted.  History entries could NOT be transferred."
      Else
        vTo_No = oRs("Memb_No")
    
        '...determine how many log entries will be transferred
        vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vFrom_No
        Set oRs = oDb.Execute (vSql)
        vFrom_Cnt = oRs("Count")
        vMsg = vMsg & "The 'From' learner contains " & vFrom_Cnt & " history entries to transfer."
      
        '...determine how many log entries will be transferred
        vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vTo_No
        Set oRs = oDb.Execute (vSql)
        vTo_Cnt = oRs("Count")
        vMsg = vMsg & "<br>The 'To' learner contains " & vTo_Cnt & " history entries before the transfer."
      
        '...do the LMS transfer
        If vFrom_Cnt > 0 Then
          vSql = "UPDATE Logs Set Logs_MembNo = " & vTo_No & ", Logs_AcctId = '" & fCustAcctId (vTo_CustId) & "' WHERE Logs_MembNo = " & vFrom_No
'         sDebug
          oDb.Execute (vSql)
          vMsg = vMsg & "<br>" & vFrom_Cnt & " history entries were transferred successfully."

          '...determine how many log entries will be transferred
          vSql = "SELECT COUNT(*) AS [Count] FROM Logs WHERE Logs_MembNo = " & vTo_No
          Set oRs = oDb.Execute (vSql)
          vFinal_Cnt = oRs("Count")
          vMsg = vMsg & "<br>The 'To' learner now contains " & vFinal_Cnt & " history entries."

          '...do the RTE transer
          Set oDbRTE = Server.CreateObject("ADODB.Connection")
          oDbRTE.ConnectionString = "Provider=SQLOLEDB.1;Application Name=V5 Platform;Password=" & svHostDbPwd & ";Persist Security Info=True;User ID=sa;Initial Catalog=vuGoldSCORM;Data Source=" & svSQL
          oDbRTE.Open
    
          Set oCmdRTE = Server.CreateObject("ADODB.Command")
          Set oCmdRTE.ActiveConnection = oDbRTE
          oCmdRTE.CommandType = adCmdStoredProc
    
          With oCmdRTE
            .CommandText = "spSessionTransfer"
            .Parameters.Append .CreateParameter("SourceMemberID", adInteger, adParamInput, , vFrom_No)
            .Parameters.Append .CreateParameter("TargetMemberID", adInteger, adParamInput, , vTo_No)
          End With
          oCmdRTE.Execute()
    
          Set oCmdRTE = Nothing
          sCloseDbRTE

        Else
          vMsg = vMsg & "<br>No history entries were transferred."
        End If    

      End If

      Set oRs = Nothing
      sCloseDb
  
    End If

  End If

%>
<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
		function validate(theForm) {

			if (theForm.vFrom_CustId.value == "") {
				alert("Please enter the FROM Cust Id.");
				theForm.vFrom_CustId.focus();
				return (false);
			}

			if (theForm.vFrom_Id.value == "") {
				alert("Please enter the FROM Learner Id.");
				theForm.vFrom_Id.focus();
				return (false);
			}

			if (theForm.vTo_CustId.value == "") {
				alert("Please enter the TO Cust Id.");
				theForm.vFrom_CustId.focus();
				return (false);
			}

			if (theForm.vTo_Id.value == "") {
				alert("Please enter the TO Learner Id.");
				theForm.vTo_Id.focus();
				return (false);
			}

			return (true);
		}
	</script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="LogsHistoryTransfer.asp" onsubmit="return validate(this)">
    <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
      <tr>
        <td align="center" valign="top">
          <h1>Transfer Learner History</h1>
          <p class="c2" align="left">This allows you to transfer learner history from one Learner to another from the LMS to the RTE.<br /><br />This issue can occur when a Learner registers more than once in an Account or when Accounts are merged.&nbsp; This service only transfers the learning history - not ecommerce.&nbsp; It assumes the TO Learner record has been setup and the FROM Learner record has or will be inactivated.</p>
          <table border="1" id="table2" cellpadding="3" style="border-collapse: collapse" bordercolor="#00FFFF">
            <tr>
              <th nowrap valign="top" colspan="2"><% If Len(vMsg) > 0 Then %>
              <p class="c5" style="border:1px solid red; margin:10px; padding:10px;"><%=vMsg%></p>
              <% End If %> </th>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Transfer history FROM Cust Id :</th>
              <td valign="top">&nbsp;<input type="text" name="vFrom_CustId" size="10" value="<%=vFrom_CustId%>"></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Learner Id :</th>
              <td valign="top">&nbsp;<input type="text" name="vFrom_Id" size="36" value="<%=vFrom_Id%>"></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">&nbsp;</th>
              <td valign="top">&nbsp;</td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Transfer history TO Cust Id :</th>
              <td valign="top">&nbsp;<input type="text" name="vTo_CustId" size="10" value="<%=vTo_CustId%>"></td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Learner Id :</th>
              <td valign="top">&nbsp;<input type="text" name="vTo_Id" size="36" value="<%=vTo_Id%>"></td>
            </tr>
          </table>
          <p class="c2"><b>Note: this action cannot be reversed!</b></p>
          <input type="submit" value="Continue" name="bContinue" class="button">
        </td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
