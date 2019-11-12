<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include file = "Cineplex_Routines.asp"-->

<%
  Dim vMessage, vScore, bOk
  bOk = False
  vMessage = ""

  If Request("bDelete").Count > 0 Then

    vMemb_No        = Request("vMemb_No")
    vMemb_Id        = Ucase(Request("vMemb_Id"))
    vMemb_FirstName = fUnquote(Request("vMemb_FirstName"))
    vMods_Id        = fUnquote(Request("vMods_Id"))
    vLogs_Posted    = fFormatDate(Request("vLogs_Posted"))
    vScore          = Cint(Request("vScore"))

    bOk = True
    vSql = "DELETE Logs WHERE Logs_No = " & Request("vLogs_No")
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb
  
  ElseIf Request("bAdd").Count > 0 Then

    '...ensure values are valid
    vMemb_Id        = Ucase(Request("vMemb_Id"))
    vMemb_FirstName = fUnquote(Request("vMemb_FirstName"))
    vMods_Id        = fUnquote(Request("vMods_Id"))
    vLogs_Posted    = fFormatDate(Request("vLogs_Posted"))
    vScore          = Cint(Request("vScore"))

    vSql = "SELECT Memb_No, Memb_FirstName FROM Memb WITH (NOLOCK) WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & vMemb_Id & "'"
    sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      vMessage = vMessage & "<br>That Learner Id is not on file."
    Else
      If vMemb_FirstName <> oRs("Memb_FirstName") Then
        vMessage = vMessage & "<br>The Learner Id you entered does not have a matching First Name."
      Else
        vMemb_No = oRs("Memb_No")
      End If
    End If
    Set oRs = Nothing      
    sCloseDb

    vSql = "SELECT Mods_Id FROM Mods WITH (NOLOCK) WHERE Mods_Id= '" & vMods_Id & "'"
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If oRsBase.Eof Then vMessage = vMessage & "<br>Please enter a valid Assessment/Module Id."
    Set oRsBase = Nothing
    sCloseDbBase    

    If vLogs_Posted = " " Then 
      vMessage = vMessage & "<br>Please enter a valid date."
    End If

    If vScore < 0 Or vScore > 100 Then 
      vMessage = vMessage & "<br>Please enter a score from 1-100."
    End If

    If vMessage = "" Then
      bOk = True
      If vScore > 0 Then    
        sOpenDb
        vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_MembNo, Logs_Posted, Logs_Item) VALUES (" _
             & "'" & svCustAcctId & "', 'T', " & vMemb_No & ", '" & vLogs_Posted & "', '" & vMods_Id & "_" & Right("000" & vScore, 3) & "')"
'       sDebug
        oDb.Execute (vSql)
        sCloseDb
      End If
    End If

  End If 
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>

    var reAlphaNumeric = new RegExp(/^[0-9A-Za-z]+$/)
    var reAlpha        = new RegExp(/^[A-Za-z]+$/)
    var reNumeric      = new RegExp(/^[0-9]+$/)


    function Validate(theForm) {
      if (theForm.vMemb_Id.value.length == 0) {
        alert('Please enter the Learner\'s ID.')
        document.fForm.vMemb_Id.focus()
        return (false);
      }
      if (theForm.vMemb_FirstName.value.length == 0) {
        alert('Please enter the Learner\'s First Name\n as it appears on the system.')
        document.fForm.vMemb_FirstName.focus()
        return (false);
      }
      if (theForm.vLogs_Posted.value.length == 0) {
        alert('Please enter the Date the Assessment was taken.')
        document.fForm.vLogs_Posted.focus()
        return (false);
      }
      if (theForm.vMods_Id.value.length != 6) {
        alert('Please enter a valid Assessment/Module ID.')
        document.fForm.vMods_Id.focus()
        return (false);
      }
      if (theForm.vScore.value.match(reNumeric)==null) { 
        alert('Please enter a Score between 1 and 100.')
        document.fForm.vScore.focus()
        return (false);
      }
      return (true);
    }
    
  </script>
</head>

<body>

<% Server.Execute vShellHi %>
<table width="100%" border="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
  <tr>
    <td align="center">
    <h1><br>Completion Utilities - Edit Learner Scores</h1>
    <h2>This allows you to Add or Delete an existing Learner&#39;s Score. This is typically used to capture scores attained from previous systems or classrooms.&nbsp; <font color="#FF0000">Note: by entering a score of 0 you can see existing scores without adding a new entry.</font></h2>
    <% If Len(vMessage) > 0 Then %><span class="c5"><%=vMessage%><br>&nbsp;</span><% End If %> </td>
  </tr>
  <tr>
    <td align="center">
    <div align="center">
      <form method="POST" action="Completion_AddScores.asp" onsubmit="return Validate(this)" name="fForm">
        <table border="1" style="border-collapse: collapse" id="table1" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
          <tr>
            <th nowrap align="right" valign="top" width="35%">Learner Id :</th>
            <td valign="top" width="65%"><input type="text" name="vMemb_Id" size="32" value="<%=vMemb_Id%>" class="c2"> <br>
            Enter exactly as it appears on file.</td>
          </tr>
          <tr>
            <th nowrap align="right" valign="top" width="35%">Learner First Name :</th>
            <td valign="top" width="65%"><input type="text" name="vMemb_FirstName" size="32" value="<%=vMemb_FirstName%>" class="c2"> <br>
            ie John. Enter exactly as it appears on file.</td>
          </tr>
          <tr>
            <th nowrap align="right" valign="top" width="35%">Assessment Date :</th>
            <td valign="top" width="65%"><input type="text" name="vLogs_Posted" size="16" maxlength="22" value="<%=vLogs_Posted%>" class="c2"> <br>
            ie Jan 1, 2011 (using English Date format)</td>
          </tr>
          <tr>
            <th nowrap align="right" valign="top" width="35%">Assessment Id :</th>
            <td valign="top" width="65%"><input type="text" name="vMods_Id" size="10" value="<%=vMods_Id%>" class="c2"> <br>
            ie 1234EN. Note: this is a Module Id, NOT a Program Id.</td>
          </tr>
          <tr>
            <th nowrap align="right" valign="top" width="35%">Score :</th>
            <td valign="top" width="65%"><input type="text" name="vScore" size="4" value="<%=vScore%>" class="c2"> <br>
            ie 1-100. <font color="#FF0000">Note: if you enter zero then you can view existing scores that are on the log file without updating the Log Table.</font></td>
          </tr>
          <tr>
            <td align="center" colspan="2"><br><input type="submit" value="Add" name="bAdd" class="button070"><p>&nbsp;</p>
            </td>
          </tr>
        </table>
      </form>
      <%  
      	If bOk Then 
      %>
        <table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="2">
          <tr>
            <th colspan="4" bgcolor="#DDEEF9">Exiting Scores</th>
          </tr>
          <tr>
            <th>Assessment Id</th>
            <th>Assessment Date</th>
            <th>Score</th>
            <th>&nbsp;</th>
          </tr>
          <%
            sOpenDb
            vSql = "SELECT * FROM Logs WITH (NOLOCK) WHERE Logs_AcctId = '" & svCustAcctId & "' AND Logs_MembNo = " & vMemb_No & " AND Logs_Type = 'T' ORDER BY Logs_Posted "
'           sDebug
            Set oRs = oDb.Execute (vSql)
            While Not oRs.Eof
          %>
          <form method="POST" action="Completion_AddScores.asp" onsubmit="return Validate(this)" name="fForm_<%=vMemb_No%>">
          <input type="hidden" name="vLogs_No" value="<%=oRs("Logs_No")%>">
          <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">
          <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">
          <input type="hidden" name="vMemb_FirstName" value="<%=vMemb_FirstName%>">
          <input type="hidden" name="vLogs_Posted" value="<%=vLogs_Posted%>">
          <input type="hidden" name="vMods_Id" value="<%=vMods_Id%>">
          <input type="hidden" name="vScore" value="<%=vScore%>">
          <tr>
            <td align="center"><%=Left(oRs("Logs_Item"), 6)%></td>
            <td align="center"><%=fFormatDate(oRs("Logs_Posted"))%></td>
            <td align="center"><%=Cint(Right(oRs("Logs_Item"), 3))%></td>
            <td align="center"><input type="submit" value="Delete" name="bDelete" class="button070"></td>
          </tr>
          </form>
          <% 
              oRs.MoveNext
            Wend
            sCloseDb
          %>
        </table>
      <%
        End If 
      %>
      <p>&nbsp;</p>
    </div>
    </td>
  </tr>
</table>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Cineplex_Footer.asp"-->

</body>

</html>


