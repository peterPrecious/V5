<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Dim vCustId, vInactivate, vAddProgs, vUseGroups, vTemp, aTemp, vRow, vCol, vCnt, vError, vLevel
  Dim oFs, oFile, vFile, vRecord, aRecord, vTabs, vHeader
  Dim b_act2, b_ina2, a_act2, a_ina2, act3, ina3, act4, ina4

	Server.ScriptTimeout = 60 * 30 '...allow 30 minutes for monster uploads

  vTabs = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab '...dummy tabs to add to input file
  vCustId      = Request("vCustId")
  sGetCust vCustId
  vUseGroups  = fIf(fCritOk (vCust_AcctId), "y", "n")
  vInactivate = Request("vInactivate")
  vAddProgs   = Request("vAddProgs")

  Sub sMessage (vError)
    Dim vErr
    vErr = "Record " & vRow & " " & vError & "<br><br>" 
    If vRow > 1 Then
	    vErr = vErr & "<table>"
	    For i = 0 To 7
	      vErr = vErr & "<tr><td class='C6'>Column " & i & " - " & aRecord(i) & "</td></tr>"
	    Next    
	    vErr = vErr & "</table>"
	  End If
    Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vInactivate=" & vInactivate & "&vReturn=/V5/Repository/Upload/Upload_Advanced.asp&vErr=" & vErr
  End Sub

  Function fProgs(vNew)
    Dim aOld, aNew  
    If vAddProgs = "n" Then
      fProgs = vNew
    Else  
      fProgs = ""  
      vSql = "SELECT Memb_Programs FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Id = '" & vMemb_Id & "'"
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      If Not oRs.Eof Then fProgs = Trim(oRs("Memb_Programs"))
      Set oRs = Nothing      
      sCloseDb
      '...ignore any new progs that are already on file
      If Len(fProgs) > 0 And Len(vNew) > 0 Then
        aOld = Split(fProgs)
        aNew = Split(vNew)      
        For i = 0 To Ubound(aNew)
          If InStr(fProgs, aNew(i)) = 0 Then
            fProgs = fProgs & " " & aNew(i)
          End If
        Next
      Else
        fProgs = fProgs & " " & vNew     
      End If
    End If
  End Function


  '...count learners before update
  spMembCountAll vCust_AcctId, b_act2, b_ina2, act3, ina3, act4, ina4

  
 '...imported names into a table
  Const ForReading = 1
  Set oFs = CreateObject("Scripting.FileSystemObject")   
  vFile = Server.MapPath(vCustId & "_Learners.txt")

  '...first check all records for integrity
  Set oFile = oFs.OpenTextFile(vFile, ForReading)

  Do While oFile.AtEndOfStream <> True
    vRecord = Trim(oFile.ReadLine)
    
    '...ignore empty records (easy to pass in from excel)
    If Len(Trim(Replace(vRecord, vbTab, ""))) > 0 Then

      vRow = vRow + 1

      '...check for valid header on first valid record (compressed)
      '   strip out tabs and spaces on header
      If vRow = 1 Then
        vHeader = Replace(vRecord, vbTab, "")
        vHeader = Replace(vHeader, " ", "")
        vHeader = Ucase(vHeader)        
        If vHeader <> "LEARNERIDGROUPFIRSTNAMELASTNAMEEMAILADDRESSPASSWORDPROGRAMSMEMOJOBS" Then
  		  	sMessage "did not contain a valid Header."
  		  End If      
      Else

        '...check remaining rows
        vCol = -1
 
        '...assemble fields (sometimes there are insufficient tabs at end - ignore and assume fields are empty)
        aRecord = Split(vRecord & vTabs, vbTab)

        '...1) learner id
        vCol = vCol + 1 : vTemp = Ucase((aRecord(vCol)))
  			If Len(vTemp) = 0 Then 
  		  	sMessage "did not contain a Learner Id."
  		  End If
  		  For i = 1 to Len(vTemp) 
          If (Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-@.", Mid(vTemp, i, 1))) = 0 Then
            sMessage "contained a invalid Learner Id [" & vTemp & "]<br>Learner Id can only contain A-Z, 0-9 and _-@."
          End If
        Next	
  
        '...2) group id 
        vCol  = vCol + 1 : vTemp = aRecord(vCol) 
        If vUseGroups = "n" And Len(vTemp) > 0 Then
          sMessage "contains a Group ID [" & vTemp & "].<br>You specified that you do NOT use Groups."
        ElseIf vUseGroups = "y" And Len(vTemp) = 0 Then
          sMessage "does not contain a Group Id.<br>[You specified that you use Groups]"
        End If
        If vUseGroups = "y" Then
          If fCriteriaNo (vCust_AcctId, vTemp) = 0 Then
            sMessage "contains a Group ID [" & vTemp  & "] that is not setup on file<br>Groups must be setup before you Upload Users."
          End If
        End If  
  
        '...3) first name
        vCol = vCol + 1 : vTemp = aRecord(vCol)
  			If Len(vTemp) = 0 Then 
  		  	sMessage "is missing the First Name."
  		  End If
  
        '...4) last name
        vCol = vCol + 1 : vTemp = aRecord(vCol)
  			If Len(vTemp) = 0 Then 
  		  	sMessage "is missing the Last Name (" & vTemp & ")"
  		  End If
  
        '...5) password
        vCol = vCol + 2 : vTemp = Ucase((aRecord(vCol)))
  			If vCust_Pwd And Len(vTemp) = 0 Then 
  		  	sMessage "is missing the Password (" & vTemp & ")"
  		  End If

        '...6) programs
        vCol = vCol + 1 : vTemp = Ucase((aRecord(vCol)))
  			If Len(vTemp) > 0 Then 
  			  aTemp = Split(vTemp)
  			  For i = 0 To Ubound(aTemp)
  			    If Len(aTemp(i)) <> 7 Then
      		  	sMessage "contains an invalid Program Id (" & vTemp & ").<br>Program IDs should be entered like: P1234EN P3443EN."
      		  ElseIf Not fProgOk (aTemp(i)) Then
      		  	sMessage "contains a Program Id that is not on file (" & vTemp & ")."
      		  End If
      		Next
  		  End If
  
      End If  

    End If
  Loop
  oFile.Close
  Set oFile = Nothing

  '...Before Updating inactive all learners
  If vInactivate = "y" Then      
    spMembInactivate vCust_AcctId
  End If

  '...Update  
  Set oFile = oFs.OpenTextFile(vFile, ForReading)

  vRow = 0
  Do While oFile.AtEndOfStream <> True
    vRecord = Trim(oFile.ReadLine)
    If Len(Trim(Replace(vRecord, vbTab, ""))) > 0 Then
      vRow = vRow + 1
      If vRow > 1 Then '...skip header
        vCol = -1 
        aRecord = Split(vRecord & vTabs, vbTab)
        For i = 0 To Ubound(aRecord) : aRecord(i) = fUnquote(Trim(aRecord(i))) : Next  
        sMemb_Empty
        vCol = vCol + 1 : vMemb_Id          = Ucase((aRecord(vCol)))
        vCol = vCol + 1 : vMemb_Criteria    = fCriteriaNo (vCust_AcctId, aRecord(vCol))
        vCol = vCol + 1 : vMemb_FirstName   = aRecord(vCol)
        vCol = vCol + 1 : vMemb_LastName    = aRecord(vCol)
        vCol = vCol + 1 : vMemb_Email       = Ucase(aRecord(vCol))
        vCol = vCol + 1 : vMemb_Pwd         = Ucase(aRecord(vCol))
        vCol = vCol + 1 : vMemb_Programs    = fProgs(Ucase(aRecord(vCol)))
        vCol = vCol + 1 : vMemb_Memo        = Ucase(aRecord(vCol))        
        vCol = vCol + 1 : vMemb_Jobs        = Ucase(aRecord(vCol))        
        vMemb_Active = 1
        sAddMemb vCust_AcctId 
      End If      
    End If
  Loop
  oFile.Close
  Set oFile = Nothing

  '...count learners after update
  spMembCountAll vCust_AcctId, a_act2, a_ina2, act3, ina3, act4, ina4


%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <title>Import</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
    <tr>
      <td align="center">
      <h1><br>Upload Learner Profiles (Advanced)</h1>
      <table border="1" id="table2" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9">
        <tr>
          <th align="right">&nbsp;</th>
          <th align="center" colspan="2">Totals</th>
        </tr>
        <tr>
          <th align="right">Total Uploaded :</th>
          <td align="center" class="c2" colspan="2"><%=vRow - 1%></td>
        </tr>
        <tr>
          <th align="right" colspan="3">&nbsp;</th>
        </tr>
        <tr>
          <th align="right">&nbsp;</th>
          <th align="center">Before</th>
          <th align="center">&nbsp;After&nbsp; </th>
        </tr>
        <tr>
          <th align="right">&nbsp; Active Learners :</th>
          <td align="center"><%=b_act2%></td>
          <td align="center"><%=a_act2%></td>
        </tr>
        <tr>
          <th align="right">&nbsp; Inactive Learners :</th>
          <td align="center"><%=b_ina2%></td>
          <td align="center"><%=a_ina2%></td>
        </tr>
        <tr>
          <th align="right">&nbsp;Total Learners :</th>
          <td align="center"><%=b_act2 + b_ina2%></td>
          <td align="center"><%=a_act2 + a_ina2%></td>
        </tr>
        <tr>
          <th align="right" colspan="3">&nbsp;</th>
        </tr>
        <tr>
          <th align="right">&nbsp; Active Facilitators :</th>
          <td align="center" colspan="2"><%=act3%></td>
        </tr>
        <tr>
          <th align="right">&nbsp; Inactive Facilitators :</th>
          <td align="center" colspan="2"><%=ina3%></td>
        </tr>
        <tr>
          <th align="right">&nbsp;Total Facilitators :</th>
          <td align="center" colspan="2"><%=act3 + ina3%></td>
        </tr>
        <tr>
          <th align="right" colspan="3">&nbsp;</th>
        </tr>
        <tr>
          <th align="right">&nbsp; Active Managers :</th>
          <td align="center" colspan="2"><%=act4%></td>
        </tr>
        <tr>
          <th align="right">&nbsp; Inactive Managers :</th>
          <td align="center" colspan="2"><%=ina4%></td>
        </tr>
        <tr>
          <th align="right">&nbsp;Total Managers :</th>
          <td align="center" colspan="2"><%=act4 + ina4%></td>
        </tr>
      </table>
      <p><br>&nbsp;<a class="c2" href="Upload_Advanced.asp">Restart Upload</a><%=f10%><a class="c2" href="/V5/Code/Users.asp">Learner List</a></p>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
