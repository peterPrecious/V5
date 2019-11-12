<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Dim vCustId, vAction, vDaysOk, vAddProgs, vUseGroups, vReportsTo
  Dim vTemp, aTemp, vRow, vCol, vCnt, vError, vLevel
  Dim oFs, oFile, vFile, vRecord, aRecord, vTabs, vHeader
  Dim b_act2, b_ina2, a_act2, a_ina2, act3, ina3, act4, ina4

	Server.ScriptTimeout = 60 * 10 '...allow up to 10 minutes
	
  vTabs = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab '...dummy tabs to add to input file

  vAction      			= Request("vAction")
  vDaysOk           = Request("vDaysOk")
  vAddProgs   			= Request("vAddProgs")
  vUseGroups  			= Request("vUserGroups")
  vReportsTo        = Request("vReportsTo")

  vCustId      			= Request("vCustId")
  sGetCust vCustId


  Sub sMessage (vError)
    Dim vErr
    vErr = "Line " & vRow & " " & vError & "<br><br>" 
    If vRow > 1 Then
	    vErr = vErr & "<table>"
	    For i = 0 To 7
	      vErr = vErr & "<tr><td class='C6'>Column " & i + 1 & " - " & aRecord(i) & "</td></tr>"
	    Next    
	    vErr = vErr & "</table>"
	  End If
    Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vAction=" & vAction & "&vDaysOk=" & vDaysOk & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload3/Upload3.asp&vErr=" & vErr
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
        If vHeader <> "LEARNERIDGROUPFIRSTNAMELASTNAMEEMAILADDRESSPASSWORDPROGRAMSMEMOJOBSREPORTSTO" Then
  		  	sMessage "did not contain a valid Header."
  		  End If      
      Else

        '...check remaining rows
        vCol = -1
 
        '...assemble fields (sometimes there are insufficient tabs at end - ignore and assume fields are empty)
        vRecord = Replace(vRecord, """", "")
        aRecord = Split(vRecord & vTabs, vbTab)

        '...1) learner id
        vCol = vCol + 1 : vTemp = Ucase(Trim(aRecord(vCol)))
  			If Len(vTemp) = 0 Then 
  		  	sMessage "did not contain a Learner Id."
  		  End If
  		  For i = 1 to Len(vTemp) 
          If (Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@$%^*()_+-{}[];,.:", Mid(vTemp, i, 1))) = 0 Then
            sMessage "contained an invalid Learner Id [" & Server.URLEncode(vTemp) & "]<br>Use only A-Z, 0-9 and !@$%^*()_+-{}[];,.:"
          End If
        Next	
  
        '...2) group id 
        vCol  = vCol + 1 : vTemp = Trim(aRecord(vCol))
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
        vCol = vCol + 1 : vTemp = Trim(aRecord(vCol))
  			If Len(vTemp) = 0 Then 
  		  	sMessage "is missing the First Name."
  		  End If
  
        '...4) last name
        vCol = vCol + 1 : vTemp = Trim(aRecord(vCol))
  			If Len(vTemp) = 0 Then 
  		  	sMessage "is missing the Last Name (" & vTemp & ")"
  		  End If
  
        '...6) password
        vCol = vCol + 2 : vTemp = Ucase((aRecord(vCol)))
  			If vCust_Pwd And Len(vTemp) = 0 Then 
  		  	sMessage "is missing the Password (" & vTemp & ")"
  		  End If
  		  For i = 1 to Len(vTemp) 
          If (Instr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@$%^*()_+-{}[];<>,.:", Mid(vTemp, i, 1))) = 0 Then
            sMessage "contained an invalid Password [" & Server.URLEncode(vTemp) & "]<br>Use only A-Z, 0-9 and !@$%^*()_+-{}[];<>,."
          End If
        Next	

        '...7) programs
        vCol = vCol + 1 : vTemp = Ucase(Trim(aRecord(vCol)))
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

        '...10) reports to
        vCol = vCol + 3 : vTemp = Ucase(Trim(aRecord(vCol)))
        If Len(vTemp) > 0 Then 
          If fFacMembNoById (svCustAcctId, vTemp) = 0 Then
    		  	sMessage "contains an invalid Facilitator Id (" & vTemp & ").<br>Facilitator must exist and be active."
          End If
        End If

      End If  

    End If
  Loop
  oFile.Close
  Set oFile = Nothing

  '...Before Updating inactive all, some or delete all learners
  If vAction = "i" Then spMembInactivate vCust_AcctId, 0
  If vDaysOk >  0  Then spMembInactivate vCust_AcctId, vDaysOk
  If vAction = "d" Then spMembDeleteLearners vCust_AcctId

  '...Update  
  Set oFile = oFs.OpenTextFile(vFile, ForReading)
  vRow = 0

  Do While oFile.AtEndOfStream <> True
    vRecord = Trim(oFile.ReadLine)
    If Len(Trim(Replace(vRecord, vbTab, ""))) > 0 Then
      vRow = vRow + 1
      If vRow > 1 Then '...skip header
        vCol = -1 
        vRecord = Replace(vRecord, """", "")
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
        vCol = vCol + 1 : vMemb_Group3      = fFacMembNoById (svCustAcctId, Ucase(aRecord(vCol)))  

				spMembUpload vCust_AcctId, vMemb_Id, vMemb_Criteria, vMemb_FirstName, vMemb_LastName, vMemb_Email, vMemb_Pwd, vMemb_Programs, vMemb_Memo, vMemb_Jobs, vMemb_Group3
      End If      
    End If
  Loop

' Set oCmd = Nothing
' sCloseDb

  oFile.Close
  Set oFile = Nothing

  '...count learners after update
  spMembCountAll vCust_AcctId, a_act2, a_ina2, act3, ina3, act4, ina4

    '...Create / Update Learners via Upload Advanced
  '   update by acctid/id and assume ACTIVE = 1
  '   NOTE: you need to open and close before using this function
  Function spMembUpload3 (vAcctId, vId, vCriteria, vFirstName, vLastName, vEmail, vPwd, vPrograms, vMemo, vJobs, vGroup3)
  	sOpenCmd
    With oCmd
      .CommandText = "spMembUpload3"
      .Parameters.Append .CreateParameter("@Memb_AcctId",  	  adVarChar, adParamInput,        4, vAcctId)
      .Parameters.Append .CreateParameter("@Memb_Id",  	  		adVarChar, adParamInput,      128, vId)
      .Parameters.Append .CreateParameter("@Memb_Criteria",  	adInteger, adParamInput,         , vCriteria)
      .Parameters.Append .CreateParameter("@Memb_FirstName",  adVarChar, adParamInput,       32, vFirstName)
      .Parameters.Append .CreateParameter("@Memb_LastName",   adVarChar, adParamInput,       64, vLastName)
      .Parameters.Append .CreateParameter("@Memb_Email",  	  adVarChar, adParamInput,      128, vEmail)
      .Parameters.Append .CreateParameter("@Memb_Pwd",  	  	adVarChar, adParamInput,       64, vPwd)
      .Parameters.Append .CreateParameter("@Memb_Programs",  	adVarChar, adParamInput,     8000, vPrograms)
      .Parameters.Append .CreateParameter("@Memb_Memo",  	  	adVarChar, adParamInput,      512, vMemo)
      .Parameters.Append .CreateParameter("@Memb_Jobs",  	  	adVarChar, adParamInput,     8000, vJobs)
      .Parameters.Append .CreateParameter("@Memb_Group3",     adInteger, adParamInput,         , vGroup3)
      .Parameters.Append .CreateParameter("@Memb_AlteredBy",  adInteger, adParamInput,         , svMembNo)
    End With
    oCmd.Execute()
	  Set oCmd = Nothing
	  sCloseDb
  End Function


%>
<html>

<head>
  <title>Upload3_Ok</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><br>Upload Learner Profiles (Advanced with Reports To)</h1>

  <table style="width: 600px; margin: auto;">
    <tr>
      <th>&nbsp;</th>
      <th style="text-align: center" colspan="2">Totals</th>
    </tr>
    <tr>
      <th>Total Uploaded :</th>
      <td style="text-align: center" class="c2" colspan="2"><%=vRow - 1%></td>
    </tr>
    <tr>
      <th colspan="3">&nbsp;</th>
    </tr>
    <tr>
      <th>&nbsp;</th>
      <th style="text-align: center">Before</th>
      <th style="text-align: center">&nbsp;After&nbsp; </th>
    </tr>
    <tr>
      <th>&nbsp; Active Learners :</th>
      <td style="text-align: center"><%=b_act2%></td>
      <td style="text-align: center"><%=a_act2%></td>
    </tr>
    <tr>
      <th>&nbsp; Inactive Learners :</th>
      <td style="text-align: center"><%=b_ina2%></td>
      <td style="text-align: center"><%=a_ina2%></td>
    </tr>
    <tr>
      <th>&nbsp;Total Learners :</th>
      <td style="text-align: center"><%=b_act2 + b_ina2%></td>
      <td style="text-align: center"><%=a_act2 + a_ina2%></td>
    </tr>
    <tr>
      <th colspan="3">&nbsp;</th>
    </tr>
    <tr>
      <th>&nbsp; Active Facilitators :</th>
      <td style="text-align: center" colspan="2"><%=act3%></td>
    </tr>
    <tr>
      <th>&nbsp; Inactive Facilitators :</th>
      <td style="text-align: center" colspan="2"><%=ina3%></td>
    </tr>
    <tr>
      <th>&nbsp;Total Facilitators :</th>
      <td style="text-align: center" colspan="2"><%=act3 + ina3%></td>
    </tr>
    <tr>
      <th colspan="3">&nbsp;</th>
    </tr>
    <tr>
      <th>&nbsp; Active Managers :</th>
      <td style="text-align: center" colspan="2"><%=act4%></td>
    </tr>
    <tr>
      <th>&nbsp; Inactive Managers :</th>
      <td style="text-align: center" colspan="2"><%=ina4%></td>
    </tr>
    <tr>
      <th>&nbsp;Total Managers :</th>
      <td style="text-align: center" colspan="2"><%=act4 + ina4%></td>
    </tr>
  </table>

  <p><br>&nbsp;<a class="c2" href="Upload3.asp">Restart Upload</a><%=f10%><a class="c2" href="/V5/Code/Users.asp">Learner List</a></p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
