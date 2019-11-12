<%
  Dim vLogs_No, vLogs_AcctID, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo
  Dim vLogs_Module, vLogs_Grade, vDetails, vCurList, vMaxList
  
  
  '____ Logs  ________________________________________________________________________

  '...Get Logs
  Sub sGetLogs_rs (vCrit1, vCrit2, vFirstName, vLastName, vPosted)
    vSql = "SELECT "
'   vSql = vSql & " TOP " & vMaxList
    vSql = vSql & " Memb.Memb_Criteria, Memb.Memb_FirstName, Memb.Memb_LastName "
    '...details of grades or maximum grade
    If vDetails = "y" Then    
      vSql = vSql & ",  Deal.Deal_Title + ' - ' + Left(Logs.Logs_Item, 6) as Logs_Module, Right(Logs.Logs_Item,3) as Logs_Grade, Logs.Logs_Posted "
    Else
      vSql = vSql & ",  Deal.Deal_Title + ' - ' + Left(Logs.Logs_Item, 6) as Logs_Module, MAX(Right(Logs.Logs_Item,3)) as Logs_Grade, MAX(Logs.Logs_Posted) as Logs_Posted " 
    End If
    vSql = vSql & " FROM Logs WITH (NOLOCK) INNER JOIN Memb WITH (NOLOCK) ON Logs_MembNo = Memb_No "
    vSql = vSql & " WHERE Logs_AcctID= '" & svCustAcctId & "' AND Logs_Type = 'T'"
    '...only get passed values (ie > 90%)
 '   vSql = vSql & " AND (SUBSTRING(Logs_Item, 8, 2) = '10' OR SUBSTRING(Logs_Item, 8, 2) = '09' OR SUBSTRING(Logs_Item, 8, 2) = '08') " 
    '...ignore vu administrators
    vSql = vSql & " AND Memb_Level < 4 "
    '...start at last member (for subsequent reads...)
    vSql = vSql & " AND Memb_Criteria  >= '" & vCrit1     & "'"
    vSql = vSql & " AND Memb_Crit2     >= '" & vCrit2     & "'"
    vSql = vSql & " AND Memb_FirstName >= '" & vFirstName & "'"
    vSql = vSql & " AND Memb_LastName  >= '" & vLastName  & "'"
    vSql = vSql & " AND Logs_Posted    >= '" & vPosted    & "'"
    '...details of grades or maximum grade
    If vDetails = "y" Then    
      vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_Crit2, Memb.Memb_LastName, Memb.Memb_FirstName, Logs.Logs_Posted "
    Else
      vSql = vSql & " GROUP BY Memb.Memb_Criteria, Memb.Memb_Crit2, Memb.Memb_LastName, Memb.Memb_FirstName,  Deal.Deal_Title + ' - ' + Left(Logs.Logs_Item, 6) "
    End If
'   sDebug
    sOpenDB
    Set oRs = oDB.Execute(vSql)
  End Sub

  '...get the current fields from the current record in the record set (just a subset of the full fields)
  Sub sReadLogsMemb
    vLogs_Module                = oRs("Logs_Module")
    vLogs_Grade                 = oRs("Logs_Grade")
    vLogs_Posted                = oRs("Logs_Posted")
    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Criteria              = oRs("Memb_Criteria")
    vMemb_Crit2                 = oRs("Memb_Crit2")
  End Sub

  Sub sReadLogs
    vLogs_No                    = oRs("Logs_No")
    vLogs_AcctID                = oRs("Logs_AcctID")
    vLogs_Type                  = oRs("Logs_Type")
    vLogs_MembNo                = oRs("Logs_MembNo")
    vLogs_Item                  = oRs("Logs_Item")
    vLogs_Posted                = oRs("Logs_Posted")
  End Sub


  Function fBookmarkLogs
    fBookmarkLogs = False
    '...see if any bookmarks for this user
    vSql = "SELECT Logs_No FROM Logs WITH (NOLOCK) WHERE "    
    vSql = vSql & "Logs_AcctId = '" & svCustAcctId & "' AND Logs_MembNo = '" & svMembNo & "' AND Logs_Type = 'B' AND Logs_Posted > '" & fFormatSqlDate(Now - 30) & "'"
'   sDebug  
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fBookmarkLogs = True
    sCloseDb
  End Function


  Sub sLogTestResults
    sOpenDb
    vSql = "INSERT INTO Logs"
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & svMembNo & ")"
'   sDebug
    oDb.Execute(vSql)
  End Sub


  '...same as above except we use vPosted rather than default NOW - modified Jul 24 2012
  Sub sLogTestResults2 (vPosted)
    sOpenDb
    vSql = "INSERT INTO Logs"
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo, Logs_Posted) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vLogs_Item & "', " & svMembNo & ", '" & vPosted & "')"
'   sDebug
    oDb.Execute(vSql)
  End Sub


  Sub sLogSurveyResults
    sOpenDb
    vSql = "INSERT INTO Logs"
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'U', '" & vLogs_Item & "', " & svMembNo & ")"
    oDb.Execute(vSql)
  End Sub


  Sub sLogAssessmentResults '...these are scorm questions and answers
    sOpenDb
    vSql = "INSERT INTO Logs"
    vSql = vSql & "(Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'A', '" & vLogs_Item & "', " & svMembNo & ")"
'   sDebug
    oDb.Execute(vSql)
  End Sub


  '...Get All Surveys
  Sub sGetSurveyLogs_Rs (vDays)
    vSql = "SELECT Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Level, Memb.Memb_Criteria, Logs.Logs_Posted, Logs.Logs_Item "
    vSql = vSql & " FROM Logs WITH (NOLOCK) INNER JOIN Memb ON Logs_MembNo = Memb_No "
    vSql = vSql & " WHERE Logs_AcctId= '" & svCustAcctId & "'"
    vSql = vSql & " AND Logs_Type = 'U'"
    vSql = vSql & " AND Logs.Logs_Posted > '" & vDays & "'"
    vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id, Logs.Logs_Posted "
    sDebug
    sOpenDb
    Set oRs = oDB.Execute(vSql)
  End Sub


  '...get the current fields from the current record in the record set (just a subset of the full fields)
  Sub sReadSurveyLogs
    vLogs_Item                  = oRs("Logs_Item")
    vLogs_Posted                = oRs("Logs_Posted")
    vMemb_Id                    = oRs("Memb_Id")
    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Level                 = oRs("Memb_Level")
    vMemb_Criteria              = oRs("Memb_Criteria")
  End Sub

%>