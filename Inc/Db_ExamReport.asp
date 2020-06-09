<%
  Dim vLogs_No, vLogs_AcctId, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo
  Dim vLogs_Module, vLogs_Grade, vDetails, vCurList, vMaxList, vLogs_Assess
  
  
  '____ Logs  (for Assessments) ________________________________________________________________________

  '...Get Log Entry for single member
  Sub sGetLog_Rs (vMembNo)
    vSql = "SELECT Left(Logs.Logs_Item, 6) AS Logs_Module, Right(Logs.Logs_Item,3) AS Logs_Grade, Logs.Logs_Posted, Logs.Logs_Memo "
    vSql = vSql & " FROM Logs "
    vSql = vSql & " WHERE Logs_AcctId= '" & svCustAcctId & "' AND Logs_Type = 'T'" & " AND Logs_MembNo=" & vMembNo
    vSql = vSql & " ORDER BY Logs_Module ASC, Logs.Logs_Posted DESC"
'   sDebug
    sOpenDB
    Set oRs = oDB.Execute(vSql)
  End Sub


  '...get the current fields from the current record in the record set (just a subset of the full fields)
  Sub sReadProcessLogsMemb
    vLogs_Item                  = oRs("Logs_Item")
    vLogs_Module                = oRs("Logs_Module")
    vLogs_Grade                 = oRs("Logs_Grade")
    vLogs_Posted                = oRs("Logs_Posted")
    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Criteria              = oRs("Memb_Criteria")
    vMemb_Memo                  = oRs("Memb_Memo")
  End Sub


  '...get the current fields from the current record in the record set (just a subset of the full fields) - for Assessment Report
  Sub sReadLogsMemb
    vLogs_Module                = oRs("Logs_Module")
    vLogs_Grade                 = oRs("Logs_Grade")
    vLogs_Assess                = oRs("Logs_Assess")
    vLogs_Posted                = oRs("Logs_Posted")
    vMemb_No                    = oRs("Memb_No")
    vMemb_Id                    = oRs("Memb_Id")
    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Criteria              = oRs("Memb_Criteria")
    vMemb_Memo                  = oRs("Memb_Memo")
  End Sub


  '...get the current fields from the current record in the record set (just a subset of the full fields) - for Assessment Report
  Sub sReadLogsMembSurvey
'   vLogs_Module                = oRs("Logs_Module")
    vLogs_Item                  = oRs("Logs_Item")
    vLogs_Posted                = oRs("Logs_Posted")
    vMemb_No                    = oRs("Memb_No")
    vMemb_Id                    = oRs("Memb_Id")
    vMemb_FirstName             = oRs("Memb_FirstName")
    vMemb_LastName              = oRs("Memb_LastName")
    vMemb_Criteria              = oRs("Memb_Criteria")
    vMemb_Memo                  = oRs("Memb_Memo")
  End Sub

  
  '...get the current fields from the current record in the record set (just a subset of the full fields)
  Sub sReadLogMemb
    vLogs_Module                = oRs("Logs_Module")
    vLogs_Grade                 = oRs("Logs_Grade")
    vLogs_Posted                = oRs("Logs_Posted")
  End Sub


 '...get the exam title 
  Function fExamTitle (vModId)
    fExamTitle = ""
    sOpenDbBase2
    vSql = "SELECT TstH_Title FROM TstH WHERE TstH_Id = '" & vModId & "'" 
    Set oRsBase2 = oDbBase2.Execute(vSql)    
    If Not oRsBase2.Eof Then fExamTitle = Trim(oRsBase2("TstH_Title"))
    sCloseDbBase2
    '...if no title then return module no
    If Len(fExamTitle) = 0 Then fExamTitle = vModId 
  End Function


  '...this routine is used in PostAssessmentScores.asp
  Sub sInsertLogs (vMembNo, vItem)
    sOpenDb
    vSql = "INSERT INTO Logs (Logs_AcctId, Logs_Type, Logs_Item, Logs_MembNo) VALUES "
    vSql = vSql & "('" & svCustAcctId & "', 'T', '" & vItem & "', " & vMembNo & ")"
'   sDebug  
    oDB.Execute(vSql)
    sCloseDb
  End Sub


  '...this routine is used in PostAssessmentScores.asp
  Sub sDeleteLogs (vMembNo, vModId)
    sOpenDb
    vSql = "DELETE Logs WHERE Logs_MembNo = " & vMembNo & " AND Left(Logs_Item, 6) = '" & vModId & "'"
'   sDebug  
    oDB.Execute(vSql)
    sCloseDb
  End Sub

%>