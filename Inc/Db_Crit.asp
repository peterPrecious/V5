<%
  Dim vCrit_AcctId, vCrit_No, vCrit_Id, vCrit_JobsId
  Dim vCrit_Eof, vCriteriaListCnt

  vCriteriaListCnt = 2

  '...Get Crit Recordset
  Sub sGetCrit_Rs (vCritAcctId)
    vSql = "SELECT * FROM Crit WHERE Crit_AcctId = '" & vCritAcctId & "' ORDER BY Crit_Id"
'   sDebug
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
  End Sub

  Sub sGetCrit (vCritNo)
    vCrit_Eof = True
    If IsNumeric(vCritNo) Then
      vSql = "SELECT * FROM Crit WHERE Crit_No = " & vCritNo
      sOpenDb2    
      Set oRs2 = oDb2.Execute(vSql)
      If Not oRs2.Eof Then 
        sReadCrit
        vCrit_Eof = False
      End If
      Set oRs2 = Nothing
      sCloseDb2    
    End If
  End Sub

  Sub sReadCrit
    vCrit_AcctId     = oRs2("Crit_AcctId")
    vCrit_No         = oRs2("Crit_No")
    vCrit_Id         = oRs2("Crit_Id")
    vCrit_JobsId     = oRs2("Crit_JobsId")
  End Sub

  Sub sExtractCrit
    vCrit_AcctId     = Request.Form("vCrit_AcctId")
    vCrit_No         = Request.Form("vCrit_No")
    vCrit_Id         = fUnquote(Request.Form("vCrit_Id"))
    vCrit_JobsId     = Replace(Request.Form("vCrit_JobsId"), ",", "")
  End Sub
  
  Sub sInsertCrit (vCritAcctId)
    vSql = "INSERT INTO Crit "
    vSql = vSql & "(Crit_AcctId, Crit_Id, Crit_JobsId)"
    vSql = vSql & " VALUES ('" & vCritAcctId & "', '" & vCrit_Id & "', '" & vCrit_JobsId & "')"
'   sDebug
    sOpenDb2
    On Error Resume Next
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub


  '...used in SignIn.asp for auto-enroll 
  Function fSignInCriteria (vCritAcctId, vCriteria)

    '...invalid criteria (Crit_Id)?
    If Len(Trim(vCriteria)) = 0 Or Trim(vCriteria) = "0" Then 
      fSignInCriteria = 0
  		Exit Function
    End If
    
    '...if on table return criteria no
    vCrit_Eof = False
    vSql = "SELECT Crit_No FROM Crit WHERE Crit_Id = '" & vCriteria & "' AND Crit_AcctId = '" & vCritAcctId & "'"
    sOpenDb2    
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      fSignInCriteria = oRs2("Crit_No")
      Set oRs2 = Nothing
      sCloseDb2    
      Exit Function
    End If

    '...if not on the table then add to table
    vSql =        " SET NOCOUNT ON"
    vSql = vSql & " INSERT INTO Crit"
    vSql = vSql & " (Crit_AcctId, Crit_Id)"
    vSql = vSql & " VALUES ('" & vCritAcctId & "', '" & vCriteria & "')"
    vSql = vSql + " SELECT vNo=@@IDENTITY"
    vSql = vSql + " SET NOCOUNT OFF"
'   sDebug
    Set oRs2 = oDb2.Execute(vSql)
    fSignInCriteria = oRs2("vNo")
    Set oRs2 = Nothing      
    sCloseDb2

  End Function


  Sub sUpdateCrit
    vSql = "UPDATE Crit SET"
    vSql = vSql & " Crit_Id              = '" & vCrit_Id                & "', " 
    vSql = vSql & " Crit_JobsId          = '" & vCrit_JobsId            & "'  " 
    vSql = vSql & " WHERE Crit_No        =  " & vCrit_No
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub

  
  Sub sDeleteCrit
    vSql = "DELETE FROM Crit WHERE Crit_No = " & vCrit_No
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub



  '...structure of the list is dependent on who is calling it (vSource)
  '   this can be from Memb or Memb:Fac, TskH, criteria (UserBulkInput) or "Rept:Criteria"
  '...this has been modified

  Function fCriteriaList (vCustAcctId, vSource)

    Dim bCheck, bOk, aCrit
    

    '... do not allow facilitators to access/assign ALL/None (Users.asp, for ex - also "KIDS" for Worksmart custom report)
    If vSource <> "Memb:Fac" And Left(vSource, 4) <> "KIDS" And svMembLevel > 3 Then
      vCrit_No = 0
      fCriteriaList = vbCrLf & "<option" & fSelectCriteria(vSource) & " value='0'>All</option>" & vbCrLf
    Else
      fCriteriaList = ""
    End If

    If Len(vSource) > 5 And (svMembLevel < 4 Or Left(vSource, 4) = "KIDS") Then
      aCrit = Split(Mid(vSource, 6))
      bCheck = True
    Else
      bCheck = False      
    End If

    '...added Aug 08, 2015 because MEVT2747 oddly stopped working
    If vSource = "KIDS:0" Then bCheck = False

    vCriteriaListCnt = 0

    sGetCrit_Rs (vCustAcctId)
    Do While Not oRs2.Eof 
      sReadCrit

      '...normally display all unless for reports, only display what user can see
      If bCheck Then
        bOk = False
        For i = 0 To Ubound(aCrit)
          If IsNumeric(aCrit(i)) Then
            If vCrit_No = Clng(aCrit(i)) Then
              bOk = True
              Exit For
            End If
          End If
        Next    
      Else
        bOk = True
      End If
      
      If bOk Then
        fCriteriaList = fCriteriaList & "<option" & fSelectCriteria(vSource) & " value='" & vCrit_No & "'>" & vCrit_Id & "</option>" & vbCrLf
        vCriteriaListCnt = vCriteriaListCnt + 1
      End If

      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2    

    If vCriteriaListCnt > 50 Then
      vCriteriaListCnt = 12
    ElseIf vCriteriaListCnt > 8 Then
      vCriteriaListCnt = 8
    End If

  End Function


  Function fSelectCriteria (vSource)
    fSelectCriteria = ""
    If vSource = "TskH" Then
      If Instr(" " & vTskH_Criteria & " ", " " & vCrit_No & " ") > 0 Then
        fSelectCriteria = " selected"
      End If 
    ElseIf vSource = "Memb" Then
      If Instr(" " & vMemb_Criteria & " ", " " & vCrit_No & " ") > 0 Then
        fSelectCriteria = " selected"
      End If     
    ElseIf Len(vSource) > 5 Then

      '...if adding a new learner do not highlight anything
      If Left(vSource, 4) = "Memb" Then
        If vMemb_No = 0 Then
          Exit Function
        ElseIf vMemb_No <> svMembNo Then
          If Instr(" " & vMemb_Criteria & " ", " " & vCrit_No & " ") > 0 Then
            fSelectCriteria = " selected"
          End If     
        End If
      Else
        If Instr(" " & Mid(vSource, 6) & " ", " " & vCrit_No & " ") > 0 Then
          fSelectCriteria = " selected"
        End If     
      End If

    End If
  End Function


  Function fCriteria (vCriteria)
    Dim aCriteria
    aCriteria = Split(vCriteria)
    fCriteria = ""
    For i = 0 to Ubound(aCriteria)       
      If i > 0 Then fCriteria = fCriteria & " + "
      If aCriteria(i) = "0" Then 
        fCriteria = fCriteria & fIf(svLang = "FR", "Tous", "All")
      Else
        sGetCrit (aCriteria(i))        
        If Not fNoValue(vCrit_Id) Then fCriteria = fCriteria & vCrit_Id
      End If
    Next
  End Function 
  
  
  Function fCriteriaJobs (vCriteria)
    Dim aCriteria
    fCriteriaJobs = ""
    If vCriteria = "0" Then Exit Function
    aCriteria = Split(vCriteria)
    For i = 0 to Ubound(aCriteria)       
      If i > 0 And fCriteriaJobs <> "" Then fCriteriaJobs = fCriteriaJobs & " + "
      sGetCrit (aCriteria(i))        
      If Not fNoValue(vCrit_JobsId) Then fCriteriaJobs = fCriteriaJobs & vCrit_JobsId
    Next
  End Function 

  

  '...get Criteria (vLevel)
  Function fCriteriaByLevel
    Dim oRs
    fCustOptions = ""
    sOpenDb
    vSql = "Select * FROM Cust "
    Set oRs = oDb.Execute(vSql)    
    Do While Not oRs.EOF 
      fCustOptions = fCustOptions & "<option>" & oRs("Cust_Id") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop
    sCloseDb           
  End Function  


  '...get Criteria No using Criteria Id
  Function fCriteriaNo (vCritAcctId, vCritId)
    fCriteriaNo = 0
    Dim oRs
    sOpenDb
    vSql = "SELECT Crit_No FROM Crit WHERE (Crit_AcctId = '" & vCritAcctId & "') AND (Crit_Id = '" & vCritId & "')"
    Set oRs = oDb.Execute(vSql)    
    If Not oRs.Eof Then fCriteriaNo = oRs("Crit_No")
    sCloseDb    
  End Function
  
  
  '...criteria issues ok for reports? 
  '   svCrit (person signed in) can be: 270 30 and vCrit (member being reported) can be: 230 70
  '   add spaces around vCrit(" 230 70 ", " 30 ") so "30" does not pickup "230"
  Function fCriteriaOk (svCrit, vCrit)
    Dim aCrit 
    fCriteriaOk = True
    If svCrit = "0" Then Exit Function
    If Instr(svCrit, " ") = 0 Then '...single criteria
      If Instr(" " & vCrit & " " , " " & svCrit & " ") = 0 Then fCriteriaOk = False
      Exit Function
    Else   '...multi-criteria
      aCrit = Split(svCrit)
      For i = 0 To Ubound(aCrit)
        If aCrit(i) = "0" Or Instr(" " & vCrit, " " & aCrit(i) & " ") > 0 Then Exit Function
      Next
      fCriteriaOk = False
    End If
  End Function  



  '...does this account use Group 1 values (used in user.asp and /repository/upload_advanced.asp)
  Function fCritOk (vCustAcctId)
    sOpenDb3
    vSql = "SELECT Count(*) AS [Crit_Count] FROM Crit WHERE (Crit_AcctId = '" & vCustAcctId & "')"
    Set oRs3 = oDb3.Execute(vSql)    
    fCritOk = False
    If Not oRs3.Eof Then 
      If Cint(oRs3("Crit_Count")) > 0 Then
        fCritOk = True
      End If
    End If 
    sCloseDb3  
  End Function  


  
%>