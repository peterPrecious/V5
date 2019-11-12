

<%
  Dim vResT_No, vResT_AcctId, vResT_MembNo, vResT_Id, vResT_FromName, vResT_FromEmail, vResT_Subject_EN, vResT_Subject_FR, vResT_Body_EN, vResT_Body_FR, vResT_Note_EN, vResT_Note_FR


  Sub sGetResTbyNo (vNo)      
    vSql = "SELECT * FROM ResT WHERE ResT.ResT_No = " & vNo
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then sReadResT
    Set oRs = Nothing      
    sCloseDb
  End Sub


  Sub sGetResTbyId (vId)      
    vSql = "SELECT * FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " AND ResT_Id = '" & Ucase(vId) & "'"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then sReadResT
    Set oRs = Nothing      
    sCloseDb
  End Sub


  Sub sGetResT_Rs
    vSql =  " SELECT * From FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & " AND ResT_MembNo = " & svMembNo & " AND ResT_FromName = '" & vResTFromName & "' ORDER BY Help_Body_FR, Help_AcctId DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Sub


  Function fCountResT
    fCountResT = 0
    vSql =  "SELECT Count(*) as [Count] FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo  
    sOpenDb2
'   sDebug
    Set oRs2 = oDb2.Execute(vSql)
    If Not oRs2.Eof Then 
      fCountResT = Cint(oRs2("Count"))
    Else
      fCountResT = 0
    End If
    Set oRs2 = Nothing
    sCloseDb2 
  End Function


  Function fIsNewResT (vId)
    fIsNewResT = False
    vSql =  "SELECT Count(*) as [Count] FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " AND Rest_Id = '" & vId & "'"
    sOpenDb2
'   sDebug
    Set oRs2 = oDb2.Execute(vSql)
    If Cint(oRs2("Count")) > 0 Then fIsNewResT = True
    Set oRs2 = Nothing
    sCloseDb2 
  End Function


  Sub sReadResT
    vResT_No            = oRs("ResT_No")
    vResT_AcctId        = oRs("ResT_AcctId")
    vResT_MembNo        = oRs("ResT_MembNo")
    vResT_Id            = oRs("ResT_Id")
    vResT_FromName      = oRs("ResT_FromName")
    vResT_FromEmail     = oRs("ResT_FromEmail")
    vResT_Subject_EN    = oRs("ResT_Subject_EN")
    vResT_Subject_FR    = oRs("ResT_Subject_FR")
    vResT_Body_EN       = oRs("ResT_Body_EN")
    vResT_Body_FR       = oRs("ResT_Body_FR")
    vResT_Note_EN       = oRs("ResT_Note_EN")
    vResT_Note_FR       = oRs("ResT_Note_FR")
  End Sub


  Sub sExtractResT
    vResT_No            = Request("vResT_No")
    vResT_Id            = Ucase(Request("vResT_Id"))
    vResT_FromName      = fUnquote(Request("vResT_FromName"))
    vResT_FromEmail     = fUnquote(Request("vResT_FromEmail"))
    vResT_Subject_EN    = fUnquote(Request("vResT_Subject_EN"))
    vResT_Subject_FR    = fUnquote(Request("vResT_Subject_FR"))
    vResT_Body_EN       = fUnquote(Request("vResT_Body_EN"))
    vResT_Body_FR       = fUnquote(Request("vResT_Body_FR"))
    vResT_Note_EN       = fUnquote(Request("vResT_Note_EN"))
    vResT_Note_FR       = fUnquote(Request("vResT_Note_FR"))
  End Sub


  Sub sInsertResT
    vSql = "INSERT INTO ResT (ResT_AcctId, ResT_Id, ResT_FromName, ResT_FromEmail, ResT_Subject_EN, ResT_Subject_FR, ResT_Body_EN, ResT_Body_FR, ResT_Note_EN, ResT_Note_FR, ResT_MembNo) " _   
         & "VALUES ('" & svCustAcctId & "', '" & vResT_Id & "', '" & vResT_FromName & "', '" & vResT_FromEmail & "', '" & vResT_Subject_EN & "', '" & vResT_Subject_FR & "', '" & vResT_Body_EN & "', '" & vResT_Body_FR & "', '" & vResT_Note_EN & "', '" & vResT_Note_FR & "', " & svMembNo & ") " 
    sOpenDb
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  Sub sUpdateResT
    vSql = "UPDATE ResT SET " _
         & "ResT_Id            = '" & vResT_Id            & "', " _
         & "ResT_FromName      = '" & vResT_FromName      & "', " _
         & "ResT_FromEmail     = '" & vResT_FromEmail     & "', " _
         & "ResT_Subject_EN    = '" & vResT_Subject_EN    & "', " _
         & "ResT_Subject_FR    = '" & vResT_Subject_FR    & "', " _
         & "ResT_Body_EN       = '" & vResT_Body_EN       & "', " _
         & "ResT_Body_FR       = '" & vResT_Body_FR       & "', " _
         & "ResT_Note_EN       = '" & vResT_Note_EN       & "', " _
         & "ResT_Note_FR       = '" & vResT_Note_FR       & "'  " _
         & "WHERE ResT_No      =  " & vRest_No
    sOpenDb
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub


  '...delete ResT
  Sub sDeleteResT
    sOpenDb
    vSql = "DELETE FROM ResT WHERE ResT_No = " & vResT_No
'   sDebug
    oDb.Execute(vSql)
    sCloseDb
  End Sub
  

  Function fResTOptionsById (vResTId)
    fResTOptionsById = ""
    vSql =  "SELECT ResT_Id FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " ORDER BY ResT_No DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      fResTOptionsById = fResTOptionsById & "<option value=" & Chr(34) & oRs("ResT_Id") & Chr(34) & fSelect(oRs("ResT_Id"), vRestId) & ">" & oRs("ResT_Id") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Function


  Function fResTOptionsByNo (vResTNo)
    fResTOptionsByNo = ""
    vSql =  "SELECT ResT_No, ResT_Id FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " ORDER BY ResT_No DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof
      fResTOptionsByNo = fResTOptionsByNo & "<option value=" & Chr(34) & oRs("ResT_No") & Chr(34) & fSelect(oRs("ResT_No"), vResTNo) & ">" & oRs("ResT_Id") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb 
  End Function


  Function fTopResTNo
    fTopResTNo = -1
    vSql =  "SELECT TOP 1 ResT_No FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " ORDER BY ResT_No DESC"
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fTopResTNo = oRs("ResT_No")
    Set oRs = Nothing
    sCloseDb 
  End Function


  Function fLastResTId ()      
    fLastResTId = ""
    vSql = "SELECT Top 1 ResT_Id FROM ResT WHERE ResT_AcctId = '" & svCustAcctId & "' AND ResT_MembNo = " & svMembNo & " ORDER BY ResT_No DESC"
'   sDebug
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then fLastResTId = oRs("ResT_Id")
    Set oRs = Nothing      
    sCloseDb
  End Function


  '...not sure if following function is used anywhere
  Function fResTId (vResTSubject_FR)
    fResTId = ""
    If Len(vResTSubject_FR) > 3 Then 
      vSql =  " SELECT ResT_Id FROM ResT WHERE ResT_Subject_FR = '" & vResTSubject_FR & "' ORDER BY ResT_AcctId DESC"
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        fResTId = fResTId & oRs("ResT_Id") & ", "    
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
      If Len(fResTId) > 6 Then 
        fResTId = Left(fResTId, Len(fResTId) - 2)
      End If
    End If
  End Function


  '...clone the default master on first visit to the site
  Sub sCloneResT
    vSql  = "SET ANSI_WARNINGS ON " _
          & "INSERT INTO ResT " _
          & "(ResT_AcctId, ResT_Id, ResT_FromName, ResT_FromEmail, ResT_Subject_EN, ResT_Subject_FR, ResT_Body_EN, ResT_Body_FR, ResT_Note_EN, ResT_Note_FR, ResT_MembNo) " _
          & "(SELECT '" & svCustAcctId & "' AS ResT_AcctId, ResT_Id, ResT_FromName, ResT_FromEmail, ResT_Subject_EN, ResT_Subject_FR, ResT_Body_EN, ResT_Body_FR, ResT_Note_EN, ResT_Note_FR, " & svMembNo & " AS ResT_MembNo " _
          & "FROM ResT WHERE ResT_AcctId = '0000')"
'   sDebug
    sOpenDb 
    oDb.Execute(vSql)
    sCloseDb    
  End Sub


%>