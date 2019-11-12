<%
  Dim vScript, vCnt, vPrev_L1, vPrev_L0, vPrev_HO, vNewTitle, vAction, aL1, vSelectNo, vRoleCnt, vMsg, vTitle, bOk, vRole , vJobs

  '...store roles in arrays plus an empty array for jobs   
  Dim aRoleHO, aRoleXX, aJobsHO, aJobsXX 
  aRoleHO = Split(Session("Completion_Roles_HO"))
  aRoleXX = Split(Session("Completion_Roles_XX"))

  '...create just to define array, then initialize 
  aJobsHO = Split(Session("Completion_Roles_HO"))
  aJobsXX = Split(Session("Completion_Roles_XX"))
  sInitJobs()

  vAction      = ""
  vUnit_L1     = ""
  vUnit_Active = fDefault(Request("vUnit_Active"), 1)


  Sub sInitJobs()
    For i = 0 To Ubound(aJobsHO) : aJobsHO(i) = "-" : Next
    For i = 0 To Ubound(aJobsXX) : aJobsXX(i) = "-" : Next
  End Sub


  '...Update Crit Table with modified Jobs
  Sub sSql_Mod (vJobsId, vRole)
    vSql = "" _
         & " UPDATE Crit"_
         & " SET Crit_JobsId = '" & Trim(vJobsId) & "'" _ 
         & " WHERE Crit_AcctId = '" & svCustAcctId & "' AND Crit_Id = '" & vUnit_L1 & " " & vUnit_L0 & " " & Trim(vRole) & "' AND Crit_JobsId <> '" & Trim(vJobsId) & "'"
    sCompletion_Debug
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
  End Sub


  Function fL1s (vWhich)
    Dim vSelected
    vSelectNo = 0     '...global parm
    fL1s = ""
    vSql = " SELECT DISTINCT "_
         & "  Unit.Unit_L1, "_
         & "  Unit_L1Title, "_
         & "  Unit_HO "_
         & " FROM "_
         & "   V5_Comp.dbo.Unit AS Unit WITH (NOLOCK) "_
         & " WHERE "_
         & "   Unit.Unit_AcctId = '" & svCustAcctId & "'"_
         & " ORDER BY "_
         & "   Unit.Unit_L1 "
    sCompletion_Debug
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    Do While Not oRs2.Eof
      '...do not include HO when moving L0s around
      If vWhich = "All" Or (vWhich = "Move" And oRs2("Unit_HO") = 0 And vUnit_L1 <> oRs2("Unit_L1") ) Then   

        vSelected = "" 
        If vUnit_L1 = oRs2("Unit_L1") And vWhich = "All" Then vSelected = " selected" 
        fL1s = fL1s & "<option value='" & oRs2("Unit_L1") & "|" & oRs2("Unit_L1Title") & "|" & fIf(oRs2("Unit_HO"), "1", "0") & "'"  & Chr(34) & vSelected & ">" & oRs2("Unit_L1") & " (" & oRs2("Unit_L1Title") & ")</option>" & vbCrLf
        vSelectNo = vSelectNo + 1   

      End If
      oRs2.MoveNext
    Loop
    Set oRs2 = Nothing
    sCloseDb2
  End Function


  '...get all Active Jobs
  Function fJobs (vId)
    Dim vSelected, aMods, vProgs, vClass
    fJobs = "<option>Select</option>"
    fJobs = ""
    sGetJobs_Rs
    Do While Not oRs3.Eof 
      sReadJobs
      If vJobs_Active Then vClass = "" Else vClass = " class='c6'"
      vSelected = "" : If Instr(vId, vJobs_Id) > 0  Then vSelected = " selected" 
      fJobs = fJobs  & "<option value=" & Chr(34) & vJobs_Id & Chr(34) & vSelected & vClass & ">" & vJobs_Id & " : " & vJobs_Title & fIf(vJobs_Active, "", " (Inactive)") & "</option>" & vbCrLf
      oRs3.MoveNext
    Loop      
    sCloseDb3           
    Set oRs3 = Nothing
  End Function  


  '...ensure there are no learners assigned to this Location
  Function fUnitDeleteOk (vL1, vL0)
    fUnitDeleteOk = False
    sOpenDb
    vSql = " "_
         & " SELECT"_     
         & "   COUNT(Memb.Memb_No) AS [Count]"_
         & " FROM"_         
         & "   V5_Vubz.dbo.Memb AS Memb WITH (NOLOCK) INNER JOIN"_
         & "   V5_Vubz.dbo.Crit AS Crit WITH (NOLOCK) ON Memb.Memb_Criteria = Crit.Crit_No AND Memb.Memb_AcctId = Crit.Crit_AcctId"_
         & " WHERE"_     
         & "   (Memb.Memb_AcctId = '" & svCustAcctId & "') AND (Crit.Crit_Id LIKE '" & vL1 & " " & vL0 & "%') AND (ISNUMERIC(Memb_Criteria)=1)"
    sCompletion_Debug
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      fUnitDeleteOk = True
    ElseIf oRs("Count").Value = 0 Then
      fUnitDeleteOk = True
    End If
    sCloseDb   
  End Function


%>

